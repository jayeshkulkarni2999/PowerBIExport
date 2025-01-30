using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Npgsql;

namespace PowerBIExportUAT
{
    class Program
    {
        private static readonly string tenantId = ConfigurationManager.AppSettings["tenantId"];
        private static readonly string clientId = ConfigurationManager.AppSettings["clientId"];
        private static readonly string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        private static readonly string groupId = ConfigurationManager.AppSettings["groupId"]; 
        private static readonly string reportId = ConfigurationManager.AppSettings["reportId"];
        private static readonly string scope = ConfigurationManager.AppSettings["scope"]; 
        private static readonly string localhost = ConfigurationManager.AppSettings["localhost"];
        private static readonly string state = ConfigurationManager.AppSettings["state"];
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["PostgresConnection"].ConnectionString;
        private static readonly string tableName = ConfigurationManager.AppSettings["tableName"];
        private static int retryCount = 1;
        static async Task Main(string[] args)
        {
            try
            {
                string[] indReportID = reportId.Split(",");
                List<EntityData> entityData = new List<EntityData>();
                bool isExportComplete, isFinalInsert = false;
                byte[] fileContent;
                string fileName = string.Empty;
                DataTable dataTable = new DataTable();

                for (int i = 0; i<indReportID.Count(); i++)
                {
                    Logger.Log($"Data Extraction starting for report ID: {indReportID[i].ToString()}");
                    int DataCount = entityData.Count();
                    retryCount = 1;
                    while (retryCount <= 10)
                    {
                        Logger.Log($"Application Try Count: {retryCount}");
                        
                        Logger.Log("Process GetVerificationToken() started");
                        var verificationCode = await GetVerificationToken();
                        Logger.Log("Process GetVerificationToken() completed");

                        Logger.Log("Process GetAccessToken() started");
                        var accessToken = await GetAccessToken(verificationCode);
                        //var accessToken = await GetVerificationAccessToken();
                        Logger.Log("Process GetAccessToken() completed");

                        Logger.Log("Process StartExportReport() started");
                        var exportID = await StartExportReport(accessToken,indReportID[i].ToString());
                        Logger.Log("Process StartExportReport() completed");

                        Logger.Log("Process MonitorExportStatusAsync() started");
                        (isExportComplete, fileName) = await MonitorExportStatusAsync(exportID, accessToken, indReportID[i].ToString());
                        Logger.Log("Process MonitorExportStatusAsync() completed");

                        if (isExportComplete)
                        {
                            Logger.Log("Process DownloadExportedFileAsync() started");
                            fileContent = await DownloadExportedFileAsync(exportID, accessToken, fileName, indReportID[i].ToString());
                            Logger.Log("Process DownloadExportedFileAsync() completed");
                        }
                        else
                        {
                            Logger.Log("Process DownloadExportedFileAsync() not hit as fileContent is null");
                            fileContent = null;
                        }

                        if (fileContent == null || fileContent.Length == 0)
                        {
                            Logger.Log("Process ReadExcelFromByteArray() not hit as fileContent is null");
                        }
                        else
                        {
                            Logger.Log("Process ReadExcelFromByteArray() started");
                            dataTable = await ReadExcelFromByteArray(fileContent);
                            Logger.Log("Process ReadExcelFromByteArray() completed");
                        }

                        if (dataTable.Rows.Count == 0 || fileContent == null || fileContent.Length == 0)
                        {
                            Logger.Log("Process MapToEntityDataList() not hit as datatable is null");
                        }
                        else
                        {
                            Logger.Log("Process MapToEntityDataList() started");
                            entityData = MapToEntityDataList(dataTable, entityData);
                            Logger.Log("Process MapToEntityDataList() completed");
                        }

                        if(entityData.Count() == DataCount)
                        {
                            Logger.Log($"No data added for report ID: { indReportID[i].ToString()}");
                            Logger.Log($"Data Extraction completed for report ID: { indReportID[i].ToString()}");
                        }
                        else
                        {
                            Logger.Log($"{entityData.Count - DataCount} data added for report ID: { indReportID[i].ToString()}");
                            Logger.Log($"Data Extraction completed for report ID: { indReportID[i].ToString()}");
                        }
                    }
                }

                if(entityData.Count == 0)
                {
                    Logger.Log("Process InsertEntityDataAsync() not hit as datatable is null");
                }
                else
                {
                    Logger.Log("Process BulkInsertEntityDataAsync() started");
                    isFinalInsert = await BulkInsertEntityDataAsync(entityData);
                    //isFinalInsert = await InsertEntityDataAsync(entityData);
                    Logger.Log("Process BulkInsertEntityDataAsync() completed");
                }

                if(isFinalInsert == true)
                {
                    Logger.Log("Data inserted to Postgres successfully!!!");
                }
                else
                {
                    Logger.Log("Data failed to insert in Postgres");
                }


            }
            catch (Exception ex)
            {
                Logger.Log($"Main Program failed : {ex.Message}");
            }
        }

        #region Verification and Access Token Generation
        private static async Task<string> GetVerificationToken()
        {
            try
            {
                return await Task.Run(() =>
                {
                    string authorizationUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authorize?client_id={clientId}&response_type=code&redirect_uri={Uri.EscapeDataString(localhost)}&scope={Uri.EscapeDataString(scope)}&state={state}";

                    var httpListener = new HttpListener();
                    httpListener.Prefixes.Add(localhost + "/");
                    httpListener.Start();
                    Console.WriteLine("Listening for redirect...");
                    Logger.Log("Listening for redirect...");

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = authorizationUrl,
                        UseShellExecute = true
                    });

                    var context = httpListener.GetContext();
                    var query = context.Request.QueryString;
                    string code = query["code"];
                    string receivedState = query["state"];

                    var response = context.Response;
                    string responseString = @"
                    <html>
                        <body>
                            <p>Authorization code received. You can close this tab.</p>
                            <script>
                                // Close the browser tab after a short delay
                                setTimeout(() => {
                                    window.close();
                                }, 1000);
                            </script>
                        </body>
                    </html>";
                    byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                    response.ContentLength64 = buffer.Length;
                    response.OutputStream.Write(buffer, 0, buffer.Length);
                    response.OutputStream.Close();

                    httpListener.Stop();

                    if (receivedState != state)
                    {
                        Console.WriteLine("State mismatch. Potential CSRF attack.");
                        Logger.Log("State mismatch. Potential CSRF attack.");
                        return "";
                    }
                    return code;
                });
            }
            catch(Exception ex)
            {
                Logger.Log($"GetVerificationToken failed : {ex.Message}");
                Logger.Log($"GetVerificationToken StackTrace : {ex.StackTrace}");
                throw new Exception(ex.Message);
            }
        }

        private static async Task<string> GetAccessToken(string tokenCode)
        {
            try
            {
                var requestBody = new FormUrlEncodedContent(new[]
                            {
                            new KeyValuePair<string, string>("code",tokenCode),
                            new KeyValuePair<string, string>("grant_type", "authorization_code"),
                            new KeyValuePair<string, string>("redirect_uri", localhost),
                            new KeyValuePair<string, string>("scope",scope),
                            new KeyValuePair<string, string>("client_id", clientId),
                            new KeyValuePair<string, string>("client_secret", clientSecret)
                            // Add more key-value pairs as required
                        });

                using (var client = new HttpClient())
                {
                    var url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    HttpResponseMessage response = await client.PostAsync(url, requestBody);
                    if (response.IsSuccessStatusCode)
                    {
                        var jsonObject = JObject.Parse(await response.Content.ReadAsStringAsync());
                        string accessToken = jsonObject["access_token"].ToString();
                        return accessToken;
                    }
                    else
                    {
                        Logger.Log($"Error at GetAccessToken(): {response.StatusCode}");
                        Console.WriteLine($"Error: {response.StatusCode}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"GetAccessToken failed : {ex.Message}");
                Logger.Log($"GetAccessToken StackTrace : {ex.StackTrace}");
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region Start, Monitor and Download Export Report 
        private static async Task<string> StartExportReport(string token, string indreportID)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    var url = $"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{indreportID}/ExportTo";

                    var requestBody = new
                    {
                        format = "XLSX"  
                    };

                    var content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(url, content);
                    response.EnsureSuccessStatusCode();

                    var responseContent = await response.Content.ReadAsStringAsync();
                    var jsonResponse = JsonDocument.Parse(responseContent);

                    return jsonResponse.RootElement.GetProperty("id").GetString();
                }
            }
            catch(Exception ex)
            {
                Logger.Log($"StartExportReport failed : {ex.Message}");
                Logger.Log($"StartExportReport StackTrace : {ex.StackTrace}");
                throw new Exception(ex.Message);
            }
        }

        private static async Task<(bool,string)> MonitorExportStatusAsync(string exportId, string token, string indreportID)
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    string url = $"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{indreportID}/exports/{exportId}";

                    httpClient.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                    while (true)
                    {
                        HttpResponseMessage response = await httpClient.GetAsync(url);
                        response.EnsureSuccessStatusCode();

                        var responseContent = await response.Content.ReadAsStringAsync();
                        var jsonResponse = JsonDocument.Parse(responseContent);
                        string status = jsonResponse.RootElement.GetProperty("status").GetString();
                        string reportName = jsonResponse.RootElement.GetProperty("reportName").GetString().Replace(" ", "");
                        Console.WriteLine($"Export Status: {status}");
                        Logger.Log($"Export Status: {status}");

                        if (status == "Succeeded")
                        {
                            retryCount = 100;
                            return (true, reportName);
                        }
                        else if (status == "Failed")
                        {
                            retryCount = retryCount + 1;
                            return (false, "");
                        }

                        await Task.Delay(5000); // Wait for 5 seconds before checking again
                    }
                }
            }
            catch (Exception ex)
            {
                retryCount = retryCount + 1;
                Logger.Log($"MonitorExportStatusAsync failed : {ex.Message}");
                Logger.Log($"MonitorExportStatusAsync StackTrace : {ex.StackTrace}");
                return (false,ex.Message);
            }
        }

        private static async Task<byte[]> DownloadExportedFileAsync(string exportId, string token, string fileName, string indreportID)
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    string url = $"https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports/{indreportID}/exports/{exportId}/file";

                    httpClient.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                    HttpResponseMessage response = await httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    var contentType = response.Content.Headers.ContentType.MediaType;
                    string fileExtension = contentType switch
                    {
                        "application/pdf" => ".pdf",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
                        _ => throw new Exception($"Unsupported Content-Type: {contentType}")
                    };

                    var fileContent = await response.Content.ReadAsByteArrayAsync();
                    fileName = $"{fileName}{fileExtension}";

                    return fileContent;
                }
            }
            catch(Exception ex)
            {
                Logger.Log($"DownloadExportedFileAsync failed : {ex.Message}");
                Logger.Log($"DownloadExportedFileAsync StackTrace : {ex.StackTrace}");
                return null;
            }
        }
        #endregion

        #region Manipulate Data structure
        private static async Task<DataTable> ReadExcelFromByteArray(byte[] fileContent)
        {
            try
            {

                DataTable dataTable = new DataTable();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var memoryStream = new MemoryStream(fileContent))
                {
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        await Task.Run(() => package.Load(memoryStream));
                        var worksheet = package.Workbook.Worksheets[0];

                        if (worksheet.Dimension == null)
                        {
                            Logger.Log("The worksheet is empty.");
                            throw new Exception("The worksheet is empty.");
                        }

                        int startRow = worksheet.Dimension.Start.Row;
                        int endRow = worksheet.Dimension.End.Row;
                        int startCol = worksheet.Dimension.Start.Column;
                        int endCol = worksheet.Dimension.End.Column;

                        for (int row = startRow; row <= endRow; row++)
                        {
                            bool hasData = false;
                            for (int col = startCol; col <= endCol; col++)
                            {
                                if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                                {
                                    hasData = true;
                                    break;
                                }
                            }
                            if (hasData)
                            {
                                startRow = row;
                                break;
                            }
                        }

                        for (int col = startCol; col <= endCol; col++)
                        {
                            bool hasData = false;
                            for (int row = startRow; row <= endRow; row++)
                            {
                                if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                                {
                                    hasData = true;
                                    break;
                                }
                            }
                            if (hasData)
                            {
                                startCol = col;
                                break;
                            }
                        }

                        for (int col = startCol; col <= endCol; col++)
                        {
                            string columnName = worksheet.Cells[startRow, col].Text.Trim();
                            if (string.IsNullOrWhiteSpace(columnName))
                            {
                                columnName = $"Column{col}";
                            }
                            dataTable.Columns.Add(columnName);
                        }

                        for (int row = startRow + 1; row <= endRow; row++)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            bool isEmptyRow = true;

                            for (int col = startCol; col <= endCol; col++)
                            {
                                string cellValue = worksheet.Cells[row, col].Text.Trim();
                                dataRow[col - startCol] = cellValue;
                                if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    isEmptyRow = false;
                                }
                            }

                            if (!isEmptyRow)
                            {
                                dataTable.Rows.Add(dataRow);
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch(Exception ex)
            {
                Logger.Log($"ReadExcelFromByteArray failed : {ex.Message}");
                Logger.Log($"ReadExcelFromByteArray StackTrace : {ex.StackTrace}");
                return null;
            }
            
        }

        public static List<EntityData> MapToEntityDataList(DataTable dataTable, List<EntityData> entityDataList )
        {
            foreach (DataRow row in dataTable.Rows)
            {
                var entityData = new EntityData
                {
                    Entity = row["Entity"].ToString(),
                    Head = row["Head"].ToString(),
                    Level5 = row["Level 5"].ToString(),
                    Level4 = row["Level 4"].ToString(),
                    Level3 = row.Table.Columns.Contains("Level 3") ? row["Level 3"].ToString() : "",
                    Level2 = row.Table.Columns.Contains("Level 2") ? row["Level 2"].ToString() : "",
                    Level1 = row.Table.Columns.Contains("Level 1") ? row["Level 1"].ToString() : "",
                    Amount = row.Table.Columns.Contains("Amount") ? row["Amount"].ToString() : "",
                    FinancialYear = row["Financial Year"].ToString(),
                    Type = row["Type"].ToString(),
                    Hierarchy = row["Hierarchy"].ToString(),
                    EntityCode = row.Table.Columns.Contains("Entity Code") ? row["Entity Code"].ToString() : "",
                    EntityName = row.Table.Columns.Contains("Entity Name") ? row["Entity Name"].ToString() : ""
                };

                entityDataList.Add(entityData);
            }
            
            return entityDataList;
        }
        #endregion

        #region Final Insert to DB
        public static async Task<bool> BulkInsertEntityDataAsync(List<EntityData> entityDataList)
        {
            using var connection = new NpgsqlConnection(connectionString);
            await connection.OpenAsync();
            using var transaction = await connection.BeginTransactionAsync();

            try
            {
                // Truncate the table
                var truncateCommand = new NpgsqlCommand($"TRUNCATE TABLE {tableName} RESTART IDENTITY;", connection, transaction);
                await truncateCommand.ExecuteNonQueryAsync();

                using (var writer = connection.BeginTextImport($"COPY {tableName} (entity, head, level_5, level_4, level_3, level_2, level_1, amount, financial_year, entity_data_type, entity_data_hierarchy,entity_name,entity_code, created_by) FROM STDIN (FORMAT CSV)"))
                {
                    foreach (var data in entityDataList)
                    {
                        var line = $"\"{data.Entity ?? ""}\",\"{data.Head ?? ""}\",\"{data.Level5 ?? ""}\",\"{data.Level4 ?? ""}\", \"{data.Level3 ?? ""}\",\"{data.Level2 ?? ""}\",\"{data.Level1 ?? ""}\",\"{data.Amount ?? ""}\",\"{data.FinancialYear ?? ""}\",\"{data.Type ?? ""}\",\"{data.Hierarchy ?? ""}\",\"{data.EntityName ?? ""}\",\"{data.EntityCode ?? ""}\",\"1\"";
                        var sanitizedLine = line.Trim().Replace("\r", "").Replace("\n", "");
                        writer.WriteLine(sanitizedLine);
                    }
                }
                // Commit the transaction
                await transaction.CommitAsync();
                Logger.Log($"BulkInsertEntityDataAsync done. Records Inserted!!!");
                return true;
            }
            catch (Exception ex)
            {
                // Rollback the transaction in case of failure
                await transaction.RollbackAsync();
                Logger.Log($"InsertEntityDataAsync failed : {ex.Message}");
                Logger.Log($"InsertEntityDataAsync failed : {ex.StackTrace}");
                return false;
            }
        }


        public static async Task<bool> InsertEntityDataAsync(List<EntityData> entityDataList)
        {
            using var connection = new NpgsqlConnection(connectionString);
            await connection.OpenAsync();
            using var transaction = await connection.BeginTransactionAsync();

            try
            {
                // Truncate the table
                var truncateCommand = new NpgsqlCommand($"TRUNCATE TABLE {tableName} RESTART IDENTITY;", connection, transaction);
                await truncateCommand.ExecuteNonQueryAsync();

                // Insert data
                foreach (var data in entityDataList)
                {
                    var insertCommand = new NpgsqlCommand(@$"
                    INSERT INTO {tableName} 
                    (entity, head, level_5, level_4, level_3, level_2, level_1, amount, financial_year, entity_data_type, entity_data_hierarchy,entity_name,entity_code, created_by, created_on) 
                    VALUES 
                    (@entity, @head, @level5, @level4, @level3, @level2, @level1, @amount, @financialYear, @type, @hierarchy,@entityName,@entityCode, @createdBy, @createdOn);",
                        connection, transaction);

                    insertCommand.Parameters.AddWithValue("entity", data.Entity ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("head", data.Head ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("level5", data.Level5 ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("level4", data.Level4 ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("level3", data.Level3 ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("level2", data.Level2 ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("level1", data.Level1 ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("amount", data.Amount ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("financialYear", data.FinancialYear ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("type", data.Type ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("hierarchy", data.Hierarchy ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("entityName", data.EntityName ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("entityCode", data.EntityCode ?? (object)DBNull.Value);
                    insertCommand.Parameters.AddWithValue("createdBy", "1");
                    insertCommand.Parameters.AddWithValue("createdOn", DateTime.Now);

                    await insertCommand.ExecuteNonQueryAsync();
                }

                // Commit the transaction
                await transaction.CommitAsync();
                Logger.Log($"InsertEntityDataAsync done. Records Inserted!!!");
                return true;
            }
            catch (Exception ex)
            {
                // Rollback the transaction in case of failure
                await transaction.RollbackAsync();
                Logger.Log($"InsertEntityDataAsync failed : {ex.Message}");
                Logger.Log($"InsertEntityDataAsync failed : {ex.StackTrace}");
                return false;
            }
        }
        #endregion


        private static async Task<string> GetVerificationAccessToken()
        {
            try
            {
                var requestBody = new FormUrlEncodedContent(new[]
                            {
                            new KeyValuePair<string, string>("grant_type", "client_credentials"),
                            new KeyValuePair<string, string>("scope",scope),
                            new KeyValuePair<string, string>("client_id", clientId),
                            new KeyValuePair<string, string>("client_secret", clientSecret)
                            // Add more key-value pairs as required
                        });

                using (var client = new HttpClient())
                {
                    var url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    HttpResponseMessage response = await client.PostAsync(url, requestBody);
                    if (response.IsSuccessStatusCode)
                    {
                        var jsonObject = JObject.Parse(await response.Content.ReadAsStringAsync());
                        string accessToken = jsonObject["access_token"].ToString();
                        return accessToken;
                    }
                    else
                    {
                        Logger.Log($"Error at GetAccessToken(): {response.StatusCode}");
                        Console.WriteLine($"Error: {response.StatusCode}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"GetAccessToken failed : {ex.Message}");
                Logger.Log($"GetAccessToken StackTrace : {ex.StackTrace}");
                throw new Exception(ex.Message);
            }
        }

    }
}
