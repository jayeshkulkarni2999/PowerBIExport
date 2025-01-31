# C# Power BI Report Exporter to PostgreSQL

## Overview
**C# Power BI Report Exporter** is a C# Console Application that exports multiple paginated **Power BI reports** into a **PostgreSQL table** using **Power BI Embedded APIs**. The application follows a structured flow to retrieve, monitor, and store large report datasets efficiently.

## Features
- Uses **Power BI Embedded APIs** to export paginated reports.
- Implements **token-based authentication** to securely access Power BI services.
- Continuously monitors export status and retries failed exports if needed.
- **Efficient bulk insertion** into PostgreSQL for high-performance data storage.
- Handles large reports with entries ranging from **10,000 to 12,000+ records per report**.
- Reduces insert time by using **bulk insert** instead of single inserts (**50000 records in ~2 mins instead of 8-10 mins**).

## Prerequisites
Ensure you have the following installed:
- [.NET Core SDK](https://dotnet.microsoft.com/download) or **.NET Framework**
- [Power BI Embedded API](https://learn.microsoft.com/en-us/rest/api/power-bi/)
- [PostgreSQL Database](https://www.postgresql.org/download/)
- [Npgsql](https://www.npgsql.org/) (PostgreSQL C# Driver)
- Git (if using version control)

## Setup Instructions
1. **Clone the Repository:**
   ```sh
   git clone <repository-url>
   cd PowerBIReportExporter
   ```
2. **Restore Dependencies:**
   ```sh
   dotnet restore
   ```
3. **Build the Project:**
   ```sh
   dotnet build
   ```
4. **Run the Application:**
   ```sh
   dotnet run
   ```

## Git Setup
Make sure `App.config` is ignored in `.gitignore`:
```
App.config
```

### Git Commands to Push Code
```sh
git init
git remote add origin <repo-url>
echo App.config > .gitignore
git add .
git commit -m "Initial commit"
git push -u origin main
```

## How It Works
1. **Get Access Token:** The application first requests an **access token** from the Power BI **token API**.
2. **Export Report:** Using the retrieved token, it calls the **ExportTo API** to start the export process for a report.
3. **Monitor Export Status:**
   - Calls the **Status API** repeatedly to check export progress.
   - If the export fails, it retrieves a new **access token** and restarts the export process.
4. **Retrieve Exported File:** Once the export is 100% complete, it fetches the file **content in byte format**.
5. **Convert Byte Data to DataTable:** The file content is converted into a **DataTable** format.
6. **Bulk Insert into PostgreSQL:**
   - The DataTable is mapped to the **PostgreSQL table model**.
   - Data is stored efficiently using **bulk insert** instead of individual row inserts.
   - The application accumulates records from multiple reports and inserts them together to optimize performance.

## Configuration
- **Power BI API Credentials:** Set up in **App.config**.
- **PostgreSQL Connection String:** Defined in the database configuration file.
- **Report Handling:** The application processes one report at a time but batches multiple reports before bulk inserting into PostgreSQL.

## Testing
- Ensure Power BI API credentials are correctly configured.
- Run tests with sample reports to validate data extraction and storage.
- Verify PostgreSQL tables to ensure correct data mapping.
- Monitor logs for failed exports and retry attempts.

## License
This project is licensed under the MIT License.

