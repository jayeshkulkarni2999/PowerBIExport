using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PowerBIExportUAT
{
    public static class Logger
    {
        private static string GetLogFilePath()
        {
            // Generate file name based on current date
            string fileName = $"log_{DateTime.Now:ddMMyyyy}.txt";
            string filePath = Path.Combine(AppContext.BaseDirectory + "Logs\\");
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
                filePath = Path.Combine(filePath, fileName);
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Dispose();
                }
            }
            else
            {
                filePath = Path.Combine(filePath, fileName);
                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Dispose();
                }
            }
            return filePath;
        }
        public static void Log(string message)
        {
            try
            {
                // Get the current day's log file path
                string logFilePath = GetLogFilePath();

                // Format the log entry with timestamp
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";

                // Append the log entry to the file
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write log: {ex.Message}");
            }
        }
    }
}
