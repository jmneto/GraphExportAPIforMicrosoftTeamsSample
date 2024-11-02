//MIT License
//
//Copyright (c) 2024 Microsoft - Jose Batista-Neto.
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using Microsoft.Extensions.Configuration;

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// This class is used to read the appsettings.json file
// This class is a Singleton, it reads the appsettings.json file only once
// This class is thread-safe
// This class controls all the configurations of the application
// Please read the file readme.md for more information
internal static class AppConfigHelper
{
    // Private Members
    private static AppConfig? appConfig = null;
    private static readonly object lck = new object();

    // This method will return the appsettings.json file configuration
    public static AppConfig GetAppConfig()
    {
        // If appConfig is already initialized, return it
        // this is to avoid reading the appsettings.json file multiple times
        if (appConfig != null)
            return appConfig;

        // Initialize appConfig
        // we lock the code to avoid multiple threads reading the appsettings.json file
        lock (lck)
        {
            // Check again if appConfig is already initialized
            // this is to avoid reading the appsettings.json file multiple times
            // this is to avoid racing conditions
            if (appConfig != null)
                return appConfig;

            // Initialize appConfig
            appConfig = new AppConfig();

            try
            {
                // Read Environment Name, this is for the development environment
                string? environmentName = Environment.GetEnvironmentVariable("EnvironmentName");

                // Read appsettings.json file
                new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
                    .AddJsonFile($"appsettings.{environmentName}.json", optional: true, reloadOnChange: false)
                    .AddEnvironmentVariables()
                    .Build()
                    .Bind(appConfig);

                if (string.IsNullOrEmpty(appConfig.TenantId))
                    throw new Exception("Application TenantId is missing. Check Readme.md");

                if (string.IsNullOrEmpty(appConfig.ApplicationId))
                    throw new Exception("ApplicationId is missing. Check Readme.md");

                if (string.IsNullOrEmpty(appConfig.ClientSecret))
                    throw new Exception("ClientSecret is missing. Check Readme.md");

                if (appConfig.StartDateTimeUTC == DateTime.MinValue)
                    throw new Exception("StartDateTimeUTCString is missing. Check Readme.md");

                if (appConfig.EndDateTimeUTC == DateTime.MaxValue)
                    throw new Exception("EndDateTimeUTCString is missing. Check Readme.md");

                if (string.IsNullOrEmpty(appConfig.SqlConnectionString))
                    throw new Exception("SqlConnectionString is missing. Check Readme.md");
            }
            catch (Exception ex)
            {
                throw new IOFatalException("Error reading app settings", ex);
            }
            return appConfig;
        }
    }

    // Type used to map the appsettings.json file configuration
    public class AppConfig
    {
        // Private Members
        // /////////////////////
        private int mailBoxLoaderTaskLimit;
        private int graphLoaderTaskLimit;
        private int preProcTaskLimit;
        private int getUserMessagesBatchSize;

        // Public Members
        // /////////////////////

        // Input Locations
        public string? InputStorageAccountUrl { get; set; }
        public string? InputStorageConnectionString { get; set; }
        public string InputStorageContainer { get; set; } = "inputcontainer";

        // Log storage location
        public string? LogStorageAccountUrl { get; set; }
        public string? LogStorageConnectionString { get; set; }
        public string LogStorageContainer { get; set; } = "logs";

        // Input Files 
        public string MailBoxesWildCardPattern { get; set; } = "AllMailBoxes*.json";

        // Task Limits
        public int MailBoxLoaderTaskLimit
        {
            get => mailBoxLoaderTaskLimit;

            set => mailBoxLoaderTaskLimit = value > 0 ? value : 8;
        }

        public int GraphLoaderTaskLimit
        {
            get => graphLoaderTaskLimit;

            set => graphLoaderTaskLimit = value > 0 ? value : 8;
        }

        public int PreProcTaskLimit
        {
            get => preProcTaskLimit;

            set => preProcTaskLimit = value > 0 ? value : 8;
        }

        // Graph API Service
        public string? TenantId { get; set; }
        public string? ApplicationId { get; set; }
        public string? ClientSecret { get; set; }

        // MicrosoftTeamsAPIModel is either "A" or "B"
        private string microsoftTeamsAPIModel = "A";
        public string MicrosoftTeamsAPIModel
        {
            get => microsoftTeamsAPIModel;
            set
            {
                if (value == "A" || value == "B")
                    microsoftTeamsAPIModel = value;
            }
        }

        // Graph Extraction Interval
        public string StartDateTimeUTCString { get; set; } = DateTime.MinValue.ToString("O");
        public string EndDateTimeUTCString { get; set; } = DateTime.MaxValue.ToString("O");

        // Provide read-only DateTime properties that parse the string values
        public DateTime StartDateTimeUTC
        {
            get
            {
                if (DateTime.TryParse(StartDateTimeUTCString, out DateTime parsedDateTime))
                {
                    return DateTime.SpecifyKind(parsedDateTime, DateTimeKind.Utc);
                }
                return DateTime.MinValue;
            }
        }

        public DateTime EndDateTimeUTC
        {
            get
            {
                if (DateTime.TryParse(EndDateTimeUTCString, out DateTime parsedDateTime))
                {
                    return DateTime.SpecifyKind(parsedDateTime, DateTimeKind.Utc);
                }
                return DateTime.MaxValue;
            }
        }

        // Graph API Batch Size for Large APIs
        public int GetUserMessagesBatchSize
        {
            get => getUserMessagesBatchSize;
            set
            {
                if (value >= 10 && value <= 1000)
                    getUserMessagesBatchSize = value;
                else
                    getUserMessagesBatchSize = 800; // Default value
            }
        }

        // SQL Connection String
        public string? SqlConnectionString { get; set; }

        // DB Maintenance
        public int ResetDB { get; set; } = 0;
    }
}