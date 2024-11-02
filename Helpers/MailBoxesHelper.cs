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

using GraphExportAPIforMicrosoftTeamsSample.Types;
using GraphExportAPIforMicrosoftTeamsSample.DB;
using Microsoft.Data.SqlClient;

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// This is a helper class for the MailBoxes
// This class is responsible for loading the MailBoxes from the Blob Storage
// and storing them in the SQL Database
// It also provides a method to get the MailBox Info
// It also provides a method to process the Extended Custom Attribute
internal static class MailBoxesHelper
{
    // Get the App Config
    private static AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

    private static readonly TaskHelper tk = new TaskHelper("MailBoxLoader", cfg.MailBoxLoaderTaskLimit);

    // Private Members
    // //////////////////////////////////////////////////////////////////////////////////// 

    // Process a Blob Source
    private static async Task FunctionCallbackFromStorageTraverse(Stream stream, string containerName, string blobName)
    {
        const string QUERY = @"
MERGE INTO [dbo].[MailBox] AS target
USING (
    VALUES
        (@ExternalDirectoryObjectId, @DisplayName, @PrimarySmtpAddress)
) AS source (ExternalDirectoryObjectId, DisplayName, PrimarySmtpAddress)
ON (target.ExternalDirectoryObjectId = source.ExternalDirectoryObjectId)
WHEN MATCHED THEN 
    UPDATE SET 
        target.DisplayName = source.DisplayName,
        target.PrimarySmtpAddress = source.PrimarySmtpAddress
WHEN NOT MATCHED BY TARGET THEN
    INSERT (ExternalDirectoryObjectId, DisplayName, PrimarySmtpAddress)
    VALUES (source.ExternalDirectoryObjectId, source.DisplayName, source.PrimarySmtpAddress);";

        // Check if the stream is null
        if (stream == null)
            throw new Exception("Mailboxes file stream is null or empty");

        LoggerHelper.WriteToConsoleAndLog($"Loading Data... File:{blobName}", ConsoleColor.Magenta);

        // Deserialize the JSON
        JsonSerializerOptions jsonSerializerOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip
        };

        // Async Read Large Stream
        long startPosition = stream.Position;
        IAsyncEnumerable<MailBox?> objects = JsonSerializer.DeserializeAsyncEnumerable<MailBox>(stream, jsonSerializerOptions);

        if (objects == null)
            throw new Exception("IAsyncEnumerable<MailBoxInfo?> is null");

        using (SQLDatabaseContext db = new SQLDatabaseContext())
        {
            await foreach (MailBox? obj in objects)
            {
                // throw exception if obj is null
                if (obj == null)
                    throw new Exception("Mailboxe Object is null");

                // Check if the ExternalDirectoryObjectId is null or empty
                if (string.IsNullOrEmpty(obj.ExternalDirectoryObjectId))
                    LoggerHelper.WriteToConsoleAndLog($"Loading Data. ExternalDirectoryObjectId is null or empty. Ignored Object: {JsonSerializer.Serialize(obj)}", ConsoleColor.White);
                else
                {
                    tk.AddTaskAndManageLimit(Task.Run(() =>
                    {
                        // Add to table
                        SqlParameter[] parameters = new SqlParameter[]
                        {
                            new SqlParameter("@ExternalDirectoryObjectId", obj.ExternalDirectoryObjectId),
                            new SqlParameter("@DisplayName", obj.DisplayName ?? (object)DBNull.Value),
                            new SqlParameter("@PrimarySmtpAddress", obj.PrimarySmtpAddress ?? (object)DBNull.Value),
                        };
                        DbHelper.ExecuteSqlCommandAsync(QUERY, parameters).Wait();
                    }));
                }

                // Get the stream current position and calculate how many bytes we read
                MonitorHelper.AddStorageBytesRead(Math.Max(stream.Position - startPosition, 0));
                startPosition = stream.Position;
            }

            tk.WaitAndClearTasks();
        }

        stream.Close();
    }

    // Public Members
    // //////////////////////////////////////////////////////////////////////////////////// 

    // Main entry point to start processing the Lokkup File
    public static async Task Start()
    {
        if (string.IsNullOrEmpty(cfg.MailBoxesWildCardPattern))
            throw new Exception("Mailboxes file (pattern) not defined. Check Readme file.");

        // Get the storage manager
        StorageHelper sh = new StorageHelper(cfg.InputStorageContainer, cfg.InputStorageAccountUrl, cfg.InputStorageConnectionString);

        // Check if the container exists
        if (sh.IsContainerEmpty(cfg.InputStorageContainer))
            throw new Exception($"Data Source container is empty: {cfg.InputStorageContainer}");

        // Process the file
        try
        {
            // Write to console and Log
            LoggerHelper.WriteToConsoleAndLog($"Loading Mailbox Files according to wildcard pattern: {cfg.MailBoxesWildCardPattern}", ConsoleColor.Yellow);

            // Process the container and subfolders within.
            await sh.TraverseContainer("", cfg.MailBoxesWildCardPattern, FunctionCallbackFromStorageTraverse);

            // Write to console and Log
            LoggerHelper.WriteToConsoleAndLog("Mailboxes Files loaded.", ConsoleColor.Yellow);
        }
        catch (Exception ex)
        {
            throw new Exception("Exception processing Mailbox File", ex);
        }
    }
}
