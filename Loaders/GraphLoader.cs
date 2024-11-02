//MIT License
//
//Copyright (c) 2024 Microsoft - Jose Batista-Neto
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

using Polly;
using System.Net;
using Microsoft.Data.SqlClient;
using System.Net.Sockets;

using GraphExportAPIforMicrosoftTeamsSample.Helpers;
using GraphExportAPIforMicrosoftTeamsSample.Processors;
using GraphExportAPIforMicrosoftTeamsSample.Types;
using GraphExportAPIforMicrosoftTeamsSample.DB;

namespace GraphExportAPIforMicrosoftTeamsSample.Loaders;

// This is a helper class for the GraphLoader functions
// This class is responsible for loading the UserMessage and ChatSysMessage tables
internal static class GraphLoader
{
    private static readonly AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

    private static readonly TaskHelper tk = new TaskHelper("GraphLoader", cfg.GraphLoaderTaskLimit);

    public static async Task LoadUserMailboxes()
    {
        const string QUERYMAILBOXES = @"
SELECT
	mb.ExternalDirectoryObjectId
FROM
	MailBox mb WITH (READCOMMITTED)
	LEFT JOIN GraphCallProcessed gcp WITH (READCOMMITTED)
	ON mb.ExternalDirectoryObjectId = gcp.value AND gcp.type = 'user'
WHERE
	gcp.value IS NULL";

        const string QUERYMAILBOXESCOUNT = @"
SELECT
	count(*)
FROM
	MailBox mb WITH (READCOMMITTED)
	LEFT JOIN GraphCallProcessed gcp WITH (READCOMMITTED)
	ON mb.ExternalDirectoryObjectId = gcp.value AND gcp.type = 'user'
WHERE
	gcp.value IS NULL";

        LoggerHelper.WriteToConsoleAndLog("Loading Teams API Graph Data (user)", ConsoleColor.Green);

        // Get the total count of records in the MailBoxes table
        int totalCount = await DbHelper.ExecuteScalarIntSqlCommandAsync(QUERYMAILBOXESCOUNT);
        int processedCount = 0;

        if (totalCount == 0)
        {
            LoggerHelper.WriteToConsoleAndLog("No MailBox records to process.", ConsoleColor.Yellow);
            return;
        }

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, QUERYMAILBOXES);

                    // Wait for all tasks to complete
                    tk.WaitAndClearTasks();
                })
            .ExecuteAsync(async () =>
            {
                // Create a connection to the database
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    // Open the connection
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(QUERYMAILBOXES, connection))
                    {
                        // Execute the query and obtain a SqlDataReader to read the results
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            // Read the results
                            while (await reader.ReadAsync())
                            {
                                string externalDirectoryObjectId = reader.GetString(0);

                                tk.AddTaskAndManageLimit(Task.Run(() =>
                                {
                                    string localExternalDirectoryObjectId = externalDirectoryObjectId;

                                    // add a timespan to calculate the time it takes to process the mailbox
                                    System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

                                    // Load Main User Messages from MailBoxes
                                    GenericLoader.LoadWithContinuationToken<GetUserMessages, UserMessage>(localExternalDirectoryObjectId, localExternalDirectoryObjectId, GraphHelper.GetUserMessages, PreProcessor.Process).Wait();

                                    // Mark MailBox as processed
                                    DbHelper.ExecuteSqlCommandAsync("INSERT INTO GraphCallProcessed (type, value) VALUES ('user', @userId)", [new SqlParameter("@userId", localExternalDirectoryObjectId)]).Wait();

                                    // Calculate and display the progress
                                    MonitorHelper.SetMailBoxTableProgress(Interlocked.Increment(ref processedCount) / (double)totalCount);

                                    // calculate elapsed time
                                    watch.Stop();
                                    int elapsedms = watch.Elapsed.Milliseconds;
                                    LoggerHelper.WriteToConsoleAndLog($"Pre-processed MailBox Graph Calls for: {localExternalDirectoryObjectId}. Elapsedtime: {elapsedms}(ms)", ConsoleColor.White);

                                }));
                            }
                        }
                    }
                }
                // Wait for all tasks to complete
                tk.WaitAndClearTasks();
            });
    }
}

file static class GenericLoader
{
    private static readonly AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();
    private static readonly TaskHelper tk = new TaskHelper("PreProcessor", cfg.PreProcTaskLimit);

    public delegate Task<GraphHelper.Resp> FetchDataDelegateContinuationTokenResource0(string resourceId0, DateTime startDateTime, DateTime endDateTime, string model, string? nextLink);
    public delegate Task<GraphHelper.Resp> FetchDataDelegateResource0(string resourceId0);
    public delegate Task PreProcessMessageDelegate<T>(string resource, T message);

    public static async Task LoadWithContinuationToken<TOne, TTwo>(string resourceId0, string resourceReplacementId0, FetchDataDelegateContinuationTokenResource0 fetchData, PreProcessMessageDelegate<TTwo> callback)
    {
        // Get the last continuation link
        string typeName = typeof(TOne).Name;
        string? nextLink = null;

        // Model A or B
        string model = cfg.MicrosoftTeamsAPIModel;

        do
        {
            string? data = null;

            do
            {
                // Fetch data from the Graph API for this resource
                GraphHelper.Resp resp = await fetchData(resourceReplacementId0, cfg.StartDateTimeUTC, cfg.EndDateTimeUTC, model, nextLink);

                if (resp.StatusCode == HttpStatusCode.NotFound)
                {
                    LogHttpStatusCodeResponse("Resource not found", resourceReplacementId0, resp);
                    return;
                }

                if (resp.StatusCode == HttpStatusCode.Forbidden)
                {
                    LogHttpStatusCodeResponse("Resource Forbidden", resourceReplacementId0, resp);
                    return;
                }

                if (resp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    LogHttpStatusCodeResponse("Unauthorized", resourceReplacementId0, resp);
                    return;
                }

                if (resp.StatusCode == HttpStatusCode.PaymentRequired)
                {
                    // We allow for 1 retry with a different model
                    if (model != cfg.MicrosoftTeamsAPIModel)  // This is not the first time we are trying, so break
                        throw new InvalidOperationException(GenerateHttpStatusCodeResponseString("Failed model fallback retry Graph Call for ", resourceReplacementId0, resp));

                    LogHttpStatusCodeResponse($"Payment Required. Attempted Model: {model} (Will retry once with fallback model)", resourceReplacementId0, resp);

                    model = (model == "A" ? "B" : "A");
                    continue;
                }

                if (resp.StatusCode == HttpStatusCode.OK)
                {
                    data = resp.Content;

                    if (data == null)
                        throw new InvalidOperationException(GenerateHttpStatusCodeResponseString("Null response Graph Call for", resourceReplacementId0, resp));

                    break;
                }
                else
                    throw new InvalidOperationException(GenerateHttpStatusCodeResponseString("Failed to get data for ", resourceReplacementId0, resp));

            } while (true);

            try
            {
                // Deserialize the messages
                dynamic? mt = JsonSerializer.Deserialize<TOne>(data);
                if (mt == null)
                    throw new InvalidOperationException($"Failed to deserialize messages {typeName}");

                // Assuming the deserialized object has a property `value` that is a collection of messages
                if (mt.value is not IEnumerable<dynamic> msgs)
                    throw new InvalidOperationException("Failed to deserialize message values");

                try
                {
                    // For each message, pre-process it
                    int count = 0;
                    foreach (dynamic msg in msgs)
                    {
                        dynamic localMsg = msg;
                        string localResourceId0 = resourceId0;

                        tk.AddTaskAndManageLimit(callback(localResourceId0, localMsg));

                        count++;
                    }
                    // Wait for all tasks to complete
                    tk.WaitAndClearTasks();

                    LoggerHelper.WriteToConsoleAndLog($"Pre-processed {count} messages for Resource: {resourceId0} ResourceRepl: {resourceReplacementId0}", ConsoleColor.White);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to Pre-process messages for {resourceId0}", ex);
                }

                // Next link using dynamic typing - assume threre is always a next link in odatanextLink
                nextLink = mt.odatanextLink;

            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load Resource: {resourceId0} ResourceRepl: {resourceReplacementId0}\ndata:{data}", ex);
            }
        } while (nextLink != null);
    }

    public static async Task LoadNoContinuationToken<TOne, TTwo>(string resourceId0, string resourceReplacementId0, FetchDataDelegateResource0 fetchData, PreProcessMessageDelegate<TTwo> callback)
    {
        // Get all messages for this user
        GraphHelper.Resp resp = await fetchData(resourceReplacementId0);

        if (resp.StatusCode == HttpStatusCode.NotFound)
        {
            LogHttpStatusCodeResponse("Resource not found", resourceReplacementId0, resp);
            return;
        }

        if (resp.StatusCode == HttpStatusCode.Forbidden)
        {
            LogHttpStatusCodeResponse("Resource Forbidden", resourceReplacementId0, resp);
            return;
        }

        if (resp.StatusCode == HttpStatusCode.Unauthorized)
        {
            LogHttpStatusCodeResponse("Resource Unauthorized", resourceReplacementId0, resp);
            return;
        }

        if (resp.StatusCode != HttpStatusCode.OK)
            throw new InvalidOperationException(GenerateHttpStatusCodeResponseString("Failed to get data for ", resourceReplacementId0, resp));

        // Deserialize the messages
        string typeName = typeof(TOne).Name;
        dynamic? mt = JsonSerializer.Deserialize<TOne>(resp.Content!);
        if (mt == null)
            throw new InvalidOperationException($"Failed to deserialize messages {typeName}");

        await callback(resourceId0, mt);

        LoggerHelper.WriteToConsoleAndLog($"Pre-processed ChatDetail messages for Resource: {resourceId0} ResourceRepl: {resourceReplacementId0}", ConsoleColor.White);
    }

    private static void LogHttpStatusCodeResponse(string message, string resource, GraphHelper.Resp resp)
    {
        LoggerHelper.WriteToConsoleAndLog(GenerateHttpStatusCodeResponseString(message, resource, resp), ConsoleColor.Magenta);
    }

    private static string GenerateHttpStatusCodeResponseString(string message, string resource, GraphHelper.Resp resp)
    {
        return $"{message} {resource}\nStatus Code:\n{resp.StatusCode}\nRequest:\n{resp.Request}\nResponse:\n{resp.Response?.ToString()}\nContent:\n{resp.Content}";
    }
}