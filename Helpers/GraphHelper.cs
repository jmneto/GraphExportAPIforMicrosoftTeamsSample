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

using Polly;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Net;
using System.Net.Sockets;

using GraphExportAPIforMicrosoftTeamsSample.Utility;

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// Helper class to make calls to the Teams Export API
// this is the class that will make the calls to the Graph API
internal static class GraphHelper
{
    // Private Members
    /////////////////////////////////////////////////////////////

    // Load AppConfig
    private static AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

    // Graph API Scope
    private const string SCOPE = "https://graph.microsoft.com/.default";

    // Good Practice: Use a single HttpClient for the lifetime of the application - Thread Safe
    private static readonly HttpClient httpClient = new();

    // Token Cache
    private static AuthenticationResult? authToken = null;

    // Lock Object
    private static readonly object lck = new();

    // Main method to make the Graph API call
    // All calls to the Graph API go through this method
    private static async Task<Resp> GraphCall(string request)
    {
        // Thread Safe Token Refresh
        CheckCreateRefreshAuthToken();

        // Execute the HTTP call using the Polly retry policy with extended exception handling
        HttpResponseMessage httpResponse = await Policy
            .Handle<HttpRequestException>(ex =>
                ex.InnerException is SocketException || // Catch SocketException
                ex.InnerException is IOException) // Catch IOException for broader network-related exceptions
            .Or<IOException>() // Directly handle IOExceptions not wrapped in HttpRequestException
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .OrResult<HttpResponseMessage>(r =>
            {
                if (!r.IsSuccessStatusCode &&
                        !(
                               r.StatusCode == HttpStatusCode.NotFound
                            || r.StatusCode == HttpStatusCode.Forbidden
                            || r.StatusCode == HttpStatusCode.PaymentRequired
                            || r.StatusCode == HttpStatusCode.Unauthorized
                        )
                   )
                    return true; // Will retry
                else if (r.StatusCode == HttpStatusCode.Unauthorized) // Special handling of 401. Retry only if the token is expired, if expired get a new token and try again
                {
                    if (IsAuthTokenExpired())
                    {
                        CheckCreateRefreshAuthToken();

                        return true; // Will retry
                    }
                    else
                        return false; // Do not retry if token is not expired (means it is really unauthorized and not an expired token)
                }
                else
                    return false; // Do not retry for other cases
            })
            .WaitAndRetryAsync(Utility.Constants.HTTPRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt) + Utils.RandInteger(0, 10)),
                onRetry: (outcome, timespan, retryAttempt, context) =>
                {
                    string? message;

                    if (outcome.Exception is TaskCanceledException)
                    {
                        message = "Task was canceled.";
                        if (((TaskCanceledException)outcome.Exception).CancellationToken.IsCancellationRequested)
                        {
                            message += " Cancellation was requested by the caller.";
                        }
                        else
                        {
                            message += " Retrying due to timeout.";
                        }
                    }
                    else
                    {
                        message = outcome.Exception?.Message ?? outcome.Result?.StatusCode.ToString();
                    }

                    LoggerHelper.WriteToConsoleAndLog($"HTTP REQUEST RETRY: {message}. Waiting {timespan} before next retry. Retry attempt {retryAttempt}.", ConsoleColor.Magenta);
                })
            .ExecuteAsync(async () =>
            {
                // Create a new HttpRequestMessage for each retry attempt
                using HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Get, request);

                if (authToken == null)
                    throw new InvalidOperationException("Auth Token is null");

                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authToken.AccessToken);
                return await httpClient.SendAsync(httpRequest);
            });

        if (
              httpResponse.IsSuccessStatusCode
           || httpResponse.StatusCode == HttpStatusCode.NotFound
           || httpResponse.StatusCode == HttpStatusCode.Forbidden
           || httpResponse.StatusCode == HttpStatusCode.PaymentRequired
           || httpResponse.StatusCode == HttpStatusCode.Unauthorized
           )
            return new Resp { Request = request, Response = httpResponse, Content = await httpResponse.Content.ReadAsStringAsync(), StatusCode = httpResponse.StatusCode };
        else
            throw new InvalidOperationException($"HTTP REQUEST FAILURE: {httpResponse.StatusCode.ToString()}\n\nRequest:\n{request}\n\nResponse:\n{httpResponse}\n\nhttpResponse.Content:\n{await httpResponse.Content.ReadAsStringAsync()}");
    }

    private static void CheckCreateRefreshAuthToken()
    {
        if (IsAuthTokenExpired())
            lock (lck)
                if (IsAuthTokenExpired())
                {
                    try
                    {
                        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(cfg.ApplicationId)
                                                                        .WithClientSecret(cfg.ClientSecret)
                                                                        .WithAuthority(AzureCloudInstance.AzurePublic, cfg.TenantId)
                                                                        .Build();

                        authToken = app.AcquireTokenForClient(new[] { SCOPE }).ExecuteAsync().Result;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Failed to acquire token for client: {ex.Message}", ex);
                    }
                }
    }

    private static bool IsAuthTokenExpired()
    {
        return authToken == null || authToken.ExpiresOn <= DateTimeOffset.UtcNow.AddMinutes(1); // Adding a buffer time of 1 minute
    }


    // Public Members
    /////////////////////////////////////////////////////////////

    // the response object
    public struct Resp
    {
        public string? Request;
        public HttpResponseMessage? Response;
        public string? Content;
        public HttpStatusCode StatusCode;
    }
    public static async Task<Resp> GetUserMessages(string userId, DateTime startDateTime, DateTime endDateTime, string model, string? nextLink)
    {
        string request;

        if (!string.IsNullOrEmpty(nextLink))
            request = nextLink;
        else
            request = $"https://graph.microsoft.com/v1.0/users/{userId}/chats/getAllMessages?model={model}&$top={cfg.GetUserMessagesBatchSize}&$filter=lastModifiedDateTime%20gt%20{startDateTime:O}%20and%20lastModifiedDateTime%20lt%20{endDateTime:O}";

        return await GraphCall(request);
    }
}