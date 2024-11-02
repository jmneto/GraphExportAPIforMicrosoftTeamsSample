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

using Azure.Identity;
using Azure.Core;
using Azure.Storage;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs.Specialized;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using Polly;

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// This is a helper class for the Azure Storage Functions
internal class StorageHelper
{
    // Private Members
    ////////////////////////////

    // Azure Storage References for performance
    private BlobServiceClient? blobServiceClient = null;
    private BlobContainerClient? containerClient = null;
    private BlobOpenReadOptions? blobReadOptions = null;
    private BlobUploadOptions? blobUploadOptions = null;

    // Cache for Objects to avoid re-authentication
    private class Cache
    {
        public BlobServiceClient? blobServiceClient;
        public BlobContainerClient? containerClient;
        public BlobOpenReadOptions? blobReadOptions;
        public BlobUploadOptions? blobUploadOptions;
    }
    // key = $"{storageAccountUrl}|{storageAccountConnectionString}|{storageContainer}";
    private readonly static ConcurrentDictionary<string, Cache> cache = new();

    // Connection info
    private readonly string? storageAccountConnectionString = null;
    private readonly string? storageAccountUrl = null;
    private readonly string? storageContainer = null;

    // Other Control Fields
    private bool initialized = false;
    private readonly object lck = new();

    // Avoid Default Constructor
    StorageHelper() { }

    // Function to check/initialize the storage
    private void CheckInit()
    {
        // Return if already initialized
        if (initialized)
            return;

        lock (lck)
            try
            {
                // Return if already initialized
                if (initialized)
                    return;

                string key = $"{storageAccountUrl}|{storageAccountConnectionString}|{storageContainer}";

                if (cache.ContainsKey(key))
                {
                    blobServiceClient = cache[key].blobServiceClient;
                    containerClient = cache[key].containerClient;
                    blobReadOptions = cache[key].blobReadOptions;
                    blobUploadOptions = cache[key].blobUploadOptions;
                    initialized = true;
                    return;
                }

                LoggerHelper.WriteToConsole($"Initializing Storage Account Client Instance: {storageContainer}", ConsoleColor.DarkGray);

                // Provide the client configuration options for connecting to Azure Blob Storage
                BlobClientOptions blobServiceClientOptions = new BlobClientOptions
                {
                    Retry = {
                            Delay = TimeSpan.FromSeconds(2),
                            MaxRetries = 10,
                            Mode = RetryMode.Exponential,
                            MaxDelay = TimeSpan.FromSeconds(10),
                            NetworkTimeout = TimeSpan.FromSeconds(100)
                        }
                };

                // Azure Identity Connection
                if (!string.IsNullOrEmpty(storageAccountUrl) && string.IsNullOrEmpty(storageAccountConnectionString))
                {
                    // Configure DefaultAzureCredential with custom options
                    DefaultAzureCredentialOptions credentialOptions = new DefaultAzureCredentialOptions
                    {
                        Retry =
                            {
                                MaxRetries = 6, // Customize retries if needed
                                NetworkTimeout = TimeSpan.FromMinutes(2) // Increase the timeout from the default (usually 30s)
                            }
                    };

                    // Use DefaultAzureCredential with the custom options
                    var credential = new DefaultAzureCredential(credentialOptions);

                    blobServiceClient = new BlobServiceClient(new Uri(storageAccountUrl), credential, blobServiceClientOptions);
                }
                // Storage Account Connection String
                else if (!string.IsNullOrEmpty(storageAccountConnectionString))
                    blobServiceClient = new BlobServiceClient(storageAccountConnectionString, blobServiceClientOptions);

                if (blobServiceClient == null)
                    throw new Exception($"Storage Account Client failed to intitialize for container '{storageContainer}', check configuration settings. Please review README.md");

                // Get and cache a reference to the container
                if (!string.IsNullOrEmpty(storageContainer))
                    containerClient = blobServiceClient.GetBlobContainerClient(storageContainer);

                // Check if we have a good container
                if (containerClient == null)
                    throw new Exception("Storage Container Client failed to initialize, check configuration settings. Please review README.md");

                // Retry policy for checking container existence
                bool containerExists = false;
                Policy
                    .Handle<Exception>()
                    .WaitAndRetryAsync(Utility.Constants.STORAGERETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                        onRetry: (exception, timeSpan, retryCount, context) =>
                        {
                            LoggerHelper.WriteToConsole($"Retrying Storage Account Client. Exception:{exception} Retry in:{timeSpan}, Retry Count:{retryCount}", ConsoleColor.Magenta);
                        })
                    .ExecuteAsync(() =>
                    {
                        containerExists = containerClient.Exists();
                        return Task.CompletedTask;
                    }).Wait();

                if (!containerExists)
                    throw new Exception($"Storage Container {storageContainer} does not exist. Please review README.md");

                // Set the options for reading a blob
                blobReadOptions = new BlobOpenReadOptions(false)
                {
                    // Set the maximum length of a transfer to 50MB.
                    BufferSize = 50 * 1024 * 1024
                };

                // Set the options for uploading a blob
                blobUploadOptions = new BlobUploadOptions
                {
                    TransferOptions = new StorageTransferOptions
                    {
                        // Set the maximum number of workers that 
                        // may be used in a parallel transfer.
                        MaximumConcurrency = 4,

                        // Set the maximum length of a transfer to 50MB.
                        MaximumTransferSize = 50 * 1024 * 1024
                    }
                };

                if (!cache.ContainsKey(key))
                    cache.TryAdd(key, new Cache { blobServiceClient = blobServiceClient, containerClient = containerClient, blobReadOptions = blobReadOptions, blobUploadOptions = blobUploadOptions });

                initialized = true;
            }
            catch (Exception ex)
            {
                throw new IOFatalException("Fatal exception in StorageHelper", ex);
            }
    }


    // Custom function to check if a file name matches a wildcard pattern
    private bool IsMatchingWildcard(string fileName, string wildcardPattern)
    {
        string pattern = wildcardPattern.ToLower();
        fileName = fileName.ToLower();

        if (pattern == "*")
            return true;

        if (pattern.Contains("*"))
        {
            // Convert wildcard pattern to a regular expression pattern
            pattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";

            return Regex.IsMatch(fileName, pattern);
        }

        return fileName == pattern;
    }

    // Public Members
    // /////////////////////

    // Class Constructor
    public StorageHelper(string storageContainer, string? storageAccountUrl, string? storageAccountConnectionString)
    {
        this.storageAccountUrl = storageAccountUrl;
        this.storageAccountConnectionString = storageAccountConnectionString;
        this.storageContainer = storageContainer;
        initialized = false;
    }

    // Write a stream to a storage
    public void WriteStream(string fullFileName, string contentType, Stream stream)
    {
        // Check if we are initialized
        CheckInit();

        // Write data to a blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient blobClient = containerClient.GetBlobClient(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Rewind the memory stream back to the beginning
        stream.Seek(0, SeekOrigin.Begin);

        // Upload data
        blobClient.Upload(stream, blobUploadOptions);

        // add context-type
        blobClient.SetHttpHeaders(new BlobHttpHeaders { ContentType = contentType });

        // Add Monitor
        MonitorHelper.AddStorageBytesWriten(stream.Length);

        // Close the memory stream
        stream.Close();
        stream.Dispose();
    }

    // get a stream from storage
    public Stream? GetStream(string fullFileName)
    {
        // Check if we are initialized
        CheckInit();

        // Read data from a blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient readBlobClient = containerClient.GetBlobClient(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        if (readBlobClient.Exists())
            return readBlobClient.OpenRead(blobReadOptions);

        return null;
    }

    // get a stream from storage
    public string? ReadString(string fullFileName)
    {
        // Check if we are initialized
        CheckInit();

        // Read data from a blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient readBlobClient = containerClient.GetBlobClient(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        if (readBlobClient.Exists())
        {
            BlobDownloadInfo blobDownloadInfo = readBlobClient.Download();

            using (StreamReader reader = new StreamReader(blobDownloadInfo.Content))
                return reader.ReadToEnd();
        }

        return null;
    }

    public BlobClient GetBobClient(string fullFileName)
    {
        // Check if we are initialized
        CheckInit();

        // Read data from a blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient readBlobClient = containerClient.GetBlobClient(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        return readBlobClient;
    }

    // Write a string to a storage (Append)
    // this is slow and should only be used for log files
    private static int blockCount = 0;
    private static string currentBlobName = "";
    private const int MAXBLOCKS = 40000;
    private static readonly string logfilename = $"logfile_{Guid.NewGuid()}.log";

    public void WriteAppendToLog(string contents)
    {
        if (string.IsNullOrEmpty(contents))
            return;

        // Check if we are initialized
        CheckInit();

        // Append data to a blob
        byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(contents);
        using (MemoryStream stream = new MemoryStream(byteArray))
        {
            if (string.IsNullOrEmpty(currentBlobName) || blockCount >= MAXBLOCKS)
            {
                currentBlobName = GetNextBlobName(logfilename);
                blockCount = 0; // Reset block count for new blob
            }

#pragma warning disable CS8602 // Dereference of a possibly null reference.
            AppendBlobClient appendBlobClient = containerClient.GetAppendBlobClient(currentBlobName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            appendBlobClient.CreateIfNotExists();

            int maxBlockSize = appendBlobClient.AppendBlobMaxAppendBlockBytes;

            if (stream.Length <= maxBlockSize)
            {
                appendBlobClient.AppendBlock(stream);
                blockCount++;
            }
            else
            {
                long bytesLeft = stream.Length;

                while (bytesLeft > 0)
                {
                    int blockSize = (int)Math.Min(bytesLeft, maxBlockSize);
                    byte[] buffer = new byte[blockSize];
                    stream.Read(buffer, 0, blockSize);
                    appendBlobClient.AppendBlock(new MemoryStream(buffer));
                    blockCount++;
                    bytesLeft -= blockSize;

                    if (blockCount >= MAXBLOCKS)
                    {
                        currentBlobName = GetNextBlobName(logfilename);
                        appendBlobClient = containerClient.GetAppendBlobClient(currentBlobName);
                        appendBlobClient.CreateIfNotExists();
                        blockCount = 0;
                    }
                }
            }
        }

        MonitorHelper.AddStorageBytesWriten(byteArray.Length);

        string GetNextBlobName(string baseFileName)
        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            BlobClient blobClient = containerClient.GetBlobClient(baseFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            if (!blobClient.Exists())
            {
                return baseFileName;
            }
            else
            {
                int index = 1;
                while (true)
                {
                    string newBlobName = $"{baseFileName.Replace(".log", "")}_{index}.log";
                    blobClient = containerClient.GetBlobClient(newBlobName);
                    if (!blobClient.Exists())
                        return newBlobName;
                    index++;
                }
            }
        }
    }


    // Write a string to a storage (overwrite)
    public void WriteString(string fullFileName, string contents)
    {
        // Check if we have any contents
        if (string.IsNullOrEmpty(contents))
            return;

        // Check if we are initialized
        CheckInit();

        // Write data to a blob
        byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(contents);
        using (MemoryStream stream = new MemoryStream(byteArray))
        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            BlobClient blobClient = containerClient.GetBlobClient(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            blobClient.Upload(stream, blobUploadOptions);
        }

        // Add Monitor
        MonitorHelper.AddStorageBytesWriten(byteArray.Length);
    }

    public delegate Task DelegateTraverse(Stream reader, string containerName, string blobName);

    public async Task TraverseContainer(string prefix, string wildcardPattern, DelegateTraverse callback)
    {
        // Check if we are initialized
        CheckInit();

        // Get the blobs
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Azure.AsyncPageable<BlobItem> blobItems = containerClient.GetBlobsAsync(BlobTraits.None, BlobStates.None, prefix);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Loop through the blobs
        await foreach (BlobItem blobItem in blobItems)
        {
            // Check if the file matches the wildcard pattern
            if (!IsMatchingWildcard(blobItem.Name, wildcardPattern))
                continue;

            // This is a "virtual" folder
            if (blobItem.Properties.ContentLength == 0)
                await TraverseContainer(blobItem.Name + "/", wildcardPattern, callback);
            else
            {
                // Read data from a blob
                BlobClient readBlobClient = containerClient.GetBlobClient(blobItem.Name);

                callback(readBlobClient.OpenRead(blobReadOptions), containerClient.Name, blobItem.Name).Wait();
            }
        }
    }

    public async Task TraverseFolder(string prefix, string wildcardPattern, DelegateTraverse callback)
    {
        // Check if we are initialized
        CheckInit();

        // Get the blobs
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Azure.AsyncPageable<BlobItem> blobItems = containerClient.GetBlobsAsync(BlobTraits.None, BlobStates.None, prefix);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Loop through the blobs
        await foreach (BlobItem blobItem in blobItems)
        {
            // Check if the file matches the wildcard pattern
            if (!IsMatchingWildcard(blobItem.Name, wildcardPattern))
                continue;

            // Read data from a blob
            BlobClient readBlobClient = containerClient.GetBlobClient(blobItem.Name);

            callback(readBlobClient.OpenRead(blobReadOptions), containerClient.Name, blobItem.Name).Wait();
        }
    }

    // Check if the file exists
    public bool CheckFileExists(string fullFileName)
    {
        // Check if we are initialized
        CheckInit();

#pragma warning disable CS8602 // Dereference of a possibly null reference.
        if (containerClient.Exists())
            if (containerClient.GetBlobClient(fullFileName).Exists())
                return true;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        return false;
    }

    // Check if the container exists
    public bool ContainerExists(string containerName)
    {
        // Check if we are initialized
        CheckInit();

#pragma warning disable CS8602 // Dereference of a possibly null reference.
        return containerClient.Exists();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
    }

    // Check if the container is empty
    public bool IsContainerEmpty(string containerName)
    {
        // Check if we are initialized
        CheckInit();

        // Get the list of blobs in the container
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        Azure.Pageable<BlobItem> blobs = containerClient.GetBlobs();
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Check if the list is empty
        return (blobs.Count() == 0);
    }

    // Delere a file
    public void DeleteFile(string fullFileName)
    {
        // Check if we are initialized
        CheckInit();

        // Delete the file
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        if (CheckFileExists(fullFileName))
            containerClient.DeleteBlob(fullFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

    }

    // Move a file to another location
    public async Task MoveFile(string sourceFileName, string destinationFileName)
    {
        if (string.IsNullOrEmpty(sourceFileName))
            throw new ArgumentNullException(nameof(sourceFileName));

        if (string.IsNullOrEmpty(destinationFileName))
            throw new ArgumentNullException(nameof(destinationFileName));

        if (sourceFileName.Equals(destinationFileName))
            return;

        // Check if we are initialized
        CheckInit();

        // Get references to source and destination blobs
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient sourceBlobClient = containerClient.GetBlobClient(sourceFileName);
        BlobClient destinationBlobClient = containerClient.GetBlobClient(destinationFileName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // check if destination file already exists
        if (destinationBlobClient.Exists())
            throw new Exception($"Destination file {destinationFileName} already exists");

        // Check if the source file exists
        if (sourceBlobClient.Exists())
        {
            // Start copying from the source to the destination
            await destinationBlobClient.StartCopyFromUriAsync(sourceBlobClient.Uri);

            // Wait until the copy operation is completed
            while (destinationBlobClient.GetProperties().Value.CopyStatus == CopyStatus.Pending)
                await Task.Delay(100); // You may adjust the sleep duration based on your needs

            // Check if the copy operation was successful
            if (destinationBlobClient.GetProperties().Value.CopyStatus == CopyStatus.Success)
            {
                // Delete the source blob after successful copy
                sourceBlobClient.Delete();
            }
            else
            {
                throw new Exception($"Failed to move file from {sourceFileName} to {destinationFileName}");
            }

        }
        else
        {
            throw new FileNotFoundException($"Source file {sourceFileName} not found");
        }
    }

    // Method to acquire a lease on a blob
    public string AcquireBlobLease(string blobName, TimeSpan? leaseDuration = null, string? existingLeaseId = null)
    {
        if (string.IsNullOrEmpty(blobName))
            throw new ArgumentNullException(nameof(blobName));

        // Check if we are initialized
        CheckInit();

        // Default lease duration is 60 seconds if not specified
        leaseDuration ??= TimeSpan.FromSeconds(60);

        // Get the blob client for the specified blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient blobClient = containerClient.GetBlobClient(blobName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Create a blob lease client for the blob
        BlobLeaseClient leaseClient = blobClient.GetBlobLeaseClient(existingLeaseId);

        // Try to acquire a lease on the blob
        BlobLease lease = leaseClient.Acquire(leaseDuration.Value);

        // Return the lease ID
        return lease.LeaseId;
    }

    public void ReleaseBlobLease(string blobName, string existingLeaseId)
    {
        if (string.IsNullOrEmpty(blobName))
            throw new ArgumentNullException(nameof(blobName));

        // Check if we are initialized
        CheckInit();

        // Get the blob client for the specified blob
#pragma warning disable CS8602 // Dereference of a possibly null reference.
        BlobClient blobClient = containerClient.GetBlobClient(blobName);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

        // Create a blob lease client for the blob using the provided lease ID
        BlobLeaseClient leaseClient = blobClient.GetBlobLeaseClient(existingLeaseId);

        // Release the lease
        leaseClient.Release();
    }
}