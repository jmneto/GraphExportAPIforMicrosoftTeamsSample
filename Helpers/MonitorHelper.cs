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

using System.Text;
using System.Diagnostics;

using GraphExportAPIforMicrosoftTeamsSample.Utility;

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// Print Real Time Monitor Info and Execution statistics data
internal static class MonitorHelper
{
    // Progress
    private static double mailBoxTableProgress = 0;

    public static void SetMailBoxTableProgress(double progress)
    {
        Interlocked.Exchange(ref mailBoxTableProgress, progress);
        UpdateMonitor();
    }

    // Pre Processor
    private static long itemPreProcessedCnt = 0;
    private static readonly Stopwatch itemPreProcessedstopwatch = new Stopwatch();
    private static readonly SemaphoreSlim itemPreProcessedsemaphore = new SemaphoreSlim(1, 1);
    private static long itemPreProcessedCntLast = 0;
    private static int itemPreProcessedCntPerSecond = 0;
    private static int itemPreProcessedCntPerSecondAvg = 0;
    private static Averager itemPreProcessedCntPerSecondAvgAverager = new Averager();

    public static void AddUserMessageProcessedCnt()
    {
        Interlocked.Increment(ref itemPreProcessedCnt);
        Interlocked.Increment(ref itemPreProcessedCntLast);

        CalculateRateSeconds(itemPreProcessedstopwatch, itemPreProcessedsemaphore, ref itemPreProcessedCntLast, ref itemPreProcessedCntPerSecond, ref itemPreProcessedCntPerSecondAvg, itemPreProcessedCntPerSecondAvgAverager);
        UpdateMonitor();
    }

    // SQL Server Stats
    private static long sqlQueriesInitiated = 0;
    private static long sqlQueriesCompleted = 0;

    public static void AddSQLQueriesInitiated()
    {
        Interlocked.Increment(ref sqlQueriesInitiated);
        UpdateMonitor();
    }

    public static void AddSQLQueriesCompleted()
    {
        Interlocked.Increment(ref sqlQueriesCompleted);
        UpdateMonitor();
    }

    // Storage Read Stats
    private static long storageBytesRead = 0;
    private static Stopwatch storageReadsStopWatch = new Stopwatch();
    private static readonly SemaphoreSlim storageReadsSemaphore = new SemaphoreSlim(1, 1);
    private static long storageBytesReadLast = 0;
    private static int storageBytesReadsec = 0;
    private static int storageBytesReadsecAvg = 0;
    private static Averager storageBytesReadsecAvgAvgAverager = new Averager();

    public static void AddStorageBytesRead(long bytes)
    {
        Interlocked.Add(ref storageBytesRead, bytes);
        Interlocked.Add(ref storageBytesReadLast, bytes);

        CalculateRateSeconds(storageReadsStopWatch, storageReadsSemaphore, ref storageBytesReadLast, ref storageBytesReadsec, ref storageBytesReadsecAvg, storageBytesReadsecAvgAvgAverager);
        UpdateMonitor();
    }

    // Storage Write Stats
    private static long storageBytesWriten = 0;
    private static Stopwatch storageWritesStopwatch = new Stopwatch();
    private static readonly SemaphoreSlim storageWritesSemaphore = new SemaphoreSlim(1, 1);
    private static long storageBytesWriteLast = 0;
    private static int storageBytesWritesec = 0;
    private static int storageBytesWritesecAvg = 0;
    private static Averager storageBytesWritesecAvgAverager = new Averager();

    public static void AddStorageBytesWriten(long bytes)
    {
        Interlocked.Add(ref storageBytesWriten, bytes);
        Interlocked.Add(ref storageBytesWriteLast, bytes);

        CalculateRateSeconds(storageWritesStopwatch, storageWritesSemaphore, ref storageBytesWriteLast, ref storageBytesWritesec, ref storageBytesWritesecAvg, storageBytesWritesecAvgAverager);
        UpdateMonitor();
    }

    // Monitor Methods
    public static void SummaryMonitor()
    {
        LoggerHelper.WriteToConsoleAndLog("\nProcessing complete.", ConsoleColor.White);
        LoggerHelper.WriteToConsoleAndLog("Final Report:", ConsoleColor.White);
        string output = DisplayProgressInfo();
        LoggerHelper.WriteToConsoleAndLog(output, ConsoleColor.Blue);
        LoggerHelper.WriteToConsoleAndLog($"\nStart time: {MonitorHelper.startTime:yyyy-MM-ddTHH:mm:ss.fff}", ConsoleColor.White);
        LoggerHelper.WriteToConsoleAndLog($"End time: {DateTime.UtcNow:yyyy-MM-ddTHH:mm:ss.fff}", ConsoleColor.White);
        LoggerHelper.WriteToConsoleAndLog("\nSample Application Complete", ConsoleColor.Green);
    }

    public static void InitMonitor()
    {
        LoggerHelper.WriteToConsoleAndLog(CollectConfig(), ConsoleColor.Yellow);
    }

    // calculate rate per second
    private static void CalculateRateSeconds(Stopwatch stopwatch, SemaphoreSlim semaphore, ref long counter, ref int ratesec, ref int ratesecAvg, Averager averager)
    {
        if (semaphore.Wait(0)) // Only the one thread can enter this block, all other threads walk around
        {
            try
            {
                if (stopwatch.IsRunning == false)
                    stopwatch.Start();

                TimeSpan elapsedTime = stopwatch.Elapsed;
                if (elapsedTime.Milliseconds >= 100)
                {
                    double ratedbl = ((double)counter / elapsedTime.Milliseconds * 1000);

                    ratesecAvg = (int)averager.AddValueGetAverage(ratedbl);

                    ratesec = (int)ratedbl;

                    Interlocked.Exchange(ref counter, 0);
                    stopwatch.Restart();
                }
            }
            finally
            {
                semaphore.Release();
            }
        }
    }

    // Averager
    private class Averager()
    {
        // Circular Cache to calculate average in a super efficient way - genius jmneto
        private const int MAXTICKRECORDS = 10;
        private double[]? _tickRecord;
        private int _lastIdx;

        public double AddValueGetAverage(double value)
        {
            // first time only initialize the array
            if (_tickRecord == null)
            {
                _tickRecord = new double[MAXTICKRECORDS].Select(h => value).ToArray(); // Trick to Init all elements the first time
                _lastIdx = -1;
            }

            // values Circular cache
            _tickRecord[++_lastIdx == MAXTICKRECORDS ? _lastIdx = 0 : _lastIdx] = value;

            return _tickRecord.Average();
        }
    }

    // Memory Usage
    private static double GetProcessMemoryUsageInMB()
    {
        // Get the current process
        Process currentProcess = Process.GetCurrentProcess();

        // Get the working set size (physical memory usage) in bytes
        long privateMemorySize64 = currentProcess.PrivateMemorySize64;

        // Convert bytes to megabytes
        double memoryUsageMB = privateMemorySize64 / (1024.0 * 1024.0);

        return memoryUsageMB;
    }

    // Task Info
    private static List<TaskInfo> tskInfo = new List<TaskInfo>();
    private struct TaskInfo
    {
        public string TaskName;
        public int Limit;
        public int Count;
    }
    public static void AddTaskInfo(string taskName, int limit, int count)
    {
        // add or update to TaskInfo base on TaskName
        lock (tskInfo)
        {
            TaskInfo task = tskInfo.FirstOrDefault(x => x.TaskName == taskName);
            if (task.TaskName == null)
            {
                tskInfo.Add(new TaskInfo { TaskName = taskName, Limit = limit, Count = count });
            }
            else
            {
                task.Limit = limit;
                task.Count = count;

                // update back to the list
                tskInfo[tskInfo.FindIndex(x => x.TaskName == taskName)] = task;
            }
        }
    }

    // Collect Info Method
    private static string CollectConfig()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{Constants.SAMPLENAME} - Code Version:{Constants.VERSION} - Required Database Version:{Constants.DBVERSION}");

        sb.AppendLine($"\nGraph API Configuration");
        sb.AppendLine($"API Tenant Id: {cfg.TenantId}");
        sb.AppendLine($"API Application Id: {cfg.ApplicationId}");
        sb.AppendLine($"API Licensing Model: {cfg.MicrosoftTeamsAPIModel}");
        sb.AppendLine($"API Extraction Interval (shown as localtime): From {cfg.StartDateTimeUTC:G} to {cfg.EndDateTimeUTC:G}");
        sb.AppendLine($"API Batch Size (https://graph.microsoft.com/v1.0/users/{{userId}}/chats/getAllMessages): {cfg.GetUserMessagesBatchSize}");


        sb.AppendLine($"\nMultitasking Configuration");
        sb.AppendLine($"Mailbox Loader Task Limit: {cfg.MailBoxLoaderTaskLimit} Tasks");
        sb.AppendLine($"Graph Loader Task Limit: {cfg.GraphLoaderTaskLimit} Tasks");
        sb.AppendLine($"PreProcessor Task Limit: {cfg.PreProcTaskLimit}  Tasks");

        sb.AppendLine($"ResetDB: {cfg.ResetDB}");

        return sb.ToString();
    }

    // Collect Progress Method
    private static string DisplayProgressInfo()
    {
        // Cache some info
        currenttime = DateTime.UtcNow;
        lastPrivateBytes = GetProcessMemoryUsageInMB();

        // Prepare to write to Screen
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"\nPRE|Items Processed: {itemPreProcessedCnt:N0}");
        sb.AppendLine($"PRE|Items Processed/Sec Avg: {itemPreProcessedCntPerSecondAvg:N0}");

        sb.AppendLine($"\nSQL|Queries Initiated: {sqlQueriesInitiated:N0}");
        sb.AppendLine($"SQL|Queries Completed: {sqlQueriesCompleted:N0}");

        sb.AppendLine($"\nStorage|BytesRead: {storageBytesRead:N0}");
        sb.AppendLine($"Storage|BytesRead/Sec Avg: {storageBytesReadsecAvg:N0}");
        sb.AppendLine($"Storage|BytesWritten: {storageBytesWriten:N0}");
        sb.AppendLine($"Storage|BytesWritten/Sec Avg: {storageBytesWritesecAvg:N0}");

        sb.AppendLine($"\nProcess|PrivateMemorySize: {lastPrivateBytes:N0} MB");

        lock (tskInfo)
            foreach (TaskInfo task in tskInfo)
                sb.AppendLine($"Process|Task: {task.TaskName} Task Limit: {task.Limit} Task Count: {task.Count}");

        sb.AppendLine($"\nSample Application Progress|PRE|Mailbox-Table % of work done: {mailBoxTableProgress:P2}");
        sb.AppendLine($"\nSample Application Progress|Total Elapsed Time: {currenttime.Subtract(startTime).ToString(@"d\ hh\:mm\:ss")}");
        return sb.ToString();
    }

    // cache private bytes
    private static double lastPrivateBytes = 0;

    // Monitor Update 
    private static DateTime lastUpdateTime = DateTime.MinValue;
    private static readonly SemaphoreSlim semaphoreUpdMonitor = new SemaphoreSlim(1, 1);
    private static void UpdateMonitor()
    {
        if (semaphoreUpdMonitor.Wait(0))
        {
            if (DateTime.Now - lastUpdateTime >= TimeSpan.FromSeconds(5))
            {
                lastUpdateTime = DateTime.Now;
                string output = DisplayProgressInfo();
                LoggerHelper.WriteToConsoleAndLog(output, ConsoleColor.Blue);
            }

            // Release Semaphore
            semaphoreUpdMonitor.Release();
        }
    }

    // Get app Settings
    private readonly static AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

    // local fields
    private static readonly DateTime startTime = DateTime.UtcNow;
    private static DateTime currenttime = DateTime.UtcNow;
}