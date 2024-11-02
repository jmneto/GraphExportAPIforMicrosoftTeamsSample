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

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// This class provides methods to write to the Log Files in the input container and to console
internal static class LoggerHelper
{
    // Read parameters
    private static readonly AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();
    private static readonly object lckWriteToLog = new object();
    private static readonly object lckWriteToConsole = new object();

    // Reference to StorageManagerHelper
    private static readonly StorageHelper sh = new StorageHelper(cfg.LogStorageContainer, cfg.LogStorageAccountUrl, cfg.LogStorageConnectionString);

    // This method will write the message just to the log file
    public static void WriteToLog(string message)
    {
        lock (lckWriteToLog)
        {
            try
            {
                sh.WriteAppendToLog(string.IsNullOrEmpty(message) ? " \n" : $"{DateTime.UtcNow:s} {message}\n");
            }
            catch (Exception ex)
            {
                throw new IOFatalException("Error writing to the Log file:", ex);
            }
        }
    }

    // This method will write the message just to the console
    public static void WriteToConsole(string message, ConsoleColor color = ConsoleColor.White)
    {
        lock (lckWriteToConsole)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(string.IsNullOrEmpty(message) ? string.Empty : $"{DateTime.UtcNow:s} {message}");
            Console.ResetColor();
        }
    }

    // This method will write the message to the console and to the log file
    public static void WriteToConsoleAndLog(string message, ConsoleColor color = ConsoleColor.White)
    {
        WriteToConsole(message, color);
        WriteToLog(message);
    }
}

