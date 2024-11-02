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

using GraphExportAPIforMicrosoftTeamsSample.DB;
using GraphExportAPIforMicrosoftTeamsSample.Helpers;
using GraphExportAPIforMicrosoftTeamsSample.Loaders;
using GraphExportAPIforMicrosoftTeamsSample.Utility;

namespace GraphExportAPIforMicrosoftTeamsSample;

internal class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // Load AppConfig
            AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

            // Launch Monitor Task - Write Status to Console and Log File
            MonitorHelper.InitMonitor();

            LoggerHelper.WriteToConsoleAndLog($"Initiating {Constants.SAMPLENAME}", ConsoleColor.Cyan);

            // Check for DB Reset request
            if (cfg.ResetDB == 1)
            {
                LoggerHelper.WriteToConsoleAndLog("ResetDB == 1. Resetting Database...\n", ConsoleColor.Magenta);

                // Reset the database
                await DbHelper.ResetDB();

                // Create the database tables
                using (SQLDatabaseContext db = new SQLDatabaseContext())
                    db.Database.EnsureCreated();

                // Update Database Version
                await DbHelper.UpdSertDBVersion(Constants.DBVERSION);
            }

            // Check DB Version
            // Throw exception if not expected version
            await DbHelper.CheckDBVersion();

            // Load MailBox Files (allMalbox.json) into the database
            await MailBoxesHelper.Start();

            // Load Data From Teams Export Graph API into the database
            await GraphLoader.LoadUserMailboxes();

            // Stop Monitor Task / Print Final Status
            MonitorHelper.SummaryMonitor();

            // Set the exit code to indicate success
            Environment.ExitCode = 0;
        }
        catch (IOFatalException ex)
        {
            // Fatal IO Exception Handler
            LoggerHelper.WriteToConsole($"Exception:\n{ex.Message}\n", ConsoleColor.Red);
            LoggerHelper.WriteToConsole($"Stack:\n{ex.StackTrace}\n", ConsoleColor.Blue);

            // Handling InnerException(s)
            Exception? innerException = ex.InnerException;
            while (innerException != null)
            {
                LoggerHelper.WriteToConsole($"Inner Exception:\n{innerException.Message}\n", ConsoleColor.Red);
                LoggerHelper.WriteToConsole($"Stack:\n{innerException.StackTrace}\n", ConsoleColor.Blue);
                innerException = innerException.InnerException;
            }

            // Set the exit code to indicate failure
            Environment.ExitCode = 1;
        }
        catch (Exception ex)
        {
            // Global System Exception Handler
            LoggerHelper.WriteToConsoleAndLog($"Exception:\n{ex.Message}\n", ConsoleColor.Red);
            LoggerHelper.WriteToConsoleAndLog($"Stack:\n{ex.StackTrace}\n", ConsoleColor.Blue);

            // Handling InnerException(s)
            Exception? innerException = ex.InnerException;
            while (innerException != null)
            {
                LoggerHelper.WriteToConsoleAndLog($"Inner Exception:\n{innerException.Message}\n", ConsoleColor.Red);
                LoggerHelper.WriteToConsoleAndLog($"Stack:\n{innerException.StackTrace}\n", ConsoleColor.Blue);
                innerException = innerException.InnerException;
            }

            // Set the exit code to indicate failure
            Environment.ExitCode = 1;
        }
    }
}
