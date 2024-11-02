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
using System.Data;
using Microsoft.Data.SqlClient;
using System.Net.Sockets;

using GraphExportAPIforMicrosoftTeamsSample.Helpers;
using GraphExportAPIforMicrosoftTeamsSample.Utility;

namespace GraphExportAPIforMicrosoftTeamsSample.DB;

// This class is a Database Helper class
// It implements some general database functions
// Plus functions to check and update the database version. I use the database version to keep track of the database schema version and to auto update the database schema when needed
// All database objects are created using Entity Framework Code First
internal class DbHelper
{
    private static readonly AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

    // Log Retry Messager
    public static void LogSQLRetryMessage(Exception exception, TimeSpan timeSpan, int retryCount, string? sqlCommandText = null)
    {
        if (!string.IsNullOrEmpty(sqlCommandText))
            LoggerHelper.WriteToConsoleAndLog($"SQL Retry {retryCount} Message: {exception.Message}. Waiting {timeSpan} before next retry. SQL Command:\n{sqlCommandText}\n", ConsoleColor.Magenta);
        else
            LoggerHelper.WriteToConsoleAndLog($"SQL Retry {retryCount} Message: {exception.Message}. Waiting {timeSpan} before next retry.", ConsoleColor.Magenta);
    }

    // Update Database Statistics for a table
    public static async Task UpdateDBStatistics(string tableName)
    {
        string sqlCommandText = $"UPDATE STATISTICS {tableName};";
        LoggerHelper.WriteToConsoleAndLog($"Updating statistics for {tableName}...", ConsoleColor.Yellow);
        await ExecuteSqlCommandAsync(sqlCommandText);
    }

    // Create a new database version or update the existing one
    // The version is stored in the database as an extended property
    public static async Task UpdSertDBVersion(string version)
    {
        string? currentVersion = await GetDBVersion();
        string sqlCommandText;

        if (string.IsNullOrEmpty(currentVersion))
        {
            // If the version does not exist, insert it
            sqlCommandText = $"EXEC sp_addextendedproperty N'Schema_Version', N'{version}', NULL, NULL, NULL, NULL, NULL;";
        }
        else
        {
            // If the version exists, update it
            sqlCommandText = $"EXEC sp_updateextendedproperty N'Schema_Version', N'{version}';";
        }

        await ExecuteSqlCommandAsync(sqlCommandText);
    }

    // Delete the database version
    // The version is stored in the database as an extended property
    public static async Task DeleteDBVersion()
    {
        const string sqlCommandText = @"
IF EXISTS (
    SELECT 1 
    FROM 
        sys.extended_properties 
    WHERE 
        class_desc = 'DATABASE' 
    AND name = 'Schema_Version'
)
BEGIN
    EXEC sp_dropextendedproperty @name = N'Schema_Version';
END";
        await ExecuteSqlCommandAsync(sqlCommandText);
    }

    // Get the database version
    // The version is stored in the database as an extended property
    public static async Task<string?> GetDBVersion()
    {
        const string QUERYGETDBVERSION = @"
SELECT 
    value as Schema_Version 
FROM 
    sys.extended_properties 
WHERE 
    class_desc = 'DATABASE' 
AND name = 'Schema_Version'";

        string? version = null;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, QUERYGETDBVERSION);
                })
            .ExecuteAsync(async () =>
                {
                    // Generates User RawChat
                    using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                    {
                        // Open the connection
                        await connection.OpenAsync();

                        using (SqlCommand command = new SqlCommand(QUERYGETDBVERSION, connection))
                        {
                            // Execute the query and obtain a SqlDataReader to read the results
                            using (SqlDataReader reader = await command.ExecuteReaderAsync())
                            {

                                while (await reader.ReadAsync())
                                {
                                    version = reader.GetString(0);
                                }
                            }
                        }
                    }
                });

        return version;
    }

    // Check the database version
    // Throw an exception if the version is not the expected one
    public static async Task CheckDBVersion()
    {
        LoggerHelper.WriteToConsoleAndLog($"Checking Database...", ConsoleColor.Green);

        string? dbversion = await DbHelper.GetDBVersion();
        if (dbversion != Constants.DBVERSION)
        {
            LoggerHelper.WriteToConsoleAndLog($"Database Version Mismatch. Expected: ({Constants.DBVERSION}) - Found: ({dbversion})", ConsoleColor.Red);
            LoggerHelper.WriteToConsoleAndLog($"Please run the Sample Application with option ResetDB = 1. All data in the DB will be deleted", ConsoleColor.Red);
            throw new Exception("Database Version Mismatch. Check Readme.md");
        }

        LoggerHelper.WriteToConsoleAndLog($"Database OK. DB Version {dbversion}", ConsoleColor.Green);
    }

    // Reset the database
    // Drop all tables and foreign key constraints
    public static async Task ResetDB()
    {
        const string sqlCommandText = @"
-- Declare variables to hold dynamically generated SQL commands
DECLARE @sql NVARCHAR(MAX)

-- Drop all Foreign Key constraints
SET @sql = ''
SELECT @sql = @sql + 'ALTER TABLE ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + '.' + QUOTENAME(t.name) + ' DROP CONSTRAINT ' + QUOTENAME(fk.name) + ';'
FROM sys.foreign_keys AS fk
INNER JOIN sys.tables AS t 
    ON fk.parent_object_id = t.object_id

EXEC sp_executesql @sql

-- Drop all tables
WHILE EXISTS (SELECT 1 FROM sys.tables where type = 'U')
BEGIN
    SELECT TOP 1 @sql = 'DROP TABLE ' + QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name)
    FROM sys.tables

    EXEC sp_executesql @sql
END";
        await ExecuteSqlCommandAsync(sqlCommandText);

        await DeleteDBVersion();
    }

    // Execute a SQL command
    // This function is used to execute SQL commands that do not return a result set
    public static async Task<int> ExecuteSqlCommandAsync(string sqlCommandText)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        int rowsAffected = 0;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
                {
                    using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                    {
                        await connection.OpenAsync();

                        using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                        {
                            rowsAffected = await command.ExecuteNonQueryAsync();
                        }
                    }
                });

        MonitorHelper.AddSQLQueriesCompleted();

        return rowsAffected;
    }

    // Execute a SQL Command with parameters
    // This function is used to execute SQL commands that do not return a result set
    public static async Task<int> ExecuteSqlCommandAsync(string sqlCommandText, SqlParameter[]? parameters)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        int rowsAffected = 0;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
                   {
                       using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                       {
                           await connection.OpenAsync();

                           using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                           {
                               if (parameters != null)
                                   foreach (var param in parameters)
                                       command.Parameters.Add(((ICloneable)param).Clone());

                               rowsAffected = await command.ExecuteNonQueryAsync();

                               command.Parameters.Clear();
                           }
                       }
                   });

        MonitorHelper.AddSQLQueriesCompleted();

        return rowsAffected;
    }

    // Execute a SQL Command with parameters and return a dataset
    // This function is used to execute SQL commands that return a result set
    public static async Task<DataSet> ExecuteSqlCommandAsync(string sqlCommandText, SqlParameter[]? parameters, string datasetTableName)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        DataSet ds = new DataSet();

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
            {
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                    {
                        if (parameters != null)
                            foreach (var param in parameters)
                                command.Parameters.Add(((ICloneable)param).Clone());

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(ds, datasetTableName);
                        }

                        command.Parameters.Clear();
                    }
                }
            });

        MonitorHelper.AddSQLQueriesCompleted();

        return ds;
    }

    // Execute a SQL Scalar SQL command
    // This function is used to execute SQL commands that return a scalar value
    public static async Task<int> ExecuteScalarIntSqlCommandAsync(string sqlCommandText)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        int result = 0;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
            {
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                result = reader.GetInt32(0);
                            }
                        }
                    }
                }
            });

        MonitorHelper.AddSQLQueriesCompleted();

        return result;
    }


    // Execute a SQL Scalar SQL command with parameters
    // This function is used to execute SQL commands that return a scalar value
    public static async Task<int> ExecuteScalarIntSqlCommandAsync(string sqlCommandText, SqlParameter[]? parameters)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        int result = 0;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
            {
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                    {
                        if (parameters != null)
                            foreach (var param in parameters)
                                command.Parameters.Add(((ICloneable)param).Clone());

                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                result = reader.GetInt32(0);
                            }
                        }

                        command.Parameters.Clear();
                    }
                }
            });

        MonitorHelper.AddSQLQueriesCompleted();

        return result;
    }

    // Execute a SQL Scalar SQL command
    // This function is used to execute SQL commands that return a scalar value
    public static async Task<DateTime?> ExecuteScalarDateTimeSqlCommandAsync(string sqlCommandText)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        DateTime? result = null;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
            {
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                    {
                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                result = reader.GetDateTime(0);
                            }
                        }
                    }
                }
            });

        MonitorHelper.AddSQLQueriesCompleted();

        return result;
    }


    // Execute a SQL Scalar SQL command with parameters
    // This function is used to execute SQL commands that return a scalar value
    public static async Task<DateTime?> ExecuteScalarDateTimeSqlCommandAsync(string sqlCommandText, SqlParameter[]? parameters)
    {
        MonitorHelper.AddSQLQueriesInitiated();

        DateTime? result = null;

        // Define a Polly retry policy for handling transient faults
        await Policy
            .Handle<SqlException>()  // Handle SQL Exceptions which could indicate transient faults
            .Or<IOException>() // Handle IOExceptions which could indicate TCP/IP level errors
            .Or<SocketException>() // Specifically handle SocketExceptions which can occur during network issues
            .Or<TaskCanceledException>() // Handle TaskCanceledException
            .WaitAndRetryAsync(Utility.Constants.SQLRETRYCOUNT, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    DbHelper.LogSQLRetryMessage(exception, timeSpan, retryCount, sqlCommandText);
                })
            .ExecuteAsync(async () =>
            {
                using (SqlConnection connection = new SqlConnection(cfg.SqlConnectionString))
                {
                    await connection.OpenAsync();

                    using (SqlCommand command = new SqlCommand(sqlCommandText, connection))
                    {
                        if (parameters != null)
                            foreach (var param in parameters)
                                command.Parameters.Add(((ICloneable)param).Clone());

                        using (SqlDataReader reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                result = reader.GetDateTime(0);
                            }
                        }

                        command.Parameters.Clear();
                    }
                }
            });

        MonitorHelper.AddSQLQueriesCompleted();

        return result;
    }
}