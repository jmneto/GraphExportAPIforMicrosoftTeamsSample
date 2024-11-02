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

using Microsoft.EntityFrameworkCore;

using GraphExportAPIforMicrosoftTeamsSample.Helpers;
using GraphExportAPIforMicrosoftTeamsSample.Types;

namespace GraphExportAPIforMicrosoftTeamsSample.DB;

// This class is used to create the Entity Framework Core Database Context
// We use Entity Framework Core just to create the database objects
// We use ADO.Net to interact with the database
// We may use entity framework core to read from a few tables, but the bulk of the work is done using ADO.Net
// Mostly the SQL work is done using parameterized SQL queries and ADO.Net
internal class SQLDatabaseContext : DbContext
{
    // Database Objects
    internal DbSet<UserMessage> UserMessage { get; set; }
    internal DbSet<MailBox> MailBox { get; set; }
    internal DbSet<GraphCallProcessed> GraphCallProcessed { get; set; }

    // Configure DB
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        AppConfigHelper.AppConfig cfg = AppConfigHelper.GetAppConfig();

        // Configuring Connection
        // Connection String is stored in AppSettings. See readme.md for details
        optionsBuilder.UseSqlServer(cfg.SqlConnectionString,
                sqlServerOptionsAction: sqlOptions =>
                {
                    sqlOptions.EnableRetryOnFailure(
                    maxRetryCount: 10,
                    maxRetryDelay: TimeSpan.FromSeconds(30),
                    errorNumbersToAdd: null);
                }
            );

        // Adding logging
        // optionsBuilder.LogTo(Console.WriteLine, new[] { DbLoggerCategory.Database.Command.Name }, LogLevel.Information).EnableSensitiveDataLogging();
    }

    // Configuring Database Model
    // Used to create the database objects
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // Messages
        modelBuilder.Entity<UserMessage>(
            ad =>
            {
                // Configure columns
                ad.Ignore(c => c.id);
                ad.Ignore(c => c.replyToId);
                ad.Ignore(c => c.etag);
                ad.Ignore(c => c.messageType);
                ad.Ignore(c => c.lastEditedDateTime);
                ad.Ignore(c => c.deletedDateTime);
                ad.Ignore(c => c.subject);
                ad.Ignore(c => c.summary);
                ad.Ignore(c => c.importance);
                ad.Ignore(c => c.locale);
                ad.Ignore(c => c.webUrl);
                ad.Ignore(c => c.from);
                ad.Ignore(c => c.body);
                ad.Ignore(c => c.attachments);
                ad.Ignore(c => c.mentions);
                ad.Ignore(c => c.reactions);
                ad.Ignore(c => c.messageHistory);
                ad.Ignore(c => c.eventDetail);
                ad.Property(c => c.chatId).HasColumnType("nvarchar(200)");
                ad.Property(c => c.mailbox).HasColumnType("nvarchar(50)");
                ad.Property(c => c.rawJson).HasColumnType("nvarchar(max)");
                ad.Property(c => c.fromid).HasColumnType("nvarchar(100)");
                ad.Property(c => c.fromdisplayname).HasColumnType("nvarchar(200)");
                ad.Property(c => c.fromidentitytype).HasColumnType("nvarchar(100)");
                ad.Property(c => c.fromtenantid).HasColumnType("nvarchar(36)");
                ad.Property(c => c.fromtype).HasColumnType("nvarchar(100)");

                // Configure Keys
                ad.HasKey(c => new { c.mailbox, c.chatId, c.idKey });
            });

        modelBuilder.Entity<GraphCallProcessed>(
          ad =>
          {
              // Configure columns
              ad.Property(c => c.type).HasColumnType("nvarchar(10)");
              ad.Property(c => c.value).HasColumnType("nvarchar(200)");

              // Configure Keys
              ad.HasKey(m => new { m.value, m.type });
          });

        // MailBoxes
        modelBuilder.Entity<MailBox>(
            ad =>
            {
                // Configure columns
                ad.Property(c => c.DisplayName).HasColumnType("nvarchar(100)");
                ad.Property(c => c.ExternalDirectoryObjectId).HasColumnType("nvarchar(50)");
                ad.Property(c => c.PrimarySmtpAddress).HasColumnType("nvarchar(100)");

                // Configure Keys
                ad.HasKey(m => new { m.ExternalDirectoryObjectId });
            });
    }
}
