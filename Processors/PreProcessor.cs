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

using System.Data;
using Microsoft.Data.SqlClient;

using GraphExportAPIforMicrosoftTeamsSample.DB;
using GraphExportAPIforMicrosoftTeamsSample.Helpers;
using GraphExportAPIforMicrosoftTeamsSample.Types;

namespace GraphExportAPIforMicrosoftTeamsSample.Processors;

internal static class PreProcessor
{
    // Pre-Process the user messages
    public static async Task Process(string mailbox, UserMessage m)
    {
        const string QUERY = @"
MERGE INTO dbo.UserMessage AS target
USING (
VALUES
    (@chatId, @mailbox, @idKey, @createdDateTime, @lastModifiedDateTime, @rawJson, @folderDate, @fromid, @fromdisplayname, @fromidentitytype, @fromtenantid, @fromtype)
) AS source (chatId, mailbox, idKey, createdDateTime, lastModifiedDateTime, rawJson, folderDate, fromid, fromdisplayname, fromidentitytype, fromtenantid, fromtype)
ON (target.mailbox = source.mailbox AND target.chatId = source.chatId AND target.idKey = source.idKey)
WHEN NOT MATCHED BY TARGET THEN
    INSERT (chatId, mailbox, idKey, createdDateTime, lastModifiedDateTime, rawJson, folderDate, fromid, fromdisplayname, fromidentitytype, fromtenantid, fromtype)
    VALUES (source.chatId, source.mailbox, source.idKey, source.createdDateTime, source.lastModifiedDateTime, source.rawJson, source.folderDate, source.fromid, source.fromdisplayname, source.fromidentitytype, source.fromtenantid, source.fromtype);";

        // Validate the input
        if (string.IsNullOrEmpty(mailbox))
            throw new ArgumentNullException("Mailbox");

        if (m.chatId == null)
            throw new ArgumentNullException("ChatId");

        if (m.id == null)
            throw new ArgumentNullException("Id");

        // agument the object with the raw json and mailbox
        m.rawJson = JsonSerializer.Serialize(m);
        m.mailbox = mailbox;

        // Add to Database
        SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@chatId", m.chatId),
                    new SqlParameter("@mailbox", m.mailbox),
                    new SqlParameter("@idKey", m.idKey),
                    new SqlParameter("@createdDateTime", SqlDbType.DateTime2) { Value = m.createdDateTime },
                    new SqlParameter("@lastModifiedDateTime", SqlDbType.DateTime2) { Value = m.lastModifiedDateTime },
                    new SqlParameter("@rawJson", m.rawJson),
                    new SqlParameter("@folderDate", SqlDbType.DateTime2) { Value = m.folderDate },
                    new SqlParameter("@fromid", m.fromid ?? (object)DBNull.Value),
                    new SqlParameter("@fromdisplayname", m.fromdisplayname ?? (object)DBNull.Value),
                    new SqlParameter("@fromidentitytype", m.fromidentitytype ?? (object)DBNull.Value),
                    new SqlParameter("@fromtenantid", m.fromtenantid ?? (object)DBNull.Value),
                    new SqlParameter("@fromtype", m.fromtype ?? (object)DBNull.Value)
                };
        await DbHelper.ExecuteSqlCommandAsync(QUERY, parameters);

        // Increment the monitor counter
        MonitorHelper.AddUserMessageProcessedCnt();
    }
}