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

using System.Text.Json.Serialization;

namespace GraphExportAPIforMicrosoftTeamsSample.Types;

// User API call type
internal class UserMessage
{
    public string? id { get; set; }
    public string? replyToId { get; set; }
    public string? etag { get; set; }
    public string? messageType { get; set; }
    public DateTime? createdDateTime { get; set; }
    public DateTime? lastModifiedDateTime { get; set; }
    public DateTime? lastEditedDateTime { get; set; }
    public DateTime? deletedDateTime { get; set; }
    public string? subject { get; set; }
    public string? summary { get; set; }
    public string? chatId { get; set; }
    public string? importance { get; set; }
    public string? locale { get; set; }
    public string? webUrl { get; set; }

    public EventDetail eventDetail { get; set; } = new EventDetail();

    public From from { get; set; } = new From();
    public Body? body { get; set; }
    public List<Attachment>? attachments { get; set; }
    public List<Reaction>? reactions { get; set; }
    public List<Mention>? mentions { get; set; }
    public List<MessageHistory>? messageHistory { get; set; }

    // Auxiliar properties
    // PK: mailbox, chatId, id
    [JsonIgnore]
    public string? mailbox { get; set; }
    [JsonIgnore]
    public string? rawJson { get; set; }

    // Derived Fields
    [JsonIgnore]
    public long idKey
    {
        get
        {
            if (long.TryParse(id, out long parsedId))
            {
                return parsedId;
            }
            return 0;
        }
        set
        {
            _ = value;
        }
    }

    [JsonIgnore]
    public DateTime? folderDate
    {
        get
        {
            return lastModifiedDateTime?.Date;
        }
        set
        {
            _ = value;
        }
    }

    [JsonIgnore]
    public string? fromid
    {
        get
        {
            if (from != null)
                if (from.user != null)
                    return from.user.id;
                else if (from.application != null)
                    return from.application.id;

            return null;
        }
        set
        {
            _ = value;
        }
    }

    [JsonIgnore]
    public string? fromdisplayname
    {
        get
        {
            if (from != null)
                if (from.user != null)
                    return from.user.displayName;
                else if (from.application != null)
                    return from.application.displayName;

            return null;
        }
        set
        {
            _ = value;
        }
    }

    [JsonIgnore]
    public string? fromidentitytype
    {
        get
        {
            if (from != null)
                if (from.user != null)
                    return from.user.userIdentityType;
                else if (from.application != null)
                    return from.application.applicationIdentityType;

            return null;
        }
        set
        {
            _ = value;
        }
    }

    [JsonIgnore]
    public string? fromtenantid
    {
        get
        {
            if (from != null)
                if (from.user != null)
                    return from.user.tenantId;

            return null;
        }
        set
        {
            _ = value;
        }
    }


    [JsonIgnore]
    public string? fromtype
    {
        get
        {
            if (from != null)
                if (from.user != null)
                    return "user";
                else if (from.application != null)
                    return "application";
                else if (from.device != null)
                    return "device";

            return null;
        }
        set
        {
            _ = value;
        }
    }
}

internal class GetUserMessages
{
    [JsonPropertyName("@odata.context")]
    public string? odatacontext { get; set; }

    [JsonPropertyName("@odata.count")]
    public int odatacount { get; set; }

    [JsonPropertyName("@odata.nextLink")]
    public string? odatanextLink { get; set; }
    public List<UserMessage>? value { get; set; }
}

