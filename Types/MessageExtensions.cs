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

// stores the types used in the GetMessages method, plus is an output to the Logic App
internal class EventDetail
{
    [JsonPropertyName("@odata.type")]
    public string? odatatype { get; set; }
    public DateTime? visibleHistoryStartDateTime { get; set; }
    public UserX initiator { get; set; } = new UserX();
    public List<Member> members { get; set; } = new List<Member>();
    public List<Participant> callParticipants { get; set; } = new List<Participant>();
}

internal class Member
{
    public string? id { get; set; }
    public string? displayName { get; set; }
    public string? userIdentityType { get; set; }
    public string? tenantId { get; set; }
}

internal class Participant
{
    public UserX? participant { get; set; }
}


internal class From
{
    public Application? application { get; set; }
    public string? device { get; set; }
    public User? user { get; set; }
}

internal class Application
{
    public string? odatatype { get; set; }
    public string? id { get; set; }
    public string? displayName { get; set; }
    public string? applicationIdentityType { get; set; }
}

internal class Body
{
    public string? contentType { get; set; }
    public string? content { get; set; }
}

internal class Attachment
{
    public string? id { get; set; }
    public string? contentType { get; set; }

    // Content URL can be dynamically created or not
    public string? contentUrl
    {
        get
        {
            // If _contentUrl is null or empty, generate it from content
            switch (contentType)
            {
                case "application/vnd.microsoft.card.codesnippet":
                    {
                        if (string.IsNullOrEmpty(content))
                            break;

                        var contentx = JsonDocument.Parse(content).RootElement;

                        if (contentx.ValueKind != JsonValueKind.Object)
                            break;

                        var codeSnippetUrl = contentx.GetProperty("codeSnippetUrl").GetString();

                        _contentUrl = codeSnippetUrl;
                        break;
                    }
                default:
                    break;
            }

            return _contentUrl;
        }
        set
        {
            _contentUrl = value;
        }
    }

    public string? content { get; set; }
    public string? name { get; set; }
    public string? thumbnailUrl { get; set; }
    public string? teamsAppId { get; set; }

    public string GetUniqueKey()
    {
        return $"{id}_{contentType}";
    }

    // privates
    private string? _contentUrl;
}

internal class Mention
{
    public int id { get; set; }
    public string? mentionText { get; set; }
    public Mentioned mentioned { get; set; } = new Mentioned();
}

internal class Mentioned
{
    public User user { get; set; } = new User();
}

internal class Reaction
{
    public string? reactionType { get; set; }
    public DateTime? createdDateTime { get; set; }
    public UserX user { get; set; } = new UserX();

    public string GetUniqueKey()
    {
        return $"{createdDateTime}_{user.user?.id}_{reactionType}";
    }
}

internal class MessageHistory
{
    public DateTime? modifiedDateTime { get; set; }
    public string? actions { get; set; }
    public MessageHistoryReaction reaction { get; set; } = new MessageHistoryReaction();
}

internal class MessageHistoryReaction
{
    public string? reactionType { get; set; }
    public DateTime? createdDateTime { get; set; }
    public UserX user { get; set; } = new UserX();
}

internal class UserX
{
    [JsonPropertyName("@odata.type")]
    public string? odatatype { get; set; }
    public object? application { get; set; }
    public object? device { get; set; }
    public User? user { get; set; }
}

internal class User
{
    [JsonPropertyName("@odata.type")]
    public string? odatatype { get; set; }
    public string? id { get; set; }
    public string? displayName { get; set; }
    public string? userIdentityType { get; set; }
    public string? tenantId { get; set; }
    public string? application { get; set; }
    public string? device { get; set; }
}