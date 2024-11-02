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

namespace GraphExportAPIforMicrosoftTeamsSample.Utility;

// Class to hold utility functions
internal static class Utils
{
    // Identify Conversation Types from ConversationID
    // 19meetingMDg5Njg3MDQtNTQzYSOOMmNkLTg2YzktZjBmOTlxMmZmZDY4threadv2                        Type:meeting                Query for:19meeting*
    // 1932732179d35d43f7bb505253644ec1bcthreadtacv2                                            Type:ChannelConversation    Query for:*threadtacv2
    // 1978d3890731&U7adb1985ff7cfbf2c06threadv2                                                Type:GroupChat              Query for:*threadv2
    // 192ea267cb-1496-442c-ae13-25dc5c7a9019d16a7eee-b7a54e97-836f-Oflae8e7973funqgblspaces    Type:Chat                   Query for:*unqgblspaces
    internal static ConversationType FigureConversationTypeFromChatId(string threadId)
    {
        ConversationType ct = ConversationType.Undefined;

        string lowerChatId = threadId.ToLower().Trim();

        if (lowerChatId.Contains("19:meeting") && lowerChatId.EndsWith("thread.v2"))
        {
            ct = ConversationType.Meeting;
        }
        else if (lowerChatId.EndsWith("threadtac.v2"))
        {
            ct = ConversationType.Channelconversation;
        }
        else if (lowerChatId.EndsWith("thread.v2"))
        {
            ct = ConversationType.Group;
        }
        else if (lowerChatId.EndsWith("unq.gbl.spaces"))
        {
            ct = ConversationType.OneOnOne;
        }
        else if (lowerChatId == "48:notes")
        {
            ct = ConversationType.Notes;
        }

        return ct;
    }

    // Identify Conversation Types from ConversationID
    internal enum ConversationType
    {
        Undefined,
        Meeting,
        Channelconversation,
        Group,
        OneOnOne,
        Notes
    }

    // UnixTimeStampToDateTime
    internal static DateTime UnixTimeStampMillisecondsToDateTime(long unixTime)
    {
        DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        return unixEpoch.AddMilliseconds(unixTime);
    }

    // DateTimeToUnixTimeStamp
    internal static long DateTimeToUnixTimeStampMilliseconds(DateTime? dateTime)
    {
        if (dateTime == null)
            return 0;

        DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        DateTime dt = (DateTime)dateTime;
        TimeSpan timeSpan = dt.ToUniversalTime() - unixEpoch;
        return (long)timeSpan.TotalMilliseconds;
    }

    // Specify UTC Kind
    internal static DateTime SpecifyUTC(DateTime utcDateTime)
    {
        // Check if the DateTime Kind is Unspecified, assume it's Utc in that case
        if (utcDateTime.Kind == DateTimeKind.Unspecified)
        {
            utcDateTime = DateTime.SpecifyKind(utcDateTime, DateTimeKind.Utc);
        }
        return utcDateTime;
    }

    // Check if a string is a GUID
    internal static bool IsStringAGuid(string input)
    {
        return Guid.TryParse(input, out _);
    }

     // Random Number Generator
    private static readonly Random random = new Random();

    internal static int RandInteger(int minValue, int maxValue)
    {
        return random.Next(minValue, maxValue);
    }
}