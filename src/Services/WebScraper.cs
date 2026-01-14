using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;


namespace SchedulerApp.Services
{
    public class WebScraper
    {
        public class EventData
        {
            public string EventTitle { get; set; } = "Untitled Event";
            public string SourceUrl { get; set; } = "";
            public string TimeZone { get; set; } = "UTC";
            public DateTime ScrapedAt { get; set; }
            public DateTime StartDate { get; set; } = DateTime.Today;
            public DateTime EndDate { get; set; } = DateTime.Today.AddDays(1);
            public List<Participant> Participants { get; set; } = new List<Participant>();
        }

        public class Participant
        {
            public string Name { get; set; } = "";
            public List<TimeSlot> AvailableSlots { get; set; } = new List<TimeSlot>();
        }

        public class TimeSlot
        {
            public DateTime Date { get; set; }
            public DateTime EndDate { get; set; }
            public TimeSpan Time { get; set; }
            public TimeSpan EndTime { get; set; }
            public TimeSpan Duration { get; set; }
            public DateTime ParsedStart { get; set; }
            public DateTime ParsedEnd { get; set; }
            public string? OriginalStartString { get; set; }
            public string? OriginalEndString { get; set; }
        }

        public static async Task<EventData?> ExtractFromLettuceMeet(string url)
        {
            try
            {
                string? eventId = ExtractEventIdFromUrl(url);
                if (string.IsNullOrEmpty(eventId))
                {
                    return null;
                }

                var result = await TryDirectApiRequest(eventId);

                if (result != null)
                {
                    return result;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static string? ExtractEventIdFromUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return null;

            url = url.Trim();

            if (url.Contains("?"))
            {
                url = url.Substring(0, url.IndexOf("?"));
            }

            string[] patterns = {
                "lettucemeet.com/l/",
                "lettucemeet.com/events/",
                "lettucemeet.com/"
            };

            foreach (var pattern in patterns)
            {
                if (url.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int startIndex = url.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) + pattern.Length;
                    if (startIndex < url.Length)
                    {
                        string remaining = url.Substring(startIndex);
                        var parts = remaining.Split(new char[] { '/', '?' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 0)
                        {
                            string candidate = parts[0];
                            if (candidate.Length >= 4 && candidate.Length <= 12 &&
                                !candidate.Contains(".") && !candidate.Contains("-"))
                            {
                                return candidate;
                            }
                        }
                    }
                }
            }

            var urlParts = url.Split(new char[] { '/', '?' }, StringSplitOptions.RemoveEmptyEntries);
            if (urlParts.Length > 0)
            {
                string lastPart = urlParts[urlParts.Length - 1];
                if (lastPart.Length >= 4 && lastPart.Length <= 12 &&
                    !lastPart.Contains(".") && !lastPart.Contains("-"))
                {
                    return lastPart;
                }
            }

            return null;
        }

        private static async Task<EventData?> TryDirectApiRequest(string eventId)
        {
            try
            {
                var queries = new List<object>
                {
                    new
                    {
                        operationName = "GetPoll",
                        query = @"query GetPoll($id: ID!) {
                            event(id: $id) {
                                id
                                title
                                pollDates
                                timeZone
                                pollResponses {
                                    user {
                                        ... on User { name email }
                                        ... on AnonymousUser { name email }
                                    }
                                    availabilities {
                                        start
                                        end
                                    }
                                }
                            }
                        }",
                        variables = new { id = eventId }
                    },
                    new
                    {
                        operationName = "GetEvent",
                        query = @"query GetEvent($id: ID!) {
                            event(id: $id) {
                                id
                                title
                                dates
                                timeZone
                                responses {
                                    participant {
                                        name
                                        email
                                    }
                                    slots {
                                        start
                                        end
                                    }
                                }
                            }
                        }",
                        variables = new { id = eventId }
                    },
                    new
                    {
                        query = @"query {
                            event(id: """ + eventId + @""") {
                                id
                                title
                                timeZone
                                pollResponses {
                                    user { name }
                                    availabilities { start end }
                                }
                            }
                        }"
                    }
                };

                string[] endpoints = {
                    "https://api.lettucemeet.com/graphql",
                    "https://api.lettucemeet.com/v1/graphql"
                };

                foreach (var endpoint in endpoints)
                {
                    for (int i = 0; i < queries.Count; i++)
                    {
                        try
                        {
                            var jsonContent = JsonSerializer.Serialize(queries[i]);

                            using var httpClient = new HttpClient();
                            httpClient.Timeout = TimeSpan.FromSeconds(30);

                            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
                            httpClient.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");
                            httpClient.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.9");
                            httpClient.DefaultRequestHeaders.Add("Origin", "https://lettucemeet.com");
                            httpClient.DefaultRequestHeaders.Add("Referer", $"https://lettucemeet.com/l/{eventId}");

                            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                            var response = await httpClient.PostAsync(endpoint, content);
                            var responseContent = await response.Content.ReadAsStringAsync();

                            if (!string.IsNullOrEmpty(responseContent) && responseContent.Length > 50)
                            {
                                if (responseContent.TrimStart().StartsWith("<!") ||
                                    responseContent.TrimStart().StartsWith("<html") ||
                                    responseContent.TrimStart().StartsWith("<!doctype"))
                                {
                                    continue;
                                }

                                if (response.IsSuccessStatusCode)
                                {
                                    var result = ParseApiResponse(responseContent, eventId);
                                    if (result != null)
                                    {
                                        return result;
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }

                        await Task.Delay(500);
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static EventData? ParseApiResponse(string responseContent, string eventId)
        {
            try
            {
                using var doc = JsonDocument.Parse(responseContent);
                var root = doc.RootElement;

                if (root.TryGetProperty("errors", out var errors))
                {
                    return null;
                }

                JsonElement eventElement = default;

                if (root.TryGetProperty("data", out var data) && data.TryGetProperty("event", out eventElement))
                {
                }
                else if (root.TryGetProperty("event", out eventElement))
                {
                }
                else if (root.TryGetProperty("poll", out var pollElement))
                {
                    eventElement = pollElement;
                }
                else
                {
                    return null;
                }

                if (eventElement.ValueKind == JsonValueKind.Null)
                {
                    return null;
                }

                var result = new EventData
                {
                    EventTitle = eventElement.TryGetProperty("title", out var titleProp) ?
                        titleProp.GetString() ?? "Untitled Event" : "Untitled Event",
                    SourceUrl = $"https://lettucemeet.com/l/{eventId}",
                    ScrapedAt = DateTime.Now
                };

                string eventTimeZoneId = "UTC";
                if (eventElement.TryGetProperty("timeZone", out var tzProp))
                {
                    eventTimeZoneId = tzProp.GetString() ?? "UTC";
                    result.TimeZone = eventTimeZoneId;
                }

                if (eventElement.TryGetProperty("pollDates", out var pollDatesProp))
                {
                    var dates = new List<DateTime>();
                    foreach (var dateElement in pollDatesProp.EnumerateArray())
                    {
                        if (DateTime.TryParse(dateElement.GetString(), out var date))
                        {
                            dates.Add(date);
                        }
                    }

                    if (dates.Any())
                    {
                        result.StartDate = dates.Min();
                        result.EndDate = dates.Max();
                    }
                }

                JsonElement responsesProp = default;

                if (eventElement.TryGetProperty("pollResponses", out responsesProp))
                {
                }
                else if (eventElement.TryGetProperty("responses", out responsesProp))
                {
                }
                else if (eventElement.TryGetProperty("participants", out responsesProp))
                {
                }
                else if (eventElement.TryGetProperty("availabilities", out responsesProp))
                {
                }
                else
                {
                    return result;
                }

                foreach (var responseElement in responsesProp.EnumerateArray())
                {
                    var participant = new Participant { Name = "Unknown" };

                    if (responseElement.TryGetProperty("user", out var userElement))
                    {
                        if (userElement.TryGetProperty("name", out var nameProp))
                        {
                            participant.Name = nameProp.GetString() ?? "Unknown";
                        }
                    }
                    else if (responseElement.TryGetProperty("participant", out var participantElement))
                    {
                        if (participantElement.TryGetProperty("name", out var nameProp))
                        {
                            participant.Name = nameProp.GetString() ?? "Unknown";
                        }
                    }
                    else if (responseElement.TryGetProperty("name", out var nameProp))
                    {
                        participant.Name = nameProp.GetString() ?? "Unknown";
                    }

                    JsonElement availabilitiesProp = default;

                    if (responseElement.TryGetProperty("availabilities", out availabilitiesProp))
                    {
                    }
                    else if (responseElement.TryGetProperty("slots", out availabilitiesProp))
                    {
                    }
                    else if (responseElement.TryGetProperty("availability", out availabilitiesProp))
                    {
                    }
                    else
                    {
                        result.Participants.Add(participant);
                        continue;
                    }

                    foreach (var availabilityElement in availabilitiesProp.EnumerateArray())
                    {
                        string? startStr = null;
                        string? endStr = null;

                        if (availabilityElement.TryGetProperty("start", out var startProp))
                        {
                            startStr = startProp.GetString();
                        }

                        if (availabilityElement.TryGetProperty("end", out var endProp))
                        {
                            endStr = endProp.GetString();
                        }

                        if (!string.IsNullOrEmpty(startStr) && !string.IsNullOrEmpty(endStr))
                        {
                            string startWithoutZ = startStr.TrimEnd('Z');
                            string endWithoutZ = endStr.TrimEnd('Z');

                            if (DateTime.TryParse(startWithoutZ, CultureInfo.InvariantCulture,
                                                  DateTimeStyles.AssumeLocal, out var startDateTime) &&
                                DateTime.TryParse(endWithoutZ, CultureInfo.InvariantCulture,
                                                  DateTimeStyles.AssumeLocal, out var endDateTime))
                            {
                                var timeSlot = new TimeSlot
                                {
                                    Date = startDateTime.Date,
                                    Time = startDateTime.TimeOfDay,
                                    EndTime = endDateTime.TimeOfDay,
                                    Duration = endDateTime - startDateTime,
                                    ParsedStart = startDateTime,
                                    ParsedEnd = endDateTime,
                                    OriginalStartString = startStr,
                                    OriginalEndString = endStr,
                                    EndDate = endDateTime.Date
                                };

                                participant.AvailableSlots.Add(timeSlot);
                            }
                        }
                    }

                    result.Participants.Add(participant);
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        public static List<MainWindow.Employee> ConvertToEmployees(EventData eventData)
        {
            var employees = new List<MainWindow.Employee>();

            foreach (var participant in eventData.Participants)
            {
                if (participant.AvailableSlots.Count == 0)
                    continue;

                var employee = new MainWindow.Employee
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = participant.Name,
                    DateRange = $"{eventData.StartDate:MMM dd} to {eventData.EndDate:MMM dd}",
                    TimeRange = "Extracted from LettuceMeet",
                    AvailabilitySummary = $"{participant.AvailableSlots.Count} time slots available"
                };

                employees.Add(employee);
            }

            return employees;
        }
    }
}