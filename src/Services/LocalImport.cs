using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;


namespace SchedulerApp.Services
{
    public class LocalImport
    {
        public class ImportResult
        {
            public bool Success { get; set; }
            public List<PersonData> People { get; set; } = new List<PersonData>();
            public string ErrorMessage { get; set; }
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            public string EventTitle { get; set; } = "Imported Schedule";
        }

        public class PersonData
        {
            public string Name { get; set; }
            public List<AvailabilitySlot> AvailableSlots { get; set; } = new List<AvailabilitySlot>();
        }

        public class AvailabilitySlot
        {
            public DateTime ParsedStart { get; set; }
            public DateTime ParsedEnd { get; set; }
        }

        public static ImportResult ImportFromFile(string filePath)
        {
            var result = new ImportResult();

            try
            {
                if (!File.Exists(filePath))
                {
                    result.Success = false;
                    result.ErrorMessage = "File does not exist.";
                    return result;
                }

                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".csv")
                {
                    return ImportFromCsv(filePath);
                }

                result.Success = false;
                result.ErrorMessage = $"Unsupported file format: {extension}. Only CSV files are supported.";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error importing file: {ex.Message}";
                return result;
            }
        }

        private static ImportResult ImportFromCsv(string filePath)
        {
            var result = new ImportResult();

            try
            {
                var lines = File.ReadAllLines(filePath);

                if (lines.Length == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = "CSV file is empty";
                    return result;
                }

                string firstLine = lines[0].ToLower();

                if (firstLine.Contains("name") && firstLine.Contains("date") &&
                    (firstLine.Contains("start") || firstLine.Contains("time")))
                {
                    return ParseCsvFormat1(lines, result);
                }
                else if (firstLine.Contains("participant") && firstLine.Contains("availability"))
                {
                    return ParseCsvFormat2(lines, result);
                }
                else if (lines[0].Split(',').Length >= 4 && lines[0].Split(',')[1].Contains("/"))
                {
                    return ParseCsvFormat1(lines, result);
                }
                else
                {
                    return ParseCsvFormat2(lines, result);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error parsing CSV: {ex.Message}";
                return result;
            }
        }

        private static ImportResult ParseCsvFormat1(string[] lines, ImportResult result)
        {
            var peopleDict = new Dictionary<string, PersonData>();
            var allDates = new HashSet<DateTime>();
            var duplicateSlots = new Dictionary<string, HashSet<string>>();
            var errors = new List<string>();

            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                var parts = line.Split(',');
                if (parts.Length < 4)
                {
                    errors.Add($"Line {i + 1}: Insufficient columns (expected 4, got {parts.Length})");
                    continue;
                }

                string name = parts[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    errors.Add($"Line {i + 1}: Missing name");
                    continue;
                }

                if (!DateTime.TryParse(parts[1].Trim(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    errors.Add($"Line {i + 1}: Invalid date format: '{parts[1]}'");
                    continue;
                }

                if (!TryParseTime(parts[2].Trim(), out TimeSpan startTime))
                {
                    errors.Add($"Line {i + 1}: Invalid start time format: '{parts[2]}'");
                    continue;
                }

                if (!TryParseTime(parts[3].Trim(), out TimeSpan endTime))
                {
                    errors.Add($"Line {i + 1}: Invalid end time format: '{parts[3]}'");
                    continue;
                }

                if (startTime >= endTime && endTime != TimeSpan.Zero)
                {
                    errors.Add($"Line {i + 1}: End time must be after start time: {startTime} - {endTime}");
                    continue;
                }

                DateTime startDateTime = date.Date.Add(startTime);
                DateTime endDateTime = date.Date.Add(endTime);

                if (endDateTime <= startDateTime)
                {
                    endDateTime = endDateTime.AddDays(1);
                }

                if (!peopleDict.ContainsKey(name))
                {
                    peopleDict[name] = new PersonData { Name = name };
                    duplicateSlots[name] = new HashSet<string>();
                }

                string slotKey = $"{startDateTime:yyyy-MM-dd HH:mm}-{endDateTime:yyyy-MM-dd HH:mm}";
                if (!duplicateSlots[name].Add(slotKey))
                {
                    errors.Add($"Line {i + 1}: Duplicate time slot for {name} on {date:yyyy-MM-dd}");
                    continue;
                }

                var newSlot = new AvailabilitySlot
                {
                    ParsedStart = startDateTime,
                    ParsedEnd = endDateTime
                };

                bool hasOverlap = false;
                foreach (var existingSlot in peopleDict[name].AvailableSlots)
                {
                    if (IsOverlapping(newSlot, existingSlot))
                    {
                        hasOverlap = true;
                        errors.Add($"Line {i + 1}: Overlapping time slot for {name} on {date:yyyy-MM-dd}");
                        break;
                    }
                }

                if (!hasOverlap)
                {
                    peopleDict[name].AvailableSlots.Add(newSlot);
                    allDates.Add(date.Date);
                }
            }

            if (peopleDict.Count == 0)
            {
                result.Success = false;
                result.ErrorMessage = "No valid data found in CSV";
                if (errors.Count > 0)
                {
                    result.ErrorMessage += $"\nErrors:\n{string.Join("\n", errors.Take(5))}";
                    if (errors.Count > 5) result.ErrorMessage += $"\n... and {errors.Count - 5} more errors";
                }
                return result;
            }

            result.People = peopleDict.Values.ToList();
            result.StartDate = allDates.Min();
            result.EndDate = allDates.Max();
            result.Success = true;

            if (errors.Count > 0)
            {
                result.ErrorMessage = $"Import completed with {errors.Count} warning(s). First 3 warnings:\n{string.Join("\n", errors.Take(3))}";
            }

            return result;
        }

        private static ImportResult ParseCsvFormat2(string[] lines, ImportResult result)
        {
            var people = new List<PersonData>();
            var allDates = new HashSet<DateTime>();
            var errors = new List<string>();

            string[] headers = lines[0].Split(',');
            var dateColumns = new List<(int index, DateTime date)>();

            for (int i = 1; i < headers.Length; i++)
            {
                if (DateTime.TryParse(headers[i].Trim(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    dateColumns.Add((i, date));
                    allDates.Add(date);
                }
                else
                {
                    errors.Add($"Header column {i + 1}: Invalid date format '{headers[i]}'");
                }
            }

            if (dateColumns.Count == 0)
            {
                result.Success = false;
                result.ErrorMessage = "No valid dates found in CSV header";
                return result;
            }

            for (int row = 1; row < lines.Length; row++)
            {
                string line = lines[row].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                {
                    errors.Add($"Row {row + 1}: Insufficient columns");
                    continue;
                }

                string name = parts[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    errors.Add($"Row {row + 1}: Missing participant name");
                    continue;
                }

                var person = new PersonData { Name = name };
                var duplicateSlots = new HashSet<string>();

                foreach (var (colIndex, date) in dateColumns)
                {
                    if (colIndex >= parts.Length) continue;

                    string timeRange = parts[colIndex].Trim();
                    if (string.IsNullOrEmpty(timeRange) ||
                        timeRange.Equals("na", StringComparison.OrdinalIgnoreCase) ||
                        timeRange.Equals("unavailable", StringComparison.OrdinalIgnoreCase) ||
                        timeRange.Equals("-", StringComparison.OrdinalIgnoreCase) ||
                        timeRange.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var timeParts = timeRange.Split('-');
                    if (timeParts.Length != 2)
                    {
                        errors.Add($"Row {row + 1}, Column {colIndex + 1}: Invalid time range format '{timeRange}'");
                        continue;
                    }

                    string startTimeStr = timeParts[0].Trim();
                    string endTimeStr = timeParts[1].Trim();

                    if (!TryParseTime(startTimeStr, out TimeSpan startSpan))
                    {
                        errors.Add($"Row {row + 1}, Column {colIndex + 1}: Invalid start time '{startTimeStr}'");
                        continue;
                    }

                    if (!TryParseTime(endTimeStr, out TimeSpan endSpan))
                    {
                        errors.Add($"Row {row + 1}, Column {colIndex + 1}: Invalid end time '{endTimeStr}'");
                        continue;
                    }

                    if (startSpan >= endSpan && endSpan != TimeSpan.Zero)
                    {
                        errors.Add($"Row {row + 1}, Column {colIndex + 1}: End time must be after start time: {startSpan} - {endSpan}");
                        continue;
                    }

                    DateTime startDateTime = date.Date.Add(startSpan);
                    DateTime endDateTime = date.Date.Add(endSpan);

                    if (endDateTime <= startDateTime)
                    {
                        endDateTime = endDateTime.AddDays(1);
                    }

                    string slotKey = $"{startDateTime:yyyy-MM-dd HH:mm}-{endDateTime:yyyy-MM-dd HH:mm}";
                    if (!duplicateSlots.Add(slotKey))
                    {
                        errors.Add($"Row {row + 1}, Column {colIndex + 1}: Duplicate time slot for {name} on {date:yyyy-MM-dd}");
                        continue;
                    }

                    var newSlot = new AvailabilitySlot
                    {
                        ParsedStart = startDateTime,
                        ParsedEnd = endDateTime
                    };

                    bool hasOverlap = false;
                    foreach (var existingSlot in person.AvailableSlots)
                    {
                        if (IsOverlapping(newSlot, existingSlot))
                        {
                            hasOverlap = true;
                            errors.Add($"Row {row + 1}, Column {colIndex + 1}: Overlapping time slot for {name} on {date:yyyy-MM-dd}");
                            break;
                        }
                    }

                    if (!hasOverlap)
                    {
                        person.AvailableSlots.Add(newSlot);
                    }
                }

                if (person.AvailableSlots.Count > 0)
                {
                    people.Add(person);
                }
            }

            if (people.Count == 0)
            {
                result.Success = false;
                result.ErrorMessage = "No valid participant data found";
                if (errors.Count > 0)
                {
                    result.ErrorMessage += $"\nErrors:\n{string.Join("\n", errors.Take(5))}";
                    if (errors.Count > 5) result.ErrorMessage += $"\n... and {errors.Count - 5} more errors";
                }
                return result;
            }

            result.People = people;
            result.StartDate = allDates.Min();
            result.EndDate = allDates.Max();
            result.Success = true;

            if (errors.Count > 0)
            {
                result.ErrorMessage = $"Import completed with {errors.Count} warning(s). First 3 warnings:\n{string.Join("\n", errors.Take(3))}";
            }

            return result;
        }

        private static bool TryParseTime(string timeString, out TimeSpan timeSpan)
        {
            timeSpan = TimeSpan.Zero;

            if (string.IsNullOrEmpty(timeString))
                return false;

            timeString = timeString.Trim().ToUpper();

            if (TimeSpan.TryParse(timeString, out timeSpan))
                return true;

            if (DateTime.TryParse(timeString, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime time))
            {
                timeSpan = time.TimeOfDay;
                return true;
            }

            if (timeString.EndsWith("AM") || timeString.EndsWith("PM"))
            {
                string formattedTime = timeString.Replace("AM", "").Replace("PM", "").Trim();
                if (TimeSpan.TryParse(formattedTime, out timeSpan))
                {
                    if (timeString.EndsWith("PM") && timeSpan.Hours < 12)
                        timeSpan = timeSpan.Add(TimeSpan.FromHours(12));
                    else if (timeString.EndsWith("AM") && timeSpan.Hours == 12)
                        timeSpan = timeSpan.Subtract(TimeSpan.FromHours(12));
                    return true;
                }
            }

            if (timeString.Contains(":"))
            {
                var parts = timeString.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[0], out int hours) && int.TryParse(parts[1], out int minutes))
                {
                    if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59)
                    {
                        timeSpan = new TimeSpan(hours, minutes, 0);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsOverlapping(AvailabilitySlot slot1, AvailabilitySlot slot2)
        {
            return slot1.ParsedStart < slot2.ParsedEnd && slot2.ParsedStart < slot1.ParsedEnd;
        }

        public static string GetSampleCsvTemplate()
        {
            return @"Name,Date,Start_Time,End_Time,Notes
John Doe,2024-01-15,09:00,17:00,Available all day
Jane Smith,2024-01-15,13:00,21:00,Afternoon shift
Bob Johnson,2024-01-16,08:00,16:00,Morning shift
Alice Brown,2024-01-16,10:00,18:00,Day shift
John Doe,2024-01-16,09:00,17:00,
Jane Smith,2024-01-17,08:00,16:00,
Bob Johnson,2024-01-17,12:00,20:00,Afternoon only
Alice Brown,2024-01-17,14:00,22:00,Late shift";
        }

        public static string GetGridCsvTemplate()
        {
            return @"Participant,2024-01-15,2024-01-16,2024-01-17,2024-01-18
John Doe,09:00-17:00,09:00-17:00,09:00-17:00,09:00-17:00
Jane Smith,13:00-21:00,13:00-21:00,08:00-16:00,NA
Bob Johnson,08:00-16:00,12:00-20:00,12:00-20:00,08:00-16:00
Alice Brown,10:00-18:00,10:00-18:00,14:00-22:00,10:00-18:00";
        }
    }
}