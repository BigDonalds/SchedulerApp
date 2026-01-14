using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SchedulerApp.Services
{
    public class PersonAvailability
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public double TotalAvailableHours { get; set; }
    }

    public class ScheduleConfig
    {
        public TimeSpan OpeningTime { get; set; }
        public TimeSpan ClosingTime { get; set; }
        public TimeSpan ShiftLength { get; set; }
        public int PeoplePerShift { get; set; }
        public HashSet<DayOfWeek> ClosedDays { get; set; } = new HashSet<DayOfWeek>();
        public bool AllowShiftStacking { get; set; } = true;
    }

    public class ScheduleResult
    {
        public List<Shift> Shifts { get; set; } = new List<Shift>();
        public bool HasUnfilledShifts =>
            Shifts.Any(s => s.AssignedPeople.Count < s.PeopleNeeded);
    }

    public class Shift
    {
        public Guid Id { get; } = Guid.NewGuid();
        public DateTime Date { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public double DurationHours => (End - Start).TotalHours;
        public int PeopleNeeded { get; set; } = 1;
        public List<string> AssignedPeople { get; set; } = new List<string>();
        public int PositionInDay { get; set; } // 0 = first shift, 1 = second shift, etc.
        public bool IsLastShiftOfDay { get; set; }
        public bool IsFirstShiftOfDay { get; set; }
    }

    public class HitMap
    {
        public DateTime Date { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public int CandidateCount { get; set; }
        public List<string> Candidates { get; set; } = new List<string>();
        public Shift Shift { get; set; }
    }

    public class Scheduler
    {
        private Dictionary<string, double> _assignedHours = new Dictionary<string, double>();
        private Dictionary<string, double> _availableHours = new Dictionary<string, double>();
        private Dictionary<string, Dictionary<DateTime, List<Shift>>> _personShifts = new Dictionary<string, Dictionary<DateTime, List<Shift>>>();
        private List<PersonAvailability> _allAvailabilities;
        private Dictionary<string, double> _weeklyHours = new Dictionary<string, double>();
        private int _totalUnderstaffedShifts = 0;
        private Dictionary<DateTime, List<Shift>> _shiftsByDate = new Dictionary<DateTime, List<Shift>>();

        public ScheduleResult GenerateSchedule(
            List<PersonAvailability> availabilities,
            ScheduleConfig config)
        {
            WriteDebugLog($"=== SCHEDULER START ===");
            WriteDebugLog($"Config - Opening: {config.OpeningTime}, Closing: {config.ClosingTime}");
            WriteDebugLog($"Shift Length: {config.ShiftLength}, People per shift: {config.PeoplePerShift}");
            WriteDebugLog($"Closed Days: {(config.ClosedDays.Count > 0 ? string.Join(", ", config.ClosedDays) : "None")}");
            WriteDebugLog($"Total availabilities: {availabilities.Count}");

            _allAvailabilities = availabilities;
            LogAllAvailabilities();
            CalculateTotalAvailableHours();
            InitializeTracking();

            var shifts = GenerateShiftsByDate(availabilities, config);
            WriteDebugLog($"\nGenerated {shifts.Count} shifts");

            // NEW: Mark shift positions and first/last shifts
            MarkShiftPositions(shifts);

            // Group shifts by date for easy access
            _shiftsByDate = shifts.GroupBy(s => s.Date.Date).ToDictionary(g => g.Key, g => g.OrderBy(s => s.Start).ToList());

            var peopleByAvailability = _availableHours
                .OrderBy(k => k.Value)
                .Select(k => k.Key)
                .ToList();

            WriteDebugLog($"\n=== PEOPLE ORDERED BY AVAILABILITY (lowest first) ===");
            foreach (var person in peopleByAvailability)
            {
                WriteDebugLog($"  {person}: {_availableHours[person]:F2} hours available");
            }

            WriteDebugLog($"\n=== HIT MAP ANALYSIS ===");
            var hitMaps = BuildHitMaps(shifts, config);
            AnalyzeHitMaps(hitMaps);

            WriteDebugLog($"\n=== PHASE 1: ASSIGNING CRITICAL SHIFTS ===");
            AssignCriticalShifts(shifts, hitMaps, peopleByAvailability, config);

            WriteDebugLog($"\n=== PHASE 2: ENFORCE BACK-TO-BACK SHIFTS ===");
            EnforceBackToBackShifts(shifts, config);

            WriteDebugLog($"\n=== PHASE 3: FILL REMAINING SHIFTS ===");
            FillRemainingShifts(shifts, peopleByAvailability, config);

            WriteDebugLog($"\n=== PHASE 4: DYNAMIC HOUR BALANCING ===");
            DynamicHourBalancing(shifts, config);

            WriteDebugLog($"\n=== PHASE 5: FINAL BACK-TO-BACK CLEANUP ===");
            FinalBackToBackCleanup(shifts, config);

            WriteDebugLog($"\n=== FINAL SCHEDULE ===");
            FinalizeSchedule(shifts, config);

            return new ScheduleResult { Shifts = shifts };
        }

        private void MarkShiftPositions(List<Shift> shifts)
        {
            // Group by date
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();

                for (int i = 0; i < dayShifts.Count; i++)
                {
                    dayShifts[i].PositionInDay = i;
                    dayShifts[i].IsFirstShiftOfDay = (i == 0);
                    dayShifts[i].IsLastShiftOfDay = (i == dayShifts.Count - 1);
                }
            }
        }

        private List<HitMap> BuildHitMaps(List<Shift> shifts, ScheduleConfig config)
        {
            var hitMaps = new List<HitMap>();

            foreach (var shift in shifts)
            {
                var candidates = new List<string>();

                foreach (var person in _availableHours.Keys)
                {
                    if (IsAvailable(person, shift) && CanAssign(person, shift, config))
                    {
                        candidates.Add(person);
                    }
                }

                hitMaps.Add(new HitMap
                {
                    Date = shift.Date,
                    Start = shift.Start,
                    End = shift.End,
                    CandidateCount = candidates.Count,
                    Candidates = candidates,
                    Shift = shift
                });
            }

            return hitMaps;
        }

        private void AnalyzeHitMaps(List<HitMap> hitMaps)
        {
            var groupedByCandidateCount = hitMaps
                .GroupBy(h => h.CandidateCount)
                .OrderBy(g => g.Key)
                .ToList();

            WriteDebugLog($"\nHit Map Analysis by Candidate Count:");
            foreach (var group in groupedByCandidateCount)
            {
                WriteDebugLog($"  {group.Key} candidates: {group.Count()} shifts");
                if (group.Key <= 5 && group.Key > 0)
                {
                    foreach (var hitMap in group.Take(5))
                    {
                        WriteDebugLog($"    {hitMap.Date:yyyy-MM-dd} {hitMap.Start.Hours}:00-{hitMap.End.Hours}:00");
                        if (group.Key <= 3)
                        {
                            WriteDebugLog($"      Candidates: {string.Join(", ", hitMap.Candidates)}");
                        }
                    }
                    if (group.Count() > 5)
                        WriteDebugLog($"    ... and {group.Count() - 5} more");
                }
            }

            // Special case for shifts with 0 candidates
            var zeroCandidateShifts = hitMaps.Where(h => h.CandidateCount == 0).ToList();
            if (zeroCandidateShifts.Any())
            {
                WriteDebugLog($"\n  WARNING: {zeroCandidateShifts.Count} shifts have NO candidates!");
                foreach (var hitMap in zeroCandidateShifts.Take(10))
                {
                    WriteDebugLog($"    {hitMap.Date:yyyy-MM-dd} {hitMap.Start.Hours}:00-{hitMap.End.Hours}:00");
                }
                if (zeroCandidateShifts.Count > 10)
                    WriteDebugLog($"    ... and {zeroCandidateShifts.Count - 10} more");
            }
        }

        private void AssignCriticalShifts(
            List<Shift> shifts,
            List<HitMap> hitMaps,
            List<string> orderedPeople,
            ScheduleConfig config)
        {
            // Get shifts with candidate count <= people needed
            var criticalShifts = hitMaps
                .Where(h => h.CandidateCount > 0 && h.CandidateCount <= h.Shift.PeopleNeeded)
                .OrderBy(h => h.CandidateCount)
                .ThenBy(h => h.Date)
                .ThenBy(h => h.Start)
                .ToList();

            WriteDebugLog($"\nFound {criticalShifts.Count} critical shifts (candidates <= needed)");

            foreach (var hitMap in criticalShifts)
            {
                var shift = hitMap.Shift;

                if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                    continue;

                WriteDebugLog($"\n  Processing critical shift: {shift.Date:yyyy-MM-dd} {shift.Start.Hours}:00-{shift.End.Hours}:00");
                WriteDebugLog($"    Has {hitMap.CandidateCount} candidate(s), needs {shift.PeopleNeeded}");

                // Sort candidates by availability (lowest first)
                var sortedCandidates = hitMap.Candidates
                    .OrderBy(c => _availableHours[c])
                    .ThenBy(c => _assignedHours.ContainsKey(c) ? _assignedHours[c] : 0)
                    .ToList();

                int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;
                int toAssign = Math.Min(needed, sortedCandidates.Count);

                for (int i = 0; i < toAssign; i++)
                {
                    var candidate = sortedCandidates[i];
                    if (CanAssign(candidate, shift, config))
                    {
                        Assign(candidate, shift);
                        WriteDebugLog($"    Assigned {candidate} (availability: {_availableHours[candidate]:F2}h)");
                    }
                }

                if (shift.AssignedPeople.Count < shift.PeopleNeeded)
                {
                    WriteDebugLog($"    WARNING: Still understaffed after critical assignment");
                }
            }

            // Now process shifts with candidate count == people needed + 1 (next level)
            var nextLevelShifts = hitMaps
                .Where(h => h.CandidateCount == h.Shift.PeopleNeeded + 1)
                .OrderBy(h => h.Date)
                .ThenBy(h => h.Start)
                .ToList();

            if (nextLevelShifts.Any())
            {
                WriteDebugLog($"\nProcessing {nextLevelShifts.Count} shifts with candidate count = needed + 1");

                foreach (var hitMap in nextLevelShifts)
                {
                    var shift = hitMap.Shift;

                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                        continue;

                    // For these shifts, assign the person with lowest availability
                    var sortedCandidates = hitMap.Candidates
                        .OrderBy(c => _availableHours[c])
                        .ThenBy(c => _assignedHours.ContainsKey(c) ? _assignedHours[c] : 0)
                        .ToList();

                    int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;
                    for (int i = 0; i < needed && i < sortedCandidates.Count; i++)
                    {
                        var candidate = sortedCandidates[i];
                        if (CanAssign(candidate, shift, config))
                        {
                            Assign(candidate, shift);
                            WriteDebugLog($"  Assigned {candidate} to {shift.Date:yyyy-MM-dd} {shift.Start.Hours}:00-{shift.End.Hours}:00");
                        }
                    }
                }
            }
        }

        private void EnforceBackToBackShifts(List<Shift> shifts, ScheduleConfig config)
        {
            WriteDebugLog($"\n=== ENFORCING BACK-TO-BACK SHIFTS ===");

            // Group shifts by date
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                WriteDebugLog($"\nProcessing {dayGroup.Key:yyyy-MM-dd}:");

                for (int i = 0; i < dayShifts.Count - 1; i++)
                {
                    var currentShift = dayShifts[i];
                    var nextShift = dayShifts[i + 1];

                    // For each person in current shift, try to assign them to next shift
                    foreach (var person in currentShift.AssignedPeople.ToList())
                    {
                        // Skip if next shift already has enough people
                        if (nextShift.AssignedPeople.Count >= nextShift.PeopleNeeded)
                            continue;

                        // Check if person is available for next shift
                        if (IsAvailable(person, nextShift) && CanAssign(person, nextShift, config))
                        {
                            // Remove from current shift if this creates a split shift
                            if (i > 0 && !IsPersonInShift(person, dayShifts[i - 1]))
                            {
                                // This would create a split shift, so don't assign
                                continue;
                            }

                            Assign(person, nextShift);
                            WriteDebugLog($"  {person} assigned to consecutive shift {nextShift.Start.Hours}:00-{nextShift.End.Hours}:00");
                        }
                    }
                }
            }
        }

        private bool IsPersonInShift(string person, Shift shift)
        {
            return shift.AssignedPeople.Contains(person);
        }

        private void FillRemainingShifts(
            List<Shift> shifts,
            List<string> orderedPeople,
            ScheduleConfig config)
        {
            var remaining = shifts
                .Where(s => s.AssignedPeople.Count < s.PeopleNeeded)
                .OrderBy(s => GetCandidateCount(s, config))
                .ThenBy(s => s.IsLastShiftOfDay ? 0 : 1) // Prioritize last shifts
                .ThenBy(s => s.Start)
                .ToList();

            WriteDebugLog($"\n{remaining.Count} shifts still need staffing");

            foreach (var shift in remaining)
            {
                int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;
                if (needed <= 0) continue;

                WriteDebugLog($"\n  Shift {shift.Date:yyyy-MM-dd} {shift.Start.Hours}:00-{shift.End.Hours}:00 (Position: {shift.PositionInDay})");
                WriteDebugLog($"    Needs {needed} more people");

                // Special handling for last shift of day
                if (shift.IsLastShiftOfDay)
                {
                    WriteDebugLog($"    LAST SHIFT OF DAY - looking for shift stacking...");

                    // Find the shift that ends when this one starts
                    var previousShift = shifts.FirstOrDefault(s =>
                        s.Date.Date == shift.Date.Date &&
                        s.End == shift.Start &&
                        s.AssignedPeople.Any());

                    if (previousShift != null)
                    {
                        WriteDebugLog($"    Found previous shift ending at {previousShift.End.Hours}:00 with {previousShift.AssignedPeople.Count} people");

                        // Prioritize people from previous shift
                        var stackCandidates = previousShift.AssignedPeople
                            .Where(p => IsAvailable(p, shift) && CanAssign(p, shift, config) && !shift.AssignedPeople.Contains(p))
                            .ToList();

                        WriteDebugLog($"    {stackCandidates.Count} candidates from previous shift");

                        foreach (var candidate in stackCandidates)
                        {
                            if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                                break;

                            Assign(candidate, shift);
                            WriteDebugLog($"    Stacked {candidate} from previous shift");
                            needed--;
                        }
                    }
                }

                if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                    continue;

                needed = shift.PeopleNeeded - shift.AssignedPeople.Count;

                // Find candidates who can work this shift without creating split shifts
                var candidates = orderedPeople
                    .Where(p => !shift.AssignedPeople.Contains(p))
                    .Where(p => IsAvailable(p, shift))
                    .Where(p => CanAssignWithoutSplit(p, shift, config))
                    .OrderBy(p => CalculateFairnessScore(p, shift))
                    .Take(needed * 2)
                    .ToList();

                // If not enough candidates without split shifts, relax the requirement
                if (candidates.Count < needed)
                {
                    candidates.AddRange(orderedPeople
                        .Where(p => !shift.AssignedPeople.Contains(p) && !candidates.Contains(p))
                        .Where(p => IsAvailable(p, shift))
                        .Where(p => CanAssign(p, shift, config))
                        .OrderBy(p => CalculateFairnessScore(p, shift))
                        .Take(needed - candidates.Count));
                }

                foreach (var candidate in candidates.Take(needed))
                {
                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                        break;

                    Assign(candidate, shift);
                    WriteDebugLog($"    Assigned {candidate} (fairness score: {CalculateFairnessScore(candidate, shift):F2})");
                }

                if (shift.AssignedPeople.Count < shift.PeopleNeeded)
                {
                    WriteDebugLog($"    WARNING: Still understaffed ({shift.AssignedPeople.Count}/{shift.PeopleNeeded})");
                    _totalUnderstaffedShifts++;
                }
            }
        }

        private double CalculateFairnessScore(string person, Shift shift)
        {
            double assignedHours = _assignedHours.ContainsKey(person) ? _assignedHours[person] : 0;
            double weeklyHours = GetWeeklyHours(person);
            double availabilityHours = _availableHours[person];

            // Lower score = better candidate
            double score = 0;

            // Prefer people with fewer assigned hours
            score += assignedHours * 10;

            // Prefer people whose availability is being underutilized
            double utilization = availabilityHours > 0 ? assignedHours / availabilityHours : 1;
            score += utilization * 5;

            // Slight penalty for weekly hours to spread work across week
            score += weeklyHours * 2;

            return score;
        }

        private void DynamicHourBalancing(List<Shift> shifts, ScheduleConfig config)
        {
            WriteDebugLog($"\n=== DYNAMIC HOUR BALANCING ===");

            // Calculate current weekly hours for everyone
            var weeklyHours = CalculateAllWeeklyHours();

            if (weeklyHours.Count == 0)
            {
                WriteDebugLog($"No weekly hours to balance");
                return;
            }

            // Find the minimum weekly hours that's realistically achievable
            double minAchievableHours = double.MaxValue;
            double maxAchievableHours = 0;

            foreach (var person in _availableHours.Keys)
            {
                double weekly = weeklyHours.ContainsKey(person) ? weeklyHours[person] : 0;
                double available = _availableHours[person];

                if (weekly > 0 && weekly < minAchievableHours)
                    minAchievableHours = weekly;

                if (weekly > maxAchievableHours)
                    maxAchievableHours = weekly;

                WriteDebugLog($"  {person}: {weekly:F2}h weekly, {available:F2}h available");
            }

            WriteDebugLog($"\nCurrent range: {minAchievableHours:F2}h to {maxAchievableHours:F2}h");

            // Try to bring everyone up to the current maximum
            double targetHours = maxAchievableHours;
            bool improved = true;
            int iterations = 0;

            while (improved && iterations < 10)
            {
                iterations++;
                improved = false;

                // Find people below target
                var belowTarget = weeklyHours
                    .Where(kvp => kvp.Value < targetHours)
                    .OrderBy(kvp => kvp.Value)
                    .Select(kvp => kvp.Key)
                    .ToList();

                // Find people above target (or at target with extra capacity)
                var aboveTarget = weeklyHours
                    .Where(kvp => kvp.Value >= targetHours * 0.9)
                    .OrderByDescending(kvp => kvp.Value)
                    .Select(kvp => kvp.Key)
                    .ToList();

                WriteDebugLog($"\nIteration {iterations}: Target = {targetHours:F2}h");
                WriteDebugLog($"Below target: {belowTarget.Count} people");
                WriteDebugLog($"Above/at target: {aboveTarget.Count} people");

                foreach (var underPerson in belowTarget)
                {
                    foreach (var overPerson in aboveTarget)
                    {
                        if (underPerson == overPerson)
                            continue;

                        if (TryTransferHours(underPerson, overPerson, shifts, config))
                        {
                            improved = true;
                            WriteDebugLog($"  Transferred hours from {overPerson} to {underPerson}");
                            break;
                        }
                    }

                    // Update weekly hours after potential transfer
                    weeklyHours = CalculateAllWeeklyHours();
                }

                // Update target based on new distribution
                if (weeklyHours.Values.Any())
                {
                    double newMax = weeklyHours.Values.Max();
                    if (newMax > targetHours)
                    {
                        targetHours = newMax;
                        WriteDebugLog($"  New target hours: {targetHours:F2}");
                    }
                }
            }

            WriteDebugLog($"\nFinal weekly hour distribution:");
            var finalWeekly = CalculateAllWeeklyHours();
            foreach (var kvp in finalWeekly.OrderBy(k => k.Value))
            {
                WriteDebugLog($"  {kvp.Key}: {kvp.Value:F2}h");
            }
        }

        private Dictionary<string, double> CalculateAllWeeklyHours()
        {
            var weeklyHours = new Dictionary<string, double>();

            foreach (var person in _availableHours.Keys)
            {
                weeklyHours[person] = GetWeeklyHours(person);
            }

            return weeklyHours;
        }

        private bool TryTransferHours(string underPerson, string overPerson, List<Shift> shifts, ScheduleConfig config)
        {
            // Find shifts where overPerson is working that underPerson could take
            var transferableShifts = shifts
                .Where(s => s.AssignedPeople.Contains(overPerson))
                .Where(s => !s.AssignedPeople.Contains(underPerson))
                .Where(s => IsAvailable(underPerson, s))
                .Where(s => CanAssignWithoutSplit(underPerson, s, config))
                .OrderByDescending(s => s.DurationHours)
                .ToList();

            foreach (var shift in transferableShifts)
            {
                // Don't transfer if it would leave shift understaffed
                if (shift.AssignedPeople.Count <= 1)
                    continue;

                // Don't transfer if it would create a split shift for underPerson
                if (WouldCreateSplitShift(underPerson, shift))
                    continue;

                // Perform the transfer
                shift.AssignedPeople.Remove(overPerson);
                shift.AssignedPeople.Add(underPerson);

                // Update tracking
                _assignedHours[overPerson] -= shift.DurationHours;
                _assignedHours[underPerson] += shift.DurationHours;

                var date = shift.Date.Date;

                // Update person shifts
                if (_personShifts.ContainsKey(overPerson) && _personShifts[overPerson].ContainsKey(date))
                {
                    _personShifts[overPerson][date].Remove(shift);
                }

                if (!_personShifts.ContainsKey(underPerson))
                    _personShifts[underPerson] = new Dictionary<DateTime, List<Shift>>();

                if (!_personShifts[underPerson].ContainsKey(date))
                    _personShifts[underPerson][date] = new List<Shift>();

                _personShifts[underPerson][date].Add(shift);

                return true;
            }

            return false;
        }

        private bool WouldCreateSplitShift(string person, Shift shift)
        {
            if (!_personShifts.ContainsKey(person) || !_personShifts[person].ContainsKey(shift.Date.Date))
                return false;

            var existingShifts = _personShifts[person][shift.Date.Date];

            // If person has no shifts on this day, it's fine
            if (existingShifts.Count == 0)
                return false;

            // Check if this shift would be consecutive with existing shifts
            bool isConsecutive = existingShifts.Any(s =>
                s.End == shift.Start || s.Start == shift.End);

            // If not consecutive and person already has a shift, it's a split shift
            return !isConsecutive;
        }

        private void FinalBackToBackCleanup(List<Shift> shifts, ScheduleConfig config)
        {
            WriteDebugLog($"\n=== FINAL BACK-TO-BACK CLEANUP ===");

            // Find and fix any remaining split shifts
            foreach (var person in _personShifts.Keys)
            {
                var personDays = _personShifts[person];

                foreach (var dateShifts in personDays)
                {
                    var date = dateShifts.Key;
                    var shiftsOnDay = dateShifts.Value.OrderBy(s => s.Start).ToList();

                    if (shiftsOnDay.Count > 1)
                    {
                        // Check for gaps between shifts
                        for (int i = 0; i < shiftsOnDay.Count - 1; i++)
                        {
                            var current = shiftsOnDay[i];
                            var next = shiftsOnDay[i + 1];

                            if (current.End != next.Start)
                            {
                                WriteDebugLog($"  {person} has split shift on {date:yyyy-MM-dd}: {current.Start.Hours}:00-{current.End.Hours}:00 and {next.Start.Hours}:00-{next.End.Hours}:00");

                                // Try to remove the earlier shift if possible
                                if (CanRemoveShift(person, current, shifts, config))
                                {
                                    RemoveShift(person, current);
                                    WriteDebugLog($"    Removed earlier shift to eliminate split");
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool CanRemoveShift(string person, Shift shift, List<Shift> allShifts, ScheduleConfig config)
        {
            // Don't remove if it would leave shift understaffed
            if (shift.AssignedPeople.Count <= shift.PeopleNeeded)
                return false;

            // Check if there are other candidates for this shift
            var otherCandidates = _availableHours.Keys
                .Where(p => p != person)
                .Where(p => IsAvailable(p, shift))
                .Where(p => CanAssignWithoutSplit(p, shift, config))
                .ToList();

            return otherCandidates.Count > 0;
        }

        private void RemoveShift(string person, Shift shift)
        {
            shift.AssignedPeople.Remove(person);
            _assignedHours[person] -= shift.DurationHours;

            var date = shift.Date.Date;
            if (_personShifts.ContainsKey(person) && _personShifts[person].ContainsKey(date))
            {
                _personShifts[person][date].Remove(shift);
            }
        }

        private bool CanAssign(string person, Shift shift, ScheduleConfig config)
        {
            if (!IsAvailable(person, shift))
            {
                return false;
            }

            // Check for overlapping shifts
            if (HasOverlap(person, shift))
            {
                return false;
            }

            return true;
        }

        private bool CanAssignWithoutSplit(string person, Shift shift, ScheduleConfig config)
        {
            if (!CanAssign(person, shift, config))
                return false;

            // Additional check: wouldn't create a split shift
            return !WouldCreateSplitShift(person, shift);
        }

        private bool HasOverlap(string person, Shift shift)
        {
            if (!_personShifts.ContainsKey(person) || !_personShifts[person].ContainsKey(shift.Date.Date))
                return false;

            var existingShifts = _personShifts[person][shift.Date.Date];
            return existingShifts.Any(s =>
                (s.Start < shift.End && s.End > shift.Start));
        }

        private int GetCandidateCount(Shift shift, ScheduleConfig config)
        {
            int count = 0;
            foreach (var person in _availableHours.Keys)
            {
                if (IsAvailable(person, shift) && CanAssign(person, shift, config))
                {
                    count++;
                }
            }
            return count;
        }

        private void Assign(string person, Shift shift)
        {
            shift.AssignedPeople.Add(person);

            if (!_assignedHours.ContainsKey(person))
                _assignedHours[person] = 0;

            _assignedHours[person] += shift.DurationHours;

            var date = shift.Date.Date;

            if (!_personShifts.ContainsKey(person))
                _personShifts[person] = new Dictionary<DateTime, List<Shift>>();

            if (!_personShifts[person].ContainsKey(date))
                _personShifts[person][date] = new List<Shift>();

            _personShifts[person][date].Add(shift);
        }

        private double GetWeeklyHours(string person)
        {
            if (!_personShifts.ContainsKey(person))
                return 0;

            double total = 0;
            foreach (var dateShifts in _personShifts[person])
            {
                total += dateShifts.Value.Sum(s => s.DurationHours);
            }
            return total;
        }

        private bool IsAvailable(string person, Shift shift)
        {
            return _allAvailabilities.Any(a =>
                a.Name == person &&
                a.Date.Date == shift.Date.Date &&
                a.Start <= shift.Start &&
                a.End >= shift.End);
        }

        private void CalculateTotalAvailableHours()
        {
            WriteDebugLog($"\n=== CALCULATING AVAILABLE HOURS ===");

            foreach (var g in _allAvailabilities.GroupBy(a => a.Name))
            {
                WriteDebugLog($"\nProcessing {g.Key}:");
                double totalHours = 0;

                foreach (var avail in g.OrderBy(a => a.Date).ThenBy(a => a.Start))
                {
                    double slotHours = (avail.End - avail.Start).TotalHours;
                    totalHours += slotHours;
                    WriteDebugLog($"  {avail.Date:yyyy-MM-dd} {avail.Start:hh\\:mm}-{avail.End:hh\\:mm}: {slotHours:F2}h");
                }

                _availableHours[g.Key] = totalHours;
                WriteDebugLog($"  TOTAL: {totalHours:F2}h");

                foreach (var avail in g)
                {
                    avail.TotalAvailableHours = totalHours;
                }
            }
        }

        private void LogAllAvailabilities()
        {
            WriteDebugLog($"\n=== ALL AVAILABILITIES DATA ===");
            WriteDebugLog($"Total records: {_allAvailabilities.Count}");

            foreach (var person in _allAvailabilities.Select(a => a.Name).Distinct())
            {
                var personAvail = _allAvailabilities.Where(a => a.Name == person).ToList();
                WriteDebugLog($"\n{person}: {personAvail.Count} availability blocks");

                foreach (var avail in personAvail.OrderBy(a => a.Date).ThenBy(a => a.Start))
                {
                    WriteDebugLog($"  {avail.Date:yyyy-MM-dd} {avail.Start:hh\\:mm}-{avail.End:hh\\:mm} ({(avail.End - avail.Start).TotalHours:F2}h)");
                }

                double total = personAvail.Sum(a => (a.End - a.Start).TotalHours);
                WriteDebugLog($"  Total available hours: {total:F2}");
            }
        }

        private void InitializeTracking()
        {
            _assignedHours.Clear();
            _personShifts.Clear();
            _weeklyHours.Clear();

            foreach (var person in _availableHours.Keys)
            {
                _assignedHours[person] = 0;
                _personShifts[person] = new Dictionary<DateTime, List<Shift>>();
            }
        }

        private List<Shift> GenerateShiftsByDate(
            List<PersonAvailability> availabilities,
            ScheduleConfig config)
        {
            var shifts = new List<Shift>();

            var dates = availabilities
                .Select(a => a.Date.Date)
                .Distinct()
                .OrderBy(d => d);

            foreach (var date in dates)
            {
                if (config.ClosedDays.Contains(date.DayOfWeek))
                    continue;

                var current = config.OpeningTime;

                while (current < config.ClosingTime)
                {
                    TimeSpan shiftLength;
                    var timeLeft = config.ClosingTime - current;

                    if (timeLeft < config.ShiftLength)
                    {
                        shiftLength = timeLeft;
                    }
                    else
                    {
                        shiftLength = config.ShiftLength;
                    }

                    shifts.Add(new Shift
                    {
                        Date = date,
                        Start = current,
                        End = current + shiftLength,
                        PeopleNeeded = config.PeoplePerShift
                    });

                    current += shiftLength;
                }
            }

            return shifts;
        }

        private void FinalizeSchedule(List<Shift> shifts, ScheduleConfig config)
        {
            WriteDebugLog($"\n=== FINAL SCHEDULE SUMMARY ===");
            WriteDebugLog($"Total shifts: {shifts.Count}");

            int filledShifts = shifts.Count(s => s.AssignedPeople.Count > 0);
            int emptyShifts = shifts.Count(s => s.AssignedPeople.Count == 0);
            int fullyStaffedShifts = shifts.Count(s => s.AssignedPeople.Count >= s.PeopleNeeded);

            WriteDebugLog($"Filled shifts: {filledShifts}");
            WriteDebugLog($"Empty shifts: {emptyShifts}");
            WriteDebugLog($"Fully staffed shifts: {fullyStaffedShifts}/{shifts.Count}");
            WriteDebugLog($"Understaffed shifts: {_totalUnderstaffedShifts}");

            WriteDebugLog($"\nDetailed shift assignments:");
            foreach (var shift in shifts.OrderBy(s => s.Date).ThenBy(s => s.Start))
            {
                string status = shift.AssignedPeople.Count > 0 ?
                    $"Assigned: {string.Join(", ", shift.AssignedPeople)} ({shift.AssignedPeople.Count}/{shift.PeopleNeeded})" :
                    "NO ASSIGNMENTS";
                WriteDebugLog($"  {shift.Date:yyyy-MM-dd} ({shift.Date.DayOfWeek}) {shift.Start.Hours}:00-{shift.End.Hours}:00: {status}");
            }

            WriteDebugLog($"\nWeekly hours per person:");
            var weeklyHours = CalculateAllWeeklyHours();
            foreach (var kvp in weeklyHours.OrderBy(k => k.Value))
            {
                var person = kvp.Key;
                double totalAssigned = _assignedHours[person];
                double available = _availableHours[person];

                WriteDebugLog($"  {person}: {kvp.Value:F2}h weekly, {totalAssigned:F2}h total / {available:F2}h available");
            }

            WriteDebugLog($"\nCoverage analysis:");
            double totalShiftHoursNeeded = shifts.Sum(s => s.DurationHours * s.PeopleNeeded);
            double totalAssignedHours = shifts.Sum(s => s.DurationHours * s.AssignedPeople.Count);
            double coveragePercentage = totalShiftHoursNeeded > 0 ? (totalAssignedHours / totalShiftHoursNeeded) * 100 : 0;

            WriteDebugLog($"  Total shift hours needed: {totalShiftHoursNeeded:F2}");
            WriteDebugLog($"  Total assigned hours: {totalAssignedHours:F2}");
            WriteDebugLog($"  Coverage: {coveragePercentage:F1}%");

            WriteDebugLog($"\nShift continuity analysis:");
            foreach (var person in _personShifts.Keys)
            {
                var splitShifts = new List<string>();
                foreach (var dateShifts in _personShifts[person])
                {
                    var shiftsOnDay = dateShifts.Value.OrderBy(s => s.Start).ToList();
                    if (shiftsOnDay.Count > 1)
                    {
                        for (int i = 0; i < shiftsOnDay.Count - 1; i++)
                        {
                            if (shiftsOnDay[i].End != shiftsOnDay[i + 1].Start)
                            {
                                splitShifts.Add($"{dateShifts.Key:yyyy-MM-dd}: gap between {shiftsOnDay[i].End.Hours}:00 and {shiftsOnDay[i + 1].Start.Hours}:00");
                            }
                        }
                    }
                }

                if (splitShifts.Count > 0)
                {
                    WriteDebugLog($"  {person} has {splitShifts.Count} split shift day(s):");
                    foreach (var split in splitShifts.Take(3))
                    {
                        WriteDebugLog($"    {split}");
                    }
                    if (splitShifts.Count > 3)
                        WriteDebugLog($"    ... and {splitShifts.Count - 3} more");
                }
            }
        }

        private void WriteDebugLog(string message)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string debugFilePath = Path.Combine(desktopPath, "scheduler_debug.txt");

                using (StreamWriter writer = new StreamWriter(debugFilePath, true))
                {
                    writer.WriteLine($"[Scheduler {DateTime.Now:HH:mm:ss.fff}] {message}");
                }
            }
            catch
            {
            }
        }
    }
}