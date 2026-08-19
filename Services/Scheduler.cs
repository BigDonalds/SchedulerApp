using System;
using System.Collections.Generic;
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

    public class PersonAssignment
    {
        public string Name { get; set; }
        public TimeSpan? StartsAt { get; set; }  // null = start of shift
        public TimeSpan? EndsAt { get; set; }    // null = end of shift
        public bool IsPartial => StartsAt.HasValue || EndsAt.HasValue;

        public override string ToString()
        {
            if (!IsPartial) return Name;
            if (StartsAt.HasValue && EndsAt.HasValue)
                return $"{Name}(starts {StartsAt.Value:hh\\:mm}, ends {EndsAt.Value:hh\\:mm})";
            if (StartsAt.HasValue)
                return $"{Name}(starts {StartsAt.Value:hh\\:mm})";
            return $"{Name}(ends {EndsAt.Value:hh\\:mm})";
        }
    }

    public enum WeekType
    {
        MonToFri,
        MonToSun
    }

    public class ScheduleConfig
    {
        public TimeSpan OpeningTime { get; set; }
        public TimeSpan ClosingTime { get; set; }
        public TimeSpan ShiftLength { get; set; }
        public int PeoplePerShift { get; set; }
        public HashSet<DayOfWeek> ClosedDays { get; set; } = new HashSet<DayOfWeek>();
        public bool AllowShiftStacking { get; set; } = true;
        public WeekType WeekDefinition { get; set; } = WeekType.MonToFri;
    }

    public class ScheduleResult
    {
        public List<Shift> Shifts { get; set; } = new List<Shift>();
        public bool HasUnfilledShifts =>
            Shifts.Any(s => s.AssignedPeople.Count < s.PeopleNeeded);
        public List<UnderstaffingAlert> UnderstaffingAlerts { get; set; } = new List<UnderstaffingAlert>();
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
        public Dictionary<string, PersonAssignment> PersonAssignments { get; set; } = new Dictionary<string, PersonAssignment>();
        public int PositionInDay { get; set; } // 0 = first shift, 1 = second shift, etc.
        public bool IsLastShiftOfDay { get; set; }
        public bool IsFirstShiftOfDay { get; set; }
        public bool IsLastShiftPartial { get; set; } // True if this last shift is shorter than ShiftLength
        public bool IsLocked { get; set; } // True if this shift should not be modified by any phase
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
        private Dictionary<(string person, DateTime date), double> _dailyAssignedHours = new Dictionary<(string, DateTime), double>();
        private Dictionary<(string person, DateTime date), int> _consecutiveShiftCount = new Dictionary<(string, DateTime), int>();
        private const double MaxDailyHoursRatio = 0.45;
        private const double DailyPenaltyThreshold = 0.30;
        // Stores pre-identified candidates for last-two-shifts stacking per day
        private Dictionary<DateTime, List<string>> _stackingCandidates = new Dictionary<DateTime, List<string>>();

        public ScheduleResult GenerateSchedule(
            List<PersonAvailability> availabilities,
            ScheduleConfig config)
        {
            _allAvailabilities = availabilities;
            CalculateTotalAvailableHours();
            InitializeTracking();

            var shifts = GenerateShiftsByDate(availabilities, config);

            // Mark shift positions and first/last shifts
            MarkShiftPositions(shifts, config);

            // Group shifts by date for easy access
            _shiftsByDate = shifts.GroupBy(s => s.Date.Date).ToDictionary(g => g.Key, g => g.OrderBy(s => s.Start).ToList());

            var peopleByAvailability = _availableHours
                .OrderBy(k => k.Value)
                .Select(k => k.Key)
                .ToList();

            // PHASE 0: Pre-identify last-two-shifts stacking candidates (Rule 3)
            // Before any scheduling, identify days where the last shift is partial,
            // find candidates available for BOTH last two shifts, and lock those shifts
            PreIdentifyLastTwoShifts(shifts, config);

            var hitMaps = BuildHitMaps(shifts, config);

            // PHASE 1: Tiered Heatmap Assignment (Rule 1 + Rule 2 combined)
            // Progressively assigns shifts in tiers of shiftLength * (tier+1)
            // At each tier, processes shifts by heatmap priority (fewest candidates first)
            // but caps each person's hours at the current tier target
            AssignTieredHeatmap(shifts, hitMaps, peopleByAvailability, config);

            // PHASE 2: Last short shift stacking (Rule 3) - runs after heatmap so previous shifts are filled
            // If last shift is shorter than normal, prioritize people from previous shift to take it
            AssignLastShortShiftStacking(shifts, config);

            // PHASE 3: Handle partial availability overlaps (Rule 6)
            // For remaining unfilled shifts, find candidates with partial overlap
            AssignPartialOverlaps(shifts, config);

            // PHASE 4: Fill remaining shifts with progressive daily cap penalty (Rule 4)
            FillRemainingShifts(shifts, peopleByAvailability, config);

            // PHASE 5: Low-hour extras (Rule 7)
            // People with < 4 hours total weekly availability get added as one extra per shift
            AssignLowHourExtras(shifts, config);

            // PHASE 6: Same-day gap prevention sweep (Rule 5)
            PreventSameDayGaps(shifts, config);

            // PHASE 7: Smart selection for last-two-shifts stacking
            // After all other scheduling is done, pick the best candidates for the locked last two shifts
            AssignLastTwoShiftsSmart(shifts, config);

            // PHASE 8: Zero-hour pity assignments
            // People with 0 weekly hours get added as extra 3rd person on shifts they partially overlap
            AssignZeroHourPity(shifts, config);

            // PHASE 9: Partial coverage gap filler
            // Fix gaps where partial people leave early and no one covers the remaining time
            FillPartialCoverageGaps(shifts, config);

            // PHASE 10: Half-hour understaffing scan and fix
            // After all other phases, walk every half-hour and verify the shift has
            // enough people. Try to fix understaffed windows with available candidates.
            var understaffingAlerts = CheckAndFixUnderstaffing(shifts, config);

            // PHASE 11: Overwork/Underwork Balance Exchange
            // Transfer full shifts from people above the weekly average to those below it.
            // Handles stacked blocks (full+partial / partial+full / partial+partial).
            BalanceWorkload(shifts, config);

            // PHASE 12: Re-scan understaffing after balancing
            // Transfers can leave half-hour gaps uncovered (e.g. a person who was
            // covering an early window is moved off, leaving only a late-starting
            // partial behind). Re-run the scan so alerts reflect the FINAL state
            // and any newly-created windows are fixed or reported.
            understaffingAlerts = CheckAndFixUnderstaffing(shifts, config);

            return new ScheduleResult
            {
                Shifts = shifts,
                UnderstaffingAlerts = understaffingAlerts
            };
        }

        private void MarkShiftPositions(List<Shift> shifts, ScheduleConfig config)
        {
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();

                for (int i = 0; i < dayShifts.Count; i++)
                {
                    dayShifts[i].PositionInDay = i;
                    dayShifts[i].IsFirstShiftOfDay = (i == 0);
                    dayShifts[i].IsLastShiftOfDay = (i == dayShifts.Count - 1);
                    // Mark if this is the last shift and it's shorter than the configured shift length
                    dayShifts[i].IsLastShiftPartial = dayShifts[i].IsLastShiftOfDay &&
                        dayShifts[i].DurationHours < config.ShiftLength.TotalHours;
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

        // ===================================================================
        // PHASE 2: Partial Availability Assignment (Rule 6)
        // Handle candidates who are available for only part of a shift
        // ===================================================================
        private void AssignPartialOverlaps(List<Shift> shifts, ScheduleConfig config)
        {
            var unfilledShifts = shifts
                .Where(s => s.AssignedPeople.Count < s.PeopleNeeded)
                .Where(s => !s.IsLocked) // Skip locked shifts (Phase 0/8 handles them)
                .OrderBy(s => s.Date)
                .ThenBy(s => s.Start)
                .ToList();

            foreach (var shift in unfilledShifts)
            {
                int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;
                if (needed <= 0) continue;

                // Find candidates with partial overlap who aren't already assigned
                // NOTE: Uses CanAssignPartial which only checks overlap + same-day gaps,
                // NOT full availability (since partial by definition doesn't cover full shift)
                var partialCandidates = _availableHours.Keys
                    .Where(p => !shift.AssignedPeople.Contains(p))
                    .Where(p => IsPartiallyAvailable(p, shift))
                    .Where(p => CanAssignPartial(p, shift, config))
                    .OrderBy(p => _availableHours[p])
                    .ThenBy(p => _assignedHours.ContainsKey(p) ? _assignedHours[p] : 0)
                    .ToList();

                if (partialCandidates.Count == 0)
                {
                    continue;
                }

                foreach (var candidate in partialCandidates.Take(needed))
                {
                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                        break;

                    var assignment = GetOverlappingPersonAssignment(candidate, shift);
                    if (assignment == null) continue;

                    // Check if this partial assignment would create a same-day gap (Rule 5)
                    if (WouldCreateGapOnSameDay(candidate, shift))
                    {
                        continue;
                    }

                    // Assign with partial details
                    AssignPartial(candidate, shift, assignment);
                }
            }
        }

        // ===================================================================
        // PHASE 3: Last Short Shift Stacking (Rule 3)
        // If the last shift of the day is shorter than ShiftLength,
        // prioritize people from the previous shift to take it
        // ===================================================================
        private void AssignLastShortShiftStacking(List<Shift> shifts, ScheduleConfig config)
        {
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                if (dayShifts.Count < 2) continue;

                var lastShift = dayShifts.Last();

                // Only apply if last shift is shorter than configured shift length
                // and still needs people
                if (!lastShift.IsLastShiftPartial) continue;
                if (lastShift.AssignedPeople.Count >= lastShift.PeopleNeeded) continue;

                var previousShift = dayShifts[dayShifts.Count - 2];

                // Skip if this day's last two shifts are locked (Phase 0 handles it)
                if (previousShift.IsLocked || lastShift.IsLocked) continue;

                // Prioritize people from the previous shift
                var stackCandidates = previousShift.AssignedPeople
                    .Where(p => IsAvailable(p, lastShift) && CanAssign(p, lastShift, config) && !lastShift.AssignedPeople.Contains(p))
                    .ToList();

                foreach (var candidate in stackCandidates)
                {
                    if (lastShift.AssignedPeople.Count >= lastShift.PeopleNeeded)
                        break;

                    Assign(candidate, lastShift);
                }

                // If still unfilled, try others who are available
                if (lastShift.AssignedPeople.Count < lastShift.PeopleNeeded)
                {
                    int needed = lastShift.PeopleNeeded - lastShift.AssignedPeople.Count;
                    var otherCandidates = _availableHours.Keys
                        .Where(p => !lastShift.AssignedPeople.Contains(p) && !previousShift.AssignedPeople.Contains(p))
                        .Where(p => IsAvailable(p, lastShift))
                        .Where(p => CanAssign(p, lastShift, config))
                        .OrderBy(p => _availableHours[p])
                        .ThenBy(p => _assignedHours.ContainsKey(p) ? _assignedHours[p] : 0)
                        .Take(needed)
                        .ToList();

                    foreach (var candidate in otherCandidates)
                    {
                        if (lastShift.AssignedPeople.Count >= lastShift.PeopleNeeded)
                            break;

                        Assign(candidate, lastShift);
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 4: Fill Remaining Shifts with Daily Cap Penalty (Rule 4)
        // ===================================================================
        private void FillRemainingShifts(
            List<Shift> shifts,
            List<string> orderedPeople,
            ScheduleConfig config)
        {
            var remaining = shifts
                .Where(s => s.AssignedPeople.Count < s.PeopleNeeded)
                .Where(s => !s.IsLocked) // Skip locked shifts (Phase 0/8 handles them)
                .OrderBy(s => GetCandidateCount(s, config))
                .ThenBy(s => s.IsLastShiftOfDay ? 0 : 1)
                .ThenBy(s => s.Start)
                .ToList();

            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;

            foreach (var shift in remaining)
            {
                int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;
                if (needed <= 0) continue;

                // Find candidates with progressive daily cap penalty
                var candidates = orderedPeople
                    .Where(p => !shift.AssignedPeople.Contains(p))
                    .Where(p => IsAvailable(p, shift))
                    .Where(p => CanAssign(p, shift, config))
                    .OrderBy(p => CalculateFairnessScore(p, shift, maxDailyHours))
                    .Take(needed * 2)
                    .ToList();

                // If not enough candidates, relax the split-shift restriction
                if (candidates.Count < needed)
                {
                    candidates.AddRange(orderedPeople
                        .Where(p => !shift.AssignedPeople.Contains(p) && !candidates.Contains(p))
                        .Where(p => IsAvailable(p, shift))
                        .Where(p => CanAssign(p, shift, config))
                        .OrderBy(p => CalculateFairnessScore(p, shift, maxDailyHours))
                        .Take(needed - candidates.Count));
                }

                foreach (var candidate in candidates.Take(needed))
                {
                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded)
                        break;

                    // Check for same-day gap (Rule 5)
                    if (WouldCreateGapOnSameDay(candidate, shift))
                    {
                        continue;
                    }

                    Assign(candidate, shift);
                }

                if (shift.AssignedPeople.Count < shift.PeopleNeeded)
                {
                    _totalUnderstaffedShifts++;
                }
            }
        }

        // ===================================================================
        // PHASE 5: Low-Hour Extra Assignments (Rule 7)
        // People with < 4 hours total weekly availability get added as one extra
        // ===================================================================
        private void AssignLowHourExtras(List<Shift> shifts, ScheduleConfig config)
        {
            // Identify low-hour people (less than 4 hours total weekly availability)
            var lowHourPeople = _availableHours
                .Where(kvp => kvp.Value < 4.0)
                .Select(kvp => kvp.Key)
                .ToList();

            if (lowHourPeople.Count == 0)
            {
                return;
            }

            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;

            foreach (var shift in shifts)
            {
                // Skip locked shifts (handled by Phase 0/8)
                if (shift.IsLocked) continue;

                // Only add one extra person per shift, and only if shift is already staffed
                if (shift.AssignedPeople.Count < shift.PeopleNeeded)
                    continue;

                // Find a low-hour person who is available and not already assigned to this shift
                var extraCandidate = lowHourPeople
                    .Where(p => !shift.AssignedPeople.Contains(p))
                    .Where(p => IsAvailable(p, shift))
                    .Where(p => CanAssign(p, shift, config))
                    .Where(p => !WouldCreateGapOnSameDay(p, shift))
                    .OrderBy(p => GetDailyAssignedHours(p, shift.Date) / maxDailyHours) // Prefer those with lower daily load
                    .FirstOrDefault();

                if (extraCandidate != null)
                {
                    // Check if this would exceed daily cap (Rule 4 applies to extras too)
                    double currentDaily = GetDailyAssignedHours(extraCandidate, shift.Date);
                    double newDaily = currentDaily + shift.DurationHours;

                    if (newDaily <= maxDailyHours || currentDaily < maxDailyHours * 0.3)
                    {
                        Assign(extraCandidate, shift);
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 6: Same-Day Gap Prevention (Rule 5)
        // Proactively prevent gaps between shifts on the same day
        // ===================================================================
        private void PreventSameDayGaps(List<Shift> shifts, ScheduleConfig config)
        {
            foreach (var person in _personShifts.Keys.ToList())
            {
                var personDays = _personShifts[person];

                foreach (var dateEntry in personDays.ToList())
                {
                    var date = dateEntry.Key;
                    var shiftsOnDay = dateEntry.Value.OrderBy(s => s.Start).ToList();

                    if (shiftsOnDay.Count <= 1) continue;

                    // Check for gaps between shifts
                    for (int i = 0; i < shiftsOnDay.Count - 1; i++)
                    {
                        var current = shiftsOnDay[i];
                        var next = shiftsOnDay[i + 1];

                        if (current.End != next.Start)
                        {
                            // Try to remove the person from one of the shifts
                            // Prefer removing from the earlier shift if possible
                            if (CanRemovePersonFromShift(person, current, shifts, config))
                            {
                                RemoveShift(person, current);
                            }
                            else if (CanRemovePersonFromShift(person, next, shifts, config))
                            {
                                RemoveShift(person, next);
                            }
                        }
                    }
                }
            }
        }

        // ===================================================================
        // Partial Availability Helpers (Rule 6)
        // ===================================================================

        /// <summary>
        /// Checks if a person has ANY overlap with a shift (not necessarily full coverage)
        /// </summary>
        private bool IsPartiallyAvailable(string person, Shift shift)
        {
            return _allAvailabilities.Any(a =>
                a.Name == person &&
                a.Date.Date == shift.Date.Date &&
                a.Start < shift.End &&
                a.End > shift.Start);
        }

        /// <summary>
        /// Gets the overlapping portion of a person's availability with a shift.
        /// Returns null if no overlap.
        /// </summary>
        private PersonAssignment GetOverlappingPersonAssignment(string person, Shift shift)
        {
            var availability = _allAvailabilities.FirstOrDefault(a =>
                a.Name == person &&
                a.Date.Date == shift.Date.Date &&
                a.Start < shift.End &&
                a.End > shift.Start);

            if (availability == null) return null;

            TimeSpan overlapStart = availability.Start > shift.Start ? availability.Start : shift.Start;
            TimeSpan overlapEnd = availability.End < shift.End ? availability.End : shift.End;

            // Only assign if there's meaningful overlap (at least 30 minutes)
            if ((overlapEnd - overlapStart).TotalMinutes < 30)
                return null;

            var assignment = new PersonAssignment
            {
                Name = person,
                StartsAt = availability.Start > shift.Start ? availability.Start : (TimeSpan?)null,
                EndsAt = availability.End < shift.End ? availability.End : (TimeSpan?)null
            };

            return assignment;
        }

        /// <summary>
        /// Gets the number of hours a person overlaps with a shift (for sorting)
        /// </summary>
        private double GetOverlappingHours(string person, Shift shift)
        {
            var availability = _allAvailabilities.FirstOrDefault(a =>
                a.Name == person &&
                a.Date.Date == shift.Date.Date &&
                a.Start < shift.End &&
                a.End > shift.Start);

            if (availability == null) return 0;

            TimeSpan overlapStart = availability.Start > shift.Start ? availability.Start : shift.Start;
            TimeSpan overlapEnd = availability.End < shift.End ? availability.End : shift.End;

            return Math.Max(0, (overlapEnd - overlapStart).TotalHours);
        }

        /// <summary>
        /// Assigns a person to a shift with partial availability details
        /// </summary>
        private void AssignPartial(string person, Shift shift, PersonAssignment assignment)
        {
            shift.AssignedPeople.Add(person);
            shift.PersonAssignments[person] = assignment;

            if (!_assignedHours.ContainsKey(person))
                _assignedHours[person] = 0;

            // Calculate actual hours worked (overlap duration)
            TimeSpan actualStart = assignment.StartsAt ?? shift.Start;
            TimeSpan actualEnd = assignment.EndsAt ?? shift.End;
            double actualHours = (actualEnd - actualStart).TotalHours;

            _assignedHours[person] += actualHours;

            var date = shift.Date.Date;

            if (!_personShifts.ContainsKey(person))
                _personShifts[person] = new Dictionary<DateTime, List<Shift>>();

            if (!_personShifts[person].ContainsKey(date))
                _personShifts[person][date] = new List<Shift>();

            _personShifts[person][date].Add(shift);

            // Track daily hours
            var key = (person, date);
            if (!_dailyAssignedHours.ContainsKey(key))
                _dailyAssignedHours[key] = 0;
            _dailyAssignedHours[key] += actualHours;
        }

        // ===================================================================
        // Daily Cap & Fairness (Rule 4)
        // ===================================================================

        private double GetDailyAssignedHours(string person, DateTime date)
        {
            var key = (person, date.Date);
            return _dailyAssignedHours.ContainsKey(key) ? _dailyAssignedHours[key] : 0;
        }

        private double CalculateDailyPenalty(string person, Shift shift, double maxDailyHours)
        {
            double currentDaily = GetDailyAssignedHours(person, shift.Date);
            double newDaily = currentDaily + shift.DurationHours;
            double ratio = newDaily / maxDailyHours;

            // Below threshold: no penalty
            if (ratio < DailyPenaltyThreshold)
                return 0;

            // Progressive penalty: ratio^2 * 100
            // At 30%: penalty = 9
            // At 45%: penalty = 20.25
            // At 60%: penalty = 36
            // At 100%: penalty = 100
            return ratio * ratio * 100;
        }

        private double CalculateFairnessScore(string person, Shift shift, double maxDailyHours)
        {
            double assignedHours = _assignedHours.ContainsKey(person) ? _assignedHours[person] : 0;
            double weeklyHours = GetWeeklyHours(person);
            double availabilityHours = _availableHours[person];

            double score = 0;

            // Prefer people with fewer assigned hours
            score += assignedHours * 10;

            // Prefer people whose availability is being underutilized
            double utilization = availabilityHours > 0 ? assignedHours / availabilityHours : 1;
            score += utilization * 5;

            // Slight penalty for weekly hours to spread work across week
            score += weeklyHours * 2;

            // Daily cap penalty (Rule 4) - progressive
            double dailyPenalty = CalculateDailyPenalty(person, shift, maxDailyHours);
            score += dailyPenalty;

            // Consecutive shift penalty (Rule 5) - avoid overusing consecutive stacking
            var consecKey = (person, shift.Date.Date);
            if (_consecutiveShiftCount.ContainsKey(consecKey))
            {
                int consecCount = _consecutiveShiftCount[consecKey];
                if (consecCount >= 2)
                {
                    score += consecCount * 15; // Heavy penalty for 3+ consecutive
                }
                else if (consecCount >= 1)
                {
                    score += 5; // Mild penalty for 2 consecutive
                }
            }

            return score;
        }

        // ===================================================================
        // Same-Day Gap Prevention (Rule 5)
        // ===================================================================

        /// <summary>
        /// Checks if assigning a person to a shift would create a gap on the same day.
        /// A gap means the person has other shifts on the same day but this new shift
        /// is not adjacent to any of them.
        /// </summary>
        private bool WouldCreateGapOnSameDay(string person, Shift shift)
        {
            if (!_personShifts.ContainsKey(person) || !_personShifts[person].ContainsKey(shift.Date.Date))
                return false;

            var existingShifts = _personShifts[person][shift.Date.Date];

            // If no existing shifts on this day, no gap possible
            if (existingShifts.Count == 0)
                return false;

            // Check if this shift is adjacent to any existing shift
            bool isAdjacent = existingShifts.Any(s =>
                s.End == shift.Start || s.Start == shift.End);

            // If not adjacent and person already has shifts, it creates a gap
            return !isAdjacent;
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

            // Track daily hours
            var key = (person, date);
            if (!_dailyAssignedHours.ContainsKey(key))
                _dailyAssignedHours[key] = 0;
            _dailyAssignedHours[key] += shift.DurationHours;

            // Track consecutive shifts
            var consecKey = (person, date);
            if (!_consecutiveShiftCount.ContainsKey(consecKey))
                _consecutiveShiftCount[consecKey] = 0;

            // Check if this shift is consecutive with an existing shift
            var existingShifts = _personShifts[person][date];
            bool isConsecutive = existingShifts.Any(s =>
                s.Id != shift.Id && (s.End == shift.Start || s.Start == shift.End));

            if (isConsecutive)
            {
                _consecutiveShiftCount[consecKey]++;
            }
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

            // Check for same-day gap (Rule 5)
            if (WouldCreateGapOnSameDay(person, shift))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Checks if a person can be assigned to a shift with partial coverage.
        /// Unlike CanAssign, this does NOT require full availability coverage -
        /// it only checks for overlapping shifts and same-day gaps.
        /// The actual availability overlap is validated separately via GetOverlappingPersonAssignment.
        /// </summary>
        private bool CanAssignPartial(string person, Shift shift, ScheduleConfig config)
        {
            // Check for overlapping shifts (still can't be in two places at once)
            if (HasOverlap(person, shift))
            {
                return false;
            }

            // Check for same-day gap (Rule 5)
            if (WouldCreateGapOnSameDay(person, shift))
            {
                return false;
            }

            return true;
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

        private bool CanRemovePersonFromShift(string person, Shift shift, List<Shift> allShifts, ScheduleConfig config)
        {
            // Don't remove if it would leave shift understaffed
            if (shift.AssignedPeople.Count <= shift.PeopleNeeded)
                return false;

            // Check if there are other candidates for this shift
            var otherCandidates = _availableHours.Keys
                .Where(p => p != person)
                .Where(p => IsAvailable(p, shift))
                .Where(p => CanAssign(p, shift, config))
                .ToList();

            return otherCandidates.Count > 0;
        }

        private void RemoveShift(string person, Shift shift)
        {
            shift.AssignedPeople.Remove(person);

            // Also remove partial assignment if exists
            if (shift.PersonAssignments.ContainsKey(person))
            {
                shift.PersonAssignments.Remove(person);
            }

            double hoursToRemove = shift.DurationHours;
            if (_assignedHours.ContainsKey(person))
            {
                _assignedHours[person] -= hoursToRemove;
            }

            var date = shift.Date.Date;
            if (_personShifts.ContainsKey(person) && _personShifts[person].ContainsKey(date))
            {
                _personShifts[person][date].Remove(shift);
            }

            // Update daily hours
            var key = (person, date);
            if (_dailyAssignedHours.ContainsKey(key))
            {
                _dailyAssignedHours[key] -= hoursToRemove;
                if (_dailyAssignedHours[key] <= 0)
                    _dailyAssignedHours.Remove(key);
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

        private void CalculateTotalAvailableHours()
        {
            foreach (var g in _allAvailabilities.GroupBy(a => a.Name))
            {
                double totalHours = 0;

                foreach (var avail in g.OrderBy(a => a.Date).ThenBy(a => a.Start))
                {
                    double slotHours = (avail.End - avail.Start).TotalHours;
                    totalHours += slotHours;
                }

                _availableHours[g.Key] = totalHours;

                foreach (var avail in g)
                {
                    avail.TotalAvailableHours = totalHours;
                }
            }
        }

        private void InitializeTracking()
        {
            _assignedHours.Clear();
            _personShifts.Clear();
            _weeklyHours.Clear();
            _dailyAssignedHours.Clear();
            _consecutiveShiftCount.Clear();

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

        // ===================================================================
        // PHASE 0: Pre-identify last-two-shifts stacking candidates
        // Before any scheduling, find people available for BOTH last two shifts
        // when the last shift is partial, and lock those shifts
        // ===================================================================
        private void PreIdentifyLastTwoShifts(List<Shift> shifts, ScheduleConfig config)
        {
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                if (dayShifts.Count < 2) continue;

                var lastShift = dayShifts.Last();
                var prevShift = dayShifts[dayShifts.Count - 2];

                // Only apply if last shift is shorter than configured shift length
                if (!lastShift.IsLastShiftPartial) continue;

                // Find people available for BOTH the last two shifts
                var candidates = _availableHours.Keys
                    .Where(p => IsAvailable(p, prevShift) && IsAvailable(p, lastShift))
                    .ToList();

                // Store candidates for later smart selection
                _stackingCandidates[lastShift.Date.Date] = candidates;

                // Lock both shifts so other phases don't modify them
                prevShift.IsLocked = true;
                lastShift.IsLocked = true;
            }
        }

        // ===================================================================
        // PHASE 8: Smart last-two-shifts assignment
        // After all other scheduling is done, pick the best candidates
        // 1. Remove candidates with same-day gaps
        // 2. Pick 2 with lowest weekly hours
        // ===================================================================
        private void AssignLastTwoShiftsSmart(List<Shift> shifts, ScheduleConfig config)
        {
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                if (dayShifts.Count < 2) continue;

                var lastShift = dayShifts.Last();
                var prevShift = dayShifts[dayShifts.Count - 2];

                // Only process if we have pre-identified candidates for this day
                if (!_stackingCandidates.ContainsKey(lastShift.Date.Date)) continue;
                if (!lastShift.IsLastShiftPartial) continue;

                var candidates = _stackingCandidates[lastShift.Date.Date];

                // Step 1: Remove candidates who would have same-day gaps
                // (already assigned to other shifts on this day that aren't adjacent to these)
                var noGapCandidates = candidates
                    .Where(p => !WouldCreateGapOnSameDay(p, prevShift))
                    .Where(p => !WouldCreateGapOnSameDay(p, lastShift))
                    .ToList();

                if (noGapCandidates.Count == 0)
                {
                    noGapCandidates = candidates.ToList();
                }

                // Step 2: Pick 2 with lowest weekly hours
                var selected = noGapCandidates
                    .OrderBy(p => GetWeeklyHours(p))
                    .Take(config.PeoplePerShift)
                    .ToList();

                // Assign selected candidates to both shifts
                foreach (var person in selected)
                {
                    if (prevShift.AssignedPeople.Count < prevShift.PeopleNeeded)
                    {
                        Assign(person, prevShift);
                    }
                    if (lastShift.AssignedPeople.Count < lastShift.PeopleNeeded)
                    {
                        Assign(person, lastShift);
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 9: Zero-Hour Pity Assignments
        // People with 0h weekly get added as extra (3rd) person on shifts
        // they partially overlap. Gives up to 3h total (preferring 1.5h chunks)
        // distributed across the week or as consecutive shifts.
        // ===================================================================
        private void AssignZeroHourPity(List<Shift> shifts, ScheduleConfig config)
        {
            // Find people with 0 weekly assigned hours
            var zeroHourPeople = _availableHours.Keys
                .Where(p => GetWeeklyHours(p) < 0.01)
                .ToList();

            if (zeroHourPeople.Count == 0)
            {
                return;
            }

            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;
            double shiftLen = config.ShiftLength.TotalHours;
            double maxWeeklyPity = Math.Min(shiftLen * 2, 3.0); // Up to 3h or 2 shift lengths

            foreach (var person in zeroHourPeople)
            {
                // Find all eligible shifts for this person, ordered by date then time
                var eligibleShifts = shifts
                    .Where(s => !s.AssignedPeople.Contains(person))
                    .Where(s => !s.IsLocked)
                    .Where(s => IsPartiallyAvailable(person, s))
                    .Where(s => CanAssignPartial(person, s, config))
                    .Where(s => !WouldCreateGapOnSameDay(person, s))
                    .OrderBy(s => s.Date)
                    .ThenBy(s => s.Start)
                    .ToList();

                double personTotal = 0;

                // Strategy 1: Try to find ONE big chunk (1.5h or more) first - prefer this
                var bigChunks = eligibleShifts
                    .Select(s => new { Shift = s, Overlap = GetOverlappingHours(person, s) })
                    .Where(x => x.Overlap >= shiftLen * 0.8) // At least 80% of a full shift
                    .OrderByDescending(x => x.Overlap)
                    .ToList();

                foreach (var chunk in bigChunks)
                {
                    if (personTotal >= shiftLen - 0.01) break; // Already have 1.5h, try for more
                    if (personTotal >= maxWeeklyPity - 0.01) break;

                    var assignment = GetOverlappingPersonAssignment(person, chunk.Shift);
                    if (assignment == null) continue;

                    double overlapHours = (assignment.EndsAt ?? chunk.Shift.End).TotalHours - (assignment.StartsAt ?? chunk.Shift.Start).TotalHours;
                    double dailyCheck = GetDailyAssignedHours(person, chunk.Shift.Date) + overlapHours;
                    if (dailyCheck > maxDailyHours && GetDailyAssignedHours(person, chunk.Shift.Date) >= maxDailyHours * 0.3)
                        continue;

                    AssignPartial(person, chunk.Shift, assignment);
                    personTotal += overlapHours;

                    if (personTotal >= maxWeeklyPity - 0.01) break;
                }

                // Strategy 2: Try consecutive shifts on the same day (e.g., 13:00-14:30 + 14:30-16:00)
                if (personTotal < maxWeeklyPity - 0.01)
                {
                    var shiftsByDate = eligibleShifts
                        .Where(s => !bigChunks.Any(b => b.Shift.Id == s.Id)) // Skip already used
                        .GroupBy(s => s.Date.Date)
                        .OrderBy(g => g.Key);

                    foreach (var dayGroup in shiftsByDate)
                    {
                        if (personTotal >= maxWeeklyPity - 0.01) break;

                        var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                        double dayTotal = 0;

                        // Look for consecutive shifts on this day
                        for (int i = 0; i < dayShifts.Count - 1; i++)
                        {
                            if (personTotal + dayTotal >= maxWeeklyPity - 0.01) break;

                            var s1 = dayShifts[i];
                            var s2 = dayShifts[i + 1];

                            // Check if they're consecutive (s1.End == s2.Start)
                            if (s1.End != s2.Start) continue;

                            var a1 = GetOverlappingPersonAssignment(person, s1);
                            var a2 = GetOverlappingPersonAssignment(person, s2);
                            if (a1 == null || a2 == null) continue;

                            double o1 = (a1.EndsAt ?? s1.End).TotalHours - (a1.StartsAt ?? s1.Start).TotalHours;
                            double o2 = (a2.EndsAt ?? s2.End).TotalHours - (a2.StartsAt ?? s2.Start).TotalHours;
                            double total = o1 + o2;

                            if (total < 0.5) continue;

                            double dailyCheck = GetDailyAssignedHours(person, s1.Date) + total;
                            if (dailyCheck > maxDailyHours && GetDailyAssignedHours(person, s1.Date) >= maxDailyHours * 0.3)
                                continue;

                            AssignPartial(person, s1, a1);
                            AssignPartial(person, s2, a2);
                            personTotal += total;
                            dayTotal += total;

                            i++; // Skip next since we used it
                        }
                    }
                }

                // Strategy 3: Fill remaining with single smaller partial shifts
                if (personTotal < maxWeeklyPity - 0.01)
                {
                    var usedShifts = new HashSet<Guid>();
                    if (bigChunks.Any()) usedShifts.UnionWith(bigChunks.Select(b => b.Shift.Id));

                    var remaining = eligibleShifts
                        .Where(s => !usedShifts.Contains(s.Id))
                        .OrderBy(s => s.Date)
                        .ThenBy(s => s.Start)
                        .ToList();

                    foreach (var shift in remaining)
                    {
                        if (personTotal >= maxWeeklyPity - 0.01) break;

                        var assignment = GetOverlappingPersonAssignment(person, shift);
                        if (assignment == null) continue;

                        double overlapHours = (assignment.EndsAt ?? shift.End).TotalHours - (assignment.StartsAt ?? shift.Start).TotalHours;
                        if (overlapHours < 0.5) continue; // Skip very small overlaps

                        double dailyCheck = GetDailyAssignedHours(person, shift.Date) + overlapHours;
                        if (dailyCheck > maxDailyHours && GetDailyAssignedHours(person, shift.Date) >= maxDailyHours * 0.3)
                            continue;

                        AssignPartial(person, shift, assignment);
                        personTotal += overlapHours;
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 1: Tiered Heatmap Assignment (Rule 1 + Rule 2)
        // Progressively assigns shifts in tiers: 1.5h, 3.0h, 4.5h, 6.0h, ...
        // At each tier, processes shifts by heatmap priority (fewest candidates first)
        // but caps each person at the current tier target to ensure fair distribution.
        // ===================================================================
        private void AssignTieredHeatmap(
            List<Shift> shifts,
            List<HitMap> hitMaps,
            List<string> orderedPeople,
            ScheduleConfig config)
        {
            double shiftLen = config.ShiftLength.TotalHours;
            int maxTier = 8;

            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;

            for (int tier = 0; tier < maxTier; tier++)
            {
                double targetHours = shiftLen * (tier + 1);
                var weeklyHours = CalculateAllWeeklyHours();

                // Find unfilled, unlocked shifts sorted by heatmap priority
                var unfilledShifts = hitMaps
                    .Where(h => h.Shift.AssignedPeople.Count < h.Shift.PeopleNeeded)
                    .Where(h => !h.Shift.IsLocked)
                    .Where(h => !h.Shift.IsLastShiftPartial)
                    .OrderBy(h => h.CandidateCount) // Heatmap: fewest candidates first
                    .ThenBy(h => h.Date)
                    .ThenBy(h => h.Start)
                    .ToList();

                if (unfilledShifts.Count == 0)
                {
                    continue;
                }

                foreach (var hitMap in unfilledShifts)
                {
                    var shift = hitMap.Shift;
                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded) continue;

                    int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;

                    // Find candidates below tier target, sorted by availability (lowest first)
                    var candidates = hitMap.Candidates
                        .Where(c => weeklyHours.ContainsKey(c) && weeklyHours[c] < targetHours)
                        .Where(c => !shift.AssignedPeople.Contains(c))
                        .Where(c => CanAssign(c, shift, config))
                        .OrderBy(c => weeklyHours[c]) // Fewest hours first
                        .ThenBy(c => _availableHours[c]) // Lowest availability first
                        .Take(needed)
                        .ToList();

                    foreach (var candidate in candidates)
                    {
                        if (shift.AssignedPeople.Count >= shift.PeopleNeeded) break;

                        double dailyCheck = GetDailyAssignedHours(candidate, shift.Date) + shift.DurationHours;
                        if (dailyCheck > maxDailyHours && GetDailyAssignedHours(candidate, shift.Date) >= maxDailyHours * 0.3)
                            continue;

                        Assign(candidate, shift);
                        weeklyHours = CalculateAllWeeklyHours(); // Refresh
                    }
                }

                // If there are still unfilled shifts and people below target, try partial assignments
                var stillUnfilled = hitMaps
                    .Where(h => h.Shift.AssignedPeople.Count < h.Shift.PeopleNeeded)
                    .Where(h => !h.Shift.IsLocked)
                    .ToList();

                foreach (var hitMap in stillUnfilled)
                {
                    var shift = hitMap.Shift;
                    if (shift.AssignedPeople.Count >= shift.PeopleNeeded) continue;

                    int needed = shift.PeopleNeeded - shift.AssignedPeople.Count;

                    var partialCandidates = _availableHours.Keys
                        .Where(p => !shift.AssignedPeople.Contains(p))
                        .Where(p => weeklyHours.ContainsKey(p) && weeklyHours[p] < targetHours)
                        .Where(p => IsPartiallyAvailable(p, shift))
                        .Where(p => CanAssignPartial(p, shift, config))
                        .OrderBy(p => weeklyHours[p])
                        .ThenBy(p => _availableHours[p])
                        .Take(needed)
                        .ToList();

                    foreach (var candidate in partialCandidates)
                    {
                        if (shift.AssignedPeople.Count >= shift.PeopleNeeded) break;

                        var assignment = GetOverlappingPersonAssignment(candidate, shift);
                        if (assignment == null) continue;

                        if (WouldCreateGapOnSameDay(candidate, shift)) continue;

                        double overlapHours = (assignment.EndsAt ?? shift.End).TotalHours - (assignment.StartsAt ?? shift.Start).TotalHours;
                        double dailyCheck = GetDailyAssignedHours(candidate, shift.Date) + overlapHours;
                        if (dailyCheck > maxDailyHours && GetDailyAssignedHours(candidate, shift.Date) >= maxDailyHours * 0.3)
                            continue;

                        // RULE: Partial assignments must bring the person's DAILY session to at least 1 full shift length (1.5h).
                        // This prevents fragmented 0.5h-1h orphan days while still
                        // allowing meaningful bundled sessions (e.g. 0.5h + 1.5h = 2h same day).
                        double currentDaily = GetDailyAssignedHours(candidate, shift.Date);
                        double projectedDaily = currentDaily + overlapHours;

                        // Also check if the person has adjacent same-day shifts that would make this a meaningful session
                        bool hasAdjacentShift = _personShifts.ContainsKey(candidate) &&
                            _personShifts[candidate].ContainsKey(shift.Date.Date) &&
                            _personShifts[candidate][shift.Date.Date].Any(s =>
                                s.End == shift.Start || s.Start == shift.End);

                        if (projectedDaily < shiftLen - 0.01 && !hasAdjacentShift)
                        {
                            continue;
                        }

                        AssignPartial(candidate, shift, assignment);
                        weeklyHours = CalculateAllWeeklyHours();
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 9: Partial Coverage Gap Filler
        // Fixes gaps where partial people leave early and no one covers the
        // remaining time. Prioritizes connecting to the next shift's people.
        // ===================================================================
        private void FillPartialCoverageGaps(List<Shift> shifts, ScheduleConfig config)
        {
            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;

            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();

                for (int i = 0; i < dayShifts.Count; i++)
                {
                    var shift = dayShifts[i];

                    // Check if this shift has partials that leave early
                    foreach (var person in shift.AssignedPeople.ToList())
                    {
                        if (!shift.PersonAssignments.ContainsKey(person)) continue;
                        var pa = shift.PersonAssignments[person];
                        if (!pa.EndsAt.HasValue) continue; // Person stays till end

                        TimeSpan gapStart = pa.EndsAt.Value;
                        TimeSpan gapEnd = shift.End;

                        // Gap must be at least 30 min to matter
                        if ((gapEnd - gapStart).TotalMinutes < 30) continue;

                        // Strategy 1: Try people from the NEXT shift (they can start earlier)
                        if (i + 1 < dayShifts.Count)
                        {
                            var nextShift = dayShifts[i + 1];
                            foreach (var nextPerson in nextShift.AssignedPeople)
                            {
                                if (shift.AssignedPeople.Contains(nextPerson)) continue;

                                // Check if nextPerson is available for the gap portion
                                if (IsPartiallyAvailable(nextPerson, shift))
                                {
                                    var assignment = new PersonAssignment
                                    {
                                        Name = nextPerson,
                                        StartsAt = gapStart,
                                        EndsAt = null // Stays till end of this shift
                                    };

                                    // Verify actual availability
                                    var avail = _allAvailabilities.FirstOrDefault(a =>
                                        a.Name == nextPerson &&
                                        a.Date.Date == shift.Date.Date &&
                                        a.Start <= gapStart &&
                                        a.End >= gapEnd);

                                    if (avail != null && !shift.AssignedPeople.Contains(nextPerson) && !WouldCreateGapOnSameDay(nextPerson, shift))
                                    {
                                        AssignPartial(nextPerson, shift, assignment);
                                        break;
                                    }
                                }
                            }
                        }

                        // Strategy 2: If still unfilled, try anyone available for the gap
                        if (shift.AssignedPeople.Count < shift.PeopleNeeded)
                        {
                            var filler = _availableHours.Keys
                                .Where(p => !shift.AssignedPeople.Contains(p))
                                .Where(p => _allAvailabilities.Any(a =>
                                    a.Name == p &&
                                    a.Date.Date == shift.Date.Date &&
                                    a.Start <= gapStart &&
                                    a.End >= gapEnd))
                                .Where(p => !WouldCreateGapOnSameDay(p, shift))
                                .Where(p => !HasOverlap(p, shift))
                                .OrderBy(p => GetWeeklyHours(p))
                                .FirstOrDefault();

                            if (filler != null)
                            {
                                double gapHours = (gapEnd - gapStart).TotalHours;
                                double currentDaily = GetDailyAssignedHours(filler, shift.Date);
                                double projectedDaily = currentDaily + gapHours;

                                // RULE: Only fill gap if it brings the person's DAILY session to at least 1 full shift length (1.5h)
                                if (projectedDaily >= config.ShiftLength.TotalHours - 0.01)
                                {
                                    var assignment = new PersonAssignment
                                    {
                                        Name = filler,
                                        StartsAt = gapStart,
                                        EndsAt = null
                                    };
                                    AssignPartial(filler, shift, assignment);
                                }
                            }
                        }
                    }
                }
            }
        }

        // ===================================================================
        // PHASE 10: Half-Hour Understaffing Scan & Fix
        // Walks every half-hour across the schedule's open hours and verifies
        // the shift has enough people. Tries to fix understaffed windows with
        // available candidates. Returns alerts for any unresolved windows.
        // ===================================================================
        private List<UnderstaffingAlert> CheckAndFixUnderstaffing(List<Shift> shifts, ScheduleConfig config)
        {
            var alerts = new List<UnderstaffingAlert>();
            var shiftsByDate = shifts.GroupBy(s => s.Date.Date);

            foreach (var dayGroup in shiftsByDate)
            {
                var dayShifts = dayGroup.OrderBy(s => s.Start).ToList();
                if (dayShifts.Count == 0) continue;

                var date = dayGroup.Key;
                var dayStart = config.OpeningTime;
                var dayEnd = config.ClosingTime;

                // Walk every half-hour
                var current = dayStart;
                while (current < dayEnd)
                {
                    var slotEnd = current + TimeSpan.FromMinutes(30);
                    if (slotEnd > dayEnd) slotEnd = dayEnd;

                    // Count people ACTUALLY PRESENT during this slot. A partial
                    // assignment (starts/ends mid-shift) only counts for the time
                    // the person is really there, so gaps left by early-leavers
                    // are detected correctly.
                    int actual = CountPeoplePresent(dayShifts, current, slotEnd);

                    int required = config.PeoplePerShift;

                    if (actual < required)
                    {
                        // Find the contiguous understaffed window
                        var windowStart = current;
                        var windowEnd = slotEnd;

                        // Extend forward while still understaffed
                        var probe = slotEnd;
                        while (probe < dayEnd)
                        {
                            var probeEnd = probe + TimeSpan.FromMinutes(30);
                            if (probeEnd > dayEnd) probeEnd = dayEnd;

                            int probeActual = CountPeoplePresent(dayShifts, probe, probeEnd);

                            if (probeActual < required)
                            {
                                windowEnd = probeEnd;
                                probe = probeEnd;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // Try to fix this window
                        bool wasFixed = TryFixUnderstaffedWindow(dayShifts, date, windowStart, windowEnd, required, config);

                        var alert = new UnderstaffingAlert
                        {
                            Date = date,
                            Start = windowStart,
                            End = windowEnd,
                            Required = required,
                            Actual = actual,
                            WasFixed = wasFixed
                        };

                        // Collect candidates who could cover this window
                        foreach (var person in _availableHours.Keys)
                        {
                            if (CanCoverWindow(person, date, windowStart, windowEnd))
                            {
                                var avail = _allAvailabilities.FirstOrDefault(a =>
                                    a.Name == person &&
                                    a.Date.Date == date.Date &&
                                    a.Start < windowEnd &&
                                    a.End > windowStart);

                                if (avail != null)
                                {
                                    alert.Candidates.Add(new AlertCandidate
                                    {
                                        Name = person,
                                        AvailableFrom = avail.Start < windowStart ? (TimeSpan?)null : avail.Start,
                                        AvailableTo = avail.End > windowEnd ? (TimeSpan?)null : avail.End
                                    });
                                }
                            }
                        }

                        alerts.Add(alert);

                        // Skip past this window
                        current = windowEnd;
                        continue;
                    }

                    current = slotEnd;
                }
            }

            return alerts;
        }

        /// <summary>
        /// Counts how many people are ACTUALLY present in the given time window.
        /// Partial assignments (with StartsAt/EndsAt) only count for the exact
        /// time they are really working, so early-leaver gaps are detected.
        /// </summary>
        private int CountPeoplePresent(List<Shift> dayShifts, TimeSpan from, TimeSpan to)
        {
            int count = 0;
            foreach (var shift in dayShifts)
            {
                if (shift.Start >= to || shift.End <= from) continue;

                foreach (var person in shift.AssignedPeople)
                {
                    TimeSpan presentStart = shift.Start;
                    TimeSpan presentEnd = shift.End;

                    if (shift.PersonAssignments.ContainsKey(person))
                    {
                        var pa = shift.PersonAssignments[person];
                        if (pa.StartsAt.HasValue) presentStart = pa.StartsAt.Value;
                        if (pa.EndsAt.HasValue) presentEnd = pa.EndsAt.Value;
                    }

                    // Person is present for the slot if their coverage overlaps it
                    if (presentStart < to && presentEnd > from)
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// Attempts to fix an understaffed window by finding a candidate who can
        /// cover a 1.5h shift that fully contains the window.
        /// </summary>
        private bool TryFixUnderstaffedWindow(
            List<Shift> dayShifts,
            DateTime date,
            TimeSpan windowStart,
            TimeSpan windowEnd,
            int required,
            ScheduleConfig config)
        {
            double shiftLen = config.ShiftLength.TotalHours;

            // Find the shift(s) covering this window
            var coveringShifts = dayShifts
                .Where(s => s.Start <= windowStart && s.End >= windowEnd)
                .ToList();

            if (coveringShifts.Count == 0) return false;

            // Try to add people to the covering shift(s).
            // IMPORTANT: Use ACTUAL presence (CountPeoplePresent) rather than the raw
            // AssignedPeople.Count, because partial people (starts/ends mid-shift) only
            // cover part of the window.
            bool anyFixed = false;
            foreach (var shift in coveringShifts)
            {
                int presentInWindow = CountPeoplePresent(dayShifts, windowStart, windowEnd);
                if (presentInWindow >= required) continue;

                int needed = required - presentInWindow;

                // Find candidates who can cover a full 1.5h shift containing the window
                var candidates = _availableHours.Keys
                    .Where(p => !shift.AssignedPeople.Contains(p))
                    .Where(p => CanCoverFullShift(p, date, windowStart, windowEnd, shiftLen))
                    .Where(p => CanAssign(p, shift, config))
                    .Where(p => !WouldCreateGapOnSameDay(p, shift))
                    .OrderBy(p => GetWeeklyHours(p))
                    .ThenBy(p => _availableHours[p])
                    .Take(needed)
                    .ToList();

                foreach (var candidate in candidates)
                {
                    if (CountPeoplePresent(dayShifts, windowStart, windowEnd) >= required) break;

                    // Check daily cap
                    double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
                    double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;
                    double currentDaily = GetDailyAssignedHours(candidate, date);
                    double newDaily = currentDaily + shift.DurationHours;

                    if (newDaily > maxDailyHours && currentDaily >= maxDailyHours * 0.3)
                        continue;

                    Assign(candidate, shift);
                    anyFixed = true;
                }
            }

            return anyFixed;
        }

        /// <summary>
        /// Checks if a person can cover a full 1.5h shift that fully contains the
        /// understaffed window. The shift can start anywhere from (windowEnd - 1.5h)
        /// to windowStart, as long as it fully covers the window.
        /// </summary>
        private bool CanCoverFullShift(string person, DateTime date, TimeSpan windowStart, TimeSpan windowEnd, double shiftLen)
        {
            // The shift must be exactly shiftLen (1.5h) and fully contain the window.
            // Valid shift starts: [windowEnd - shiftLen, windowStart]
            var earliestStart = windowEnd - TimeSpan.FromHours(shiftLen);
            var latestStart = windowStart;

            // Check if the person is available for any valid 1.5h shift covering the window
            return _allAvailabilities.Any(a =>
                a.Name == person &&
                a.Date.Date == date.Date &&
                a.End - a.Start >= TimeSpan.FromHours(shiftLen) &&
                a.Start <= latestStart &&
                a.End >= earliestStart + TimeSpan.FromHours(shiftLen));
        }

        /// <summary>
        /// Checks if a person has ANY availability overlapping the window (for reporting).
        /// </summary>
        private bool CanCoverWindow(string person, DateTime date, TimeSpan windowStart, TimeSpan windowEnd)
        {
            return _allAvailabilities.Any(a =>
                a.Name == person &&
                a.Date.Date == date.Date &&
                a.Start < windowEnd &&
                a.End > windowStart);
        }

        // ===================================================================
        // PHASE 11: Overwork/Underwork Balance Exchange
        // Transfers full shifts from people above the weekly average to those
        // below it. Handles stacked blocks (full+partial / partial+full /
        // partial+partial) by transferring the entire block together.
        // ===================================================================
        private void BalanceWorkload(List<Shift> shifts, ScheduleConfig config)
        {
            var weeklyHours = CalculateAllWeeklyHours();
            if (weeklyHours.Count == 0) return;

            double average = weeklyHours.Values.Sum() / weeklyHours.Count;

            // Sort overworked descending (most hours first)
            var overworked = weeklyHours
                .Where(kvp => kvp.Value > average + 0.01)
                .OrderByDescending(kvp => kvp.Value)
                .Select(kvp => kvp.Key)
                .ToList();

            // Sort underworked ascending (least hours first)
            var underworked = weeklyHours
                .Where(kvp => kvp.Value < average - 0.01)
                .OrderBy(kvp => kvp.Value)
                .Select(kvp => kvp.Key)
                .ToList();

            if (overworked.Count == 0 || underworked.Count == 0)
            {
                return;
            }

            double totalOpeningHours = (config.ClosingTime - config.OpeningTime).TotalHours;
            double maxDailyHours = totalOpeningHours * MaxDailyHoursRatio;

            // For each overworked person (most first), try to transfer shifts to underworked people
            foreach (var overPerson in overworked)
            {
                if (GetWeeklyHours(overPerson) <= average + 0.01) continue;

                // Find this person's shifts, grouped by day
                if (!_personShifts.ContainsKey(overPerson)) continue;

                foreach (var dateEntry in _personShifts[overPerson].ToList())
                {
                    var date = dateEntry.Key;
                    var dayShifts = dateEntry.Value.OrderBy(s => s.Start).ToList();
                    if (dayShifts.Count == 0) continue;

                    // Build the stacked block: if the last shift is partial and sticks to the
                    // previous shift, treat them as one block. Also handle partial+full and
                    // partial+partial combinations.
                    var blocks = new List<List<Shift>>();
                    var currentBlock = new List<Shift> { dayShifts[0] };

                    for (int i = 1; i < dayShifts.Count; i++)
                    {
                        if (dayShifts[i].Start == dayShifts[i - 1].End)
                        {
                            // Adjacent - same block
                            currentBlock.Add(dayShifts[i]);
                        }
                        else
                        {
                            blocks.Add(currentBlock);
                            currentBlock = new List<Shift> { dayShifts[i] };
                        }
                    }
                    blocks.Add(currentBlock);

                    // Try to transfer each block to an underworked person
                    foreach (var block in blocks)
                    {
                        if (GetWeeklyHours(overPerson) <= average + 0.01) break;

                        // Don't transfer if it would leave the shift understaffed
                        if (block.Any(s => s.AssignedPeople.Count <= 1))
                            continue;

                        // Find an underworked candidate who can take the whole block
                        var candidate = underworked
                            .Where(p => p != overPerson)
                            .Where(p => GetWeeklyHours(p) < average - 0.01)
                            .Where(p => block.All(s => IsAvailable(p, s)))
                            .Where(p => block.All(s => CanAssign(p, s, config)))
                            .Where(p => !block.Any(s => WouldCreateGapOnSameDay(p, s)))
                            .OrderBy(p => GetWeeklyHours(p))
                            .FirstOrDefault();

                        if (candidate == null) continue;

                        // Check daily cap for the candidate
                        double blockHours = block.Sum(s => s.DurationHours);
                        double currentDaily = GetDailyAssignedHours(candidate, date);
                        double newDaily = currentDaily + blockHours;
                        if (newDaily > maxDailyHours && currentDaily >= maxDailyHours * 0.3)
                            continue;

                        // GRADUAL BALANCE GUARD: Don't let a transfer push the recipient
                        // above the weekly average. This stops "going all out" on one
                        // person. Instead, overworked people come down gradually and evenly.
                        double candidateWeekly = GetWeeklyHours(candidate);
                        double candidateAfter = candidateWeekly + blockHours;
                        double overAfter = GetWeeklyHours(overPerson) - blockHours;
                        if (candidateAfter > average + 0.01 && candidateAfter > overAfter + 0.01)
                            continue;

                        // Perform the transfer of the whole block
                        foreach (var shift in block)
                        {
                            // Remove overPerson
                            shift.AssignedPeople.Remove(overPerson);
                            if (shift.PersonAssignments.ContainsKey(overPerson))
                            {
                                shift.PersonAssignments.Remove(overPerson);
                            }

                            // Add candidate
                            shift.AssignedPeople.Add(candidate);

                            // Update tracking
                            _assignedHours[overPerson] -= shift.DurationHours;
                            _assignedHours[candidate] += shift.DurationHours;

                            if (_personShifts.ContainsKey(overPerson) && _personShifts[overPerson].ContainsKey(date))
                            {
                                _personShifts[overPerson][date].Remove(shift);
                            }

                            if (!_personShifts.ContainsKey(candidate))
                                _personShifts[candidate] = new Dictionary<DateTime, List<Shift>>();

                            if (!_personShifts[candidate].ContainsKey(date))
                                _personShifts[candidate][date] = new List<Shift>();

                            _personShifts[candidate][date].Add(shift);

                            // Update daily hours
                            var overKey = (overPerson, date);
                            var underKey = (candidate, date);
                            if (_dailyAssignedHours.ContainsKey(overKey))
                            {
                                _dailyAssignedHours[overKey] -= shift.DurationHours;
                            }
                            if (!_dailyAssignedHours.ContainsKey(underKey))
                                _dailyAssignedHours[underKey] = 0;
                            _dailyAssignedHours[underKey] += shift.DurationHours;
                        }
                    }
                }
            }

            // =================================================================
            // SECOND PASS: Backfill-aware transfers
            // Some overworked people couldn't transfer because removing them would
            // leave a shift understaffed. In this pass, we allow the transfer IF
            // we can find a backfill candidate to fill the vacated shift slot.
            // =================================================================
            foreach (var overPerson in overworked)
            {
                if (GetWeeklyHours(overPerson) <= average + 0.01) continue;
                if (!_personShifts.ContainsKey(overPerson)) continue;

                foreach (var dateEntry in _personShifts[overPerson].ToList())
                {
                    var date = dateEntry.Key;
                    var dayShifts = dateEntry.Value.OrderBy(s => s.Start).ToList();
                    if (dayShifts.Count == 0) continue;

                    // Build blocks (adjacent shifts grouped)
                    var blocks = new List<List<Shift>>();
                    var currentBlock = new List<Shift> { dayShifts[0] };
                    for (int i = 1; i < dayShifts.Count; i++)
                    {
                        if (dayShifts[i].Start == dayShifts[i - 1].End)
                            currentBlock.Add(dayShifts[i]);
                        else
                        {
                            blocks.Add(currentBlock);
                            currentBlock = new List<Shift> { dayShifts[i] };
                        }
                    }
                    blocks.Add(currentBlock);

                    foreach (var block in blocks)
                    {
                        if (GetWeeklyHours(overPerson) <= average + 0.01) break;

                        // Find an underworked candidate who can take the whole block
                        var candidate = underworked
                            .Where(p => p != overPerson)
                            .Where(p => GetWeeklyHours(p) < average - 0.01)
                            .Where(p => block.All(s => IsAvailable(p, s)))
                            .Where(p => block.All(s => CanAssign(p, s, config)))
                            .Where(p => !block.Any(s => WouldCreateGapOnSameDay(p, s)))
                            .OrderBy(p => GetWeeklyHours(p))
                            .FirstOrDefault();

                        if (candidate == null) continue;

                        // Check daily cap for the candidate
                        double blockHours = block.Sum(s => s.DurationHours);
                        double currentDaily = GetDailyAssignedHours(candidate, date);
                        double newDaily = currentDaily + blockHours;
                        if (newDaily > maxDailyHours && currentDaily >= maxDailyHours * 0.3)
                            continue;

                        // GRADUAL BALANCE GUARD: Don't let a transfer push the recipient
                        // above the weekly average (same rule as the first pass).
                        double candidateWeekly = GetWeeklyHours(candidate);
                        double candidateAfter = candidateWeekly + blockHours;
                        double overAfter = GetWeeklyHours(overPerson) - blockHours;
                        if (candidateAfter > average + 0.01 && candidateAfter > overAfter + 0.01)
                            continue;

                        // For each shift in the block, if removing overPerson would
                        // leave it understaffed, find a backfill candidate.
                        bool canBackfill = true;
                        var backfills = new Dictionary<Shift, string>();
                        foreach (var shift in block)
                        {
                            if (shift.AssignedPeople.Count - 1 >= shift.PeopleNeeded)
                                continue; // Still staffed after removal

                            // Need a backfill for this shift
                            var backfill = _availableHours.Keys
                                .Where(p => p != overPerson && p != candidate)
                                .Where(p => !shift.AssignedPeople.Contains(p))
                                .Where(p => IsAvailable(p, shift))
                                .Where(p => CanAssign(p, shift, config))
                                .Where(p => !WouldCreateGapOnSameDay(p, shift))
                                .OrderBy(p => GetWeeklyHours(p))
                                .FirstOrDefault();

                            if (backfill == null)
                            {
                                canBackfill = false;
                                break;
                            }
                            backfills[shift] = backfill;
                        }

                        if (!canBackfill) continue;

                        // Perform the transfer + backfills
                        foreach (var shift in block)
                        {
                            shift.AssignedPeople.Remove(overPerson);
                            if (shift.PersonAssignments.ContainsKey(overPerson))
                                shift.PersonAssignments.Remove(overPerson);

                            shift.AssignedPeople.Add(candidate);

                            _assignedHours[overPerson] -= shift.DurationHours;
                            _assignedHours[candidate] += shift.DurationHours;

                            if (_personShifts.ContainsKey(overPerson) && _personShifts[overPerson].ContainsKey(date))
                                _personShifts[overPerson][date].Remove(shift);

                            if (!_personShifts.ContainsKey(candidate))
                                _personShifts[candidate] = new Dictionary<DateTime, List<Shift>>();
                            if (!_personShifts[candidate].ContainsKey(date))
                                _personShifts[candidate][date] = new List<Shift>();
                            _personShifts[candidate][date].Add(shift);

                            var overKey = (overPerson, date);
                            var underKey = (candidate, date);
                            if (_dailyAssignedHours.ContainsKey(overKey))
                                _dailyAssignedHours[overKey] -= shift.DurationHours;
                            if (!_dailyAssignedHours.ContainsKey(underKey))
                                _dailyAssignedHours[underKey] = 0;
                            _dailyAssignedHours[underKey] += shift.DurationHours;

                            // Apply backfill if needed
                            if (backfills.ContainsKey(shift))
                            {
                                var backfill = backfills[shift];
                                shift.AssignedPeople.Add(backfill);
                                _assignedHours[backfill] += shift.DurationHours;

                                if (!_personShifts.ContainsKey(backfill))
                                    _personShifts[backfill] = new Dictionary<DateTime, List<Shift>>();
                                if (!_personShifts[backfill].ContainsKey(date))
                                    _personShifts[backfill][date] = new List<Shift>();
                                _personShifts[backfill][date].Add(shift);

                                var backKey = (backfill, date);
                                if (!_dailyAssignedHours.ContainsKey(backKey))
                                    _dailyAssignedHours[backKey] = 0;
                                _dailyAssignedHours[backKey] += shift.DurationHours;
                            }
                        }
                    }
                }
            }
        }
    }
}