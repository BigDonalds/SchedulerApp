using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SchedulerApp.Views
{
    public partial class SetupPage : UserControl
    {
        private AppState _state;

        public SetupPage()
        {
            InitializeComponent();
        }

        public void Initialize(AppState state)
        {
            _state = state;
            PopulateDropdowns();
            UpdateBatchComboBox();
        }

        public void UpdateBatchComboBox()
        {
            BatchComboBox.Items.Clear();
            if (_state.Batches.Count == 0)
            {
                BatchComboBox.Items.Add("No batches available");
                BatchComboBox.SelectedIndex = 0;
                BatchComboBox.IsEnabled = false;
                return;
            }
            BatchComboBox.IsEnabled = true;
            foreach (var batch in _state.Batches)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = $"{batch.Name} ({batch.Count} people)";
                item.Tag = batch.Id;
                BatchComboBox.Items.Add(item);
            }
            if (BatchComboBox.Items.Count > 0)
            {
                BatchComboBox.SelectedIndex = 0;
            }
        }

        private void PopulateDropdowns()
        {
            OpeningTimeBox.Items.Clear();
            ClosingTimeBox.Items.Clear();
            ShiftLengthBox.Items.Clear();
            PeoplePerShiftBox.Items.Clear();

            for (int hour = 6; hour <= 23; hour++)
            {
                string label;
                if (hour < 12)
                    label = hour + ":00 AM";
                else if (hour == 12)
                    label = "12:00 PM";
                else
                    label = (hour - 12) + ":00 PM";

                OpeningTimeBox.Items.Add(label);
                ClosingTimeBox.Items.Add(label);
            }

            OpeningTimeBox.SelectedIndex = 0;
            ClosingTimeBox.SelectedIndex = ClosingTimeBox.Items.Count - 1;

            for (double hours = 0.5; hours <= 12; hours += 0.5)
            {
                ShiftLengthBox.Items.Add($"{hours:F1} hours");
            }
            ShiftLengthBox.SelectedIndex = 7;

            for (int i = 1; i <= 4; i++)
            {
                PeoplePerShiftBox.Items.Add($"{i} person{(i > 1 ? "s" : "")}");
            }
            PeoplePerShiftBox.SelectedIndex = 0;

            OpeningTimeBox.SelectionChanged += UpdateShiftLengthOptions;
            ClosingTimeBox.SelectionChanged += UpdateShiftLengthOptions;
        }

        private void UpdateShiftLengthOptions(object sender, SelectionChangedEventArgs e)
        {
            if (OpeningTimeBox.SelectedIndex == -1 || ClosingTimeBox.SelectedIndex == -1)
                return;

            int openingHour = 6 + OpeningTimeBox.SelectedIndex;
            int closingHour = 6 + ClosingTimeBox.SelectedIndex;
            int totalHours = closingHour - openingHour;

            if (totalHours <= 0)
                return;

            ShiftLengthBox.Items.Clear();

            for (double hours = 0.5; hours <= totalHours; hours += 0.5)
            {
                ShiftLengthBox.Items.Add($"{hours:F1} hours");
            }

            if (ShiftLengthBox.Items.Count > 0)
            {
                ShiftLengthBox.SelectedIndex = Math.Min(ShiftLengthBox.Items.Count - 1, 7);
            }
        }

        private void BatchComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private List<AvailabilityEntry> FilterBatchByTimeInterval(
            Batch batch,
            int startHour,
            int endHour)
        {
            var result = new List<AvailabilityEntry>();

            foreach (var availability in batch.EmployeeData)
            {
                bool dateOverlap = availability.EndDate >= batch.StartDate &&
                                  availability.StartDate <= batch.EndDate;

                if (!dateOverlap)
                {
                    continue;
                }

                bool timeOverlap = availability.StartHour < endHour &&
                                  availability.EndHour > startHour;

                if (!timeOverlap)
                {
                    continue;
                }

                result.Add(availability);
            }

            return result;
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            if (OpeningTimeBox.SelectedIndex == -1 || ClosingTimeBox.SelectedIndex == -1 ||
                ShiftLengthBox.SelectedIndex == -1 || PeoplePerShiftBox.SelectedIndex == -1)
            {
                _state?.ShowNotification?.Invoke("Please fill in all settings before generating the schedule.",
                    "Incomplete Settings", "warning", 4000);
                return;
            }

            if (BatchComboBox.SelectedIndex == -1 || _state.Batches.Count == 0)
            {
                _state?.ShowNotification?.Invoke("Please create and select a batch to use for scheduling.",
                    "No Batch Selected", "warning", 4000);
                return;
            }

            var selectedBatch = GetSelectedBatch();
            if (selectedBatch == null) return;

            string shiftLengthStr = ShiftLengthBox.SelectedItem?.ToString() ?? "1.0 hours";

            if (shiftLengthStr.EndsWith("hours"))
            {
                shiftLengthStr = shiftLengthStr.Replace("hours", "").Trim();
            }
            double shiftLengthHours = double.Parse(shiftLengthStr);
            int openingHourIndex = OpeningTimeBox.SelectedIndex + 6;
            int closingHourIndex = ClosingTimeBox.SelectedIndex + 6;
            double totalHours = closingHourIndex - openingHourIndex;
            int shiftIntervals = (int)Math.Ceiling(totalHours / shiftLengthHours);
            bool includeWeekends = IncludeWeekendsCheckBox.IsChecked ?? false;
            int currentPeoplePerShift = PeoplePerShiftBox.SelectedIndex + 1;
            var filteredBatchEmployees = FilterBatchByTimeInterval(selectedBatch, openingHourIndex, closingHourIndex);

            if (filteredBatchEmployees.Count == 0)
            {
                _state?.ShowNotification?.Invoke("No employees in the batch are available within the selected time interval.",
                    "No Available Employees", "warning", 4000);
                return;
            }

            var schedulerAvailabilities = ConvertToSchedulerFormat(filteredBatchEmployees, selectedBatch, includeWeekends);
            var schedulerConfig = new ScheduleConfig
            {
                OpeningTime = TimeSpan.FromHours(openingHourIndex),
                ClosingTime = TimeSpan.FromHours(closingHourIndex),
                ShiftLength = TimeSpan.FromHours(shiftLengthHours),
                PeoplePerShift = currentPeoplePerShift,
                ClosedDays = includeWeekends ? new HashSet<DayOfWeek>() :
                    new HashSet<DayOfWeek> { DayOfWeek.Saturday, DayOfWeek.Sunday }
            };

            var scheduler = new Scheduler();
            var scheduleResult = scheduler.GenerateSchedule(schedulerAvailabilities, schedulerConfig);
            string scheduleName = $"Schedule from {selectedBatch.Name}";

            int count = 1;
            string baseName = scheduleName;
            while (_state.Schedules.Any(s => s.Name.Equals(scheduleName, StringComparison.OrdinalIgnoreCase)))
            {
                scheduleName = $"{baseName} ({count})";
                count++;
            }

            DateTime scheduleStartDate = selectedBatch.StartDate;
            DateTime scheduleEndDate = selectedBatch.EndDate;

            if (!includeWeekends)
            {
                int weekdays = 0;
                DateTime currentDate = scheduleStartDate;
                List<DateTime> validDates = new List<DateTime>();

                while (currentDate <= scheduleEndDate)
                {
                    if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
                    {
                        weekdays++;
                        validDates.Add(currentDate);
                    }
                    currentDate = currentDate.AddDays(1);
                }

                if (weekdays == 0)
                {
                    _state?.ShowNotification?.Invoke("The selected batch contains no weekdays. Please include weekends or select a different batch.",
                        "No Weekdays Available", "warning", 4000);
                    return;
                }

                scheduleStartDate = validDates.First();
                scheduleEndDate = validDates.Last();
            }

            var schedule = new Schedule
            {
                Id = Guid.NewGuid().ToString(),
                Name = scheduleName,
                CreatedDate = DateTime.Now,
                BatchId = selectedBatch.Id,
                Status = scheduleResult.HasUnfilledShifts ? "Incomplete" : "Complete",
                StartDate = scheduleStartDate,
                EndDate = scheduleEndDate,
                OpeningHour = openingHourIndex,
                ClosingHour = closingHourIndex,
                ShiftLengthHours = shiftLengthHours,
                ShiftIntervals = shiftIntervals,
                IncludeWeekends = includeWeekends,
                PeoplePerShift = currentPeoplePerShift,
                OriginalDayCount = GetScheduleDayCountFromDates(scheduleStartDate, scheduleEndDate, includeWeekends),
                OriginalShiftIntervals = shiftIntervals,
                UnderstaffingAlerts = scheduleResult.UnderstaffingAlerts ?? new List<UnderstaffingAlert>()
            };

            InitializeCellAssignmentsFromShifts(schedule, scheduleResult.Shifts, scheduleStartDate, scheduleEndDate, includeWeekends);
            _state.Schedules.Add(schedule);
            _state.RefreshScheduleList?.Invoke();
            _state.StateSave.SaveSchedule(schedule);
            _state?.ShowNotification?.Invoke($"Schedule '{scheduleName}' created with {scheduleResult.Shifts.Count(s => s.AssignedPeople.Count > 0)} filled shifts", "Schedule Generated", "success", 4000);
            _state?.NavigateRequested?.Invoke("Schedule");
        }

        private int GetScheduleDayCountFromDates(DateTime startDate, DateTime endDate, bool includeWeekends)
        {
            int dayCount = 0;
            DateTime currentDate = startDate;

            while (currentDate <= endDate)
            {
                if (includeWeekends ||
                    (currentDate.DayOfWeek != DayOfWeek.Saturday &&
                     currentDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    dayCount++;
                }
                currentDate = currentDate.AddDays(1);
            }

            return dayCount;
        }

        private void InitializeCellAssignmentsFromShifts(Schedule schedule, List<Shift> shifts, DateTime startDate, DateTime endDate, bool includeWeekends)
        {
            var dateMapping = new Dictionary<int, DateTime>();
            int dayIndex = 0;
            DateTime currentDate = startDate;

            while (currentDate <= endDate)
            {
                if (includeWeekends || (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    dateMapping[dayIndex] = currentDate;
                    dayIndex++;
                }
                currentDate = currentDate.AddDays(1);
            }

            var shiftsByDate = shifts
                .GroupBy(s => s.Date.Date)
                .ToDictionary(g => g.Key, g => g.OrderBy(s => s.Start).ToList());

            for (int dayIdx = 0; dayIdx < dateMapping.Count; dayIdx++)
            {
                var date = dateMapping[dayIdx];
                var dateShifts = shiftsByDate.ContainsKey(date.Date) ?
                    shiftsByDate[date.Date] : new List<Shift>();

                int shiftIndex = 0;
                foreach (var shift in dateShifts)
                {
                    string cellId = $"cell_{dayIdx}_{shiftIndex}";
                    var cellAssignments = new List<string>();
                    foreach (var person in shift.AssignedPeople)
                    {
                        if (shift.PersonAssignments.ContainsKey(person))
                        {
                            cellAssignments.Add(shift.PersonAssignments[person].ToString());
                        }
                        else
                        {
                            cellAssignments.Add(person);
                        }
                    }
                    schedule.CellAssignments[cellId] = cellAssignments;
                    shiftIndex++;
                }

                while (shiftIndex < schedule.ShiftIntervals)
                {
                    string cellId = $"cell_{dayIdx}_{shiftIndex}";
                    schedule.CellAssignments[cellId] = new List<string>();
                    shiftIndex++;
                }
            }
        }

        private List<PersonAvailability> ConvertToSchedulerFormat(
            List<AvailabilityEntry> availabilities,
            Batch batch,
            bool includeWeekends)
        {
            var schedulerAvailabilities = new List<PersonAvailability>();
            int openingHourIndex = OpeningTimeBox.SelectedIndex + 6;
            int closingHourIndex = ClosingTimeBox.SelectedIndex + 6;

            List<DateTime> batchDates = new List<DateTime>();
            DateTime currentBatchDate = batch.StartDate.Date;
            while (currentBatchDate <= batch.EndDate.Date)
            {
                if (includeWeekends ||
                    (currentBatchDate.DayOfWeek != DayOfWeek.Saturday &&
                     currentBatchDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    batchDates.Add(currentBatchDate);
                }
                currentBatchDate = currentBatchDate.AddDays(1);
            }

            foreach (var availability in availabilities)
            {
                if (availability.ScheduleMatrix == null ||
                    availability.SelectedDates == null ||
                    availability.SelectedDates.Count == 0)
                {
                    foreach (var batchDate in batchDates)
                    {
                        if (batchDate >= availability.StartDate && batchDate <= availability.EndDate)
                        {
                            int start = Math.Max(availability.StartHour, openingHourIndex);
                            int end = Math.Min(availability.EndHour, closingHourIndex);

                            if (start < end)
                            {
                                schedulerAvailabilities.Add(new PersonAvailability
                                {
                                    Name = availability.Name,
                                    Date = batchDate,
                                    Start = TimeSpan.FromHours(start),
                                    End = TimeSpan.FromHours(end)
                                });
                            }
                        }
                    }
                    continue;
                }

                List<DateTime> fullDateRange = new List<DateTime>();
                DateTime currentDate = availability.StartDate.Date;
                while (currentDate <= availability.EndDate.Date)
                {
                    fullDateRange.Add(currentDate);
                    currentDate = currentDate.AddDays(1);
                }

                bool matrixMatchesFullRange = availability.ScheduleMatrix.GetLength(0) == fullDateRange.Count;
                var dateToMatrixIndex = new Dictionary<DateTime, int>();

                if (matrixMatchesFullRange)
                {
                    for (int i = 0; i < fullDateRange.Count; i++)
                    {
                        dateToMatrixIndex[fullDateRange[i]] = i;
                    }
                }
                else
                {
                    var allSelectedDates = availability.SelectedDates
                        .Select(d => d.Date)
                        .Distinct()
                        .OrderBy(d => d)
                        .ToList();

                    for (int i = 0; i < allSelectedDates.Count; i++)
                    {
                        dateToMatrixIndex[allSelectedDates[i]] = i;
                    }
                }

                foreach (var batchDate in batchDates)
                {
                    if (!dateToMatrixIndex.ContainsKey(batchDate))
                    {
                        continue;
                    }

                    int matrixRowIndex = dateToMatrixIndex[batchDate];

                    if (matrixRowIndex >= availability.ScheduleMatrix.GetLength(0))
                    {
                        continue;
                    }

                    int totalMatrixColumns = availability.ScheduleMatrix.GetLength(1);
                    int col = 0;

                    while (col < totalMatrixColumns)
                    {
                        if (!availability.ScheduleMatrix[matrixRowIndex, col])
                        {
                            col++;
                            continue;
                        }

                        int startCol = col;

                        while (col < totalMatrixColumns &&
                               availability.ScheduleMatrix[matrixRowIndex, col])
                        {
                            col++;
                        }

                        int endCol = col;

                        TimeSpan start =
                            TimeSpan.FromHours(availability.StartHour)
                            + TimeSpan.FromMinutes(30 * startCol);

                        TimeSpan end =
                            TimeSpan.FromHours(availability.StartHour)
                            + TimeSpan.FromMinutes(30 * endCol);

                        bool timeOverlap = start < TimeSpan.FromHours(closingHourIndex) &&
                                          end > TimeSpan.FromHours(openingHourIndex);

                        if (timeOverlap)
                        {
                            var clippedStart = start < TimeSpan.FromHours(openingHourIndex)
                                ? TimeSpan.FromHours(openingHourIndex)
                                : start;

                            var clippedEnd = end > TimeSpan.FromHours(closingHourIndex)
                                ? TimeSpan.FromHours(closingHourIndex)
                                : end;

                            if (clippedEnd > clippedStart)
                            {
                                schedulerAvailabilities.Add(new PersonAvailability
                                {
                                    Name = availability.Name,
                                    Date = batchDate,
                                    Start = clippedStart,
                                    End = clippedEnd
                                });
                            }
                        }
                    }
                }
            }

            return schedulerAvailabilities;
        }

        private Batch GetSelectedBatch()
        {
            if (BatchComboBox.SelectedItem is ComboBoxItem selectedItem && selectedItem.Tag is string batchId)
            {
                return _state.Batches.FirstOrDefault(b => b.Id == batchId);
            }
            return null;
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _state?.NavigateRequested?.Invoke("Back");
        }
    }
}
