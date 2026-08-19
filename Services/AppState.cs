using System;
using System.Collections.Generic;
using System.Windows.Media;

namespace SchedulerApp.Services
{
    /// <summary>
    /// Central shared state for the application. Passed to each page UserControl
    /// so pages can read/write common data without referencing each other directly.
    /// </summary>
    public class AppState
    {
        // Data collections
        public List<AvailabilityEntry> CombinedAvailabilities { get; set; } = new List<AvailabilityEntry>();
        public List<Batch> Batches { get; set; } = new List<Batch>();
        public List<Schedule> Schedules { get; set; } = new List<Schedule>();
        public List<Employee> ManualEmployees { get; set; } = new List<Employee>();

        // Current selection / context
        public Schedule CurrentSchedule { get; set; }
        public string SelectedCellId { get; set; }
        public int CurrentPeoplePerShift { get; set; } = 1;
        public bool IncludeWeekends { get; set; } = false;

        // Persistence
        public StateSave StateSave { get; set; }

        // Navigation callback - pages call this to request navigation
        public Action<string> NavigateRequested { get; set; }

        // Popup callback - pages call this to show notifications
        public Func<string, string, string, int, bool> ShowNotification { get; set; }
        public Func<string, string, System.Threading.Tasks.Task<bool>> ShowConfirmDialog { get; set; }
        public Func<string, string, string, System.Threading.Tasks.Task<string>> ShowInputDialog { get; set; }
        public Func<string, string, List<string>, System.Threading.Tasks.Task<string>> ShowSelectionDialog { get; set; }

        // Refresh callbacks - pages register these so MainWindow can trigger refreshes
        public Action RefreshStatistics { get; set; }
        public Action RefreshCombinedAvailabilities { get; set; }
        public Action RefreshBatchList { get; set; }
        public Action RefreshBatchComboBox { get; set; }
        public Action RefreshScheduleList { get; set; }
        public Action RefreshEmployeesList { get; set; }
        public Action ShowScheduleListView { get; set; }
        public Action ShowExportPage { get; set; }

        // Export callback
        public Action<Schedule> AddScheduleToExport { get; set; }
    }

    // Data model classes shared across pages
    public class Employee
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string DateRange { get; set; }
        public string TimeRange { get; set; }
        public string AvailabilitySummary { get; set; }
        public int SlotCount { get; set; }
        public string Source { get; set; } = "Manual";
        public Color SourceColor { get; set; } = Colors.Blue;
        public bool IsSelected { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int StartHour { get; set; }
        public int EndHour { get; set; }
    }

    public class AvailabilityEntry : Employee
    {
        public DateTime CreatedDate { get; set; }
        public string BatchId { get; set; }
        public bool[,] ScheduleMatrix { get; set; }
        public List<DateTime> SelectedDates { get; set; } = new List<DateTime>();
    }

    public class Batch
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedDate { get; set; }
        public List<string> EmployeeIds { get; set; } = new List<string>();
        public List<AvailabilityEntry> EmployeeData { get; set; } = new List<AvailabilityEntry>();
        public int Count => EmployeeIds.Count;
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int StartHour { get; set; }
        public int EndHour { get; set; }
    }

    public class Schedule
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedDate { get; set; }
        public string BatchId { get; set; }
        public string Status { get; set; } = "Active";
        public Dictionary<string, List<string>> CellAssignments { get; set; } = new Dictionary<string, List<string>>();
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int OpeningHour { get; set; }
        public int ClosingHour { get; set; }
        public double ShiftLengthHours { get; set; }
        public int ShiftIntervals { get; set; }
        public bool IncludeWeekends { get; set; } = false;
        public int PeoplePerShift { get; set; } = 1;
        public int OriginalDayCount { get; set; }
        public int OriginalShiftIntervals { get; set; }
        public List<UnderstaffingAlert> UnderstaffingAlerts { get; set; } = new List<UnderstaffingAlert>();
    }

    /// <summary>
    /// Represents a period of time on a given day where the schedule is
    /// understaffed (fewer people than required) or completely uncovered.
    /// </summary>
    public class UnderstaffingAlert
    {
        public DateTime Date { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public int Required { get; set; }
        public int Actual { get; set; }
        public bool IsUncovered => Actual <= 0;
        public bool WasFixed { get; set; }
        /// <summary>Candidate name -> available window start (or null if fully available).</summary>
        public List<AlertCandidate> Candidates { get; set; } = new List<AlertCandidate>();
    }

    public class AlertCandidate
    {
        public string Name { get; set; }
        public TimeSpan? AvailableFrom { get; set; }
        public TimeSpan? AvailableTo { get; set; }
    }
}
