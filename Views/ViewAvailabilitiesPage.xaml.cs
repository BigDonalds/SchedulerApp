using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SchedulerApp.Views
{
    public partial class ViewAvailabilitiesPage : UserControl
    {
        private AppState _state;

        public ViewAvailabilitiesPage()
        {
            InitializeComponent();
        }

        public void Initialize(AppState state)
        {
            _state = state;
        }

        public void UpdateStatistics()
        {
            int totalPeople = _state.CombinedAvailabilities.Count;
            int manualEntries = _state.CombinedAvailabilities.Count(e => e.Source == "Manual");
            int importedData = _state.CombinedAvailabilities.Count(e => e.Source != "Manual");
            TotalPeopleText.Text = totalPeople.ToString();
            ManualEntriesText.Text = manualEntries.ToString();
            ImportedDataText.Text = importedData.ToString();

            if (_state.CombinedAvailabilities.Count > 0)
            {
                var allStartDates = _state.CombinedAvailabilities.Select(e => e.StartDate).ToList();
                var allEndDates = _state.CombinedAvailabilities.Select(e => e.EndDate).ToList();
                DateTime earliestDate = allStartDates.Min();
                DateTime latestDate = allEndDates.Max();
                DateRangeText.Text = $"{earliestDate:MMM dd} to {latestDate:MMM dd}";
            }
            else
            {
                DateRangeText.Text = "No data";
            }
        }

        public void UpdateBatchList()
        {
            BatchListControl.ItemsSource = null;
            BatchListControl.ItemsSource = _state.Batches;
            NoBatchesText.Visibility = _state.Batches.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public void UpdateCombinedAvailabilities()
        {
            CombinedAvailabilitiesList.ItemsSource = null;
            CombinedAvailabilitiesList.ItemsSource = _state.CombinedAvailabilities;
            NoCombinedAvailabilitiesText.Visibility = _state.CombinedAvailabilities.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void SaveBatchButton_Click(object sender, RoutedEventArgs e)
        {
            string batchName = BatchNameTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(batchName))
            {
                _state?.ShowNotification?.Invoke("Please enter a batch name.", "Missing Batch Name", "warning", 4000);
                return;
            }

            if (_state.CombinedAvailabilities.Count == 0)
            {
                _state?.ShowNotification?.Invoke("No availabilities to save as batch.", "No Data", "warning", 4000);
                return;
            }

            if (_state.Batches.Any(b => b.Name.Equals(batchName, StringComparison.OrdinalIgnoreCase)))
            {
                _state?.ShowNotification?.Invoke("A batch with this name already exists. Please choose a different name.",
                    "Duplicate Batch Name", "warning", 4000);
                return;
            }

            var selectedEmployees = _state.CombinedAvailabilities.Where(e => e.IsSelected).ToList();
            var employeesToAdd = selectedEmployees.Count > 0 ? selectedEmployees : _state.CombinedAvailabilities;
            DateTime minDate = employeesToAdd.Min(e => e.StartDate);
            DateTime maxDate = employeesToAdd.Max(e => e.EndDate);
            int minStartHour = employeesToAdd.Min(e => e.StartHour);
            int maxEndHour = employeesToAdd.Max(e => e.EndHour);
            var employeeCopies = new List<AvailabilityEntry>();

            foreach (var employee in employeesToAdd)
            {
                var copy = new AvailabilityEntry
                {
                    Id = employee.Id,
                    Name = employee.Name,
                    DateRange = employee.DateRange,
                    TimeRange = employee.TimeRange,
                    AvailabilitySummary = employee.AvailabilitySummary,
                    SlotCount = employee.SlotCount,
                    Source = employee.Source,
                    SourceColor = employee.SourceColor,
                    CreatedDate = employee.CreatedDate,
                    BatchId = Guid.NewGuid().ToString(),
                    StartDate = employee.StartDate,
                    EndDate = employee.EndDate,
                    StartHour = employee.StartHour,
                    EndHour = employee.EndHour,
                    ScheduleMatrix = employee.ScheduleMatrix != null ?
                        (bool[,])employee.ScheduleMatrix.Clone() : null,
                    SelectedDates = employee.SelectedDates != null ?
                        new List<DateTime>(employee.SelectedDates) : new List<DateTime>()
                };
                employeeCopies.Add(copy);
            }

            var batch = new Batch
            {
                Id = Guid.NewGuid().ToString(),
                Name = batchName,
                CreatedDate = DateTime.Now,
                EmployeeIds = employeeCopies.Select(e => e.Id).ToList(),
                EmployeeData = employeeCopies,
                StartDate = minDate,
                EndDate = maxDate,
                StartHour = minStartHour,
                EndHour = maxEndHour
            };

            _state.Batches.Add(batch);
            UpdateBatchList();
            _state.StateSave.SaveBatch(batch);
            BatchNameTextBox.Text = "";
            _state?.ShowNotification?.Invoke($"Saved '{batchName}' batch with {batch.Count} employees", "Batch Saved", "success", 4000);
            _state?.RefreshBatchComboBox?.Invoke();
        }

        private async void DeleteBatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string batchId)
            {
                var batchToRemove = _state.Batches.FirstOrDefault(b => b.Id == batchId);
                if (batchToRemove != null)
                {
                    var result = await _state.ShowConfirmDialog($"Are you sure you want to delete batch '{batchToRemove.Name}'?",
                        "Confirm Delete");
                    if (!result) return;

                    _state.Batches.Remove(batchToRemove);
                    UpdateBatchList();
                    _state.StateSave.DeleteBatch(batchId);
                    _state.StateSave.CleanupOrphanedEmployees(_state.Batches);
                    _state?.ShowNotification?.Invoke($"Batch '{batchToRemove.Name}' deleted", "Batch Deleted", "info", 4000);
                    _state?.RefreshBatchComboBox?.Invoke();
                }
            }
        }

        private async void RenameBatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string batchId)
            {
                var batch = _state.Batches.FirstOrDefault(b => b.Id == batchId);
                if (batch != null)
                {
                    string newName = await _state.ShowInputDialog("Enter new batch name:", "Rename Batch", batch.Name);

                    if (string.IsNullOrWhiteSpace(newName) || newName == batch.Name)
                    {
                        return;
                    }

                    if (_state.Batches.Any(b => b.Name.Equals(newName, StringComparison.OrdinalIgnoreCase) && b.Id != batchId))
                    {
                        _state?.ShowNotification?.Invoke("A batch with this name already exists.", "Duplicate Name", "warning", 4000);
                        return;
                    }

                    batch.Name = newName;
                    UpdateBatchList();
                    _state.StateSave.SaveBatch(batch);
                    _state?.ShowNotification?.Invoke($"Batch renamed to '{newName}'", "Batch Renamed", "info", 4000);
                    _state?.RefreshBatchComboBox?.Invoke();
                }
            }
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            bool allSelected = _state.CombinedAvailabilities.All(e => e.IsSelected);
            foreach (var employee in _state.CombinedAvailabilities)
            {
                employee.IsSelected = !allSelected;
            }
            SelectAllButton.Content = allSelected ? "☑ Select All" : "☐ Unselect All";
            UpdateCombinedAvailabilities();
        }

        private async void DeleteSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedEmployees = _state.CombinedAvailabilities.Where(e => e.IsSelected).ToList();

            if (selectedEmployees.Count == 0)
            {
                _state?.ShowNotification?.Invoke("Please select at least one employee to delete.", "No Selection", "warning", 4000);
                return;
            }

            string message = selectedEmployees.Count == 1
                ? $"Are you sure you want to delete '{selectedEmployees[0].Name}'?"
                : $"Are you sure you want to delete {selectedEmployees.Count} selected employees?";

            var result = await _state.ShowConfirmDialog(message, "Confirm Delete");
            if (!result) return;

            foreach (var employee in selectedEmployees)
            {
                _state.CombinedAvailabilities.Remove(employee);

                if (employee.Source == "Manual")
                {
                    var manualEmployee = _state.ManualEmployees.FirstOrDefault(emp => emp.Id == employee.Id);
                    if (manualEmployee != null)
                    {
                        _state.ManualEmployees.Remove(manualEmployee);
                    }
                }

                _state.StateSave.DeleteEmployee(employee.Id);
            }

            UpdateCombinedAvailabilities();
            UpdateStatistics();
            _state?.ShowNotification?.Invoke($"Deleted {selectedEmployees.Count} employees", "Deletion Complete", "info", 4000);
            _state?.RefreshEmployeesList?.Invoke();
        }

        private async void DeleteCombinedEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string employeeId)
            {
                var employeeToRemove = _state.CombinedAvailabilities.FirstOrDefault(emp => emp.Id == employeeId);
                if (employeeToRemove != null)
                {
                    var result = await _state.ShowConfirmDialog($"Are you sure you want to delete '{employeeToRemove.Name}'?", "Confirm Delete");
                    if (!result) return;

                    _state.CombinedAvailabilities.Remove(employeeToRemove);

                    if (employeeToRemove.Source == "Manual")
                    {
                        var manualEmployee = _state.ManualEmployees.FirstOrDefault(emp => emp.Id == employeeId);
                        if (manualEmployee != null)
                        {
                            _state.ManualEmployees.Remove(manualEmployee);
                        }
                    }

                    _state.StateSave.DeleteEmployee(employeeId);
                    UpdateCombinedAvailabilities();
                    UpdateStatistics();
                    _state?.ShowNotification?.Invoke($"{employeeToRemove.Name} removed", "Employee Deleted", "info", 4000);
                    _state?.RefreshEmployeesList?.Invoke();
                }
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _state?.NavigateRequested?.Invoke("Back");
        }
    }
}