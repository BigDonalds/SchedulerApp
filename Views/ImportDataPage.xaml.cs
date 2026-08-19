using Microsoft.Win32;
using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SchedulerApp.Views
{
    public partial class ImportDataPage : UserControl
    {
        private AppState _state;
        private bool isLettuceMeetImport = false;
        private WebScraper.EventData lettuceMeetData = null;
        private string importedFilePath = "";

        public ImportDataPage()
        {
            InitializeComponent();
        }

        public void Initialize(AppState state)
        {
            _state = state;
            ImportSourceBox.SelectedIndex = 0;
            FileUploadArea.Visibility = Visibility.Visible;
            LettuceMeetInputArea.Visibility = Visibility.Collapsed;
        }

        private void ImportSourceBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ImportSourceBox.SelectedIndex == 0)
            {
                FileUploadArea.Visibility = Visibility.Visible;
                LettuceMeetInputArea.Visibility = Visibility.Collapsed;
                isLettuceMeetImport = false;
                lettuceMeetData = null;
                ImportedDataCard.Visibility = Visibility.Collapsed;
            }
            else
            {
                FileUploadArea.Visibility = Visibility.Collapsed;
                LettuceMeetInputArea.Visibility = Visibility.Visible;
                isLettuceMeetImport = true;
            }
        }

        private async void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            if (isLettuceMeetImport)
            {
                await ExtractFromLettuceMeet();
                return;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                Title = "Select a CSV schedule file",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                importedFilePath = openFileDialog.FileName;
                await LoadFileData(importedFilePath);
            }
        }

        private async Task LoadFileData(string filePath)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                ImportIcon.Text = "📊";
                ImportTitleText.Text = Path.GetFileName(filePath);
                ImportDetailsText.Text = $"Size: {(fileInfo.Length / 1024.0):F1} KB • Type: CSV";
                var importResult = LocalImport.ImportFromFile(filePath);
                if (importResult.Success)
                {
                    int totalSlots = importResult.People.Sum(p => p.AvailableSlots.Count);
                    ImportStatsText.Text = $"{importResult.People.Count} participants • {importResult.StartDate:MMM dd} to {importResult.EndDate:MMM dd}";
                    ImportStatusIcon.Text = "✅";
                    ImportStatusTitle.Text = "Data Extracted";
                    ImportDataButton.Content = "📥 Import";

                    if (!string.IsNullOrEmpty(importResult.ErrorMessage) && importResult.Success)
                    {
                        ImportStatsText.Text += " (with warnings)";
                        ImportStatusTitle.Text = "Import (Check Warnings)";
                    }
                }
                else
                {
                    ImportStatsText.Text = "File format not recognized. Please use the template format.";
                    ImportStatusIcon.Text = "⚠️";
                    ImportStatusTitle.Text = "Format Issue";
                    ImportDataButton.Content = "📥 Import Anyway";
                }

                ImportedDataCard.Visibility = Visibility.Visible;
                _state?.ShowNotification?.Invoke($"CSV file loaded! Found {importResult.People.Count} participants.", "File Ready", "success", 4000);
            }
            catch (Exception ex)
            {
                _state?.ShowNotification?.Invoke($"Error loading file: {ex.Message}", "File Error", "error", 4000);
            }
        }

        private async Task ExtractFromLettuceMeet()
        {
            string url = LettuceMeetUrlBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(url))
            {
                _state?.ShowNotification?.Invoke("Please enter a LettuceMeet URL.", "Missing URL", "warning", 4000);
                return;
            }

            ExtractionStatus.Visibility = Visibility.Visible;
            ExtractionStatusIcon.Text = "⏳";
            ExtractionStatusText.Text = "Connecting to LettuceMeet...";
            ExtractionStatusDetails.Text = "Please wait while we extract the schedule data";
            ExtractFromLettuceMeetButton.IsEnabled = false;

            try
            {
                var result = await WebScraper.ExtractFromLettuceMeet(url);

                if (result == null)
                {
                    ExtractionStatusIcon.Text = "❌";
                    ExtractionStatusText.Text = "Extraction Failed";
                    ExtractionStatusDetails.Text = "Could not connect to LettuceMeet or parse the data. Please check the link and try again.";
                }
                else if (result.Participants.Count == 0)
                {
                    ExtractionStatusIcon.Text = "⚠️";
                    ExtractionStatusText.Text = "No Participants Found";
                    ExtractionStatusDetails.Text = "The event was found but no participants have marked their availability yet.";
                    lettuceMeetData = result;
                    ImportedDataCard.Visibility = Visibility.Visible;
                }
                else
                {
                    lettuceMeetData = result;
                    ImportIcon.Text = "🔗";
                    ImportTitleText.Text = result.EventTitle;
                    ImportDetailsText.Text = $"From: {result.SourceUrl}";
                    ImportStatsText.Text = $"{result.Participants.Count} participants • {result.StartDate:MMM dd} to {result.EndDate:MMM dd}";
                    ImportStatusIcon.Text = "✅";
                    ImportStatusTitle.Text = "Data Extracted";
                    ImportDataButton.Content = "📥 Import";
                    ExtractionStatusIcon.Text = "✅";
                    ExtractionStatusText.Text = "Extraction Complete";
                    ExtractionStatusDetails.Text = $"Successfully extracted {result.Participants.Count} participants with schedule data";
                    ImportedDataCard.Visibility = Visibility.Visible;
                    ImportedDataCard.BringIntoView();
                }
            }
            catch (Exception ex)
            {
                ExtractionStatusIcon.Text = "❌";
                ExtractionStatusText.Text = "Extraction Error";
                ExtractionStatusDetails.Text = $"Error: {ex.Message}";
                _state?.ShowNotification?.Invoke($"Error: {ex.Message}", "Extraction Error", "error", 4000);
            }
            finally
            {
                ExtractFromLettuceMeetButton.IsEnabled = true;
            }
        }

        private void ImportDataButton_Click(object sender, RoutedEventArgs e)
        {
            if (isLettuceMeetImport && lettuceMeetData != null)
            {
                ImportLettuceMeetData();
            }
            else if (!string.IsNullOrEmpty(importedFilePath))
            {
                ImportFileData();
            }
        }

        private async void ImportLettuceMeetData()
        {
            if (lettuceMeetData == null) return;

            try
            {
                int importedCount = 0;
                DateTime overallStartDate = lettuceMeetData.StartDate;
                DateTime overallEndDate = lettuceMeetData.EndDate;
                int totalDays = (overallEndDate - overallStartDate).Days + 1;
                int earliestHour = 23;
                int earliestMinute = 59;
                int latestHour = 0;
                int latestMinute = 0;

                foreach (var participant in lettuceMeetData.Participants)
                {
                    foreach (var slot in participant.AvailableSlots)
                    {
                        if (slot.ParsedStart.Hour < earliestHour ||
                            (slot.ParsedStart.Hour == earliestHour && slot.ParsedStart.Minute < earliestMinute))
                        {
                            earliestHour = slot.ParsedStart.Hour;
                            earliestMinute = slot.ParsedStart.Minute;
                        }

                        if (slot.ParsedEnd.Hour > latestHour ||
                            (slot.ParsedEnd.Hour == latestHour && slot.ParsedEnd.Minute > latestMinute))
                        {
                            latestHour = slot.ParsedEnd.Hour;
                            latestMinute = slot.ParsedEnd.Minute;
                        }
                    }
                }

                if (earliestHour > latestHour || (earliestHour == latestHour && earliestMinute >= latestMinute))
                {
                    earliestHour = 9;
                    earliestMinute = 0;
                    latestHour = 17;
                    latestMinute = 0;
                }

                TimeSpan earliestTime = new TimeSpan(earliestHour, earliestMinute, 0);
                TimeSpan latestTime = new TimeSpan(latestHour, latestMinute, 0);
                TimeSpan totalDuration = latestTime - earliestTime;
                int totalHalfHourSlots = (int)Math.Ceiling(totalDuration.TotalMinutes / 30);

                foreach (var participant in lettuceMeetData.Participants)
                {
                    if (participant.AvailableSlots.Count == 0)
                    {
                        continue;
                    }

                    bool[,] scheduleMatrix = new bool[totalDays, totalHalfHourSlots];

                    for (int day = 0; day < totalDays; day++)
                    {
                        for (int slot = 0; slot < totalHalfHourSlots; slot++)
                        {
                            scheduleMatrix[day, slot] = false;
                        }
                    }

                    foreach (var lettuceSlot in participant.AvailableSlots)
                    {
                        int dayIndex = (lettuceSlot.ParsedStart.Date - overallStartDate.Date).Days;

                        if (dayIndex < 0 || dayIndex >= totalDays)
                        {
                            continue;
                        }

                        TimeSpan slotStartTime = new TimeSpan(lettuceSlot.ParsedStart.Hour, lettuceSlot.ParsedStart.Minute, 0);
                        TimeSpan slotEndTime = new TimeSpan(lettuceSlot.ParsedEnd.Hour, lettuceSlot.ParsedEnd.Minute, 0);
                        TimeSpan startOffset = slotStartTime - earliestTime;
                        TimeSpan endOffset = slotEndTime - earliestTime;
                        int startSlotIndex = (int)Math.Floor(startOffset.TotalMinutes / 30);
                        int endSlotIndex = (int)Math.Ceiling(endOffset.TotalMinutes / 30);
                        startSlotIndex = Math.Max(0, startSlotIndex);
                        endSlotIndex = Math.Min(totalHalfHourSlots, endSlotIndex);

                        if (startSlotIndex >= endSlotIndex)
                        {
                            continue;
                        }

                        for (int slotIndex = startSlotIndex; slotIndex < endSlotIndex; slotIndex++)
                        {
                            if (slotIndex >= 0 && slotIndex < totalHalfHourSlots)
                            {
                                scheduleMatrix[dayIndex, slotIndex] = true;
                            }
                        }
                    }

                    int availableSlotCount = 0;
                    for (int day = 0; day < totalDays; day++)
                    {
                        for (int slot = 0; slot < totalHalfHourSlots; slot++)
                        {
                            if (scheduleMatrix[day, slot]) availableSlotCount++;
                        }
                    }

                    List<DateTime> selectedDates = new List<DateTime>();
                    foreach (var slot in participant.AvailableSlots)
                    {
                        DateTime slotDate = slot.ParsedStart.Date;
                        if (!selectedDates.Contains(slotDate))
                        {
                            selectedDates.Add(slotDate);
                        }
                    }

                    if (selectedDates.Count == 0)
                    {
                        DateTime currentDate = overallStartDate;
                        while (currentDate <= overallEndDate)
                        {
                            selectedDates.Add(currentDate);
                            currentDate = currentDate.AddDays(1);
                        }
                    }

                    var employee = new AvailabilityEntry
                    {
                        Id = Guid.NewGuid().ToString(),
                        Name = participant.Name,
                        DateRange = $"{overallStartDate:MMM dd} to {overallEndDate:MMM dd}",
                        TimeRange = $"From extracted availability ({earliestHour}:{earliestMinute:D2} to {latestHour}:{latestMinute:D2})",
                        AvailabilitySummary = $"{availableSlotCount} half-hour slots available",
                        SlotCount = availableSlotCount,
                        Source = "LettuceMeet",
                        SourceColor = (Color)ColorConverter.ConvertFromString("#10B981"),
                        CreatedDate = DateTime.Now,
                        StartDate = overallStartDate,
                        EndDate = overallEndDate,
                        StartHour = earliestHour,
                        EndHour = latestHour,
                        ScheduleMatrix = scheduleMatrix,
                        SelectedDates = selectedDates
                    };

                    if (!_state.CombinedAvailabilities.Any(emp => emp.Name.Equals(employee.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        _state.CombinedAvailabilities.Add(employee);
                        _state.StateSave.SaveEmployee(employee);
                        importedCount++;
                    }
                }

                lettuceMeetData = null;
                LettuceMeetUrlBox.Text = "";
                ImportedDataCard.Visibility = Visibility.Collapsed;
                ExtractionStatus.Visibility = Visibility.Collapsed;
                _state?.ShowNotification?.Invoke($"Imported {importedCount} employees from LettuceMeet with exact timing!", "Import Complete", "success", 4000);
                _state?.NavigateRequested?.Invoke("ViewAvailabilities");
            }
            catch (Exception ex)
            {
                _state?.ShowNotification?.Invoke($"Error importing data: {ex.Message}\n\nCheck debug.txt on desktop for details.", "Import Error", "error", 0);
            }
        }

        private async void ImportFileData()
        {
            try
            {
                var importResult = LocalImport.ImportFromFile(importedFilePath);

                if (!importResult.Success)
                {
                    string errorMsg = importResult.ErrorMessage;

                    if (!string.IsNullOrEmpty(errorMsg) && importResult.Success)
                    {
                        _state?.ShowNotification?.Invoke($"Import completed with warnings:\n\n{errorMsg}", "Import Warnings", "warning", 0);
                    }
                    else
                    {
                        _state?.ShowNotification?.Invoke($"Error importing file:\n\n{errorMsg}", "Import Error", "error", 0);
                        return;
                    }
                }

                int importedCount = 0;
                int duplicateCount = 0;

                foreach (var person in importResult.People)
                {
                    DateTime startDate = importResult.StartDate;
                    DateTime endDate = importResult.EndDate;
                    int totalDays = (endDate - startDate).Days + 1;
                    int earliestHour = 23;
                    int earliestMinute = 59;
                    int latestHour = 0;
                    int latestMinute = 0;

                    foreach (var slot in person.AvailableSlots)
                    {
                        if (slot.ParsedStart.Hour < earliestHour ||
                            (slot.ParsedStart.Hour == earliestHour && slot.ParsedStart.Minute < earliestMinute))
                        {
                            earliestHour = slot.ParsedStart.Hour;
                            earliestMinute = slot.ParsedStart.Minute;
                        }

                        if (slot.ParsedEnd.Hour > latestHour ||
                            (slot.ParsedEnd.Hour == latestHour && slot.ParsedEnd.Minute > latestMinute))
                        {
                            latestHour = slot.ParsedEnd.Hour;
                            latestMinute = slot.ParsedEnd.Minute;
                        }
                    }

                    if (earliestHour > latestHour || (earliestHour == latestHour && earliestMinute >= latestMinute))
                    {
                        earliestHour = 9;
                        earliestMinute = 0;
                        latestHour = 17;
                        latestMinute = 0;
                    }

                    TimeSpan earliestTime = new TimeSpan(earliestHour, earliestMinute, 0);
                    TimeSpan latestTime = new TimeSpan(latestHour, latestMinute, 0);
                    TimeSpan totalDuration = latestTime - earliestTime;
                    int totalHalfHourSlots = (int)Math.Ceiling(totalDuration.TotalMinutes / 30);
                    bool[,] scheduleMatrix = new bool[totalDays, totalHalfHourSlots];

                    foreach (var slot in person.AvailableSlots)
                    {
                        int dayIndex = (slot.ParsedStart.Date - startDate.Date).Days;
                        if (dayIndex < 0 || dayIndex >= totalDays) continue;

                        TimeSpan slotStartTime = new TimeSpan(slot.ParsedStart.Hour, slot.ParsedStart.Minute, 0);
                        TimeSpan slotEndTime = new TimeSpan(slot.ParsedEnd.Hour, slot.ParsedEnd.Minute, 0);
                        TimeSpan startOffset = slotStartTime - earliestTime;
                        TimeSpan endOffset = slotEndTime - earliestTime;
                        int startSlotIndex = (int)Math.Floor(startOffset.TotalMinutes / 30);
                        int endSlotIndex = (int)Math.Ceiling(endOffset.TotalMinutes / 30);
                        startSlotIndex = Math.Max(0, startSlotIndex);
                        endSlotIndex = Math.Min(totalHalfHourSlots, endSlotIndex);

                        for (int slotIndex = startSlotIndex; slotIndex < endSlotIndex; slotIndex++)
                        {
                            if (slotIndex >= 0 && slotIndex < totalHalfHourSlots)
                            {
                                scheduleMatrix[dayIndex, slotIndex] = true;
                            }
                        }
                    }

                    int availableSlotCount = 0;
                    for (int day = 0; day < totalDays; day++)
                    {
                        for (int slot = 0; slot < totalHalfHourSlots; slot++)
                        {
                            if (scheduleMatrix[day, slot]) availableSlotCount++;
                        }
                    }

                    List<DateTime> selectedDates = new List<DateTime>();
                    DateTime currentDate = startDate;
                    while (currentDate <= endDate)
                    {
                        selectedDates.Add(currentDate);
                        currentDate = currentDate.AddDays(1);
                    }

                    var employee = new AvailabilityEntry
                    {
                        Id = Guid.NewGuid().ToString(),
                        Name = person.Name,
                        DateRange = $"{startDate:MMM dd} to {endDate:MMM dd}",
                        TimeRange = $"From imported data ({earliestHour}:{earliestMinute:D2} to {latestHour}:{latestMinute:D2})",
                        AvailabilitySummary = $"{availableSlotCount} half-hour slots available",
                        SlotCount = availableSlotCount,
                        Source = "CSV Import",
                        SourceColor = (Color)ColorConverter.ConvertFromString("#F59E0B"),
                        CreatedDate = DateTime.Now,
                        StartDate = startDate,
                        EndDate = endDate,
                        StartHour = earliestHour,
                        EndHour = latestHour,
                        ScheduleMatrix = scheduleMatrix,
                        SelectedDates = selectedDates
                    };

                    if (!_state.CombinedAvailabilities.Any(emp => emp.Name.Equals(employee.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        _state.CombinedAvailabilities.Add(employee);
                        _state.StateSave.SaveEmployee(employee);
                        importedCount++;
                    }
                    else
                    {
                        duplicateCount++;
                    }
                }

                string notificationMessage;
                string notificationTitle;
                string notificationType = "success";

                if (importedCount > 0 && duplicateCount == 0)
                {
                    notificationMessage = $"Successfully imported {importedCount} employees from CSV!";
                    notificationTitle = "Import Complete";
                }
                else if (importedCount > 0 && duplicateCount > 0)
                {
                    notificationMessage = $"Imported {importedCount} employees, skipped {duplicateCount} duplicates.";
                    notificationTitle = "Import Complete (Some Skipped)";
                    notificationType = "warning";
                }
                else if (importedCount == 0 && duplicateCount > 0)
                {
                    notificationMessage = $"All {duplicateCount} employees were duplicates and were skipped.";
                    notificationTitle = "No New Employees Imported";
                    notificationType = "warning";
                }
                else
                {
                    notificationMessage = "No data was imported.";
                    notificationTitle = "Import Failed";
                    notificationType = "error";
                }

                _state?.ShowNotification?.Invoke(notificationMessage, notificationTitle, notificationType, 4000);
                importedFilePath = "";
                ImportedDataCard.Visibility = Visibility.Collapsed;
                _state?.NavigateRequested?.Invoke("ViewAvailabilities");
            }
            catch (Exception ex)
            {
                _state?.ShowNotification?.Invoke($"Error importing data: {ex.Message}\n\nPlease ensure your file matches the template format.", "Import Error", "error", 0);
            }
        }

        private void UploadArea_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            BrowseButton_Click(sender, e);
        }

        private async void ExtractFromLettuceMeetButton_Click(object sender, RoutedEventArgs e)
        {
            await ExtractFromLettuceMeet();
        }

        private async void DownloadSampleButton_Click(object sender, RoutedEventArgs e)
        {
            var result = await _state.ShowSelectionDialog(
                "Which template format would you like to download?",
                "Choose Template Format",
                new List<string>
                {
                    "CSV with time slots (Recommended)",
                    "CSV with date columns grid"
                });

            if (string.IsNullOrEmpty(result)) return;

            string sampleContent;
            string fileName;

            if (result == "CSV with date columns grid")
            {
                sampleContent = LocalImport.GetGridCsvTemplate();
                fileName = "schedule_grid_template.csv";
            }
            else
            {
                sampleContent = LocalImport.GetSampleCsvTemplate();
                fileName = "schedule_template.csv";
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                Title = "Save Template",
                FileName = fileName,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, sampleContent);
                    _state?.ShowNotification?.Invoke("Template downloaded successfully!", "Download Complete", "success", 4000);

                    _state?.ShowNotification?.Invoke(
                        "Template downloaded! Here's how to use it:\n\n" +
                        "1. Open the CSV file in Excel or similar\n" +
                        "2. Fill in employee names and their availability\n" +
                        "3. Save the file\n" +
                        "4. Upload it back to Scheduler Pro\n\n" +
                        "Format: Name, Date (YYYY-MM-DD), Start_Time (HH:mm), End_Time (HH:mm)",
                        "Template Instructions",
                        "info", 0);
                }
                catch (Exception ex)
                {
                    _state?.ShowNotification?.Invoke($"Error saving template: {ex.Message}", "Error", "error", 4000);
                }
            }
        }

        private void ChangeImportButton_Click(object sender, RoutedEventArgs e)
        {
            ImportedDataCard.Visibility = Visibility.Collapsed;
            lettuceMeetData = null;
            importedFilePath = "";

            if (isLettuceMeetImport)
            {
                LettuceMeetUrlBox.Focus();
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _state?.NavigateRequested?.Invoke("Back");
        }
    }
}