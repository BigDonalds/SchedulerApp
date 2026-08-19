using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace SchedulerApp.Views
{
    public partial class ManualSchedulePage : UserControl
    {
        private AppState _state;
        private List<DateTime> selectedDates = new List<DateTime>();
        private List<Border> selectedCells = new List<Border>();
        private DateTime currentMonth = DateTime.Today;
        private Dictionary<Button, DateTime> calendarButtons = new Dictionary<Button, DateTime>();
        private DateTime? calendarDragStartDate = null;
        private bool isCalendarDragging = false;
        private bool isAdditiveSelection = false;
        private bool isSelecting = false;
        private int days = 0;
        private int hours = 0;

        public ManualSchedulePage()
        {
            InitializeComponent();
        }

        public void Initialize(AppState state)
        {
            _state = state;
            StartTimeBox.SelectedIndex = 0;
            EndTimeBox.SelectedIndex = 11;
            UpdateCalendarDisplay();
            UpdateEmployeesList();
        }

        public void Reset()
        {
            RangeSelectionCard.Visibility = Visibility.Visible;
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            selectedCells.Clear();
            EmployeeNameBox.Text = "";
        }

        public void UpdateEmployeesList()
        {
            EmployeesList.ItemsSource = null;
            EmployeesList.ItemsSource = _state.ManualEmployees;
            NoEmployeesText.Visibility = _state.ManualEmployees.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public void UpdateBackButtonVisibility()
        {
            ManualBackButton.Visibility = (ScheduleGridCard.Visibility == Visibility.Visible)
                ? Visibility.Visible : Visibility.Collapsed;
        }

        public bool IsGridVisible => ScheduleGridCard.Visibility == Visibility.Visible;

        private void UpdateCalendarDisplay()
        {
            MonthYearDisplay.Text = currentMonth.ToString("MMMM yyyy");
            DaysGrid.Children.Clear();
            calendarButtons.Clear();
            DateTime firstDayOfMonth = new DateTime(currentMonth.Year, currentMonth.Month, 1);
            int startDay = (int)firstDayOfMonth.DayOfWeek;
            int totalDays = DateTime.DaysInMonth(currentMonth.Year, currentMonth.Month);
            DateTime today = DateTime.Today;

            for (int i = 0; i < 42; i++)
            {
                int dayNumber = i - startDay + 1;
                Button dayButton = new Button();

                if (dayNumber > 0 && dayNumber <= totalDays)
                {
                    DateTime currentDate = new DateTime(currentMonth.Year, currentMonth.Month, dayNumber);
                    dayButton.Template = CreateDayButtonTemplate();
                    dayButton.Content = dayNumber.ToString();
                    dayButton.Tag = currentDate;
                    dayButton.Height = 36;
                    dayButton.Width = 36;
                    dayButton.Margin = new Thickness(2, -10, 2, 2);
                    dayButton.FontSize = 14;
                    dayButton.FontWeight = FontWeights.Medium;
                    dayButton.BorderThickness = new Thickness(1);
                    UpdateDayButtonAppearance(dayButton, currentDate, today);
                    calendarButtons[dayButton] = currentDate;
                    dayButton.PreviewMouseLeftButtonDown += DayButton_PreviewMouseLeftButtonDown;
                    dayButton.PreviewMouseMove += DayButton_PreviewMouseMove;
                    dayButton.PreviewMouseLeftButtonUp += DayButton_PreviewMouseLeftButtonUp;
                    dayButton.MouseEnter += DayButton_MouseEnter;
                    dayButton.MouseLeave += DayButton_MouseLeave;
                }
                else
                {
                    dayButton.Content = "";
                    dayButton.IsEnabled = false;
                    dayButton.Background = Brushes.Transparent;
                    dayButton.BorderThickness = new Thickness(0);
                    dayButton.Opacity = 0.3;
                }

                DaysGrid.Children.Add(dayButton);
            }
        }

        private ControlTemplate CreateDayButtonTemplate()
        {
            ControlTemplate template = new ControlTemplate(typeof(Button));
            FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            FrameworkElementFactory contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.AppendChild(contentFactory);
            template.VisualTree = borderFactory;
            return template;
        }

        private void UpdateDayButtonAppearance(Button dayButton, DateTime date, DateTime today)
        {
            var selectedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5"));
            var todayBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));
            var normalBrush = new SolidColorBrush(Colors.Transparent);
            var whiteText = new SolidColorBrush(Colors.White);
            var darkText = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2937"));

            if (selectedDates.Contains(date))
            {
                dayButton.Background = selectedBrush;
                dayButton.Foreground = whiteText;
                dayButton.BorderBrush = selectedBrush;
            }
            else if (date.Date == today.Date)
            {
                dayButton.Background = normalBrush;
                dayButton.Foreground = darkText;
                dayButton.BorderBrush = todayBrush;
                dayButton.BorderThickness = new Thickness(2);
            }
            else
            {
                dayButton.Background = normalBrush;
                dayButton.Foreground = darkText;
                dayButton.BorderBrush = Brushes.Transparent;
                dayButton.BorderThickness = new Thickness(1);
            }
        }

        private void DayButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (sender is Button dayButton && dayButton.Tag is DateTime date && !selectedDates.Contains(date))
            {
                dayButton.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB"));
            }
        }

        private void DayButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sender is Button dayButton && dayButton.Tag is DateTime date)
            {
                UpdateDayButtonAppearance(dayButton, date, DateTime.Today);
            }
        }

        private void DayButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Button dayButton && dayButton.Tag is DateTime date)
            {
                calendarDragStartDate = date;
                isCalendarDragging = true;
                isAdditiveSelection = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
                if (!isAdditiveSelection)
                {
                    selectedDates.Clear();
                }
                ToggleDateSelection(date);
                e.Handled = true;
            }
        }

        private void DayButton_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (isCalendarDragging && e.LeftButton == MouseButtonState.Pressed && calendarDragStartDate.HasValue)
            {
                if (sender is Button dayButton && dayButton.Tag is DateTime currentDate)
                {
                    DateTime start = calendarDragStartDate.Value;
                    DateTime end = currentDate;
                    if (start > end)
                    {
                        DateTime temp = start;
                        start = end;
                        end = temp;
                    }
                    if (!isAdditiveSelection)
                    {
                        selectedDates.Clear();
                    }
                    for (DateTime date = start; date <= end; date = date.AddDays(1))
                    {
                        if (!selectedDates.Contains(date))
                        {
                            selectedDates.Add(date);
                        }
                    }
                    UpdateCalendarDisplay();
                    UpdateSelectedDatesDisplay();
                }
            }
        }

        private void DayButton_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            isCalendarDragging = false;
            calendarDragStartDate = null;
            UpdateSelectedDatesDisplay();
        }

        private void ToggleDateSelection(DateTime date)
        {
            if (selectedDates.Contains(date))
            {
                selectedDates.Remove(date);
            }
            else
            {
                selectedDates.Add(date);
            }
            foreach (var kvp in calendarButtons)
            {
                UpdateDayButtonAppearance(kvp.Key, kvp.Value, DateTime.Today);
            }
            UpdateSelectedDatesDisplay();
        }

        private void PrevMonthButton_Click(object sender, RoutedEventArgs e)
        {
            currentMonth = currentMonth.AddMonths(-1);
            UpdateCalendarDisplay();
        }

        private void NextMonthButton_Click(object sender, RoutedEventArgs e)
        {
            currentMonth = currentMonth.AddMonths(1);
            UpdateCalendarDisplay();
        }

        private void UpdateSelectedDatesDisplay()
        {
            SelectedDatesPanel.Children.Clear();
            var sortedDates = selectedDates.OrderBy(d => d.Date).ToList();

            if (sortedDates.Count == 0)
            {
                SelectedDatesPanel.Children.Add(new TextBlock
                {
                    Text = "No dates selected",
                    FontSize = 12,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7280")),
                    FontStyle = FontStyles.Italic
                });
            }
            else
            {
                foreach (DateTime date in sortedDates)
                {
                    Border dateBadge = new Border
                    {
                        Style = (Style)FindResource("SelectedDateStyle"),
                        Margin = new Thickness(0, 0, 4, 4),
                        Child = new TextBlock
                        {
                            Text = date.ToString("MMM d"),
                            FontSize = 11,
                            FontWeight = FontWeights.Medium,
                            Foreground = Brushes.Black
                        }
                    };
                    SelectedDatesPanel.Children.Add(dateBadge);
                }
            }
        }

        private void CreateGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (selectedDates.Count == 0)
            {
                _state?.ShowNotification?.Invoke("Please select at least one date.", "No Dates Selected", "warning", 4000);
                return;
            }

            selectedDates = selectedDates.OrderBy(d => d.Date).ToList();

            if (StartTimeBox.SelectedIndex == -1 || EndTimeBox.SelectedIndex == -1)
            {
                _state?.ShowNotification?.Invoke("Please select both start and end times.", "Missing Times", "warning", 4000);
                return;
            }

            int startHourIndex = StartTimeBox.SelectedIndex;
            int endHourIndex = EndTimeBox.SelectedIndex;

            if (endHourIndex <= startHourIndex)
            {
                _state?.ShowNotification?.Invoke("End time must be after start time.", "Invalid Time Range", "warning", 4000);
                return;
            }

            days = selectedDates.Count;
            hours = (endHourIndex - startHourIndex) + 1;

            RangeSelectionCard.Visibility = Visibility.Collapsed;
            ScheduleGridCard.Visibility = Visibility.Visible;
            BackToDatesButton.Visibility = Visibility.Visible;
            ManualBackButton.Visibility = Visibility.Collapsed;
            CurrentEmployeeText.Text = "for: New Employee";
            EmployeeNameBox.Text = "";
            selectedCells.Clear();
            GenerateScheduleGrid();
            _state?.ShowNotification?.Invoke($"Grid created with half-hour intervals! ({hours * 2} time slots per day)", "Ready to Add Employee", "success", 4000);
        }

        private void GenerateScheduleGrid()
        {
            ScheduleGrid.Children.Clear();
            ScheduleGrid.RowDefinitions.Clear();
            ScheduleGrid.ColumnDefinitions.Clear();
            selectedCells.Clear();
            ScheduleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });

            for (int day = 0; day < days; day++)
            {
                ScheduleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            }

            ScheduleGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(35) });
            for (int halfHour = 0; halfHour < hours * 2; halfHour++)
            {
                ScheduleGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(35) });
            }

            for (int day = 0; day < days; day++)
            {
                DateTime currentDate = selectedDates[day];
                string dayLabel = currentDate.ToString("ddd\nMM/dd");
                TextBlock dayHeader = new TextBlock
                {
                    Text = dayLabel,
                    FontSize = 10,
                    FontWeight = FontWeights.SemiBold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = Brushes.White,
                    TextAlignment = TextAlignment.Center
                };
                Border headerBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5")),
                    Child = dayHeader
                };
                Grid.SetColumn(headerBorder, day + 1);
                Grid.SetRow(headerBorder, 0);
                ScheduleGrid.Children.Add(headerBorder);
            }

            int startHourIndex = StartTimeBox.SelectedIndex;
            int startHour = 6 + startHourIndex;

            for (int halfHour = 0; halfHour < hours * 2; halfHour++)
            {
                int totalMinutes = (startHour * 60) + (halfHour * 30);
                int hour = totalMinutes / 60;
                int minute = totalMinutes % 60;
                int endTotalMinutes = totalMinutes + 30;
                int endHour = endTotalMinutes / 60;
                int endMinute = endTotalMinutes % 60;
                string startTimeStr = FormatTimeForGrid(hour, minute);
                string endTimeStr = FormatTimeForGrid(endHour, endMinute);
                string timeLabel = $"{startTimeStr} to\n{endTimeStr}";
                TextBlock timeText = new TextBlock
                {
                    Text = timeLabel,
                    FontSize = 9,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7280")),
                    TextAlignment = TextAlignment.Center
                };
                Border timeBorder = new Border
                {
                    Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC")),
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                    BorderThickness = new Thickness(1),
                    Child = timeText
                };
                Grid.SetColumn(timeBorder, 0);
                Grid.SetRow(timeBorder, halfHour + 1);
                ScheduleGrid.Children.Add(timeBorder);
            }

            for (int day = 0; day < days; day++)
            {
                for (int halfHour = 0; halfHour < hours * 2; halfHour++)
                {
                    Border cell = new Border
                    {
                        Style = (Style)FindResource("ScheduleCell"),
                        Tag = $"{day},{halfHour}"
                    };
                    cell.MouseLeftButtonDown += Cell_MouseLeftButtonDown;
                    cell.MouseEnter += Cell_MouseEnter;
                    cell.MouseLeftButtonUp += Cell_MouseLeftButtonUp;
                    Grid.SetColumn(cell, day + 1);
                    Grid.SetRow(cell, halfHour + 1);
                    ScheduleGrid.Children.Add(cell);
                }
            }

            double totalGridHeight = (hours * 2 + 1) * 35;
            GridScrollViewer.MaxHeight = 500;

            if (totalGridHeight > 400)
            {
                GridScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else
            {
                GridScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            }
        }

        private string FormatTimeForGrid(int hour, int minute)
        {
            string ampm = hour >= 12 ? "PM" : "AM";
            int displayHour = hour > 12 ? hour - 12 : (hour == 0 ? 12 : hour);
            if (minute == 0)
                return $"{displayHour}:00";
            else
                return $"{displayHour}:{minute:D2}";
        }

        private void Cell_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            isSelecting = true;
            ToggleCellSelection(sender as Border);
        }

        private void Cell_MouseEnter(object sender, MouseEventArgs e)
        {
            if (isSelecting && sender is Border cell)
            {
                ToggleCellSelection(cell);
            }
        }

        private void Cell_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            isSelecting = false;
        }

        private void ToggleCellSelection(Border cell)
        {
            if (cell == null) return;
            if (selectedCells.Contains(cell))
            {
                selectedCells.Remove(cell);
                cell.Style = (Style)FindResource("ScheduleCell");
            }
            else
            {
                selectedCells.Add(cell);
                cell.Style = (Style)FindResource("SelectedScheduleCell");
            }
        }

        private void SaveAvailability_Click(object sender, RoutedEventArgs e)
        {
            string name = EmployeeNameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                _state?.ShowNotification?.Invoke("Please enter employee name.", "Missing Information", "warning", 4000);
                return;
            }

            if (selectedCells.Count == 0)
            {
                _state?.ShowNotification?.Invoke("Please select at least one time slot for availability.", "No Selection", "warning", 4000);
                return;
            }

            selectedDates = selectedDates.OrderBy(d => d.Date).ToList();
            int daysCount = selectedDates.Count;
            int halfHourCount = hours * 2;
            bool[,] scheduleMatrix = new bool[daysCount, halfHourCount];

            foreach (Border cell in selectedCells)
            {
                if (cell.Tag is string tag)
                {
                    string[] parts = tag.Split(',');
                    if (parts.Length == 2 &&
                        int.TryParse(parts[0], out int day) &&
                        int.TryParse(parts[1], out int halfHour))
                    {
                        scheduleMatrix[day, halfHour] = true;
                    }
                }
            }

            int actualStartHour = 6 + StartTimeBox.SelectedIndex;
            int actualEndHour = 6 + EndTimeBox.SelectedIndex + 1;

            var employee = new AvailabilityEntry
            {
                Id = Guid.NewGuid().ToString(),
                Name = name,
                DateRange = $"{selectedDates.First():MMM dd} to {selectedDates.Last():MMM dd}",
                TimeRange = $"{actualStartHour}:00 to {actualEndHour}:00 (half-hour intervals)",
                Source = "Manual",
                SourceColor = (Color)ColorConverter.ConvertFromString("#4F46E5"),
                StartDate = selectedDates.First(),
                EndDate = selectedDates.Last(),
                StartHour = actualStartHour,
                EndHour = actualEndHour,
                AvailabilitySummary = $"{selectedCells.Count}/{daysCount * halfHourCount} half-hour slots",
                SlotCount = selectedCells.Count,
                CreatedDate = DateTime.Now,
                ScheduleMatrix = scheduleMatrix,
                SelectedDates = new List<DateTime>(selectedDates)
            };

            _state.ManualEmployees.Add(employee);
            _state.CombinedAvailabilities.Add(employee);
            _state.StateSave.SaveEmployee(employee);
            UpdateEmployeesList();
            EmployeeNameBox.Text = "";
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            RangeSelectionCard.Visibility = Visibility.Visible;
            selectedCells.Clear();
            _state?.ShowNotification?.Invoke($"Availability saved for {name}!", "Success", "success", 4000);
            _state?.RefreshStatistics?.Invoke();
            _state?.RefreshCombinedAvailabilities?.Invoke();
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (Border cell in selectedCells)
            {
                cell.Style = (Style)FindResource("ScheduleCell");
            }
            selectedCells.Clear();
        }

        private async void DeleteEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string employeeId)
            {
                var employeeToRemove = _state.ManualEmployees.FirstOrDefault(emp => emp.Id == employeeId);
                if (employeeToRemove != null)
                {
                    var result = await _state.ShowConfirmDialog($"Are you sure you want to delete '{employeeToRemove.Name}'?", "Confirm Delete");
                    if (!result) return;

                    _state.ManualEmployees.Remove(employeeToRemove);
                    var combinedToRemove = _state.CombinedAvailabilities.FirstOrDefault(emp => emp.Id == employeeId);
                    if (combinedToRemove != null)
                    {
                        _state.CombinedAvailabilities.Remove(combinedToRemove);
                    }
                    _state.StateSave.DeleteEmployee(employeeId);
                    UpdateEmployeesList();
                    _state?.ShowNotification?.Invoke($"{employeeToRemove.Name} removed", "Employee Deleted", "info", 4000);
                    _state?.RefreshStatistics?.Invoke();
                    _state?.RefreshCombinedAvailabilities?.Invoke();
                }
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (ScheduleGridCard.Visibility == Visibility.Visible)
            {
                ScheduleGridCard.Visibility = Visibility.Collapsed;
                RangeSelectionCard.Visibility = Visibility.Visible;
                selectedCells.Clear();
                EmployeeNameBox.Text = "";
                UpdateBackButtonVisibility();
            }
        }

        private void BackToDatesButton_Click(object sender, RoutedEventArgs e)
        {
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            RangeSelectionCard.Visibility = Visibility.Visible;
            BackToDatesButton.Visibility = Visibility.Collapsed;
            selectedCells.Clear();
            EmployeeNameBox.Text = "";
            UpdateBackButtonVisibility();
        }
    }
}