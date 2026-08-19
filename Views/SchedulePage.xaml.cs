using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace SchedulerApp.Views
{
    public class WorkloadItem
    {
        public string Name { get; set; }
        public double Hours { get; set; }
        public string HoursText { get; set; }
        public string StatusLabel { get; set; }
        public string StatusColor { get; set; }
        public string BarColor { get; set; }
        public double BarWidth { get; set; }
        public string DetailText { get; set; }
    }

    public class SuggestionItem
    {
        public string Icon { get; set; }
        public string Message { get; set; }
        public string SuggestionBg { get; set; }
        public string SuggestionBorder { get; set; }
    }

    public class StaffingAlertItem
    {
        public string AlertTitle { get; set; }
        public string AlertDetail { get; set; }
        public string AlertCandidates { get; set; }
        public string AlertBg { get; set; }
        public string AlertBorder { get; set; }
        public string AlertTitleColor { get; set; }
        public Visibility CandidatesVisible { get; set; } = Visibility.Collapsed;
    }

    public partial class SchedulePage : UserControl
    {
        private AppState _state;
        private string _editingCellId = null;
        private List<string> _editingNames = new List<string>();
        private bool _isEditPanelClosing = false;
        private bool _sortByAlphabet = false;
        private List<WorkloadItem> _workloadItems = new List<WorkloadItem>();

        public SchedulePage()
        {
            InitializeComponent();
            EditNamesOverlayBackground.MouseLeftButtonDown += (s, e) => CloseEditNamesPanel();
        }

        public void Initialize(AppState state)
        {
            _state = state;
        }

        public void UpdateScheduleList()
        {
            ScheduleListControl.ItemsSource = null;
            ScheduleListControl.ItemsSource = _state.Schedules;
            ScheduleCountText.Text = $"({_state.Schedules.Count})";
            NoSchedulesPanel.Visibility = _state.Schedules.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public void ShowScheduleListView()
        {
            ScheduleListView.Visibility = Visibility.Visible;
            ScheduleDetailView.Visibility = Visibility.Collapsed;
            ScheduleBackButton.Visibility = Visibility.Collapsed;
            _editingCellId = null;
        }

        public void ShowScheduleDetailView(Schedule schedule)
        {
            _state.CurrentSchedule = schedule;
            ScheduleListView.Visibility = Visibility.Collapsed;
            ScheduleDetailView.Visibility = Visibility.Visible;
            ScheduleBackButton.Visibility = Visibility.Visible;
            ScheduleDetailTitle.Text = schedule.Name;
            CurrentScheduleName.Text = schedule.Name;
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            UpdateAnalytics();
        }

        private void ViewSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var schedule = _state.Schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (schedule != null)
                {
                    ShowScheduleDetailView(schedule);
                }
            }
        }

        private void ScheduleItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.DataContext is Schedule schedule)
            {
                ShowScheduleDetailView(schedule);
            }
        }

        private async void RenameSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var schedule = _state.Schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (schedule != null)
                {
                    string newName = await _state.ShowInputDialog("Enter new schedule name:", "Rename Schedule", schedule.Name);

                    if (string.IsNullOrWhiteSpace(newName) || newName == schedule.Name)
                    {
                        return;
                    }

                    if (_state.Schedules.Any(s => s.Name.Equals(newName, StringComparison.OrdinalIgnoreCase) && s.Id != scheduleId))
                    {
                        _state?.ShowNotification?.Invoke("A schedule with this name already exists.", "Duplicate Name", "warning", 4000);
                        return;
                    }

                    schedule.Name = newName;
                    UpdateScheduleList();
                    _state.StateSave.SaveSchedule(schedule);

                    if (_state.CurrentSchedule != null && _state.CurrentSchedule.Id == scheduleId)
                    {
                        _state.CurrentSchedule.Name = newName;
                        ScheduleDetailTitle.Text = newName;
                        CurrentScheduleName.Text = newName;
                    }

                    _state?.ShowNotification?.Invoke($"Schedule renamed to '{newName}'", "Schedule Renamed", "info", 4000);
                }
            }
        }

        private async void DeleteSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var scheduleToRemove = _state.Schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (scheduleToRemove != null)
                {
                    var result = await _state.ShowConfirmDialog($"Are you sure you want to delete schedule '{scheduleToRemove.Name}'?",
                        "Confirm Delete");
                    if (!result) return;

                    _state.Schedules.Remove(scheduleToRemove);
                    UpdateScheduleList();
                    _state.StateSave.DeleteSchedule(scheduleId);

                    if (_state.CurrentSchedule != null && _state.CurrentSchedule.Id == scheduleId)
                    {
                        ShowScheduleListView();
                        _state.CurrentSchedule = null;
                    }

                    _state?.ShowNotification?.Invoke($"Schedule '{scheduleToRemove.Name}' deleted", "Schedule Deleted", "info", 4000);
                }
            }
        }

        private void ScheduleBackButton_Click(object sender, RoutedEventArgs e)
        {
            ShowScheduleListView();
        }

        // ==================== GRID GENERATION ====================

        private void GenerateScheduleDetailGrid()
        {
            ScheduleDetailGrid.Children.Clear();
            ScheduleDetailGrid.RowDefinitions.Clear();
            ScheduleDetailGrid.ColumnDefinitions.Clear();

            if (_state.CurrentSchedule == null) return;

            var currentSchedule = _state.CurrentSchedule;
            int scheduleDays = GetScheduleDayCount(currentSchedule);
            double shiftLength = currentSchedule.ShiftLengthHours;
            List<DateTime> scheduleDates = GetScheduleDateList(currentSchedule);
            int shiftIntervals = currentSchedule.ShiftIntervals;
            int originalDays = currentSchedule.OriginalDayCount > 0 ? currentSchedule.OriginalDayCount : scheduleDays;
            int originalIntervals = currentSchedule.OriginalShiftIntervals > 0 ? currentSchedule.OriginalShiftIntervals : shiftIntervals;

            // Compact grid dimensions
            double timeColumnWidth = 100;
            double dayColumnWidth = 140;

            ScheduleDetailGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(timeColumnWidth) });

            for (int day = 0; day < scheduleDays; day++)
            {
                ScheduleDetailGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(dayColumnWidth) });
            }

            ScheduleDetailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) });

            for (int interval = 0; interval < shiftIntervals; interval++)
            {
                // Adaptive row height: estimate total lines needed based on name lengths
                int maxEstimatedLines = 0;
                for (int day = 0; day < scheduleDays; day++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    if (currentSchedule.CellAssignments.ContainsKey(cellId))
                    {
                        var names = currentSchedule.CellAssignments[cellId];
                        int estimatedLines = 0;
                        foreach (var name in names)
                        {
                            // Estimate lines: ~14 chars per line at 11px font in ~120px usable width
                            int nameLines = Math.Max(1, (int)Math.Ceiling((double)name.Length / 14));
                            estimatedLines += nameLines;
                        }
                        if (estimatedLines > maxEstimatedLines) maxEstimatedLines = estimatedLines;
                    }
                }

                double rowHeight = Math.Max(36, maxEstimatedLines * 18 + 10);
                ScheduleDetailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(rowHeight) });
            }

            // Day headers
            for (int day = 0; day < scheduleDays; day++)
            {
                bool isOriginalDay = day < originalDays;
                string headerCellId = $"dayheader_{day}";
                List<string> headerNames = new List<string>();
                if (currentSchedule.CellAssignments.ContainsKey(headerCellId))
                {
                    headerNames = currentSchedule.CellAssignments[headerCellId];
                }

                string dayHeaderText;
                Brush headerBackground;
                Brush headerForeground;

                if (isOriginalDay && day < scheduleDates.Count)
                {
                    DateTime dayDate = scheduleDates[day];
                    dayHeaderText = $"{dayDate.ToString("ddd")} {dayDate.ToString("MM/dd")}";
                    headerBackground = new LinearGradientBrush(
                        (Color)ColorConverter.ConvertFromString("#6366F1"),
                        (Color)ColorConverter.ConvertFromString("#4F46E5"),
                        90);
                    headerForeground = Brushes.White;
                }
                else
                {
                    if (headerNames.Count > 0)
                    {
                        dayHeaderText = string.Join("\n", headerNames.Take(2));
                        headerBackground = new LinearGradientBrush(
                            (Color)ColorConverter.ConvertFromString("#34D399"),
                            (Color)ColorConverter.ConvertFromString("#10B981"),
                            90);
                        headerForeground = Brushes.White;
                    }
                    else
                    {
                        dayHeaderText = "+";
                        headerBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                        headerForeground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));
                    }
                }

                Border headerBorder = new Border
                {
                    Background = headerBackground,
                    CornerRadius = new CornerRadius(6, 6, 0, 0),
                    BorderBrush = isOriginalDay ? Brushes.Transparent : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                    BorderThickness = isOriginalDay ? new Thickness(0) : new Thickness(1),
                    Padding = new Thickness(6),
                    Tag = headerCellId,
                    Cursor = Cursors.Hand
                };

                if (isOriginalDay)
                {
                    TextBlock dayHeader = new TextBlock
                    {
                        Text = dayHeaderText,
                        FontSize = 11,
                        FontWeight = FontWeights.SemiBold,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = headerForeground,
                        TextAlignment = TextAlignment.Center
                    };
                    headerBorder.Child = dayHeader;
                    headerBorder.ToolTip = "Day header (not editable)";
                }
                else
                {
                    StackPanel headerContent = new StackPanel();

                    if (headerNames.Count > 0)
                    {
                        foreach (var name in headerNames)
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = 12,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = headerForeground,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 2, 0, 2)
                            };
                            headerContent.Children.Add(nameText);
                        }

                        headerBorder.ToolTip = $"Assigned: {string.Join(", ", headerNames)}";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = dayHeaderText,
                            FontSize = 16,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = headerForeground
                        };
                        headerContent.Children.Add(emptyText);
                        headerBorder.ToolTip = "Right-click to edit names";
                    }

                    headerBorder.Child = headerContent;
                    headerBorder.MouseRightButtonDown += (sender, e) =>
                    {
                        _editingCellId = headerCellId;
                        OpenEditNamesPanel(headerCellId);
                        e.Handled = true;
                    };
                }

                Grid.SetColumn(headerBorder, day + 1);
                Grid.SetRow(headerBorder, 0);
                ScheduleDetailGrid.Children.Add(headerBorder);
            }

            // Time headers
            for (int interval = 0; interval < shiftIntervals; interval++)
            {
                bool isOriginalInterval = interval < originalIntervals;
                string timeHeaderCellId = $"timeheader_{interval}";
                List<string> headerNames = new List<string>();
                if (currentSchedule.CellAssignments.ContainsKey(timeHeaderCellId))
                {
                    headerNames = currentSchedule.CellAssignments[timeHeaderCellId];
                }

                string timeHeaderText;
                Brush timeForeground;
                Brush timeBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));

                if (isOriginalInterval)
                {
                    double startTimeInHours = currentSchedule.OpeningHour + (interval * shiftLength);
                    double endTimeInHours = Math.Min(startTimeInHours + shiftLength, currentSchedule.ClosingHour);
                    string startTimeLabel = FormatTimeFromHour(startTimeInHours);
                    string endTimeLabel = FormatTimeFromHour(endTimeInHours);
                    timeHeaderText = $"{startTimeLabel}\nto\n{endTimeLabel}";
                    timeForeground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7280"));
                }
                else
                {
                    if (headerNames.Count > 0)
                    {
                        timeHeaderText = string.Join("\n", headerNames.Take(2));
                        timeForeground = Brushes.White;
                        timeBackground = new LinearGradientBrush(
                            (Color)ColorConverter.ConvertFromString("#34D399"),
                            (Color)ColorConverter.ConvertFromString("#10B981"),
                            90);
                    }
                    else
                    {
                        timeHeaderText = "+";
                        timeForeground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));
                    }
                }

                Border timeBorder = new Border
                {
                    Background = timeBackground,
                    BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(6),
                    Tag = timeHeaderCellId,
                    Cursor = Cursors.Hand
                };

                if (isOriginalInterval)
                {
                    TextBlock timeText = new TextBlock
                    {
                        Text = timeHeaderText,
                        FontSize = 12,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = timeForeground,
                        TextAlignment = TextAlignment.Center
                    };
                    timeBorder.Child = timeText;
                    timeBorder.ToolTip = "Time header (not editable)";
                }
                else
                {
                    StackPanel timeContent = new StackPanel();

                    if (headerNames.Count > 0)
                    {
                        foreach (var name in headerNames)
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = 12,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = timeForeground,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 2, 0, 2)
                            };
                            timeContent.Children.Add(nameText);
                        }

                        timeBorder.ToolTip = $"Assigned: {string.Join(", ", headerNames)}";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = timeHeaderText,
                            FontSize = 16,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = timeForeground
                        };
                        timeContent.Children.Add(emptyText);
                        timeBorder.ToolTip = "Right-click to edit names";
                    }

                    timeBorder.Child = timeContent;
                    timeBorder.MouseRightButtonDown += (sender, e) =>
                    {
                        _editingCellId = timeHeaderCellId;
                        OpenEditNamesPanel(timeHeaderCellId);
                        e.Handled = true;
                    };
                }

                Grid.SetColumn(timeBorder, 0);
                Grid.SetRow(timeBorder, interval + 1);
                ScheduleDetailGrid.Children.Add(timeBorder);
            }

            // Data cells
            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < shiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    List<string> names = new List<string>();

                    if (currentSchedule.CellAssignments.ContainsKey(cellId))
                    {
                        names = currentSchedule.CellAssignments[cellId];
                    }

                    // Soft, muted cell colors based on fill status
                    Brush cellBackground;
                    Brush textColor;

                    if (names.Count == 0)
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));
                    }
                    else if (names.Count < currentSchedule.PeoplePerShift)
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEF9C3"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#854D0E"));
                    }
                    else
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCFCE7"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#166534"));
                    }

                    Border cell = new Border
                    {
                        Background = cellBackground,
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        CornerRadius = new CornerRadius(4),
                        Padding = new Thickness(8),
                        Tag = cellId,
                        Cursor = Cursors.Hand,
                        ContextMenu = (ContextMenu)FindResource("CellContextMenu")
                    };

                    if (names.Count > 0)
                    {
                        // Dynamic font size: smaller when more names to fit compactly
                        double nameFontSize = names.Count <= 2 ? 11 : (names.Count <= 4 ? 10 : 9);

                        // Stack names vertically, compact spacing
                        StackPanel nameStack = new StackPanel
                        {
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        foreach (var name in names)
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = nameFontSize,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = textColor,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 1, 0, 1)
                            };
                            nameStack.Children.Add(nameText);
                        }

                        cell.Child = nameStack;
                        cell.ToolTip = $"Assigned: {string.Join(", ", names)}\nRight-click to edit";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = "+",
                            FontSize = 16,
                            FontWeight = FontWeights.Light,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"))
                        };
                        cell.Child = emptyText;
                        cell.ToolTip = "Right-click to add names";
                    }

                    cell.MouseRightButtonDown += (sender, e) =>
                    {
                        _editingCellId = cellId;
                        e.Handled = true;
                    };

                    Grid.SetColumn(cell, day + 1);
                    Grid.SetRow(cell, interval + 1);
                    ScheduleDetailGrid.Children.Add(cell);
                }
            }

        }

        // ==================== MOUSE WHEEL HORIZONTAL SCROLL ====================

        private void ScheduleDetailScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is ScrollViewer scrollViewer)
            {
                // Convert vertical mouse wheel scroll to horizontal scroll
                scrollViewer.ScrollToHorizontalOffset(scrollViewer.HorizontalOffset - e.Delta);
                e.Handled = true;
            }
        }

        // ==================== RIGHT-CLICK CONTEXT MENU ====================

        private void ChangeNamesMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Parent is ContextMenu contextMenu)
            {
                if (contextMenu.PlacementTarget is Border cell && cell.Tag is string cellId)
                {
                    _editingCellId = cellId;
                    OpenEditNamesPanel(cellId);
                }
            }
        }

        // ==================== SIDE PANEL EDITING ====================

        private void OpenEditNamesPanel(string cellId)
        {
            if (_state.CurrentSchedule == null) return;

            _editingCellId = cellId;

            // Get cell info for subtitle
            string subtitle = GetCellDescription(cellId);
            EditNamesSubtitle.Text = subtitle;

            // Load current names
            _editingNames = new List<string>();
            if (_state.CurrentSchedule.CellAssignments.ContainsKey(cellId))
            {
                _editingNames = new List<string>(_state.CurrentSchedule.CellAssignments[cellId]);
            }

            RefreshEditNamesList();

            // Show panel with slide-in animation
            EditNamesOverlay.Visibility = Visibility.Visible;
            EditNamesPanelTransform.X = 440;
            var slideIn = new DoubleAnimation
            {
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3),
                DecelerationRatio = 0.9
            };
            EditNamesPanelTransform.BeginAnimation(TranslateTransform.XProperty, slideIn);

            NewNameTextBox.Focus();
        }

        private string GetCellDescription(string cellId)
        {
            if (_state.CurrentSchedule == null) return "";

            var currentSchedule = _state.CurrentSchedule;

            if (cellId.StartsWith("dayheader_"))
            {
                if (int.TryParse(cellId.Replace("dayheader_", ""), out int dayIndex))
                {
                    return $"New Day Header (Column {dayIndex + 1})";
                }
            }
            else if (cellId.StartsWith("timeheader_"))
            {
                if (int.TryParse(cellId.Replace("timeheader_", ""), out int intervalIndex))
                {
                    return $"New Time Header (Row {intervalIndex + 1})";
                }
            }
            else
            {
                string[] parts = cellId.Split('_');
                if (parts.Length >= 3 && int.TryParse(parts[1], out int day) && int.TryParse(parts[2], out int interval))
                {
                    var scheduleDates = GetScheduleDateList(currentSchedule);
                    if (day < scheduleDates.Count)
                    {
                        DateTime cellDate = scheduleDates[day];
                        string[] dayNames = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

                        double startTime = currentSchedule.OpeningHour + (interval * currentSchedule.ShiftLengthHours);
                        double endTime = Math.Min(startTime + currentSchedule.ShiftLengthHours, currentSchedule.ClosingHour);

                        string startTimeStr = FormatTimeFromHour(startTime);
                        string endTimeStr = FormatTimeFromHour(endTime);

                        return $"{dayNames[(int)cellDate.DayOfWeek]}, {cellDate:MMM dd} • {startTimeStr} to {endTimeStr}";
                    }
                }
            }

            return "";
        }

        private void RefreshEditNamesList()
        {
            EditNamesList.ItemsSource = null;
            EditNamesList.ItemsSource = _editingNames;
            NoNamesText.Visibility = _editingNames.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void AddNameButton_Click(object sender, RoutedEventArgs e)
        {
            AddNameFromTextBox();
        }

        private void NewNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddNameFromTextBox();
                e.Handled = true;
            }
        }

        private void AddNameFromTextBox()
        {
            string name = NewNameTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (_editingNames.Contains(name))
            {
                _state?.ShowNotification?.Invoke("This name is already assigned to this cell.",
                    "Duplicate Name", "warning", 4000);
                return;
            }

            _editingNames.Add(name);
            NewNameTextBox.Clear();
            RefreshEditNamesList();
            NewNameTextBox.Focus();
        }

        private async void EditNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string oldName)
            {
                string newName = await _state.ShowInputDialog($"Change '{oldName}' to:", "Rename Person", oldName);

                if (string.IsNullOrWhiteSpace(newName) || newName == oldName)
                {
                    return;
                }

                if (_editingNames.Contains(newName))
                {
                    _state?.ShowNotification?.Invoke("This name is already assigned to this cell.",
                        "Duplicate Name", "warning", 4000);
                    return;
                }

                int index = _editingNames.IndexOf(oldName);
                if (index >= 0)
                {
                    _editingNames[index] = newName;
                    RefreshEditNamesList();
                }
            }
        }

        private async void RemoveNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string nameToRemove)
            {
                var result = await _state.ShowConfirmDialog($"Are you sure you want to remove '{nameToRemove}' from this shift?",
                    "Confirm Remove");
                if (!result) return;

                _editingNames.Remove(nameToRemove);
                RefreshEditNamesList();
            }
        }

        private void ConfirmEditNames_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null || string.IsNullOrEmpty(_editingCellId)) return;

            // Apply changes
            if (_editingNames.Count > 0)
            {
                _state.CurrentSchedule.CellAssignments[_editingCellId] = new List<string>(_editingNames);
            }
            else
            {
                _state.CurrentSchedule.CellAssignments.Remove(_editingCellId);
            }

            // Save and refresh
            _state.StateSave.SaveSchedule(_state.CurrentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            UpdateScheduleCompletionStatus(_state.CurrentSchedule);

            _state?.ShowNotification?.Invoke("Shift assignments updated", "Changes Saved", "success", 4000);

            CloseEditNamesPanel();
        }

        private void CancelEditNames_Click(object sender, RoutedEventArgs e)
        {
            CloseEditNamesPanel();
        }

        private void CloseEditNamesPanel()
        {
            if (EditNamesOverlay.Visibility != Visibility.Visible || _isEditPanelClosing)
            {
                return;
            }

            _isEditPanelClosing = true;

            var slideOut = new DoubleAnimation
            {
                To = 440,
                Duration = TimeSpan.FromSeconds(0.3),
                AccelerationRatio = 0.9
            };

            slideOut.Completed += (s, e) =>
            {
                EditNamesOverlay.Visibility = Visibility.Collapsed;
                _editingCellId = null;
                _editingNames.Clear();
                NewNameTextBox.Clear();
                _isEditPanelClosing = false;
            };

            EditNamesPanelTransform.BeginAnimation(TranslateTransform.XProperty, slideOut);
        }

        // ==================== STATISTICS ====================

        private void UpdateScheduleStatistics()
        {
            if (_state.CurrentSchedule != null)
            {
                var currentSchedule = _state.CurrentSchedule;
                int totalPeople = 0;
                int totalShifts = 0;

                foreach (var cell in currentSchedule.CellAssignments.Values)
                {
                    totalPeople += cell.Count;
                    if (cell.Count > 0) totalShifts++;
                }

                int totalCells = 0;
                DateTime currentDate = currentSchedule.StartDate;
                while (currentDate <= currentSchedule.EndDate)
                {
                    if (currentSchedule.IncludeWeekends ||
                        (currentDate.DayOfWeek != DayOfWeek.Saturday &&
                         currentDate.DayOfWeek != DayOfWeek.Sunday))
                    {
                        totalCells += currentSchedule.ShiftIntervals;
                    }
                    currentDate = currentDate.AddDays(1);
                }

                double coverage = totalCells > 0 ? (totalShifts * 100.0) / totalCells : 0;
            }
        }

        // ==================== ANALYTICS ====================

        private string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return name;

            // Strip suffixes like "(starts: 9:00 AM)" or "(ends: 5:00 PM)"
            int parenIndex = name.IndexOf('(');
            if (parenIndex > 0)
            {
                string baseName = name.Substring(0, parenIndex).Trim();
                if (!string.IsNullOrWhiteSpace(baseName))
                {
                    return baseName;
                }
            }

            return name.Trim();
        }

        private void UpdateAnalytics()
        {
            if (_state.CurrentSchedule == null) return;

            var schedule = _state.CurrentSchedule;
            var workload = new Dictionary<string, double>();
            var shiftCounts = new Dictionary<string, int>();

            // Calculate hours per person from all data cells
            int scheduleDays = GetScheduleDayCount(schedule);
            double shiftLength = schedule.ShiftLengthHours;

            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < schedule.ShiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    if (schedule.CellAssignments.ContainsKey(cellId))
                    {
                        var names = schedule.CellAssignments[cellId];
                        foreach (var rawName in names)
                        {
                            string name = NormalizeName(rawName);
                            if (!workload.ContainsKey(name))
                            {
                                workload[name] = 0;
                                shiftCounts[name] = 0;
                            }
                            workload[name] += shiftLength;
                            shiftCounts[name]++;
                        }
                    }
                }
            }

            // Also count header assignments (day/time headers)
            foreach (var kvp in schedule.CellAssignments)
            {
                if (kvp.Key.StartsWith("dayheader_") || kvp.Key.StartsWith("timeheader_"))
                {
                    foreach (var rawName in kvp.Value)
                    {
                        string name = NormalizeName(rawName);
                        if (!workload.ContainsKey(name))
                        {
                            workload[name] = 0;
                            shiftCounts[name] = 0;
                        }
                    }
                }
            }

            double totalHours = workload.Values.Sum();
            int uniquePeople = workload.Count;
            double avgHours = uniquePeople > 0 ? totalHours / uniquePeople : 0;

            // Calculate shift slot stats
            int totalShiftSlots = scheduleDays * schedule.ShiftIntervals;
            int filledShifts = 0;
            int unfilledShifts = 0;
            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < schedule.ShiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    int count = 0;
                    if (schedule.CellAssignments.ContainsKey(cellId))
                    {
                        count = schedule.CellAssignments[cellId].Count;
                    }

                    if (count > 0) filledShifts++;
                    else unfilledShifts++;
                }
            }

            double coveragePct = totalShiftSlots > 0 ? (filledShifts * 100.0) / totalShiftSlots : 0;
            double avgShiftsPerPerson = uniquePeople > 0 ? (double)shiftCounts.Values.Sum() / uniquePeople : 0;

            // Update summary cards
            TotalHoursText.Text = $"{totalHours:F0} hrs";
            AvgHoursText.Text = $"{avgHours:F1} hrs";
            UniquePeopleText.Text = uniquePeople.ToString();
            AvgShiftsPerPersonText.Text = $"{avgShiftsPerPerson:F1}";
            AnalyticsCoverageText.Text = $"{coveragePct:F0}%";
            FilledShiftsText.Text = $"{filledShifts} / {totalShiftSlots}";
            UnfilledShiftsText.Text = unfilledShifts.ToString();

            // Calculate unassigned people (people in the batch/employees who got no shifts)
            int unassignedPeople = 0;
            if (_state.Batches != null && _state.CurrentSchedule != null)
            {
                var batch = _state.Batches.FirstOrDefault(b => b.Id == _state.CurrentSchedule.BatchId);
                if (batch != null)
                {
                    foreach (var employeeId in batch.EmployeeIds)
                    {
                        var employee = _state.CombinedAvailabilities.FirstOrDefault(e => e.Id == employeeId);
                        if (employee != null)
                        {
                            string empName = NormalizeName(employee.Name);
                            if (!workload.ContainsKey(empName))
                            {
                                unassignedPeople++;
                            }
                        }
                    }
                }
            }
            UnassignedPeopleText.Text = unassignedPeople.ToString();

            // Build workload list
            _workloadItems = new List<WorkloadItem>();
            double maxHours = workload.Values.Count > 0 ? workload.Values.Max() : 1;
            double minHours = workload.Values.Count > 0 ? workload.Values.Min() : 0;

            foreach (var kvp in workload)
            {
                double hours = kvp.Value;
                double pctOfMax = maxHours > 0 ? (hours / maxHours) * 100 : 0;
                double pctOfAvg = avgHours > 0 ? (hours / avgHours) * 100 : 0;

                string statusLabel;
                string statusColor;
                string barColor;
                string detailText;

                if (hours > avgHours * 1.25)
                {
                    statusLabel = "Overworked";
                    statusColor = "#EF4444";
                    barColor = "#EF4444";
                    detailText = $"{pctOfAvg:F0}% of average • {shiftCounts[kvp.Key]} shifts";
                }
                else if (hours < avgHours * 0.75 && uniquePeople > 1)
                {
                    statusLabel = "Underworked";
                    statusColor = "#F59E0B";
                    barColor = "#F59E0B";
                    detailText = $"{pctOfAvg:F0}% of average • {shiftCounts[kvp.Key]} shifts";
                }
                else
                {
                    statusLabel = "Balanced";
                    statusColor = "#10B981";
                    barColor = "#10B981";
                    detailText = $"{pctOfAvg:F0}% of average • {shiftCounts[kvp.Key]} shifts";
                }

                _workloadItems.Add(new WorkloadItem
                {
                    Name = kvp.Key,
                    Hours = hours,
                    HoursText = $"{hours:F1} hrs",
                    StatusLabel = statusLabel,
                    StatusColor = statusColor,
                    BarColor = barColor,
                    BarWidth = Math.Max(4, pctOfMax),
                    DetailText = detailText
                });
            }

            ApplyWorkloadSort();

            // Generate smart suggestions
            var suggestions = new List<SuggestionItem>();

            // Check for understaffed cells
            int understaffedCells = 0;
            int emptyCells = 0;
            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < schedule.ShiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    int count = 0;
                    if (schedule.CellAssignments.ContainsKey(cellId))
                    {
                        count = schedule.CellAssignments[cellId].Count;
                    }

                    if (count == 0) emptyCells++;
                    else if (count < schedule.PeoplePerShift) understaffedCells++;
                }
            }

            if (emptyCells > 0)
            {
                suggestions.Add(new SuggestionItem
                {
                    Icon = "⚠️",
                    Message = $"{emptyCells} shift(s) have no one assigned. Right-click those cells to add people.",
                    SuggestionBg = "#FEF2F2",
                    SuggestionBorder = "#FECACA"
                });
            }

            if (understaffedCells > 0)
            {
                suggestions.Add(new SuggestionItem
                {
                    Icon = "🔶",
                    Message = $"{understaffedCells} shift(s) are understaffed (below {schedule.PeoplePerShift} people). Consider adding more people to these shifts.",
                    SuggestionBg = "#FFFBEB",
                    SuggestionBorder = "#FDE68A"
                });
            }

            // Check for overworked people
            var overworked = _workloadItems.Where(w => w.StatusLabel == "Overworked").ToList();
            if (overworked.Count > 0)
            {
                string names = string.Join(", ", overworked.Take(3).Select(w => w.Name));
                if (overworked.Count > 3) names += $" and {overworked.Count - 3} more";
                string overworkedVerb = overworked.Count == 1 ? "is" : "are";
                suggestions.Add(new SuggestionItem
                {
                    Icon = "⚡",
                    Message = $"{names} {overworkedVerb} working significantly more than average. Consider redistributing some shifts.",
                    SuggestionBg = "#FEF2F2",
                    SuggestionBorder = "#FECACA"
                });
            }

            // Check for underworked people
            var underworked = _workloadItems.Where(w => w.StatusLabel == "Underworked").ToList();
            if (underworked.Count > 0)
            {
                string names = string.Join(", ", underworked.Take(3).Select(w => w.Name));
                if (underworked.Count > 3) names += $" and {underworked.Count - 3} more";
                string underworkedVerb = underworked.Count == 1 ? "has" : "have";
                suggestions.Add(new SuggestionItem
                {
                    Icon = "💤",
                    Message = $"{names} {underworkedVerb} fewer hours than average. Consider giving them more shifts.",
                    SuggestionBg = "#FFFBEB",
                    SuggestionBorder = "#FDE68A"
                });
            }

            // Check workload balance
            if (uniquePeople > 1 && maxHours > 0 && minHours > 0)
            {
                double ratio = maxHours / minHours;
                if (ratio > 2.0)
                {
                    suggestions.Add(new SuggestionItem
                    {
                        Icon = "⚖️",
                        Message = $"Workload is imbalanced: the busiest person works {ratio:F1}x more hours than the least busy. Consider rebalancing shifts.",
                        SuggestionBg = "#FEF2F2",
                        SuggestionBorder = "#FECACA"
                    });
                }
            }

            // Check coverage
            int totalCells = scheduleDays * schedule.ShiftIntervals;
            int filledCells = totalCells - emptyCells;
            double coverageCheck = totalCells > 0 ? (filledCells * 100.0) / totalCells : 0;

            if (coverageCheck < 50 && totalCells > 0)
            {
                suggestions.Add(new SuggestionItem
                {
                    Icon = "📉",
                    Message = $"Schedule coverage is only {coverageCheck:F0}%. Most shifts are unfilled - consider assigning more people.",
                    SuggestionBg = "#FEF2F2",
                    SuggestionBorder = "#FECACA"
                });
            }

            // Update badge
            if (suggestions.Count == 0)
            {
                AnalyticsSummaryBadge.Text = "balanced";
                AnalyticsSummaryBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#059669"));
                var badgeParent = AnalyticsSummaryBadge.Parent as Border;
                if (badgeParent != null)
                {
                    badgeParent.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1FAE5"));
                }
            }
            else if (suggestions.Count <= 2)
            {
                AnalyticsSummaryBadge.Text = "needs attention";
                AnalyticsSummaryBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D97706"));
                var badgeParent = AnalyticsSummaryBadge.Parent as Border;
                if (badgeParent != null)
                {
                    badgeParent.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEF3C7"));
                }
            }
            else
            {
                AnalyticsSummaryBadge.Text = "needs review";
                AnalyticsSummaryBadge.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626"));
                var badgeParent = AnalyticsSummaryBadge.Parent as Border;
                if (badgeParent != null)
                {
                    badgeParent.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEE2E2"));
                }
            }

            SuggestionsList.ItemsSource = null;
            SuggestionsList.ItemsSource = suggestions;
            NoSuggestionsText.Visibility = suggestions.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

            // ==================== STAFFING ALERTS ====================
            var staffingAlerts = new List<StaffingAlertItem>();

            if (schedule.UnderstaffingAlerts != null)
            {
                foreach (var alert in schedule.UnderstaffingAlerts)
                {
                    // Only show unresolved alerts. Fixed windows were already
                    // resolved by the scheduler, so no need to surface them.
                    if (alert.WasFixed) continue;

                    string dateLabel = $"{alert.Date:ddd} {alert.Date:MM/dd}";
                    string timeLabel = $"{FormatTimeFromHour(alert.Start.TotalHours)} - {FormatTimeFromHour(alert.End.TotalHours)}";
                    string statusLabel = alert.IsUncovered ? "Uncovered" : "Understaffed";
                    string titleColor;
                    string bgColor;
                    string borderColor;

                    if (alert.IsUncovered)
                    {
                        titleColor = "#DC2626";
                        bgColor = "#FEF2F2";
                        borderColor = "#FECACA";
                    }
                    else
                    {
                        titleColor = "#D97706";
                        bgColor = "#FFFBEB";
                        borderColor = "#FDE68A";
                    }

                    string detail = $"{dateLabel}  {timeLabel}  Needs {alert.Required}, has {alert.Actual}";
                    string candidatesText = "";

                    if (!alert.WasFixed && alert.Candidates.Count > 0)
                    {
                        var candidateParts = new List<string>();
                        foreach (var candidate in alert.Candidates)
                        {
                            string availText = "";
                            if (candidate.AvailableFrom.HasValue && candidate.AvailableTo.HasValue)
                            {
                                availText = $" (available {FormatTimeFromHour(candidate.AvailableFrom.Value.TotalHours)} - {FormatTimeFromHour(candidate.AvailableTo.Value.TotalHours)})";
                            }
                            else if (candidate.AvailableFrom.HasValue)
                            {
                                availText = $" (available from {FormatTimeFromHour(candidate.AvailableFrom.Value.TotalHours)})";
                            }
                            else if (candidate.AvailableTo.HasValue)
                            {
                                availText = $" (available until {FormatTimeFromHour(candidate.AvailableTo.Value.TotalHours)})";
                            }
                            candidateParts.Add($"{candidate.Name}{availText}");
                        }
                        candidatesText = "Candidates: " + string.Join(", ", candidateParts);
                    }

                    staffingAlerts.Add(new StaffingAlertItem
                    {
                        AlertTitle = $"{statusLabel}: {dateLabel} {timeLabel}",
                        AlertDetail = detail,
                        AlertCandidates = candidatesText,
                        AlertBg = bgColor,
                        AlertBorder = borderColor,
                        AlertTitleColor = titleColor,
                        CandidatesVisible = !string.IsNullOrEmpty(candidatesText) ? Visibility.Visible : Visibility.Collapsed
                    });
                }
            }

            StaffingAlertsList.ItemsSource = null;
            StaffingAlertsList.ItemsSource = staffingAlerts;
            NoAlertsText.Visibility = staffingAlerts.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ApplyWorkloadSort()
        {
            if (_sortByAlphabet)
            {
                WorkloadList.ItemsSource = null;
                WorkloadList.ItemsSource = _workloadItems.OrderBy(w => w.Name, StringComparer.OrdinalIgnoreCase).ToList();
                SortToggleButton.Content = "Sort: A-Z";
            }
            else
            {
                WorkloadList.ItemsSource = null;
                WorkloadList.ItemsSource = _workloadItems.OrderByDescending(w => w.Hours).ToList();
                SortToggleButton.Content = "Sort: Workload";
            }
        }

        private void SortToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _sortByAlphabet = !_sortByAlphabet;
            ApplyWorkloadSort();
        }

        private void UpdateScheduleCompletionStatus(Schedule schedule)
        {
            if (schedule == null) return;

            bool isComplete = true;
            int scheduleDays = GetScheduleDayCount(schedule);

            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < schedule.ShiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";

                    int assignedCount = 0;
                    if (schedule.CellAssignments.ContainsKey(cellId))
                    {
                        assignedCount = schedule.CellAssignments[cellId].Count;
                    }

                    if (assignedCount < schedule.PeoplePerShift)
                    {
                        isComplete = false;
                    }
                }
            }

            schedule.Status = isComplete ? "Complete" : "Incomplete";

            if (_state.CurrentSchedule != null && _state.CurrentSchedule.Id == schedule.Id)
            {
                _state.CurrentSchedule.Status = schedule.Status;
                UpdateScheduleStatistics();
            }

            _state.StateSave.SaveSchedule(schedule);
            UpdateScheduleList();
        }

        // ==================== HELPERS ====================

        private List<DateTime> GetScheduleDateList(Schedule schedule)
        {
            List<DateTime> dates = new List<DateTime>();
            DateTime currentDate = schedule.StartDate;

            while (currentDate <= schedule.EndDate)
            {
                if (schedule.IncludeWeekends ||
                    (currentDate.DayOfWeek != DayOfWeek.Saturday &&
                     currentDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    dates.Add(currentDate);
                }
                currentDate = currentDate.AddDays(1);
            }

            return dates;
        }

        private int GetScheduleDayCount(Schedule schedule)
        {
            int dayCount = 0;
            DateTime currentDate = schedule.StartDate;

            while (currentDate <= schedule.EndDate)
            {
                if (schedule.IncludeWeekends ||
                    (currentDate.DayOfWeek != DayOfWeek.Saturday &&
                     currentDate.DayOfWeek != DayOfWeek.Sunday))
                {
                    dayCount++;
                }
                currentDate = currentDate.AddDays(1);
            }

            return dayCount;
        }

        private string FormatTimeFromHour(double hour)
        {
            int hourInt = (int)Math.Floor(hour);
            int minutes = (int)Math.Round((hour - hourInt) * 60);

            if (minutes == 60)
            {
                hourInt++;
                minutes = 0;
            }

            string ampm = hourInt >= 12 ? "PM" : "AM";
            int displayHour = hourInt > 12 ? hourInt - 12 : (hourInt == 0 ? 12 : hourInt);

            if (minutes == 0)
                return $"{displayHour}:00 {ampm}";
            else
                return $"{displayHour}:{minutes:D2} {ampm}";
        }

        // ==================== ROW/COLUMN MANAGEMENT ====================

        private async void ExportPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null)
            {
                _state?.ShowNotification?.Invoke("Please select a schedule to export.", "No Schedule", "warning", 4000);
                return;
            }

            try
            {
                var newExport = ExportService.CreateExportItem(_state.CurrentSchedule);
                _state?.ShowNotification?.Invoke($"Schedule '{_state.CurrentSchedule.Name}' added to exports!", "Export Ready", "success", 4000);
                _state?.ShowExportPage?.Invoke();
            }
            catch (Exception ex)
            {
                _state?.ShowNotification?.Invoke($"Error creating export: {ex.Message}", "Export Error", "error", 4000);
            }
        }

        private async void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null) return;

            var result = await _state.ShowConfirmDialog(
                "Are you sure you want to add a new row to the schedule? This will add a new shift interval at the bottom.",
                "Add Row Confirmation"
            );

            if (!result) return;

            _state.CurrentSchedule.ShiftIntervals++;
            _state.StateSave.SaveSchedule(_state.CurrentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            _state?.ShowNotification?.Invoke("New row added at the bottom of schedule", "Row Added", "success", 4000);
        }

        private async void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null) return;

            var result = await _state.ShowConfirmDialog(
                "Are you sure you want to add a new column to the schedule? This will add a new day at the end.",
                "Add Column Confirmation"
            );

            if (!result) return;

            _state.CurrentSchedule.EndDate = _state.CurrentSchedule.EndDate.AddDays(1);

            if (!_state.CurrentSchedule.IncludeWeekends)
            {
                while (_state.CurrentSchedule.EndDate.DayOfWeek == DayOfWeek.Saturday ||
                       _state.CurrentSchedule.EndDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    _state.CurrentSchedule.EndDate = _state.CurrentSchedule.EndDate.AddDays(1);
                }
            }

            _state.StateSave.SaveSchedule(_state.CurrentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            _state?.ShowNotification?.Invoke("New column added at the end of schedule", "Column Added", "info", 4000);
        }

        private async void DeleteRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null || _state.CurrentSchedule.ShiftIntervals <= 1) return;

            var result = await _state.ShowConfirmDialog(
                $"Are you sure you want to delete the last row? This will remove shift interval {_state.CurrentSchedule.ShiftIntervals}.",
                "Delete Row Confirmation"
            );

            if (!result) return;

            int scheduleDays = GetScheduleDayCount(_state.CurrentSchedule);

            for (int day = 0; day < scheduleDays; day++)
            {
                string lastCellId = $"cell_{day}_{_state.CurrentSchedule.ShiftIntervals - 1}";
                _state.CurrentSchedule.CellAssignments.Remove(lastCellId);
            }

            _state.CurrentSchedule.ShiftIntervals--;
            _state.StateSave.SaveSchedule(_state.CurrentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            _state?.ShowNotification?.Invoke("Last row deleted from schedule", "Row Deleted", "info", 4000);
        }

        private async void DeleteColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state.CurrentSchedule == null) return;

            int currentDays = GetScheduleDayCount(_state.CurrentSchedule);

            if (currentDays <= 1) return;

            var result = await _state.ShowConfirmDialog(
                $"Are you sure you want to delete the last column? This will remove day {currentDays}.",
                "Delete Column Confirmation"
            );

            if (!result) return;

            int dayToRemove = currentDays - 1;

            for (int interval = 0; interval < _state.CurrentSchedule.ShiftIntervals; interval++)
            {
                string cellId = $"cell_{dayToRemove}_{interval}";
                _state.CurrentSchedule.CellAssignments.Remove(cellId);
            }

            DateTime newEndDate = _state.CurrentSchedule.EndDate;
            do
            {
                newEndDate = newEndDate.AddDays(-1);
            } while (!_state.CurrentSchedule.IncludeWeekends &&
                     (newEndDate.DayOfWeek == DayOfWeek.Saturday ||
                      newEndDate.DayOfWeek == DayOfWeek.Sunday));

            _state.CurrentSchedule.EndDate = newEndDate;
            _state.StateSave.SaveSchedule(_state.CurrentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            _state?.ShowNotification?.Invoke("Last column deleted from schedule", "Column Deleted", "info", 4000);
        }
    }
}