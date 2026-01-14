using Microsoft.Win32;
using SchedulerApp.Services;
using SchedulerApp.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace SchedulerApp
{
    public partial class MainWindow : Window
    {
        private Color currentBackgroundColor = Colors.White;
        private List<Employee> employees = new List<Employee>();
        private Employee currentEmployee = null;
        private bool isSelecting = false;
        private List<Border> selectedCells = new List<Border>();
        private Dictionary<string, bool[,]> employeeSchedules = new Dictionary<string, bool[,]>();
        private List<DateTime> selectedDates = new List<DateTime>();
        private List<AvailabilityEntry> combinedAvailabilities = new List<AvailabilityEntry>();
        private List<Batch> batches = new List<Batch>();
        private List<Schedule> schedules = new List<Schedule>();
        private Schedule currentSchedule = null;
        private string selectedCellId = null;
        private int days = 0;
        private int hours = 0;
        private string[] timeNames = { "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM", "11:00 AM",
                                       "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM",
                                       "6:00 PM", "7:00 PM", "8:00 PM", "9:00 PM", "10:00 PM", "11:00 PM", "12:00 AM", "1:00 AM" };
        private DateTime currentMonth = DateTime.Today;
        private Dictionary<Button, DateTime> calendarButtons = new Dictionary<Button, DateTime>();
        private DateTime? calendarDragStartDate = null;
        private bool isCalendarDragging = false;
        private bool isAdditiveSelection = false;
        private bool isLettuceMeetImport = false;
        private WebScraper.EventData lettuceMeetData = null;
        private string importedFilePath = "";
        private SolidColorBrush selectedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5"));
        private SolidColorBrush hoverBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB"));
        private SolidColorBrush todayBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));
        private SolidColorBrush normalBrush = new SolidColorBrush(Colors.Transparent);
        private SolidColorBrush whiteText = new SolidColorBrush(Colors.White);
        private SolidColorBrush darkText = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2937"));
        private SolidColorBrush grayText = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7280"));
        private bool includeWeekends = false;
        private StateSave stateSave;
        private int currentPeoplePerShift = 1;
        private Views.ExportPage ExportPageContent { get; set; }
        private PopupSystem popupSystem;

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
        }

        public MainWindow()
        {
            InitializeComponent();
            stateSave = new StateSave();
            popupSystem = (PopupSystem)this.FindName("PopupSystemControl");
            LoadSavedData();
            InitializeApplication();
            GridScrollViewer.PreviewMouseWheel += GridScrollViewer_PreviewMouseWheel;
            ScheduleDetailScrollViewer.PreviewMouseWheel += ScheduleDetailScrollViewer_PreviewMouseWheel;
            ExportPageContent = (Views.ExportPage)this.FindName("ExportPageControl");
        }

        private void LoadSavedData()
        {
            var savedEmployees = stateSave.LoadEmployees();
            if (savedEmployees.Count > 0)
            {
                combinedAvailabilities = savedEmployees;
                UpdateCombinedAvailabilities();
                UpdateStatistics();
                employees.Clear();
                foreach (var emp in combinedAvailabilities.Where(e => e.Source == "Manual"))
                {
                    if (!employees.Any(e => e.Id == emp.Id))
                    {
                        employees.Add(emp);
                    }
                }
                UpdateEmployeesList();
            }

            var savedBatches = stateSave.LoadBatches();
            if (savedBatches.Count > 0)
            {
                batches = savedBatches;
                UpdateBatchComboBox();
                UpdateBatchList();
            }

            var savedSchedules = stateSave.LoadSchedules();
            if (savedSchedules.Count > 0)
            {
                schedules = savedSchedules;
                UpdateScheduleList();
            }
        }

        private void InitializeApplication()
        {
            PopulateDropdowns();
            SetupNavButton.Tag = "active";
            StartTimeBox.SelectedIndex = 0;
            EndTimeBox.SelectedIndex = 11;
            UpdateGridHint();
            InitializeCalendar();
            ImportSourceBox.SelectedIndex = 0;
            FileUploadArea.Visibility = Visibility.Visible;
            LettuceMeetInputArea.Visibility = Visibility.Collapsed;
            if (batches.Count == 0)
            {
                InitializeBatches();
            }
            UpdateSelectedBatchText();
            UpdateBackButtonVisibility();
        }

        private void InitializeBatches()
        {
            batches.Clear();
            UpdateBatchComboBox();
            UpdateBatchList();
        }

        private void UpdateBatchComboBox()
        {
            BatchComboBox.Items.Clear();
            if (batches.Count == 0)
            {
                BatchComboBox.Items.Add("No batches available");
                BatchComboBox.SelectedIndex = 0;
                BatchComboBox.IsEnabled = false;
                return;
            }
            BatchComboBox.IsEnabled = true;
            foreach (var batch in batches)
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

        private void UpdateSelectedBatchText()
        {
        }

        private void UpdateBatchList()
        {
            BatchListControl.ItemsSource = null;
            BatchListControl.ItemsSource = batches;
            NoBatchesText.Visibility = batches.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateScheduleList()
        {
            ScheduleListControl.ItemsSource = null;
            ScheduleListControl.ItemsSource = schedules;
            NoSchedulesText.Visibility = schedules.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateGridHint()
        {
            GridHintText.Text = "Each selected date will create a column in the schedule grid, and the time range will create the rows. You can select dates by: 1) Clicking individual dates, 2) Holding Ctrl while clicking to select multiple non-consecutive dates, or 3) Clicking and dragging to select a range of consecutive dates. Selected dates will be displayed below.";
        }

        private void InitializeCalendar()
        {
            UpdateCalendarDisplay();
        }

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
                dayButton.Background = hoverBrush;
            }
        }

        private void DayButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sender is Button dayButton && dayButton.Tag is DateTime date)
            {
                DateTime today = DateTime.Today;
                UpdateDayButtonAppearance(dayButton, date, today);
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
            DateTime today = DateTime.Today;
            foreach (var kvp in calendarButtons)
            {
                UpdateDayButtonAppearance(kvp.Key, kvp.Value, today);
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
                    Foreground = grayText,
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

        private void Setup_Click(object sender, RoutedEventArgs e)
        {
            SetupPage.Visibility = Visibility.Visible;
            ManualSchedulePage.Visibility = Visibility.Collapsed;
            ImportDataPage.Visibility = Visibility.Collapsed;
            ViewAvailabilitiesPage.Visibility = Visibility.Collapsed;
            SchedulePage.Visibility = Visibility.Collapsed;
            ExportPage.Visibility = Visibility.Collapsed;
            UpdateNavigationButtons(SetupNavButton);
            UpdateBackButtonVisibility();
        }

        private void ManualSchedule_Click(object sender, RoutedEventArgs e)
        {
            SetupPage.Visibility = Visibility.Collapsed;
            ManualSchedulePage.Visibility = Visibility.Visible;
            ImportDataPage.Visibility = Visibility.Collapsed;
            ViewAvailabilitiesPage.Visibility = Visibility.Collapsed;
            SchedulePage.Visibility = Visibility.Collapsed;
            ExportPage.Visibility = Visibility.Collapsed;
            RangeSelectionCard.Visibility = Visibility.Visible;
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            UpdateSelectedDatesDisplay();
            UpdateNavigationButtons(ManualScheduleNavButton);
            UpdateBackButtonVisibility();
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            SetupPage.Visibility = Visibility.Collapsed;
            ManualSchedulePage.Visibility = Visibility.Collapsed;
            ImportDataPage.Visibility = Visibility.Visible;
            ViewAvailabilitiesPage.Visibility = Visibility.Collapsed;
            SchedulePage.Visibility = Visibility.Collapsed;
            ExportPage.Visibility = Visibility.Collapsed;
            UpdateNavigationButtons(ImportDataNavButton);
            UpdateBackButtonVisibility();
        }

        private void ViewAvailabilities_Click(object sender, RoutedEventArgs e)
        {
            SetupPage.Visibility = Visibility.Collapsed;
            ManualSchedulePage.Visibility = Visibility.Collapsed;
            ImportDataPage.Visibility = Visibility.Collapsed;
            ViewAvailabilitiesPage.Visibility = Visibility.Visible;
            SchedulePage.Visibility = Visibility.Collapsed;
            ExportPage.Visibility = Visibility.Collapsed;
            UpdateCombinedAvailabilities();
            UpdateStatistics();
            UpdateBatchList();
            UpdateNavigationButtons(ViewAvailabilitiesNavButton);
            UpdateBackButtonVisibility();
        }

        private void Schedule_Click(object sender, RoutedEventArgs e)
        {
            SetupPage.Visibility = Visibility.Collapsed;
            ManualSchedulePage.Visibility = Visibility.Collapsed;
            ImportDataPage.Visibility = Visibility.Collapsed;
            ViewAvailabilitiesPage.Visibility = Visibility.Collapsed;
            SchedulePage.Visibility = Visibility.Visible;
            ExportPage.Visibility = Visibility.Collapsed;
            ShowScheduleListView();
            UpdateNavigationButtons(ScheduleNavButton);
            UpdateBackButtonVisibility();
        }

        private void UpdateNavigationButtons(Button activeButton)
        {
            SetupNavButton.Style = (Style)FindResource("SidebarButton");
            ManualScheduleNavButton.Style = (Style)FindResource("SidebarButton");
            ImportDataNavButton.Style = (Style)FindResource("SidebarButton");
            ViewAvailabilitiesNavButton.Style = (Style)FindResource("SidebarButton");
            ScheduleNavButton.Style = (Style)FindResource("SidebarButton");
            ExportNavButton.Style = (Style)FindResource("SidebarButton");
            activeButton.Style = (Style)FindResource("ActiveSidebarButton");
        }

        private void UpdateBackButtonVisibility()
        {
            SetupBackButton.Visibility = Visibility.Collapsed;
            ManualBackButton.Visibility = Visibility.Collapsed;
            ImportBackButton.Visibility = Visibility.Collapsed;
            AvailabilitiesBackButton.Visibility = Visibility.Collapsed;
            ScheduleBackButton.Visibility = Visibility.Collapsed;

            if (ManualSchedulePage.Visibility == Visibility.Visible && ScheduleGridCard.Visibility == Visibility.Visible)
            {
                ManualBackButton.Visibility = Visibility.Visible;
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (ManualSchedulePage.Visibility == Visibility.Visible)
            {
                if (ScheduleGridCard.Visibility == Visibility.Visible)
                {
                    ScheduleGridCard.Visibility = Visibility.Collapsed;
                    RangeSelectionCard.Visibility = Visibility.Visible;
                    selectedCells.Clear();
                    currentEmployee = null;
                    EmployeeNameBox.Text = "";
                    UpdateBackButtonVisibility();
                }
            }
        }

        private void BackToDatesButton_Click(object sender, RoutedEventArgs e)
        {
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            RangeSelectionCard.Visibility = Visibility.Visible;
            BackToDatesButton.Visibility = Visibility.Collapsed;
            selectedCells.Clear();
            currentEmployee = null;
            EmployeeNameBox.Text = "";
            UpdateBackButtonVisibility();
        }

        private void ScheduleBackButton_Click(object sender, RoutedEventArgs e)
        {
            ShowScheduleListView();
        }

        private void ShowScheduleListView()
        {
            ScheduleListView.Visibility = Visibility.Visible;
            ScheduleDetailView.Visibility = Visibility.Collapsed;
            ScheduleBackButton.Visibility = Visibility.Collapsed;
            selectedCellId = null;
            ClearCellSelection();
        }

        private void ShowScheduleDetailView(Schedule schedule)
        {
            currentSchedule = schedule;
            ScheduleListView.Visibility = Visibility.Collapsed;
            ScheduleDetailView.Visibility = Visibility.Visible;
            ScheduleBackButton.Visibility = Visibility.Visible;
            ScheduleDetailTitle.Text = schedule.Name;
            CurrentScheduleName.Text = schedule.Name;
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            ClearCellSelection();
        }

        private async void CreateGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (selectedDates.Count == 0)
            {
                popupSystem.ShowNotification("Please select at least one date.", "No Dates Selected", "warning");
                return;
            }

            selectedDates = selectedDates.OrderBy(d => d.Date).ToList();

            if (StartTimeBox.SelectedIndex == -1 || EndTimeBox.SelectedIndex == -1)
            {
                popupSystem.ShowNotification("Please select both start and end times.", "Missing Times", "warning");
                return;
            }

            int startHourIndex = StartTimeBox.SelectedIndex;
            int endHourIndex = EndTimeBox.SelectedIndex;

            if (endHourIndex <= startHourIndex)
            {
                popupSystem.ShowNotification("End time must be after start time.", "Invalid Time Range", "warning");
                return;
            }

            days = selectedDates.Count;
            hours = (endHourIndex - startHourIndex) + 1;
            string startTimeStr = StartTimeBox.SelectedItem?.ToString();
            string endTimeStr = EndTimeBox.SelectedItem?.ToString();
            string dateRangeText = days == 1 ?
                $"{selectedDates[0]:MMM dd}" :
                $"{selectedDates[0]:MMM dd} to {selectedDates[selectedDates.Count - 1]:MMM dd}";

            RangeSelectionCard.Visibility = Visibility.Collapsed;
            ScheduleGridCard.Visibility = Visibility.Visible;
            BackToDatesButton.Visibility = Visibility.Visible;
            ManualBackButton.Visibility = Visibility.Collapsed;
            CurrentEmployeeText.Text = "for: New Employee";
            EmployeeNameBox.Text = "";
            selectedCells.Clear();
            GenerateScheduleGrid();
            popupSystem.ShowNotification($"Grid created with half-hour intervals! ({hours * 2} time slots per day)", "Ready to Add Employee", "success");
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

        private async void SaveAvailability_Click(object sender, RoutedEventArgs e)
        {
            string name = EmployeeNameBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                popupSystem.ShowNotification("Please enter employee name.", "Missing Information", "warning");
                return;
            }

            if (selectedCells.Count == 0)
            {
                popupSystem.ShowNotification("Please select at least one time slot for availability.", "No Selection", "warning");
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

            var timeRanges = new List<string>();
            foreach (var dateIndex in Enumerable.Range(0, daysCount))
            {
                var date = selectedDates[dateIndex];
                for (int i = 0; i < halfHourCount; i++)
                {
                    if (scheduleMatrix[dateIndex, i])
                    {
                        int start = i;
                        while (i + 1 < halfHourCount && scheduleMatrix[dateIndex, i + 1])
                        {
                            i++;
                        }
                        int end = i + 1;
                        int startHourIndex = StartTimeBox.SelectedIndex;
                        int startHour = 6 + startHourIndex;
                        int startTotalMinutes = (startHour * 60) + (start * 30);
                        int startHourFinal = startTotalMinutes / 60;
                        int startMinute = startTotalMinutes % 60;
                        int endTotalMinutes = (startHour * 60) + (end * 30);
                        int endHourFinal = endTotalMinutes / 60;
                        int endMinute = endTotalMinutes % 60;
                        timeRanges.Add($"{date:yyyy-MM-dd} {startHourFinal:00}:{startMinute:00}-{endHourFinal:00}:{endMinute:00}");
                    }
                }
            }

            int actualStartHour = 6 + StartTimeBox.SelectedIndex;
            int actualEndHour = 6 + EndTimeBox.SelectedIndex + 1;
            currentEmployee = new Employee
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
                EndHour = actualEndHour
            };

            int selectedSlots = selectedCells.Count;
            int totalSlots = daysCount * halfHourCount;
            double percentage = (selectedSlots * 100.0) / totalSlots;
            currentEmployee.AvailabilitySummary =
                $"{selectedSlots}/{totalSlots} half-hour slots ({percentage:F1}% availability)";
            currentEmployee.SlotCount = selectedSlots;
            employees.Add(currentEmployee);

            var availabilityEntry = new AvailabilityEntry
            {
                Id = currentEmployee.Id,
                Name = currentEmployee.Name,
                DateRange = currentEmployee.DateRange,
                TimeRange = currentEmployee.TimeRange,
                AvailabilitySummary = currentEmployee.AvailabilitySummary,
                SlotCount = currentEmployee.SlotCount,
                Source = "Manual",
                SourceColor = currentEmployee.SourceColor,
                CreatedDate = DateTime.Now,
                StartDate = currentEmployee.StartDate,
                EndDate = currentEmployee.EndDate,
                StartHour = currentEmployee.StartHour,
                EndHour = currentEmployee.EndHour,
                ScheduleMatrix = scheduleMatrix,
                SelectedDates = new List<DateTime>(selectedDates)
            };

            combinedAvailabilities.Add(availabilityEntry);
            stateSave.SaveEmployee(availabilityEntry);
            UpdateEmployeesList();
            EmployeeNameBox.Text = "";
            currentEmployee = null;
            ScheduleGridCard.Visibility = Visibility.Collapsed;
            RangeSelectionCard.Visibility = Visibility.Visible;
            selectedCells.Clear();
            popupSystem.ShowNotification($"Availability saved for {name}! ({selectedSlots} half-hour slots)", "Success", "success");
            UpdateBackButtonVisibility();
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (Border cell in selectedCells)
            {
                cell.Style = (Style)FindResource("ScheduleCell");
            }
            selectedCells.Clear();
        }

        private void UpdateEmployeesList()
        {
            EmployeesList.ItemsSource = null;
            EmployeesList.ItemsSource = employees;
            NoEmployeesText.Visibility = employees.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void DeleteEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string employeeId)
            {
                var employeeToRemove = employees.FirstOrDefault(emp => emp.Id == employeeId);
                if (employeeToRemove != null)
                {
                    var result = await popupSystem.ShowConfirmDialog($"Are you sure you want to delete '{employeeToRemove.Name}'?", "Confirm Delete");
                    if (!result) return;

                    employees.Remove(employeeToRemove);
                    employeeSchedules.Remove(employeeId);
                    var combinedToRemove = combinedAvailabilities.FirstOrDefault(emp => emp.Id == employeeId);
                    if (combinedToRemove != null)
                    {
                        combinedAvailabilities.Remove(combinedToRemove);
                    }
                    stateSave.DeleteEmployee(employeeId);
                    UpdateEmployeesList();
                    popupSystem.ShowNotification($"{employeeToRemove.Name} removed", "Employee Deleted", "info");
                }
            }
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
                    ImportDataButton.Content = $"📥 Import";

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
                popupSystem.ShowNotification($"CSV file loaded! Found {importResult.People.Count} participants.", "File Ready", "success");
            }
            catch (Exception ex)
            {
                popupSystem.ShowNotification($"Error loading file: {ex.Message}", "File Error", "error");
            }
        }

        private async Task ExtractFromLettuceMeet()
        {
            string url = LettuceMeetUrlBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(url))
            {
                popupSystem.ShowNotification("Please enter a LettuceMeet URL.", "Missing URL", "warning");
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
                popupSystem.ShowNotification($"Error: {ex.Message}", "Extraction Error", "error");
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

                    int markedSlots = 0;
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
                                markedSlots++;
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

                    if (!combinedAvailabilities.Any(emp => emp.Name.Equals(employee.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        combinedAvailabilities.Add(employee);
                        stateSave.SaveEmployee(employee);
                        importedCount++;
                    }
                }

                lettuceMeetData = null;
                LettuceMeetUrlBox.Text = "";
                ImportedDataCard.Visibility = Visibility.Collapsed;
                ExtractionStatus.Visibility = Visibility.Collapsed;
                popupSystem.ShowNotification($"Imported {importedCount} employees from LettuceMeet with exact timing!", "Import Complete", "success");
                ViewAvailabilities_Click(this, new RoutedEventArgs());
            }
            catch (Exception ex)
            {
                popupSystem.ShowNotification($"Error importing data: {ex.Message}\n\nCheck debug.txt on desktop for details.",
                    "Import Error", "error");
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
                        popupSystem.ShowNotification($"Import completed with warnings:\n\n{errorMsg}",
                            "Import Warnings", "warning");
                    }
                    else
                    {
                        popupSystem.ShowNotification($"Error importing file:\n\n{errorMsg}",
                            "Import Error", "error");
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

                    if (!combinedAvailabilities.Any(emp => emp.Name.Equals(employee.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        combinedAvailabilities.Add(employee);
                        stateSave.SaveEmployee(employee);
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

                popupSystem.ShowNotification(notificationMessage, notificationTitle, notificationType);
                importedFilePath = "";
                ImportedDataCard.Visibility = Visibility.Collapsed;
                ViewAvailabilities_Click(null, new RoutedEventArgs());
            }
            catch (Exception ex)
            {
                popupSystem.ShowNotification($"Error importing data: {ex.Message}\n\nPlease ensure your file matches the template format.", "Import Error", "error");
            }
        }

        private void UploadArea_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BrowseButton_Click(sender, e);
        }

        private async void ExtractFromLettuceMeetButton_Click(object sender, RoutedEventArgs e)
        {
            await ExtractFromLettuceMeet();
        }

        private void UpdateCombinedAvailabilities()
        {
            CombinedAvailabilitiesList.ItemsSource = null;
            CombinedAvailabilitiesList.ItemsSource = combinedAvailabilities;
            NoCombinedAvailabilitiesText.Visibility = combinedAvailabilities.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateStatistics()
        {
            int totalPeople = combinedAvailabilities.Count;
            int manualEntries = combinedAvailabilities.Count(e => e.Source == "Manual");
            int importedData = combinedAvailabilities.Count(e => e.Source != "Manual");
            TotalPeopleText.Text = totalPeople.ToString();
            ManualEntriesText.Text = manualEntries.ToString();
            ImportedDataText.Text = importedData.ToString();

            if (combinedAvailabilities.Count > 0)
            {
                var allStartDates = combinedAvailabilities.Select(e => e.StartDate).ToList();
                var allEndDates = combinedAvailabilities.Select(e => e.EndDate).ToList();
                DateTime earliestDate = allStartDates.Min();
                DateTime latestDate = allEndDates.Max();
                DateRangeText.Text = $"{earliestDate:MMM dd} to {latestDate:MMM dd}";
            }
            else
            {
                DateRangeText.Text = "No data";
            }
        }

        private async void DeleteCombinedEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string employeeId)
            {
                var employeeToRemove = combinedAvailabilities.FirstOrDefault(emp => emp.Id == employeeId);
                if (employeeToRemove != null)
                {
                    var result = await popupSystem.ShowConfirmDialog($"Are you sure you want to delete '{employeeToRemove.Name}'?", "Confirm Delete");
                    if (!result) return;

                    combinedAvailabilities.Remove(employeeToRemove);

                    if (employeeToRemove.Source == "Manual")
                    {
                        var manualEmployee = employees.FirstOrDefault(emp => emp.Id == employeeId);
                        if (manualEmployee != null)
                        {
                            employees.Remove(manualEmployee);
                            employeeSchedules.Remove(employeeId);
                        }
                    }

                    stateSave.DeleteEmployee(employeeId);
                    UpdateCombinedAvailabilities();
                    UpdateStatistics();
                    UpdateEmployeesList();
                    popupSystem.ShowNotification($"{employeeToRemove.Name} removed", "Employee Deleted", "info");
                }
            }
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            bool allSelected = combinedAvailabilities.All(e => e.IsSelected);
            foreach (var employee in combinedAvailabilities)
            {
                employee.IsSelected = !allSelected;
            }
            SelectAllButton.Content = allSelected ? "☑ Select All" : "☐ Unselect All";
            UpdateCombinedAvailabilities();
        }

        private async void DeleteSelectedButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedEmployees = combinedAvailabilities.Where(e => e.IsSelected).ToList();

            if (selectedEmployees.Count == 0)
            {
                popupSystem.ShowNotification("Please select at least one employee to delete.", "No Selection", "warning");
                return;
            }

            string message = selectedEmployees.Count == 1
                ? $"Are you sure you want to delete '{selectedEmployees[0].Name}'?"
                : $"Are you sure you want to delete {selectedEmployees.Count} selected employees?";

            var result = await popupSystem.ShowConfirmDialog(message, "Confirm Delete");
            if (!result) return;

            foreach (var employee in selectedEmployees)
            {
                combinedAvailabilities.Remove(employee);

                if (employee.Source == "Manual")
                {
                    var manualEmployee = employees.FirstOrDefault(emp => emp.Id == employee.Id);
                    if (manualEmployee != null)
                    {
                        employees.Remove(manualEmployee);
                        employeeSchedules.Remove(employee.Id);
                    }
                }

                stateSave.DeleteEmployee(employee.Id);
            }

            UpdateCombinedAvailabilities();
            UpdateStatistics();
            UpdateEmployeesList();
            popupSystem.ShowNotification($"Deleted {selectedEmployees.Count} employees", "Deletion Complete", "info");
        }

        private async void SaveBatchButton_Click(object sender, RoutedEventArgs e)
        {
            string batchName = BatchNameTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(batchName))
            {
                popupSystem.ShowNotification("Please enter a batch name.", "Missing Batch Name", "warning");
                return;
            }

            if (combinedAvailabilities.Count == 0)
            {
                popupSystem.ShowNotification("No availabilities to save as batch.", "No Data", "warning");
                return;
            }

            if (batches.Any(b => b.Name.Equals(batchName, StringComparison.OrdinalIgnoreCase)))
            {
                popupSystem.ShowNotification("A batch with this name already exists. Please choose a different name.",
                    "Duplicate Batch Name", "warning");
                return;
            }

            var selectedEmployees = combinedAvailabilities.Where(e => e.IsSelected).ToList();
            var employeesToAdd = selectedEmployees.Count > 0 ? selectedEmployees : combinedAvailabilities;
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

            batches.Add(batch);
            UpdateBatchComboBox();
            UpdateBatchList();
            stateSave.SaveBatch(batch);
            BatchNameTextBox.Text = "";
            popupSystem.ShowNotification($"Saved '{batchName}' batch with {batch.Count} employees", "Batch Saved", "success");
        }

        private async void DeleteBatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string batchId)
            {
                var batchToRemove = batches.FirstOrDefault(b => b.Id == batchId);
                if (batchToRemove != null)
                {
                    var result = await popupSystem.ShowConfirmDialog($"Are you sure you want to delete batch '{batchToRemove.Name}'?",
                        "Confirm Delete");
                    if (!result) return;

                    batches.Remove(batchToRemove);
                    UpdateBatchComboBox();
                    UpdateBatchList();
                    stateSave.DeleteBatch(batchId);
                    stateSave.CleanupOrphanedEmployees(batches);
                    popupSystem.ShowNotification($"Batch '{batchToRemove.Name}' deleted", "Batch Deleted", "info");
                }
            }
        }

        private async void RenameBatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string batchId)
            {
                var batch = batches.FirstOrDefault(b => b.Id == batchId);
                if (batch != null)
                {
                    string newName = await popupSystem.ShowInputDialog("Enter new batch name:", "Rename Batch", batch.Name);

                    if (string.IsNullOrWhiteSpace(newName) || newName == batch.Name)
                    {
                        return;
                    }

                    if (batches.Any(b => b.Name.Equals(newName, StringComparison.OrdinalIgnoreCase) && b.Id != batchId))
                    {
                        popupSystem.ShowNotification("A batch with this name already exists.", "Duplicate Name", "warning");
                        return;
                    }

                    batch.Name = newName;
                    UpdateBatchComboBox();
                    UpdateBatchList();
                    stateSave.SaveBatch(batch);
                    popupSystem.ShowNotification($"Batch renamed to '{newName}'", "Batch Renamed", "info");
                }
            }
        }

        private void BatchComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void ViewSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var schedule = schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (schedule != null)
                {
                    ShowScheduleDetailView(schedule);
                }
            }
        }

        private async void RenameSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var schedule = schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (schedule != null)
                {
                    string newName = await popupSystem.ShowInputDialog("Enter new schedule name:", "Rename Schedule", schedule.Name);

                    if (string.IsNullOrWhiteSpace(newName) || newName == schedule.Name)
                    {
                        return;
                    }

                    if (schedules.Any(s => s.Name.Equals(newName, StringComparison.OrdinalIgnoreCase) && s.Id != scheduleId))
                    {
                        popupSystem.ShowNotification("A schedule with this name already exists.", "Duplicate Name", "warning");
                        return;
                    }

                    schedule.Name = newName;
                    UpdateScheduleList();
                    stateSave.SaveSchedule(schedule);

                    if (currentSchedule != null && currentSchedule.Id == scheduleId)
                    {
                        currentSchedule.Name = newName;
                        ScheduleDetailTitle.Text = newName;
                        CurrentScheduleName.Text = newName;
                    }

                    popupSystem.ShowNotification($"Schedule renamed to '{newName}'", "Schedule Renamed", "info");
                }
            }
        }

        private async void DeleteSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string scheduleId)
            {
                var scheduleToRemove = schedules.FirstOrDefault(s => s.Id == scheduleId);
                if (scheduleToRemove != null)
                {
                    var result = await popupSystem.ShowConfirmDialog($"Are you sure you want to delete schedule '{scheduleToRemove.Name}'?",
                        "Confirm Delete");
                    if (!result) return;

                    schedules.Remove(scheduleToRemove);
                    UpdateScheduleList();
                    stateSave.DeleteSchedule(scheduleId);

                    if (currentSchedule != null && currentSchedule.Id == scheduleId)
                    {
                        ShowScheduleListView();
                        currentSchedule = null;
                    }

                    popupSystem.ShowNotification($"Schedule '{scheduleToRemove.Name}' deleted", "Schedule Deleted", "info");
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

        private void GenerateScheduleDetailGrid()
        {
            ScheduleDetailGrid.Children.Clear();
            ScheduleDetailGrid.RowDefinitions.Clear();
            ScheduleDetailGrid.ColumnDefinitions.Clear();

            if (currentSchedule == null) return;

            int scheduleDays = GetScheduleDayCount(currentSchedule);
            double shiftLength = currentSchedule.ShiftLengthHours;
            List<DateTime> scheduleDates = GetScheduleDateList(currentSchedule);
            int shiftIntervals = currentSchedule.ShiftIntervals;
            int originalDays = currentSchedule.OriginalDayCount > 0 ? currentSchedule.OriginalDayCount : scheduleDays;
            int originalIntervals = currentSchedule.OriginalShiftIntervals > 0 ? currentSchedule.OriginalShiftIntervals : shiftIntervals;
            ScheduleDetailGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });

            for (int day = 0; day < scheduleDays; day++)
            {
                ScheduleDetailGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            }

            ScheduleDetailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) });

            for (int interval = 0; interval < shiftIntervals; interval++)
            {
                ScheduleDetailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) });
            }

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
                    dayHeaderText = $"{dayDate.ToString("ddd")}\n{dayDate.ToString("MM/dd")}";
                    headerBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5"));
                    headerForeground = Brushes.White;
                }
                else
                {
                    if (headerNames.Count > 0)
                    {
                        dayHeaderText = string.Join("\n", headerNames.Take(2));
                        headerBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1FAE5"));
                        headerForeground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#059669"));
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
                    CornerRadius = new CornerRadius(4, 4, 0, 0),
                    BorderBrush = isOriginalDay ? Brushes.Transparent : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                    BorderThickness = isOriginalDay ? new Thickness(0) : new Thickness(1),
                    Padding = new Thickness(4),
                    Tag = headerCellId
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
                        foreach (var name in headerNames.Take(3))
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = headerNames.Count <= 2 ? 10 : 9,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = headerForeground,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 1, 0, 1)
                            };
                            headerContent.Children.Add(nameText);
                        }

                        if (headerNames.Count > 3)
                        {
                            TextBlock moreText = new TextBlock
                            {
                                Text = $"+{headerNames.Count - 3} more",
                                FontSize = 8,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = headerForeground,
                                FontStyle = FontStyles.Italic
                            };
                            headerContent.Children.Add(moreText);
                        }

                        headerBorder.ToolTip = $"Assigned: {string.Join(", ", headerNames)}";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = dayHeaderText,
                            FontSize = 14,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = headerForeground
                        };
                        headerContent.Children.Add(emptyText);
                        headerBorder.ToolTip = "Click to select, then use buttons below to add names";
                    }

                    headerBorder.Child = headerContent;
                    headerBorder.MouseLeftButtonDown += (sender, e) =>
                    {
                        selectedCellId = headerCellId;
                        UpdateCellSelection(headerBorder);
                        UpdateSelectedCellInfo();
                        e.Handled = true;
                    };
                }

                Grid.SetColumn(headerBorder, day + 1);
                Grid.SetRow(headerBorder, 0);
                ScheduleDetailGrid.Children.Add(headerBorder);
            }

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
                        timeForeground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#059669"));
                        timeBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1FAE5"));
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
                    Padding = new Thickness(4),
                    Tag = timeHeaderCellId
                };

                if (isOriginalInterval)
                {
                    TextBlock timeText = new TextBlock
                    {
                        Text = timeHeaderText,
                        FontSize = 10,
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
                        foreach (var name in headerNames.Take(3))
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = headerNames.Count <= 2 ? 10 : 9,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = timeForeground,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 1, 0, 1)
                            };
                            timeContent.Children.Add(nameText);
                        }

                        if (headerNames.Count > 3)
                        {
                            TextBlock moreText = new TextBlock
                            {
                                Text = $"+{headerNames.Count - 3} more",
                                FontSize = 8,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = timeForeground,
                                FontStyle = FontStyles.Italic
                            };
                            timeContent.Children.Add(moreText);
                        }

                        timeBorder.ToolTip = $"Assigned: {string.Join(", ", headerNames)}";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = timeHeaderText,
                            FontSize = 14,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = timeForeground
                        };
                        timeContent.Children.Add(emptyText);
                        timeBorder.ToolTip = "Click to select, then use buttons below to add names";
                    }

                    timeBorder.Child = timeContent;
                    timeBorder.MouseLeftButtonDown += (sender, e) =>
                    {
                        selectedCellId = timeHeaderCellId;
                        UpdateCellSelection(timeBorder);
                        UpdateSelectedCellInfo();
                        e.Handled = true;
                    };
                }

                Grid.SetColumn(timeBorder, 0);
                Grid.SetRow(timeBorder, interval + 1);
                ScheduleDetailGrid.Children.Add(timeBorder);
            }

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

                    SolidColorBrush cellBackground;
                    SolidColorBrush textColor;

                    if (names.Count == 0)
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F8FAFC"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"));
                    }
                    else if (names.Count < currentSchedule.PeoplePerShift)
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FEF3C7"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D97706"));
                    }
                    else
                    {
                        cellBackground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1FAE5"));
                        textColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#059669"));
                    }

                    Border cell = new Border
                    {
                        Background = cellBackground,
                        BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB")),
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        CornerRadius = new CornerRadius(4),
                        Padding = new Thickness(4),
                        Tag = cellId
                    };

                    StackPanel cellContent = new StackPanel();

                    if (names.Count > 0)
                    {
                        foreach (var name in names)
                        {
                            TextBlock nameText = new TextBlock
                            {
                                Text = name,
                                FontSize = names.Count <= 2 ? 10 : 9,
                                FontWeight = FontWeights.Medium,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                Foreground = textColor,
                                TextAlignment = TextAlignment.Center,
                                TextWrapping = TextWrapping.Wrap,
                                Margin = new Thickness(0, 1, 0, 1)
                            };
                            cellContent.Children.Add(nameText);
                        }

                        cell.ToolTip = $"Assigned: {string.Join(", ", names)}";
                    }
                    else
                    {
                        TextBlock emptyText = new TextBlock
                        {
                            Text = "+",
                            FontSize = 14,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#9CA3AF"))
                        };
                        cellContent.Children.Add(emptyText);
                        cell.ToolTip = "Click to select, then use buttons below to add names";
                    }

                    cell.Child = cellContent;
                    cell.MouseLeftButtonDown += ScheduleDetailCell_MouseLeftButtonDown;

                    Grid.SetColumn(cell, day + 1);
                    Grid.SetRow(cell, interval + 1);
                    ScheduleDetailGrid.Children.Add(cell);
                }
            }

            double totalGridHeight = 40 + (shiftIntervals * 50);
            ScheduleDetailScrollViewer.MaxHeight = 500;

            if (totalGridHeight > 400)
            {
                ScheduleDetailScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else
            {
                ScheduleDetailScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            }
        }

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

        private void ScheduleDetailCell_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border cell && cell.Tag is string cellId)
            {
                selectedCellId = cellId;
                UpdateCellSelection(cell);
                UpdateSelectedCellInfo();
            }
        }

        private void UpdateCellSelection(Border selectedCell)
        {
            ClearCellSelection();

            if (selectedCell != null)
            {
                selectedCell.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5"));
                selectedCell.BorderThickness = new Thickness(2);
            }
        }

        private void ClearCellSelection()
        {
            foreach (var child in ScheduleDetailGrid.Children)
            {
                if (child is Border cell)
                {
                    cell.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E5E7EB"));
                    cell.BorderThickness = new Thickness(0, 0, 1, 1);
                }
            }
        }

        private void UpdateSelectedCellInfo()
        {
            if (string.IsNullOrEmpty(selectedCellId) || currentSchedule == null)
            {
                SelectedCellInfo.Text = "No cell selected";
                SelectedCellNames.Text = "";
                SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                return;
            }

            if (selectedCellId.StartsWith("dayheader_"))
            {
                if (int.TryParse(selectedCellId.Replace("dayheader_", ""), out int dayIndex))
                {
                    SelectedCellInfo.Text = $"New Day Header (Column {dayIndex + 1})";

                    if (currentSchedule.CellAssignments.ContainsKey(selectedCellId))
                    {
                        var names = currentSchedule.CellAssignments[selectedCellId];
                        if (names.Count > 0)
                        {
                            SelectedCellNames.Text = $"Assigned: {string.Join(", ", names)}";
                            SelectedCellAssignmentsContainer.Visibility = Visibility.Visible;
                            SelectedCellAssignmentsList.ItemsSource = null;
                            SelectedCellAssignmentsList.ItemsSource = names;
                        }
                        else
                        {
                            SelectedCellNames.Text = "Empty header cell";
                            SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                        }
                    }
                    else
                    {
                        SelectedCellNames.Text = "Empty header cell";
                        SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                    }
                }
            }
            else if (selectedCellId.StartsWith("timeheader_"))
            {
                if (int.TryParse(selectedCellId.Replace("timeheader_", ""), out int intervalIndex))
                {
                    SelectedCellInfo.Text = $"New Time Header (Row {intervalIndex + 1})";

                    if (currentSchedule.CellAssignments.ContainsKey(selectedCellId))
                    {
                        var names = currentSchedule.CellAssignments[selectedCellId];
                        if (names.Count > 0)
                        {
                            SelectedCellNames.Text = $"Assigned: {string.Join(", ", names)}";
                            SelectedCellAssignmentsContainer.Visibility = Visibility.Visible;
                            SelectedCellAssignmentsList.ItemsSource = null;
                            SelectedCellAssignmentsList.ItemsSource = names;
                        }
                        else
                        {
                            SelectedCellNames.Text = "Empty header cell";
                            SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                        }
                    }
                    else
                    {
                        SelectedCellNames.Text = "Empty header cell";
                        SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                    }
                }
            }
            else
            {
                string[] parts = selectedCellId.Split('_');
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

                        SelectedCellInfo.Text = $"{dayNames[(int)cellDate.DayOfWeek]}, {cellDate:MMM dd}, {startTimeStr} to {endTimeStr}";
                    }
                }

                if (currentSchedule.CellAssignments.ContainsKey(selectedCellId))
                {
                    var names = currentSchedule.CellAssignments[selectedCellId];
                    if (names.Count > 0)
                    {
                        SelectedCellNames.Text = $"Assigned: {string.Join(", ", names)}";
                        SelectedCellAssignmentsContainer.Visibility = Visibility.Visible;
                        SelectedCellAssignmentsList.ItemsSource = null;
                        SelectedCellAssignmentsList.ItemsSource = names;
                    }
                    else
                    {
                        SelectedCellNames.Text = "Empty cell";
                        SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    SelectedCellNames.Text = "Empty cell";
                    SelectedCellAssignmentsContainer.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void UpdateScheduleStatistics()
        {
            if (currentSchedule != null)
            {
                int totalPeople = 0;
                int totalShifts = 0;

                foreach (var cell in currentSchedule.CellAssignments.Values)
                {
                    totalPeople += cell.Count;
                    if (cell.Count > 0) totalShifts++;
                }

                SchedulePeopleCount.Text = totalPeople.ToString();
                ScheduleShiftsCount.Text = totalShifts.ToString();

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
                ScheduleCoverage.Text = $"{coverage:F0}%";
            }
        }

        private async void AddNameToCell_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedCellId))
            {
                popupSystem.ShowNotification("Please select a cell first by clicking on it in the grid.",
                    "No Cell Selected", "warning");
                return;
            }

            string name = await popupSystem.ShowInputDialog("Enter name to add:", "Add Name to Cell");

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (!currentSchedule.CellAssignments.ContainsKey(selectedCellId))
            {
                currentSchedule.CellAssignments[selectedCellId] = new List<string>();
            }

            if (!currentSchedule.CellAssignments[selectedCellId].Contains(name))
            {
                currentSchedule.CellAssignments[selectedCellId].Add(name);
                GenerateScheduleDetailGrid();
                UpdateScheduleStatistics();
                UpdateSelectedCellInfo();
                UpdateScheduleCompletionStatus(currentSchedule);
                popupSystem.ShowNotification($"Added '{name}' to cell", "Name Added", "success");
            }
            else
            {
                popupSystem.ShowNotification("This name is already assigned to this cell.",
                    "Duplicate Name", "warning");
            }
        }

        private async void ChangeNameInCell_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedCellId) || currentSchedule == null ||
                !currentSchedule.CellAssignments.ContainsKey(selectedCellId) ||
                currentSchedule.CellAssignments[selectedCellId].Count == 0)
            {
                popupSystem.ShowNotification("Please select a cell with names first.",
                    "No Names to Change", "warning");
                return;
            }

            var names = currentSchedule.CellAssignments[selectedCellId];

            if (names.Count == 1)
            {
                string selectedName = names[0];
                string newName = await popupSystem.ShowInputDialog($"Change '{selectedName}' to:", "Change Name", selectedName);

                if (!string.IsNullOrWhiteSpace(newName) && newName != selectedName)
                {
                    int index = currentSchedule.CellAssignments[selectedCellId].IndexOf(selectedName);
                    if (index >= 0)
                    {
                        currentSchedule.CellAssignments[selectedCellId][index] = newName;
                        GenerateScheduleDetailGrid();
                        UpdateSelectedCellInfo();
                        UpdateScheduleCompletionStatus(currentSchedule);
                        popupSystem.ShowNotification($"Changed '{selectedName}' to '{newName}'", "Name Changed", "info");
                    }
                }
            }
            else
            {
                string selectedName = await popupSystem.ShowSelectionDialog("Select name to change:", "Select Name", names);

                if (!string.IsNullOrEmpty(selectedName))
                {
                    string newName = await popupSystem.ShowInputDialog($"Change '{selectedName}' to:", "Change Name", selectedName);

                    if (!string.IsNullOrWhiteSpace(newName) && newName != selectedName)
                    {
                        int index = currentSchedule.CellAssignments[selectedCellId].IndexOf(selectedName);
                        if (index >= 0)
                        {
                            currentSchedule.CellAssignments[selectedCellId][index] = newName;
                            GenerateScheduleDetailGrid();
                            UpdateSelectedCellInfo();
                            UpdateScheduleCompletionStatus(currentSchedule);
                            popupSystem.ShowNotification($"Changed '{selectedName}' to '{newName}'", "Name Changed", "info");
                        }
                    }
                }
            }
        }

        private async void RemoveNameFromCell_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string nameToRemove)
            {
                if (string.IsNullOrEmpty(selectedCellId) || currentSchedule == null ||
                    !currentSchedule.CellAssignments.ContainsKey(selectedCellId))
                {
                    return;
                }

                var result = await popupSystem.ShowConfirmDialog($"Are you sure you want to remove '{nameToRemove}' from this cell?",
                    "Confirm Remove");
                if (!result) return;

                currentSchedule.CellAssignments[selectedCellId].Remove(nameToRemove);
                GenerateScheduleDetailGrid();
                UpdateScheduleStatistics();
                UpdateSelectedCellInfo();
                UpdateScheduleCompletionStatus(currentSchedule);
                popupSystem.ShowNotification($"Removed '{nameToRemove}' from cell", "Name Removed", "info");
            }
        }

        private async void ExportPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            if (currentSchedule == null)
            {
                popupSystem.ShowNotification("Please select a schedule to export.", "No Schedule", "warning");
                return;
            }

            try
            {
                var newExport = ExportService.CreateExportItem(currentSchedule);
                popupSystem.ShowNotification($"Schedule '{currentSchedule.Name}' added to exports!", "Export Ready", "success");
                ShowExportPage();

                var mainContentGrid = (Grid)this.FindName("MainContentGrid");
                if (mainContentGrid != null && mainContentGrid.Children.Count > 1)
                {
                    var exportPage = mainContentGrid.Children[1] as Views.ExportPage;
                    if (exportPage != null)
                    {
                        await Task.Delay(100);
                        ExportPageControl.SelectExportById(newExport.Id);
                    }
                }
            }
            catch (Exception ex)
            {
                popupSystem.ShowNotification($"Error creating export: {ex.Message}", "Export Error", "error");
            }
        }

        private void GridScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer scv = (ScrollViewer)sender;
            scv.ScrollToHorizontalOffset(scv.HorizontalOffset - e.Delta);
            e.Handled = true;
        }

        private void ScheduleDetailScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer scv = (ScrollViewer)sender;
            scv.ScrollToHorizontalOffset(scv.HorizontalOffset - e.Delta);
            e.Handled = true;
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

        private async void Generate_Click(object sender, RoutedEventArgs e)
        {
            if (OpeningTimeBox.SelectedIndex == -1 || ClosingTimeBox.SelectedIndex == -1 ||
                ShiftLengthBox.SelectedIndex == -1 || PeoplePerShiftBox.SelectedIndex == -1)
            {
                popupSystem.ShowNotification("Please fill in all settings before generating the schedule.",
                    "Incomplete Settings", "warning");
                return;
            }

            if (BatchComboBox.SelectedIndex == -1 || batches.Count == 0)
            {
                popupSystem.ShowNotification("Please create and select a batch to use for scheduling.",
                    "No Batch Selected", "warning");
                return;
            }

            var selectedBatch = batches[BatchComboBox.SelectedIndex];
            string openingTime = OpeningTimeBox.SelectedItem?.ToString() ?? "6:00 AM";
            string closingTime = ClosingTimeBox.SelectedItem?.ToString() ?? "11:00 PM";
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
            includeWeekends = IncludeWeekendsCheckBox.IsChecked ?? false;
            currentPeoplePerShift = PeoplePerShiftBox.SelectedIndex + 1;
            var filteredBatchEmployees = FilterBatchByTimeInterval(selectedBatch, openingHourIndex, closingHourIndex);

            if (filteredBatchEmployees.Count == 0)
            {
                popupSystem.ShowNotification("No employees in the batch are available within the selected time interval.",
                    "No Available Employees", "warning");
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
            while (schedules.Any(s => s.Name.Equals(scheduleName, StringComparison.OrdinalIgnoreCase)))
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
                    popupSystem.ShowNotification("The selected batch contains no weekdays. Please include weekends or select a different batch.",
                        "No Weekdays Available", "warning");
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
                OriginalShiftIntervals = shiftIntervals
            };

            InitializeCellAssignmentsFromShifts(schedule, scheduleResult.Shifts, scheduleStartDate, scheduleEndDate, includeWeekends);
            schedules.Add(schedule);
            UpdateScheduleList();
            stateSave.SaveSchedule(schedule);
            popupSystem.ShowNotification($"Schedule '{scheduleName}' created with {scheduleResult.Shifts.Count(s => s.AssignedPeople.Count > 0)} filled shifts", "Schedule Generated", "success");
            Schedule_Click(this, new RoutedEventArgs());
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
                    schedule.CellAssignments[cellId] = shift.AssignedPeople;
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

        private void WriteDebugLog(string message)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string debugFilePath = Path.Combine(desktopPath, "scheduler_debug.txt");

            try
            {
                using (StreamWriter writer = new StreamWriter(debugFilePath, true))
                {
                    writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Debug log error: {ex.Message}");
            }
        }

        private async void DownloadSampleButton_Click(object sender, RoutedEventArgs e)
        {
            var result = await popupSystem.ShowSelectionDialog(
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
                    popupSystem.ShowNotification("Template downloaded successfully!", "Download Complete", "success");

                    popupSystem.ShowNotification(
                        "Template downloaded! Here's how to use it:\n\n" +
                        "1. Open the CSV file in Excel or similar\n" +
                        "2. Fill in employee names and their availability\n" +
                        "3. Save the file\n" +
                        "4. Upload it back to Scheduler Pro\n\n" +
                        "Format: Name, Date (YYYY-MM-DD), Start_Time (HH:mm), End_Time (HH:mm)",
                        "Template Instructions",
                        "info");
                }
                catch (Exception ex)
                {
                    popupSystem.ShowNotification($"Error saving template: {ex.Message}", "Error", "error");
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

        private void UpdateScheduleCompletionStatus(Schedule schedule)
        {
            if (schedule == null) return;

            bool isComplete = true;
            int totalCells = 0;
            int filledCells = 0;

            int scheduleDays = GetScheduleDayCount(schedule);

            for (int day = 0; day < scheduleDays; day++)
            {
                for (int interval = 0; interval < schedule.ShiftIntervals; interval++)
                {
                    string cellId = $"cell_{day}_{interval}";
                    totalCells++;

                    int assignedCount = 0;
                    if (schedule.CellAssignments.ContainsKey(cellId))
                    {
                        assignedCount = schedule.CellAssignments[cellId].Count;
                    }

                    if (assignedCount < schedule.PeoplePerShift)
                    {
                        isComplete = false;
                    }

                    if (assignedCount > 0)
                    {
                        filledCells++;
                    }
                }
            }

            schedule.Status = isComplete ? "Complete" : "Incomplete";

            if (currentSchedule != null && currentSchedule.Id == schedule.Id)
            {
                currentSchedule.Status = schedule.Status;
                UpdateScheduleStatistics();
            }

            stateSave.SaveSchedule(schedule);
            UpdateScheduleList();
        }

        private void ExportPage_Click(object sender, RoutedEventArgs e)
        {
            ShowExportPage();
        }

        private void ShowExportPage()
        {
            try
            {
                SetupPage.Visibility = Visibility.Collapsed;
                ManualSchedulePage.Visibility = Visibility.Collapsed;
                ImportDataPage.Visibility = Visibility.Collapsed;
                ViewAvailabilitiesPage.Visibility = Visibility.Collapsed;
                SchedulePage.Visibility = Visibility.Collapsed;
                ExportPage.Visibility = Visibility.Visible;

                if (ExportPageControl != null)
                {
                    var exportPageType = ExportPageControl.GetType();
                    var setReferenceMethod = exportPageType.GetMethod("SetMainWindowReference",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

                    if (setReferenceMethod != null)
                    {
                        setReferenceMethod.Invoke(ExportPageControl, new object[] { this });
                    }
                }

                if (ExportPageControl != null)
                {
                    ExportPageControl.LoadExports();
                }

                UpdateNavigationButtons(ExportNavButton);
                UpdateBackButtonVisibility();
                SetupBackButton.Visibility = Visibility.Collapsed;
                ManualBackButton.Visibility = Visibility.Collapsed;
                ImportBackButton.Visibility = Visibility.Collapsed;
                AvailabilitiesBackButton.Visibility = Visibility.Collapsed;
                ScheduleBackButton.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                popupSystem.ShowNotification($"Error loading export page: {ex.Message}", "Error", "error");
            }
        }

        private async void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSchedule == null) return;

            var result = await popupSystem.ShowConfirmDialog(
                "Are you sure you want to add a new row to the schedule? This will add a new shift interval at the bottom.",
                "Add Row Confirmation"
            );

            if (!result) return;

            currentSchedule.ShiftIntervals++;
            stateSave.SaveSchedule(currentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            popupSystem.ShowNotification("New row added at the bottom of schedule", "Row Added", "success");
        }

        private async void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSchedule == null) return;

            var result = await popupSystem.ShowConfirmDialog(
                "Are you sure you want to add a new column to the schedule? This will add a new day at the end.",
                "Add Column Confirmation"
            );

            if (!result) return;

            currentSchedule.EndDate = currentSchedule.EndDate.AddDays(1);

            if (!currentSchedule.IncludeWeekends)
            {
                while (currentSchedule.EndDate.DayOfWeek == DayOfWeek.Saturday ||
                       currentSchedule.EndDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    currentSchedule.EndDate = currentSchedule.EndDate.AddDays(1);
                }
            }

            stateSave.SaveSchedule(currentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            popupSystem.ShowNotification("New column added at the end of schedule", "Column Added", "info");
        }

        private async void DeleteRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSchedule == null || currentSchedule.ShiftIntervals <= 1) return;

            var result = await popupSystem.ShowConfirmDialog(
                $"Are you sure you want to delete the last row? This will remove shift interval {currentSchedule.ShiftIntervals}.",
                "Delete Row Confirmation"
            );

            if (!result) return;

            int scheduleDays = GetScheduleDayCount(currentSchedule);

            for (int day = 0; day < scheduleDays; day++)
            {
                string lastCellId = $"cell_{day}_{currentSchedule.ShiftIntervals - 1}";
                currentSchedule.CellAssignments.Remove(lastCellId);
            }

            currentSchedule.ShiftIntervals--;
            stateSave.SaveSchedule(currentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            popupSystem.ShowNotification("Last row deleted from schedule", "Row Deleted", "info");
        }

        private async void DeleteColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSchedule == null) return;

            int currentDays = GetScheduleDayCount(currentSchedule);

            if (currentDays <= 1) return;

            var result = await popupSystem.ShowConfirmDialog(
                $"Are you sure you want to delete the last column? This will remove day {currentDays}.",
                "Delete Column Confirmation"
            );

            if (!result) return;

            int dayToRemove = currentDays - 1;

            for (int interval = 0; interval < currentSchedule.ShiftIntervals; interval++)
            {
                string cellId = $"cell_{dayToRemove}_{interval}";
                currentSchedule.CellAssignments.Remove(cellId);
            }

            DateTime newEndDate = currentSchedule.EndDate;
            do
            {
                newEndDate = newEndDate.AddDays(-1);
            } while (!currentSchedule.IncludeWeekends &&
                     (newEndDate.DayOfWeek == DayOfWeek.Saturday ||
                      newEndDate.DayOfWeek == DayOfWeek.Sunday));

            currentSchedule.EndDate = newEndDate;
            stateSave.SaveSchedule(currentSchedule);
            GenerateScheduleDetailGrid();
            UpdateScheduleStatistics();
            popupSystem.ShowNotification("Last column deleted from schedule", "Column Deleted", "info");
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

        private async void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            string aboutMessage =
                "Scheduler Pro – Personal Project\n\n" +
                "Scheduler Pro is a solo personal project developed for the University of Windsor Peer Support Centre coordinator. It is provided free of charge for non-commercial and non-profit use by anyone who may find it helpful.\n\n" +
                "\nImportant Notes:\n" +
                "This application was created as a personal project rather than an industrial or commercial product. As a result, it should be considered a completed project rather than an actively maintained one. Future updates are unlikely, and over time some features may become outdated or stop functioning as external dependencies change.\n\n" +
                "As of the release date (01/03/2026), all core features are fully functional.\n\n" +
                "\nSupport:\n" +
                "If you encounter any issues or have concerns, you can contact the developer at:\n\n" +
                "danesh.amir2001@gmail.com";

            popupSystem.ShowNotification(aboutMessage, "About Scheduler Pro", "info", 0);
        }
    }
}