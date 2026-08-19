using SchedulerApp.Services;
using SchedulerApp.Views;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SchedulerApp
{
    public partial class MainWindow : Window
    {
        private AppState _state;
        private PopupSystem popupSystem;
        private Views.ExportPage ExportPageContent { get; set; }

        private const string PAGE_SETUP = "Setup";
        private const string PAGE_MANUAL = "Manual";
        private const string PAGE_IMPORT = "Import";
        private const string PAGE_AVAILABILITIES = "ViewAvailabilities";
        private const string PAGE_SCHEDULE = "Schedule";
        private const string PAGE_EXPORT = "Export";

        private BackgroundAnimation _backgroundAnimation;
        
        // Track window state for custom maximize behavior
        private bool _isMaximized = false;
        private double _previousLeft;
        private double _previousTop;
        private double _previousWidth;
        private double _previousHeight;

        public MainWindow()
        {
            InitializeComponent();
            _state = new AppState();
            popupSystem = (PopupSystem)this.FindName("PopupSystemControl");
            ExportPageContent = (Views.ExportPage)this.FindName("ExportPageControl");

            InitializeState();
            InitializePages();
            LoadSavedData();
            ShowPage(PAGE_SETUP);

            // Update maximize button when window state changes
            StateChanged += MainWindow_StateChanged;

            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _backgroundAnimation = new BackgroundAnimation(SnowCanvas);

            // Daytime (6 AM - 6 PM): Snow theme
            // Nighttime (6 PM - 6 AM): Starry Night
            int hour = DateTime.Now.Hour;
            if (hour >= 6 && hour < 18)
            {
                _backgroundAnimation.SetTheme(0); // Snow
            }
            else
            {
                _backgroundAnimation.SetTheme(1); // Starry Night
            }
        }

        private void MainWindow_StateChanged(object sender, EventArgs e)
        {
            if (MaximizeButton != null)
            {
                MaximizeButton.Content = _isMaximized ? "❐" : "▢";
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                MaximizeButton_Click(sender, e);
                return;
            }
            
            // If maximized, restore to normal size when dragging
            if (_isMaximized)
            {
                // Restore to previous size and position
                WindowState = WindowState.Normal;
                Left = _previousLeft;
                Top = _previousTop;
                Width = _previousWidth;
                Height = _previousHeight;
                _isMaximized = false;
                
                // Update button icon
                if (MaximizeButton != null)
                {
                    MaximizeButton.Content = "▢";
                }
            }
            
            DragMove();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isMaximized)
            {
                // Restore to previous size and position
                WindowState = WindowState.Normal;
                Left = _previousLeft;
                Top = _previousTop;
                Width = _previousWidth;
                Height = _previousHeight;
                _isMaximized = false;
            }
            else
            {
                // Save current position and size
                _previousLeft = Left;
                _previousTop = Top;
                _previousWidth = Width;
                _previousHeight = Height;
                
                // Maximize to working area (excludes taskbar)
                var workingArea = SystemParameters.WorkArea;
                Left = workingArea.Left;
                Top = workingArea.Top;
                Width = workingArea.Width;
                Height = workingArea.Height;
                _isMaximized = true;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }


        private void InitializeState()
        {
            _state.StateSave = new StateSave();

            _state.ShowNotification = (message, title, type, duration) =>
            {
                popupSystem.ShowNotification(message, title, type, duration);
                return true;
            };
            _state.ShowConfirmDialog = (message, title) => popupSystem.ShowConfirmDialog(message, title);
            _state.ShowInputDialog = (message, title, defaultValue) => popupSystem.ShowInputDialog(message, title, defaultValue);
            _state.ShowSelectionDialog = (message, title, items) => popupSystem.ShowSelectionDialog(message, title, items);

            _state.NavigateRequested = (page) => HandleNavigateRequest(page);

            _state.RefreshStatistics = () =>
            {
                if (ViewAvailabilitiesPageControl != null)
                    ViewAvailabilitiesPageControl.UpdateStatistics();
            };
            _state.RefreshCombinedAvailabilities = () =>
            {
                if (ViewAvailabilitiesPageControl != null)
                    ViewAvailabilitiesPageControl.UpdateCombinedAvailabilities();
            };
            _state.RefreshBatchList = () =>
            {
                if (ViewAvailabilitiesPageControl != null)
                    ViewAvailabilitiesPageControl.UpdateBatchList();
            };
            _state.RefreshBatchComboBox = () =>
            {
                if (SetupPageControl != null)
                    SetupPageControl.UpdateBatchComboBox();
            };
            _state.RefreshScheduleList = () =>
            {
                if (SchedulePageControl != null)
                    SchedulePageControl.UpdateScheduleList();
            };
            _state.RefreshEmployeesList = () =>
            {
                if (ManualSchedulePageControl != null)
                    ManualSchedulePageControl.UpdateEmployeesList();
            };
            _state.ShowScheduleListView = () =>
            {
                if (SchedulePageControl != null)
                    SchedulePageControl.ShowScheduleListView();
            };
            _state.ShowExportPage = () => ShowPage(PAGE_EXPORT);
        }

        private void InitializePages()
        {
            SetupPageControl.Initialize(_state);
            ManualSchedulePageControl.Initialize(_state);
            ImportDataPageControl.Initialize(_state);
            ViewAvailabilitiesPageControl.Initialize(_state);
            SchedulePageControl.Initialize(_state);
        }

        private void LoadSavedData()
        {
            var savedEmployees = _state.StateSave.LoadEmployees();
            if (savedEmployees != null && savedEmployees.Count > 0)
            {
                _state.CombinedAvailabilities = savedEmployees;
                foreach (var emp in savedEmployees.Where(e => e.Source == "Manual"))
                {
                    if (!_state.ManualEmployees.Any(e => e.Id == emp.Id))
                    {
                        _state.ManualEmployees.Add(emp);
                    }
                }
            }

            var savedBatches = _state.StateSave.LoadBatches();
            if (savedBatches != null && savedBatches.Count > 0)
            {
                _state.Batches = savedBatches;
            }

            var savedSchedules = _state.StateSave.LoadSchedules();
            if (savedSchedules != null && savedSchedules.Count > 0)
            {
                _state.Schedules = savedSchedules;
            }

            // Refresh all UI bindings after loading
            ServiceRefreshAll();
        }

        private void ServiceRefreshAll()
        {
            _state.RefreshStatistics?.Invoke();
            _state.RefreshCombinedAvailabilities?.Invoke();
            _state.RefreshBatchList?.Invoke();
            _state.RefreshBatchComboBox?.Invoke();
            _state.RefreshScheduleList?.Invoke();
            _state.RefreshEmployeesList?.Invoke();
        }

        private void HandleNavigateRequest(string page)
        {
            switch (page)
            {
                case "Setup":
                    ShowPage(PAGE_SETUP);
                    break;
                case "Manual":
                    ShowPage(PAGE_MANUAL);
                    break;
                case "Import":
                    ShowPage(PAGE_IMPORT);
                    break;
                case "ViewAvailabilities":
                    ShowPage(PAGE_AVAILABILITIES);
                    break;
                case "Schedule":
                    ShowPage(PAGE_SCHEDULE);
                    break;
                case "Export":
                    ShowPage(PAGE_EXPORT);
                    break;
                case "Back":
                    // Back logic handled per page; default to Setup
                    ShowPage(PAGE_SETUP);
                    break;
            }
        }

        private void ShowPage(string page)
        {
            SetupPageControl.Visibility = Visibility.Collapsed;
            ManualSchedulePageControl.Visibility = Visibility.Collapsed;
            ImportDataPageControl.Visibility = Visibility.Collapsed;
            ViewAvailabilitiesPageControl.Visibility = Visibility.Collapsed;
            SchedulePageControl.Visibility = Visibility.Collapsed;
            ExportPage.Visibility = Visibility.Collapsed;

            Button activeButton = SetupNavButton;

            switch (page)
            {
                case PAGE_SETUP:
                    SetupPageControl.Visibility = Visibility.Visible;
                    activeButton = SetupNavButton;
                    break;
                case PAGE_MANUAL:
                    ManualSchedulePageControl.Visibility = Visibility.Visible;
                    ManualSchedulePageControl.Reset();
                    activeButton = ManualScheduleNavButton;
                    break;
                case PAGE_IMPORT:
                    ImportDataPageControl.Visibility = Visibility.Visible;
                    activeButton = ImportDataNavButton;
                    break;
                case PAGE_AVAILABILITIES:
                    ViewAvailabilitiesPageControl.Visibility = Visibility.Visible;
                    ViewAvailabilitiesPageControl.UpdateCombinedAvailabilities();
                    ViewAvailabilitiesPageControl.UpdateStatistics();
                    ViewAvailabilitiesPageControl.UpdateBatchList();
                    activeButton = ViewAvailabilitiesNavButton;
                    break;
                case PAGE_SCHEDULE:
                    SchedulePageControl.Visibility = Visibility.Visible;
                    SchedulePageControl.UpdateScheduleList();
                    SchedulePageControl.ShowScheduleListView();
                    activeButton = ScheduleNavButton;
                    break;
                case PAGE_EXPORT:
                    ExportPage.Visibility = Visibility.Visible;
                    LoadExportPage();
                    activeButton = ExportNavButton;
                    break;
            }

            UpdateNavigationButtons(activeButton);
        }

        private void LoadExportPage()
        {
            if (ExportPageContent != null)
            {
                var exportPageType = ExportPageContent.GetType();
                var setReferenceMethod = exportPageType.GetMethod("SetMainWindowReference",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

                if (setReferenceMethod != null)
                {
                    setReferenceMethod.Invoke(ExportPageContent, new object[] { this });
                }

                ExportPageContent.LoadExports();
            }
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

        // Navigation button handlers
        private void Setup_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_SETUP);
        private void ManualSchedule_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_MANUAL);
        private void ImportData_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_IMPORT);
        private void ViewAvailabilities_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_AVAILABILITIES);
        private void Schedule_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_SCHEDULE);
        private void ExportPage_Click(object sender, RoutedEventArgs e) => ShowPage(PAGE_EXPORT);

        private async void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            string aboutMessage =
                "Scheduler Pro – Personal Project\n\n" +
                "Scheduler Pro is a solo personal project developed for the University of Windsor Peer Support Centre coordinator. It is provided free of charge for non-commercial and non-profit use by anyone who may find it helpful.\n\n" +
                "\nVersion 2.0 – Release Notes\n" +
                "This release introduces several improvements across the application:\n\n" +
                "• Scheduling Engine: Refined the scheduling algorithm to produce more balanced and reliable schedules. Workload distribution is now handled more evenly, and coverage gaps are detected and resolved more consistently.\n\n" +
                "• Export Functionality: The PowerPoint export pipeline has been rebuilt to resolve the issues present in the previous release. Exports now generate correctly with improved template handling and slide rendering.\n\n" +
                "• Application Theme: The interface has been refreshed with a new visual theme and UI elements.\n\n" +
                "\nImportant Notes:\n" +
                "This application was created as a personal project rather than an industrial or commercial product. As a result, it should be considered a completed project rather than an actively maintained one. Future updates are unlikely, and over time some features may become outdated or stop functioning as external dependencies change.\n\n" +
                "As of the release date (08/20/2026), all core features are fully functional.\n\n" +
                "\nSupport:\n" +
                "If you encounter any issues or have concerns, you can contact the developer at:\n\n" +
                "danesh.amir2001@gmail.com";

            popupSystem.ShowNotification(aboutMessage, "About Scheduler Pro", "info", 0);
        }
    }
}
