using Microsoft.Win32;
using SchedulerApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static SchedulerApp.MainWindow;

namespace SchedulerApp.Views
{
    public partial class ExportPage : UserControl
    {
        public class ExportScheduleViewModel
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public DateTime CreatedDate { get; set; }
            public bool IsReady { get; set; } = true;
            public Schedule SourceSchedule { get; set; }
            public ExportSettings Settings { get; set; }
        }

        public class SlideThumbnailViewModel
        {
            public int SlideNumber { get; set; }
            public BitmapImage PreviewImage { get; set; }
            public bool IsSelected { get; set; }
        }

        private List<ExportScheduleViewModel> _exportSchedules = new List<ExportScheduleViewModel>();
        private List<SlideThumbnailViewModel> _slideThumbnails = new List<SlideThumbnailViewModel>();
        private ExportScheduleViewModel _selectedExport = null;
        private MainWindow _mainWindow;
        private bool _isGeneratingPreview = false;

        public ExportPage()
        {
            InitializeComponent();
            ExportService.SetDispatcher(System.Windows.Threading.Dispatcher.CurrentDispatcher);
            LoadExports();

            Loaded += ExportPage_Loaded;
            SizeChanged += ExportPage_SizeChanged;
        }

        private void ExportPage_Loaded(object sender, RoutedEventArgs e)
        {
            _mainWindow = Window.GetWindow(this) as MainWindow;
            AdjustLayoutForSize(ActualWidth);
        }

        private void ExportPage_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            AdjustLayoutForSize(e.NewSize.Width);
        }

        private void AdjustLayoutForSize(double width)
        {
            if (width < 1100)
            {
                MainContainer.Margin = new Thickness(Math.Max(10, width * 0.02), 20, Math.Max(10, width * 0.02), 20);
            }
            else
            {
                MainContainer.Margin = new Thickness(20);
            }
        }

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                    return result;

                T childResult = FindVisualChild<T>(child);
                if (childResult != null)
                    return childResult;
            }
            return null;
        }

        public void SetMainWindowReference(MainWindow mainWindow)
        {
            _mainWindow = mainWindow;
        }

        public void LoadExports()
        {
            try
            {
                var exportServiceExports = ExportService.GetExportSchedules();
                _exportSchedules.Clear();

                foreach (var export in exportServiceExports)
                {
                    _exportSchedules.Add(new ExportScheduleViewModel
                    {
                        Id = export.Id,
                        Name = export.Name,
                        CreatedDate = export.CreatedDate,
                        IsReady = export.IsReady,
                        SourceSchedule = export.SourceSchedule,
                        Settings = export.Settings
                    });
                }

                UpdateExportList();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error loading exports: {ex.Message}", "Error", "error");
            }
        }

        private void UpdateExportList()
        {
            ExportsListControl.ItemsSource = null;
            ExportsListControl.ItemsSource = _exportSchedules;
            NoExportsPanel.Visibility = _exportSchedules.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ShowListView()
        {
            ExportListView.Visibility = Visibility.Visible;
            ExportDetailView.Visibility = Visibility.Collapsed;
            ExportBackButton.Visibility = Visibility.Collapsed;
        }

        private void ShowDetailView()
        {
            ExportListView.Visibility = Visibility.Collapsed;
            ExportDetailView.Visibility = Visibility.Visible;
            ExportBackButton.Visibility = Visibility.Visible;
        }

        private void ExportItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.DataContext is ExportScheduleViewModel export)
            {
                SelectExport(export);
                ShowDetailView();
            }
        }

        private void ViewExport_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string exportId)
            {
                var export = _exportSchedules.FirstOrDefault(exp => exp.Id == exportId);
                if (export != null)
                {
                    SelectExport(export);
                    ShowDetailView();
                }
            }
        }

        private void ExportBackButton_Click(object sender, RoutedEventArgs e)
        {
            ShowListView();
        }

        public void SelectExportById(string exportId)
        {
            var export = _exportSchedules.FirstOrDefault(e => e.Id == exportId);
            if (export != null)
            {
                SelectExport(export);
                ShowDetailView();
            }
        }

        private async void SelectExport(ExportScheduleViewModel export)
        {
            try
            {
                _selectedExport = export;

                ExportDetailTitle.Text = $"Adjustments: {export.Name}";
                SourceScheduleName.Text = export.SourceSchedule?.Name ?? "Unknown";
                CreatedDateText.Text = export.CreatedDate.ToString("MMM dd, yyyy HH:mm");
                ExportStatusText.Text = export.IsReady ? "Ready" : "Processing";
                ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));

                LoadExportSettings(export.Settings);
                ClearPreview();
                await GeneratePreviewAsync();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error selecting export: {ex.Message}", "Error", "error");
            }
        }

        private void LoadExportSettings(ExportSettings settings)
        {
            if (settings != null)
            {
                TextSizeSlider.Value = settings.FontSize;
                CellOpacitySlider.Value = settings.CellOpacity;

                try
                {
                    TextColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        string.IsNullOrEmpty(settings.TextColor) ? "#1F2937" : settings.TextColor));

                    NameCellColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        string.IsNullOrEmpty(settings.NameCellColor) ? "#FFFFFF" : settings.NameCellColor));

                    TimeCellColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        string.IsNullOrEmpty(settings.TimeCellColor) ? "#FFFFFF" : settings.TimeCellColor));

                    DaysRowColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        string.IsNullOrEmpty(settings.DaysRowColor) ? "#FFFFFF" : settings.DaysRowColor));

                    BackgroundColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(
                        string.IsNullOrEmpty(settings.BackgroundColor) ? "#FFFFFF" : settings.BackgroundColor));
                }
                catch
                {
                    TextColorPreview.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1F2937"));
                    NameCellColorPreview.Background = new SolidColorBrush(Colors.White);
                    TimeCellColorPreview.Background = new SolidColorBrush(Colors.White);
                    DaysRowColorPreview.Background = new SolidColorBrush(Colors.White);
                    BackgroundColorPreview.Background = new SolidColorBrush(Colors.White);
                }
            }
        }

        private void ClearPreview()
        {
            PreviewImage.Source = null;
            _slideThumbnails.Clear();
            SlideThumbnails.ItemsSource = null;
            SlideCounter.Text = "Slide 1 of 0";
            PrevSlideButton.IsEnabled = false;
            NextSlideButton.IsEnabled = false;
        }

        private async Task GeneratePreviewAsync()
        {
            if (_selectedExport == null || _isGeneratingPreview) return;

            try
            {
                _isGeneratingPreview = true;
                RefreshPreviewButton.IsEnabled = false;
                LoadingOverlay.Visibility = Visibility.Visible;
                NoPreviewOverlay.Visibility = Visibility.Collapsed;

                var exportSchedule = ExportService.GetExportScheduleById(_selectedExport.Id);
                if (exportSchedule != null)
                {
                    UpdateSettingsFromUI();
                    var previews = await ExportService.GenerateSlidePreviewsAsync(exportSchedule,
                        (current, total) => Dispatcher.Invoke(() =>
                        {
                            LoadingOverlay.Visibility = Visibility.Visible;
                            NoPreviewOverlay.Visibility = Visibility.Collapsed;
                        }));

                    if (previews != null && previews.Count > 0)
                    {
                        _selectedExport.Settings = exportSchedule.Settings;
                        exportSchedule.SlidePreviews = previews;
                        DisplaySlidePreviews(previews);
                    }
                    else
                    {
                        LoadingOverlay.Visibility = Visibility.Collapsed;
                        NoPreviewOverlay.Visibility = Visibility.Visible;
                    }
                }
            }
            catch (Exception ex)
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
                NoPreviewOverlay.Visibility = Visibility.Visible;
                ShowNotification($"Error generating preview: {ex.Message}", "Preview Error", "error");
            }
            finally
            {
                _isGeneratingPreview = false;
                RefreshPreviewButton.IsEnabled = true;
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private async void RefreshPreview_Click(object sender, RoutedEventArgs e)
        {
            await GeneratePreviewAsync();
        }

        private void DisplaySlidePreviews(List<BitmapImage> previews)
        {
            _slideThumbnails.Clear();

            for (int i = 0; i < previews.Count; i++)
            {
                _slideThumbnails.Add(new SlideThumbnailViewModel
                {
                    SlideNumber = i + 1,
                    PreviewImage = previews[i],
                    IsSelected = i == 0
                });
            }

            SlideThumbnails.ItemsSource = null;
            SlideThumbnails.ItemsSource = _slideThumbnails;

            if (previews.Count > 0)
            {
                PreviewImage.Source = previews[0];
                SlideCounter.Text = $"Slide 1 of {previews.Count}";

                PrevSlideButton.IsEnabled = false;
                NextSlideButton.IsEnabled = previews.Count > 1;
                NoPreviewOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private void PrevSlide_Click(object sender, RoutedEventArgs e)
        {
            var currentIndex = _slideThumbnails.FindIndex(t => t.IsSelected);
            if (currentIndex > 0)
            {
                _slideThumbnails[currentIndex].IsSelected = false;
                _slideThumbnails[currentIndex - 1].IsSelected = true;

                PreviewImage.Source = _slideThumbnails[currentIndex - 1].PreviewImage;
                SlideCounter.Text = $"Slide {currentIndex} of {_slideThumbnails.Count}";

                UpdateNavigationButtons();
                RefreshThumbnails();
            }
        }

        private void NextSlide_Click(object sender, RoutedEventArgs e)
        {
            var currentIndex = _slideThumbnails.FindIndex(t => t.IsSelected);
            if (currentIndex < _slideThumbnails.Count - 1)
            {
                _slideThumbnails[currentIndex].IsSelected = false;
                _slideThumbnails[currentIndex + 1].IsSelected = true;

                PreviewImage.Source = _slideThumbnails[currentIndex + 1].PreviewImage;
                SlideCounter.Text = $"Slide {currentIndex + 2} of {_slideThumbnails.Count}";

                UpdateNavigationButtons();
                RefreshThumbnails();
            }
        }

        private void SlideThumbnail_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.DataContext is SlideThumbnailViewModel thumbnail)
            {
                var index = _slideThumbnails.FindIndex(t => t == thumbnail);
                if (index >= 0)
                {
                    foreach (var t in _slideThumbnails)
                    {
                        t.IsSelected = false;
                    }

                    thumbnail.IsSelected = true;
                    PreviewImage.Source = thumbnail.PreviewImage;
                    SlideCounter.Text = $"Slide {index + 1} of {_slideThumbnails.Count}";

                    UpdateNavigationButtons();
                    RefreshThumbnails();
                }
            }
        }

        private void UpdateNavigationButtons()
        {
            var currentIndex = _slideThumbnails.FindIndex(t => t.IsSelected);
            PrevSlideButton.IsEnabled = currentIndex > 0;
            NextSlideButton.IsEnabled = currentIndex < _slideThumbnails.Count - 1;
        }

        private void RefreshThumbnails()
        {
            SlideThumbnails.ItemsSource = null;
            SlideThumbnails.ItemsSource = _slideThumbnails;
        }

        private void UpdateSettingsFromUI()
        {
            if (_selectedExport?.Settings == null) return;

            _selectedExport.Settings.FontSize = TextSizeSlider.Value;
            _selectedExport.Settings.CellOpacity = CellOpacitySlider.Value;

            _selectedExport.Settings.TextColor = GetColorHex(TextColorPreview);
            _selectedExport.Settings.NameCellColor = GetColorHex(NameCellColorPreview);
            _selectedExport.Settings.TimeCellColor = GetColorHex(TimeCellColorPreview);
            _selectedExport.Settings.DaysRowColor = GetColorHex(DaysRowColorPreview);
            _selectedExport.Settings.BackgroundColor = GetColorHex(BackgroundColorPreview);

            if (string.IsNullOrEmpty(_selectedExport.Settings.NameCellColor))
                _selectedExport.Settings.NameCellColor = "#FFFFFF";
            if (string.IsNullOrEmpty(_selectedExport.Settings.TimeCellColor))
                _selectedExport.Settings.TimeCellColor = "#FFFFFF";
            if (string.IsNullOrEmpty(_selectedExport.Settings.DaysRowColor))
                _selectedExport.Settings.DaysRowColor = "#FFFFFF";

            ExportService.UpdateExportSettings(_selectedExport.Id, _selectedExport.Settings);
        }

        private string GetColorHex(Border colorBorder)
        {
            var brush = colorBorder.Background as SolidColorBrush;
            return brush != null ? $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}" : "#FFFFFF";
        }

        private async void SettingChanged(object sender, RoutedEventArgs e)
        {
            if (_selectedExport != null)
            {
                UpdateSettingsFromUI();

                if (_slideThumbnails.Count > 0)
                {
                    await GeneratePreviewAsync();
                }
            }
        }

        private void TextSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_selectedExport != null && !_isGeneratingPreview)
            {
                UpdateSettingsFromUI();
            }
        }

        private void CellOpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_selectedExport != null && !_isGeneratingPreview)
            {
                UpdateSettingsFromUI();
            }
        }

        private async void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedExport == null) return;

            try
            {
                UpdateSettingsFromUI();

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "PowerPoint Presentations (*.pptx)|*.pptx",
                    FileName = $"{_selectedExport.Name.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx",
                    DefaultExt = ".pptx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var exportSchedule = ExportService.GetExportScheduleById(_selectedExport.Id);
                    if (exportSchedule != null)
                    {
                        ExportButton.IsEnabled = false;
                        ExportStatusText.Text = "Exporting...";
                        ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F59E0B"));

                        string exportResult = null;

                        try
                        {
                            exportResult = await ExportService.ExportToPowerPointAsync(exportSchedule,
                                new Progress<string>(message => Dispatcher.Invoke(() =>
                                {
                                    ExportStatusText.Text = message;
                                })));
                        }
                        catch (Exception ex)
                        {
                            ExportStatusText.Text = "Error";
                            ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444"));
                            ShowNotification($"Export failed: {ex.Message}", "Export Error", "error");
                            ExportButton.IsEnabled = true;
                            return;
                        }

                        if (!string.IsNullOrEmpty(exportResult) && System.IO.File.Exists(exportResult))
                        {
                            ExportStatusText.Text = "Ready";
                            ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#10B981"));

                            bool openResult = await ShowConfirmDialog(
                                "Would you like to open the exported file now?",
                                "Open File");

                            if (openResult)
                            {
                                System.Diagnostics.Process.Start(exportResult);
                            }

                            ShowNotification($"PowerPoint exported successfully to {System.IO.Path.GetFileName(exportResult)}", "Export Complete", "success");
                        }
                        else
                        {
                            ExportStatusText.Text = "Error";
                            ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444"));
                            ShowNotification("Failed to export PowerPoint file.", "Export Error", "error");
                        }

                        ExportButton.IsEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                ExportStatusText.Text = "Error";
                ExportStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EF4444"));
                ShowNotification($"Error exporting: {ex.Message}", "Export Error", "error");
                ExportButton.IsEnabled = true;
            }
        }

        private async void RenameExport_Click(object sender, RoutedEventArgs e)
        {
            string exportId = null;

            if (sender is Button button && button.Tag is string id)
            {
                exportId = id;
            }
            else if (_selectedExport != null)
            {
                exportId = _selectedExport.Id;
            }

            if (exportId == null) return;

            var export = _exportSchedules.FirstOrDefault(exp => exp.Id == exportId);
            if (export == null) return;

            string newName = await ShowInputDialog(
                "Enter new name for the export:",
                "Rename Export",
                export.Name);

            if (!string.IsNullOrWhiteSpace(newName))
            {
                try
                {
                    bool success = ExportService.RenameExportSchedule(export.Id, newName);
                    if (success)
                    {
                        export.Name = newName;

                        if (_selectedExport != null && _selectedExport.Id == export.Id)
                        {
                            ExportDetailTitle.Text = $"Adjustments: {newName}";
                        }

                        UpdateExportList();
                        ShowNotification($"Export renamed to '{newName}'", "Success", "success");
                    }
                }
                catch (Exception ex)
                {
                    ShowNotification($"Error renaming export: {ex.Message}", "Error", "error");
                }
            }
        }

        private async void DeleteExport_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string exportId)
            {
                var export = _exportSchedules.FirstOrDefault(exp => exp.Id == exportId);
                if (export != null)
                {
                    bool confirm = await ShowConfirmDialog(
                        $"Are you sure you want to delete export '{export.Name}'?",
                        "Confirm Delete");

                    if (confirm)
                    {
                        try
                        {
                            bool success = ExportService.DeleteExportSchedule(exportId);
                            if (success)
                            {
                                _exportSchedules.RemoveAll(exp => exp.Id == exportId);

                                if (_selectedExport != null && _selectedExport.Id == exportId)
                                {
                                    _selectedExport = null;
                                    ShowListView();
                                }

                                UpdateExportList();
                                ShowNotification($"Export '{export.Name}' deleted", "Success", "success");
                            }
                        }
                        catch (Exception ex)
                        {
                            ShowNotification($"Error deleting export: {ex.Message}", "Error", "error");
                        }
                    }
                }
            }
        }

        private async void TextColorPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            await ShowColorPicker(TextColorPreview, "Choose color for text");
        }

        private async void NameCellColorPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            await ShowColorPicker(NameCellColorPreview, "Choose color for name cells");
        }

        private async void TimeCellColorPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            await ShowColorPicker(TimeCellColorPreview, "Choose color for time cells");
        }

        private async void DaysRowColorPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            await ShowColorPicker(DaysRowColorPreview, "Choose color for days row");
        }

        private async void BackgroundColorPreview_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            await ShowColorPicker(BackgroundColorPreview, "Choose color for background");
        }

        private async Task ShowColorPicker(Border targetBorder, string title)
        {
            var currentColor = ((SolidColorBrush)targetBorder.Background)?.Color ?? Colors.Black;

            if (_mainWindow?.PopupSystemControl != null)
            {
                var selectedColor = await _mainWindow.PopupSystemControl.ShowColorDialog(title, currentColor);

                if (selectedColor.HasValue)
                {
                    targetBorder.Background = new SolidColorBrush(selectedColor.Value);
                    SettingChanged(targetBorder, new RoutedEventArgs());
                }
            }
            else
            {
                var mainWindow = Window.GetWindow(this) as MainWindow;
                if (mainWindow?.PopupSystemControl != null)
                {
                    var selectedColor = await mainWindow.PopupSystemControl.ShowColorDialog(title, currentColor);

                    if (selectedColor.HasValue)
                    {
                        targetBorder.Background = new SolidColorBrush(selectedColor.Value);
                        SettingChanged(targetBorder, new RoutedEventArgs());
                    }
                }
            }
        }

        private async Task<bool> ShowConfirmDialog(string message, string title = "Confirm")
        {
            if (_mainWindow?.PopupSystemControl != null)
            {
                return await _mainWindow.PopupSystemControl.ShowConfirmDialog(message, title);
            }

            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow?.PopupSystemControl != null)
            {
                return await mainWindow.PopupSystemControl.ShowConfirmDialog(message, title);
            }

            throw new InvalidOperationException("PopupSystem not available");
        }

        private async Task<string> ShowInputDialog(string message, string title = "Input", string defaultValue = "")
        {
            if (_mainWindow?.PopupSystemControl != null)
            {
                return await _mainWindow.PopupSystemControl.ShowInputDialog(message, title, defaultValue);
            }

            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow?.PopupSystemControl != null)
            {
                return await mainWindow.PopupSystemControl.ShowInputDialog(message, title, defaultValue);
            }

            throw new InvalidOperationException("PopupSystem not available");
        }

        private void ShowNotification(string message, string title = "Notification", string type = "info")
        {
            if (_mainWindow?.PopupSystemControl != null)
            {
                _mainWindow.PopupSystemControl.ShowNotification(message, title, type);
                return;
            }

            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow?.PopupSystemControl != null)
            {
                mainWindow.PopupSystemControl.ShowNotification(message, title, type);
                return;
            }

            throw new InvalidOperationException("PopupSystem not available");
        }
    }
}