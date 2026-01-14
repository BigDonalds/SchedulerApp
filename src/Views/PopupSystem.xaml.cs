using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Threading;

namespace SchedulerApp.Views
{
    public partial class PopupSystem : UserControl
    {
        public event Action<string> InputDialogClosed;
        public event Action<string> SelectionDialogClosed;
        public event Action<bool> ConfirmDialogClosed;
        public event Action<Color?> ColorDialogClosed;

        private readonly List<string> _colorPalette = new List<string>
        {
            "#1F2937", "#374151", "#4B5563", "#6B7280", "#9CA3AF",
            "#4F46E5", "#3B82F6", "#0EA5E9", "#06B6D4", "#10B981",
            "#84CC16", "#EAB308", "#F59E0B", "#F97316", "#EF4444",
            "#EC4899", "#8B5CF6", "#6366F1", "#14B8A6", "#22C55E"
        };

        private Color? _selectedColor;
        private Border _currentSelectedColorBorder;
        private string _currentDialogType = "";
        private bool _isClosing = false;
        private bool _isShowing = false;
        private List<string> _activeHandlers = new List<string>();
        private readonly Queue<Action> _dialogQueue = new Queue<Action>();
        private readonly object _queueLock = new object();

        public PopupSystem()
        {
            InitializeComponent();

            OverlayBackground.MouseLeftButtonDown += (s, e) =>
            {
                if (!_isClosing)
                {
                    HidePopup();
                }
            };
        }

        private void QueueOrExecuteDialog(Action dialogAction)
        {
            lock (_queueLock)
            {
                if (_isShowing || _isClosing)
                {
                    _dialogQueue.Enqueue(dialogAction);
                }
                else
                {
                    _isShowing = true;
                    dialogAction();
                }
            }
        }

        private void ProcessNextDialog()
        {
            lock (_queueLock)
            {
                if (_dialogQueue.Count > 0)
                {
                    var nextDialog = _dialogQueue.Dequeue();
                    Dispatcher.Invoke(() =>
                    {
                        _isShowing = true;
                        nextDialog();
                    });
                }
                else
                {
                    _isShowing = false;
                }
            }
        }

        public Task<Color?> ShowColorDialog(string title = "Choose a Color", Color? initialColor = null)
        {
            var tcs = new TaskCompletionSource<Color?>();

            QueueOrExecuteDialog(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    _currentDialogType = "ColorDialog";
                    MouseButtonEventHandler overlayHandler = null;
                    RoutedEventHandler confirmHandler = null;
                    RoutedEventHandler cancelHandler = null;

                    try
                    {
                        _selectedColor = initialColor;
                        _currentSelectedColorBorder = null;
                        SidePanelContent.Children.Clear();
                        SidePanelButtons.Children.Clear();
                        SidePanelTitle.Text = title;

                        var contentStack = new StackPanel();

                        if (title.StartsWith("Choose a color for"))
                        {
                            var messageText = new TextBlock
                            {
                                Text = title,
                                FontSize = 14,
                                Foreground = Brushes.Black,
                                TextWrapping = TextWrapping.Wrap,
                                TextAlignment = TextAlignment.Center,
                                Margin = new Thickness(0, 0, 0, 16)
                            };
                            contentStack.Children.Add(messageText);
                        }

                        var paletteContainer = new WrapPanel
                        {
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 20)
                        };

                        foreach (var colorHex in _colorPalette)
                        {
                            var colorBorder = CreateColorItem(colorHex, initialColor);
                            paletteContainer.Children.Add(colorBorder);
                        }

                        contentStack.Children.Add(paletteContainer);

                        var selectionStack = new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 20)
                        };

                        var selectionText = new TextBlock
                        {
                            Text = "Selected:",
                            FontSize = 13,
                            Foreground = Brushes.Gray,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(0, 0, 12, 0)
                        };

                        var selectedColorBorder = new Border
                        {
                            Width = 32,
                            Height = 32,
                            CornerRadius = new CornerRadius(8),
                            Background = initialColor.HasValue ?
                                new SolidColorBrush(initialColor.Value) :
                                Brushes.Transparent,
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Margin = new Thickness(0, 0, 8, 0)
                        };

                        var selectedColorText = new TextBlock
                        {
                            Text = initialColor.HasValue ?
                                $"#{initialColor.Value.R:X2}{initialColor.Value.G:X2}{initialColor.Value.B:X2}" :
                                "None",
                            FontSize = 13,
                            Foreground = Brushes.Black,
                            VerticalAlignment = VerticalAlignment.Center,
                            FontFamily = new FontFamily("Consolas")
                        };

                        selectionStack.Children.Add(selectionText);
                        selectionStack.Children.Add(selectedColorBorder);
                        selectionStack.Children.Add(selectedColorText);
                        contentStack.Children.Add(selectionStack);
                        SidePanelContent.Children.Add(contentStack);

                        var confirmButton = new Button
                        {
                            Content = "Confirm",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelButton"),
                            Margin = new Thickness(0, 0, 8, 0),
                            Name = "ColorConfirmButton"
                        };

                        var cancelButton = new Button
                        {
                            Content = "Cancel",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelSecondaryButton"),
                            Name = "ColorCancelButton"
                        };

                        confirmHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(_selectedColor);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("ColorOverlayHandler");
                                }
                                HidePopup();
                                ColorDialogClosed?.Invoke(_selectedColor);
                            }
                        };

                        cancelHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(null);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("ColorOverlayHandler");
                                }
                                HidePopup();
                                ColorDialogClosed?.Invoke(null);
                            }
                        };

                        confirmButton.Click += confirmHandler;
                        cancelButton.Click += cancelHandler;

                        if (!initialColor.HasValue)
                        {
                            confirmButton.IsEnabled = false;
                        }

                        SidePanelButtons.Children.Add(cancelButton);
                        SidePanelButtons.Children.Add(confirmButton);
                        ShowPopup();

                        overlayHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                if (e.OriginalSource is FrameworkElement element)
                                {
                                    if (element.Name != "SidePanelContainer" &&
                                        !IsChildOfSidePanel(element))
                                    {
                                        tcs.TrySetResult(null);
                                        if (overlayHandler != null)
                                        {
                                            OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                            _activeHandlers.Remove("ColorOverlayHandler");
                                        }
                                        HidePopup();
                                        ColorDialogClosed?.Invoke(null);
                                    }
                                }
                            }
                        };

                        OverlayBackground.MouseLeftButtonDown += overlayHandler;
                        _activeHandlers.Add("ColorOverlayHandler");
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                        ProcessNextDialog();
                    }
                });
            });

            return tcs.Task;
        }

        private Border CreateColorItem(string colorHex, Color? initialColor)
        {
            var color = (Color)ColorConverter.ConvertFromString(colorHex);
            var isSelected = initialColor.HasValue && initialColor.Value == color;

            var colorBorder = new Border
            {
                Width = 40,
                Height = 40,
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(color),
                Margin = new Thickness(6),
                Cursor = Cursors.Hand,
                Tag = colorHex
            };

            if (isSelected)
            {
                colorBorder.BorderBrush = Brushes.Black;
                colorBorder.BorderThickness = new Thickness(3);
                _currentSelectedColorBorder = colorBorder;
            }
            else
            {
                colorBorder.BorderBrush = Brushes.LightGray;
                colorBorder.BorderThickness = new Thickness(1);
            }

            colorBorder.ToolTip = colorHex;

            colorBorder.MouseEnter += (s, e) =>
            {
                var border = s as Border;
                if (border != _currentSelectedColorBorder)
                {
                    var scaleAnimation = new DoubleAnimation(1.1, TimeSpan.FromMilliseconds(150));
                    border.RenderTransform = new ScaleTransform();
                    border.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnimation);
                    border.RenderTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnimation);

                    border.Effect = new DropShadowEffect
                    {
                        BlurRadius = 10,
                        ShadowDepth = 2,
                        Opacity = 0.3,
                        Color = Colors.Black
                    };
                }
            };

            colorBorder.MouseLeave += (s, e) =>
            {
                var border = s as Border;
                if (border != _currentSelectedColorBorder)
                {
                    var scaleAnimation = new DoubleAnimation(1.0, TimeSpan.FromMilliseconds(150));
                    border.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnimation);
                    border.RenderTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnimation);
                    border.Effect = null;
                }
            };

            colorBorder.MouseLeftButtonDown += (s, e) =>
            {
                var border = s as Border;

                if (_currentSelectedColorBorder != null && _currentSelectedColorBorder != border)
                {
                    _currentSelectedColorBorder.BorderBrush = Brushes.LightGray;
                    _currentSelectedColorBorder.BorderThickness = new Thickness(1);
                }

                border.BorderBrush = Brushes.Black;
                border.BorderThickness = new Thickness(3);
                _currentSelectedColorBorder = border;

                try
                {
                    _selectedColor = (Color)ColorConverter.ConvertFromString(colorHex);

                    var confirmButton = SidePanelButtons.Children
                        .OfType<Button>()
                        .FirstOrDefault(b => b.Content.ToString() == "Confirm");
                    if (confirmButton != null)
                    {
                        confirmButton.IsEnabled = true;
                    }

                    UpdateSelectionIndicator(colorHex, color);
                }
                catch
                {
                    _selectedColor = Colors.Black;
                }
            };

            return colorBorder;
        }

        private void UpdateSelectionIndicator(string colorHex, Color color)
        {
            var selectionStack = SidePanelContent.Children
                .OfType<StackPanel>()
                .FirstOrDefault(sp => sp.Children.Count > 0 && sp.Children[0] is TextBlock textBlock && textBlock.Text == "Selected:");

            if (selectionStack != null && selectionStack.Children.Count >= 3)
            {
                var colorBorder = selectionStack.Children[1] as Border;
                if (colorBorder != null)
                {
                    colorBorder.Background = new SolidColorBrush(color);
                }

                var colorText = selectionStack.Children[2] as TextBlock;
                if (colorText != null)
                {
                    colorText.Text = colorHex;
                }
            }
        }

        public void ShowNotification(string message, string title = "Notification", string type = "info", int duration = 4000)
        {
            QueueOrExecuteDialog(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    _currentDialogType = "Notification";

                    try
                    {
                        SidePanelContent.Children.Clear();
                        SidePanelButtons.Children.Clear();
                        SidePanelTitle.Text = title;

                        var contentStack = new StackPanel();
                        Border iconBorder = new Border();
                        TextBlock iconText = new TextBlock
                        {
                            FontSize = 20,
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        switch (type.ToLower())
                        {
                            case "success":
                                iconBorder.Style = (Style)FindResource("NotificationIconSuccess");
                                iconText.Text = "✓";
                                iconText.Foreground = Brushes.Green;
                                break;
                            case "error":
                                iconBorder.Style = (Style)FindResource("NotificationIconError");
                                iconText.Text = "✕";
                                iconText.Foreground = Brushes.Red;
                                break;
                            case "warning":
                                iconBorder.Style = (Style)FindResource("NotificationIconWarning");
                                iconText.Text = "⚠";
                                iconText.Foreground = Brushes.Orange;
                                break;
                            default:
                                iconBorder.Style = (Style)FindResource("NotificationIconInfo");
                                iconText.Text = "ℹ";
                                iconText.Foreground = Brushes.Blue;
                                break;
                        }

                        iconBorder.Child = iconText;
                        contentStack.Children.Add(iconBorder);
                        contentStack.Children.Add(new Border { Height = 20 });

                        var messageText = new TextBlock
                        {
                            Text = message,
                            FontSize = 14,
                            Foreground = Brushes.Black,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 20)
                        };
                        contentStack.Children.Add(messageText);
                        SidePanelContent.Children.Add(contentStack);

                        var closeButton = new Button
                        {
                            Content = "Close",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelButton"),
                            Tag = "close"
                        };

                        closeButton.Click += (s, e) =>
                        {
                            HidePopup();
                        };

                        SidePanelButtons.Children.Add(closeButton);
                        ShowPopup();

                        if (duration > 0)
                        {
                            var timer = new DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(duration)
                            };
                            timer.Tick += (s, e) =>
                            {
                                timer.Stop();
                                HidePopup();
                            };
                            timer.Start();
                        }
                    }
                    catch (Exception)
                    {
                        ProcessNextDialog();
                    }
                });
            });
        }

        public Task<bool> ShowConfirmDialog(string message, string title = "Confirm")
        {
            var tcs = new TaskCompletionSource<bool>();

            QueueOrExecuteDialog(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    _currentDialogType = "ConfirmDialog";
                    MouseButtonEventHandler overlayHandler = null;
                    RoutedEventHandler yesHandler = null;
                    RoutedEventHandler noHandler = null;

                    try
                    {
                        SidePanelContent.Children.Clear();
                        SidePanelButtons.Children.Clear();
                        SidePanelTitle.Text = title;

                        var contentStack = new StackPanel();

                        var iconBorder = new Border
                        {
                            Style = (Style)FindResource("NotificationIconWarning"),
                            Width = 48,
                            Height = 48,
                            HorizontalAlignment = HorizontalAlignment.Center
                        };

                        var iconText = new TextBlock
                        {
                            Text = "?",
                            FontSize = 20,
                            FontWeight = FontWeights.Bold,
                            Foreground = Brushes.Orange,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        iconBorder.Child = iconText;
                        contentStack.Children.Add(iconBorder);
                        contentStack.Children.Add(new Border { Height = 20 });

                        var messageText = new TextBlock
                        {
                            Text = message,
                            FontSize = 14,
                            Foreground = Brushes.Black,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 20)
                        };
                        contentStack.Children.Add(messageText);
                        SidePanelContent.Children.Add(contentStack);

                        var yesButton = new Button
                        {
                            Content = "Yes",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelButton"),
                            Margin = new Thickness(0, 0, 8, 0),
                            Name = "ConfirmYesButton"
                        };

                        var noButton = new Button
                        {
                            Content = "No",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelSecondaryButton"),
                            Name = "ConfirmNoButton"
                        };

                        yesHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(true);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("ConfirmOverlayHandler");
                                }
                                HidePopup();
                                ConfirmDialogClosed?.Invoke(true);
                            }
                        };

                        noHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(false);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("ConfirmOverlayHandler");
                                }
                                HidePopup();
                                ConfirmDialogClosed?.Invoke(false);
                            }
                        };

                        yesButton.Click += yesHandler;
                        noButton.Click += noHandler;

                        SidePanelButtons.Children.Add(yesButton);
                        SidePanelButtons.Children.Add(noButton);
                        ShowPopup();

                        overlayHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                if (e.OriginalSource is FrameworkElement element)
                                {
                                    if (element.Name != "SidePanelContainer" &&
                                        !IsChildOfSidePanel(element))
                                    {
                                        tcs.TrySetResult(false);
                                        if (overlayHandler != null)
                                        {
                                            OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                            _activeHandlers.Remove("ConfirmOverlayHandler");
                                        }
                                        HidePopup();
                                        ConfirmDialogClosed?.Invoke(false);
                                    }
                                }
                            }
                        };

                        OverlayBackground.MouseLeftButtonDown += overlayHandler;
                        _activeHandlers.Add("ConfirmOverlayHandler");
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                        ProcessNextDialog();
                    }
                });
            });

            return tcs.Task;
        }

        public Task<string> ShowInputDialog(string message, string title = "Input", string defaultValue = "")
        {
            var tcs = new TaskCompletionSource<string>();

            QueueOrExecuteDialog(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    _currentDialogType = "InputDialog";
                    MouseButtonEventHandler overlayHandler = null;
                    KeyEventHandler textBoxKeyHandler = null;
                    RoutedEventHandler okHandler = null;
                    RoutedEventHandler cancelHandler = null;

                    try
                    {
                        SidePanelContent.Children.Clear();
                        SidePanelButtons.Children.Clear();
                        SidePanelTitle.Text = title;

                        var contentStack = new StackPanel();

                        var messageText = new TextBlock
                        {
                            Text = message,
                            FontSize = 14,
                            Foreground = Brushes.Black,
                            TextWrapping = TextWrapping.Wrap,
                            Margin = new Thickness(0, 0, 0, 12)
                        };
                        contentStack.Children.Add(messageText);

                        var textBox = new TextBox
                        {
                            Style = (Style)FindResource("InputDialogTextBox"),
                            Text = defaultValue,
                            Margin = new Thickness(0, 0, 0, 20),
                            Name = "InputDialogTextBox"
                        };
                        contentStack.Children.Add(textBox);
                        SidePanelContent.Children.Add(contentStack);

                        var okButton = new Button
                        {
                            Content = "OK",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelButton"),
                            Margin = new Thickness(0, 0, 8, 0),
                            Name = "InputOkButton"
                        };

                        var cancelButton = new Button
                        {
                            Content = "Cancel",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelSecondaryButton"),
                            Name = "InputCancelButton"
                        };

                        okHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(textBox.Text);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("InputOverlayHandler");
                                }
                                if (textBoxKeyHandler != null)
                                {
                                    textBox.KeyDown -= textBoxKeyHandler;
                                    _activeHandlers.Remove("InputKeyHandler");
                                }
                                HidePopup();
                                InputDialogClosed?.Invoke(textBox.Text);
                            }
                        };

                        cancelHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(null);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("InputOverlayHandler");
                                }
                                if (textBoxKeyHandler != null)
                                {
                                    textBox.KeyDown -= textBoxKeyHandler;
                                    _activeHandlers.Remove("InputKeyHandler");
                                }
                                HidePopup();
                                InputDialogClosed?.Invoke(null);
                            }
                        };

                        textBoxKeyHandler = (s, e) =>
                        {
                            if (e.Key == Key.Enter)
                            {
                                e.Handled = true;
                                okHandler(null, null);
                            }
                            else if (e.Key == Key.Escape)
                            {
                                e.Handled = true;
                                cancelHandler(null, null);
                            }
                        };

                        okButton.Click += okHandler;
                        cancelButton.Click += cancelHandler;
                        textBox.KeyDown += textBoxKeyHandler;

                        SidePanelButtons.Children.Add(okButton);
                        SidePanelButtons.Children.Add(cancelButton);
                        ShowPopup();
                        textBox.Focus();
                        textBox.SelectAll();

                        overlayHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                if (e.OriginalSource is FrameworkElement element)
                                {
                                    if (element.Name != "SidePanelContainer" &&
                                        !IsChildOfSidePanel(element))
                                    {
                                        tcs.TrySetResult(null);
                                        if (overlayHandler != null)
                                        {
                                            OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                            _activeHandlers.Remove("InputOverlayHandler");
                                        }
                                        if (textBoxKeyHandler != null)
                                        {
                                            textBox.KeyDown -= textBoxKeyHandler;
                                            _activeHandlers.Remove("InputKeyHandler");
                                        }
                                        HidePopup();
                                        InputDialogClosed?.Invoke(null);
                                    }
                                }
                            }
                        };

                        OverlayBackground.MouseLeftButtonDown += overlayHandler;
                        _activeHandlers.Add("InputOverlayHandler");
                        _activeHandlers.Add("InputKeyHandler");
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                        ProcessNextDialog();
                    }
                });
            });

            return tcs.Task;
        }

        public Task<string> ShowSelectionDialog(string message, string title, List<string> items)
        {
            var tcs = new TaskCompletionSource<string>();

            QueueOrExecuteDialog(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    _currentDialogType = "SelectionDialog";
                    MouseButtonEventHandler overlayHandler = null;
                    KeyEventHandler listBoxKeyHandler = null;
                    RoutedEventHandler selectHandler = null;
                    RoutedEventHandler cancelHandler = null;

                    try
                    {
                        SidePanelContent.Children.Clear();
                        SidePanelButtons.Children.Clear();
                        SidePanelTitle.Text = title;

                        var contentStack = new StackPanel();

                        var messageText = new TextBlock
                        {
                            Text = message,
                            FontSize = 14,
                            Foreground = Brushes.Black,
                            TextWrapping = TextWrapping.Wrap,
                            Margin = new Thickness(0, 0, 0, 12)
                        };
                        contentStack.Children.Add(messageText);

                        var listBox = new ListBox
                        {
                            Height = 150,
                            Margin = new Thickness(0, 0, 0, 20),
                            BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D1D5DB")),
                            BorderThickness = new Thickness(1),
                            Padding = new Thickness(8),
                            Name = "SelectionListBox"
                        };

                        foreach (var item in items)
                        {
                            listBox.Items.Add(item);
                        }

                        if (listBox.Items.Count > 0)
                        {
                            listBox.SelectedIndex = 0;
                        }

                        contentStack.Children.Add(listBox);
                        SidePanelContent.Children.Add(contentStack);

                        var selectButton = new Button
                        {
                            Content = "Select",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelButton"),
                            Margin = new Thickness(0, 0, 8, 0),
                            Name = "SelectionSelectButton"
                        };

                        var cancelButton = new Button
                        {
                            Content = "Cancel",
                            Width = 100,
                            Height = 40,
                            Style = (Style)FindResource("SidePanelSecondaryButton"),
                            Name = "SelectionCancelButton"
                        };

                        selectHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(listBox.SelectedItem as string);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("SelectionOverlayHandler");
                                }
                                if (listBoxKeyHandler != null)
                                {
                                    listBox.KeyDown -= listBoxKeyHandler;
                                    _activeHandlers.Remove("SelectionKeyHandler");
                                }
                                HidePopup();
                                SelectionDialogClosed?.Invoke(listBox.SelectedItem as string);
                            }
                        };

                        cancelHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                tcs.SetResult(null);
                                if (overlayHandler != null)
                                {
                                    OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                    _activeHandlers.Remove("SelectionOverlayHandler");
                                }
                                if (listBoxKeyHandler != null)
                                {
                                    listBox.KeyDown -= listBoxKeyHandler;
                                    _activeHandlers.Remove("SelectionKeyHandler");
                                }
                                HidePopup();
                                SelectionDialogClosed?.Invoke(null);
                            }
                        };

                        listBoxKeyHandler = (s, e) =>
                        {
                            if (e.Key == Key.Enter)
                            {
                                e.Handled = true;
                                selectHandler(null, null);
                            }
                            else if (e.Key == Key.Escape)
                            {
                                e.Handled = true;
                                cancelHandler(null, null);
                            }
                        };

                        selectButton.Click += selectHandler;
                        cancelButton.Click += cancelHandler;
                        listBox.KeyDown += listBoxKeyHandler;

                        SidePanelButtons.Children.Add(selectButton);
                        SidePanelButtons.Children.Add(cancelButton);
                        ShowPopup();
                        listBox.Focus();

                        overlayHandler = (s, e) =>
                        {
                            if (!tcs.Task.IsCompleted && !_isClosing)
                            {
                                if (e.OriginalSource is FrameworkElement element)
                                {
                                    if (element.Name != "SidePanelContainer" &&
                                        !IsChildOfSidePanel(element))
                                    {
                                        tcs.TrySetResult(null);
                                        if (overlayHandler != null)
                                        {
                                            OverlayBackground.MouseLeftButtonDown -= overlayHandler;
                                            _activeHandlers.Remove("SelectionOverlayHandler");
                                        }
                                        if (listBoxKeyHandler != null)
                                        {
                                            listBox.KeyDown -= listBoxKeyHandler;
                                            _activeHandlers.Remove("SelectionKeyHandler");
                                        }
                                        HidePopup();
                                        SelectionDialogClosed?.Invoke(null);
                                    }
                                }
                            }
                        };

                        OverlayBackground.MouseLeftButtonDown += overlayHandler;
                        _activeHandlers.Add("SelectionOverlayHandler");
                        _activeHandlers.Add("SelectionKeyHandler");
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                        ProcessNextDialog();
                    }
                });
            });

            return tcs.Task;
        }

        private void ShowPopup()
        {
            _isClosing = false;
            OverlayBackground.Visibility = Visibility.Visible;
            SidePanelContainer.Visibility = Visibility.Visible;

            var slideIn = new DoubleAnimation
            {
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3),
                DecelerationRatio = 0.9
            };

            SidePanelTransform.BeginAnimation(TranslateTransform.XProperty, slideIn);
        }

        private void HidePopup()
        {
            if (SidePanelContainer.Visibility != Visibility.Visible || _isClosing)
            {
                return;
            }

            _isClosing = true;

            var slideOut = new DoubleAnimation
            {
                To = 420,
                Duration = TimeSpan.FromSeconds(0.3),
                AccelerationRatio = 0.9
            };

            slideOut.Completed += (s, e) =>
            {
                SidePanelContainer.Visibility = Visibility.Collapsed;
                OverlayBackground.Visibility = Visibility.Collapsed;
                SidePanelContent.Children.Clear();
                SidePanelButtons.Children.Clear();
                _selectedColor = null;
                _currentSelectedColorBorder = null;
                _activeHandlers.Clear();
                _currentDialogType = "";
                _isClosing = false;
                ProcessNextDialog();
            };

            SidePanelTransform.BeginAnimation(TranslateTransform.XProperty, slideOut);
        }

        private bool IsChildOfSidePanel(FrameworkElement element)
        {
            try
            {
                DependencyObject parent = element;
                while (parent != null)
                {
                    if (parent == SidePanelContainer)
                        return true;
                    parent = VisualTreeHelper.GetParent(parent);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}