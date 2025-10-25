using System;
using System.IO;
using System.Data;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Configuration;
using System.Windows.Input;
using System.Windows.Media;
using System.Globalization;
using System.Windows.Controls;
using Point = System.Drawing.Point;
using System.Runtime.InteropServices;
using static PointAC.AppFileOperations;
using WindowState = System.Windows.WindowState;
using ClickType = PointAC.MouseHandler.ClickType;
using MouseButton = PointAC.MouseHandler.MouseButton;
using Orientation = System.Windows.Controls.Orientation;

namespace PointAC
{
    public partial class MainWindow : Window
    {
        #region App Main
        public static string AppVersion { get; } = "1.0";
        #endregion

        #region Window Settings
        private WindowState windowState;
        private double windowLeft, windowTop;
        private double windowWidth, windowHeight;
        #endregion

        #region App Settings
        private int loopCount = 0;
        private bool looped = true;
        private int duration = 1000;
        private bool alwaysOnTop = true;
        private string theme = "System";
        private string button = "System";
        private string menuStyle = "Top";
        private string language = "System";
        private string clickType = "Single";
        private HashSet<Key> toggleKey = new() { Key.F6 };
        #endregion

        #region Variables
        private bool updateAvailable;
        private MouseButton currentButton;
        private ClickType currentClickType;
        private bool canChangeHotkey = false;
        private string currentMode = "Normal";

        private Point lastMousePosition;
        private PointEntry? lastHovered;
        private readonly PointManager pointManager = new();
        private readonly HashSet<Key> currentlyPressedKeys = new();

        private Task? runningTask;
        private bool isRunning = false;
        private readonly object runLock = new();

        private Configuration? configuration;
        private KeyValueConfigurationCollection? appSettings;
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                appSettings = configuration.AppSettings.Settings;

                theme = GetSetting("Theme", theme);
                button = GetSetting("Button", button);
                looped = GetSetting("Looped", looped);
                language = GetSetting("Language", language);
                duration = GetSetting("Duration", duration);
                menuStyle = GetSetting("MenuStyle", menuStyle);
                loopCount = GetSetting("LoopCount", loopCount);
                clickType = GetSetting("ClickType", clickType);
                toggleKey = GetSetting("ToggleHotkey", toggleKey);
                alwaysOnTop = GetSetting("AlwaysOnTop", alwaysOnTop);
                windowState = GetSetting("WindowState", WindowState.Normal);
                var size = GetSetting("WindowSize", new Size(Width, Height));
                var location = GetSetting("WindowLocation", new Point((int)Left, (int)Top));
                
                windowTop = location.Y;
                windowLeft = location.X;
                windowWidth = size.Width;
                windowHeight = size.Height;

                this.ThemeMode = theme switch
                {
                    "Light" => ThemeMode.Light,
                    "Dark" => ThemeMode.Dark,
                    _ => ThemeMode.System
                };
                
                Top = windowTop;
                Left = windowLeft;
                Width = windowWidth;
                Height = windowHeight;
                ApplyLanguage(language);
                WindowState = windowState;
                currentButton = MouseHandler.GetMouseButtonFromString(button);
                currentClickType = MouseHandler.GetClickTypeFromString(clickType);

                configuration.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Config Error] {ex.Message}");
            }

            PointsList.ItemsSource = pointManager.Points;

            pointManager.Points.CollectionChanged += (_, e) =>
            {
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        PointsScrollViewer.ScrollToEnd();
                    }), System.Windows.Threading.DispatcherPriority.Background);
                }
            };

            MouseHandler.ListenToMouseClick(screenPoint =>
            {
                if (IsClickInsideAppWindow(screenPoint))
                {
                    var addButtonBounds = GetElementScreenBounds(AddButton);
                    if (addButtonBounds.Contains(screenPoint))
                        return false;

                    return false;
                }

                if (currentMode == "Add")
                {
                    pointManager.AddPoint(screenPoint, currentButton, currentClickType, duration);
                    return true;
                }
                else if (currentMode == "Remove")
                {
                    if (pointManager.RemoveNearest(screenPoint))
                    {
                        return true;
                    }
                }

                return false;
            });

            MouseHandler.ListenToMouseMove(screenPoint =>
            {
                if (pointManager.IsRuntimeMode || currentMode == "Add")
                    return;

                if (Math.Abs(screenPoint.X - lastMousePosition.X) < 2 &&
                    Math.Abs(screenPoint.Y - lastMousePosition.Y) < 2)
                    return;

                lastMousePosition = screenPoint;

                var nearest = pointManager.Points
                    .OrderBy(p => Distance(p.Position, screenPoint))
                    .FirstOrDefault();

                double minDist = nearest == null ? double.MaxValue : Distance(nearest.Position, screenPoint);
                const int hoverRadius = 15;

                bool hoverChanged = nearest != lastHovered || (nearest != null && minDist > hoverRadius);

                if (!hoverChanged)
                    return;

                lastHovered = (minDist <= hoverRadius) ? nearest : null;

                if (!Application.Current.Dispatcher.CheckAccess())
                {
                    Application.Current.Dispatcher.BeginInvoke(() => UpdateHoverState(nearest, minDist, hoverRadius));
                }
                else
                {
                    UpdateHoverState(nearest, minDist, hoverRadius);
                }
            });

            KeyboardHandler.ListenToKeyDown(key =>
            {
                bool isModifier = key == Key.LeftCtrl || key == Key.RightCtrl ||
                                  key == Key.LeftAlt || key == Key.RightAlt ||
                                  key == Key.LeftShift || key == Key.RightShift;

                if (!currentlyPressedKeys.Contains(key))
                    currentlyPressedKeys.Add(key);

                if (canChangeHotkey)
                {
                    var modifiers = currentlyPressedKeys.Where(k =>
                        k == Key.LeftCtrl || k == Key.RightCtrl ||
                        k == Key.LeftAlt || k == Key.RightAlt ||
                        k == Key.LeftShift || k == Key.RightShift).ToList();

                    var normalKeys = currentlyPressedKeys.Except(modifiers).ToList();

                    if (normalKeys.Count > 1 || modifiers.Count > 3)
                        return;

                    var formattedKeys = modifiers.Concat(normalKeys)
                        .Select(k => FormatKeyName(k))
                        .ToList();

                    Dispatcher.Invoke(() =>
                    {
                        CurrentHotkeyButtons.Text = string.Join(" + ", formattedKeys);
                    });

                    toggleKey = new HashSet<Key>(modifiers.Concat(normalKeys));
                    SaveSetting("ToggleHotkey", GetFormattedToggleKey());
                    return;
                }

                if (IsShortcutPressed(new KeyEventArgs(
                    Keyboard.PrimaryDevice,
                    PresentationSource.FromVisual(Application.Current.MainWindow),
                    0, key), toggleKey))
                {
                    ToggleButton_Click(null!, null!);
                }
            });

            KeyboardHandler.ListenToKeyUp(key =>
            {
                currentlyPressedKeys.Remove(key);
            });
        }

        #region Events Methods
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LoopButton_Click(null, e);
                LoopCountTextBox.Text = loopCount.ToString();
                GlobalDurationTextBox.Text = duration.ToString();
                CurrentHotkeyButtons.Text = string.Join("+", toggleKey
            .OrderBy(k => k.ToString())
            .Select(k => FormatKeyName(k)));

                CheckForUpdatesButton_Click(null!, null!);
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }

            try
            {
                foreach (ComboBoxItem item in ClickTypeSelector.Items)
                {
                    if (item.Tag?.ToString() == clickType)
                    {
                        ClickTypeSelector.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in MouseButtonSelector.Items)
                {
                    if (item.Tag?.ToString() == button)
                    {
                        MouseButtonSelector.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in MenuStyleSelector.Items)
                {
                    if (item.Tag?.ToString() == menuStyle)
                    {
                        SetMenuDock();
                        MenuStyleSelector.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in AlwaysOnTopSelector.Items)
                {
                    if (item.Tag?.ToString()?.ToLower() == alwaysOnTop.ToString().ToLower())
                    {
                        Topmost = alwaysOnTop;
                        AlwaysOnTopSelector.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in AppThemeSelector.Items)
                {
                    if (item.Tag?.ToString() == theme)
                    {
                        AppThemeSelector.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in AppLanguageSelector.Items)
                {
                    if (item.Tag?.ToString() == language)
                    {
                        AppLanguageSelector.SelectedItem = item;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            pointManager.ClearAll();
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && Path.GetExtension(files[0]).Equals(FileType, StringComparison.OrdinalIgnoreCase))
                {
                    e.Effects = DragDropEffects.Copy;
                }
            }

            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            currentMode = "Normal";
            UpdateModeVisuals();

            try
            {
                if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                    return;

                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 0)
                    return;

                string filePath = files[0];

                if (!Path.GetExtension(filePath).Equals(FileType, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                var data = await LoadFromFileAsync(filePath);

                LoadPoints(data);
            }
            catch
            {

            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
                return;

            windowState = WindowState;
            SaveSetting("WindowState", windowState.ToString());
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState != WindowState.Normal) return;

            windowWidth = e.NewSize.Width;
            windowHeight = e.NewSize.Height;
            SaveSetting("WindowSize", $"{Math.Round(windowWidth)},{Math.Round(windowHeight)}");
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            if (WindowState != WindowState.Normal) return;

            windowTop = Top;
            windowLeft = Left;
            SaveSetting("WindowLocation", $"{windowLeft},{windowTop}");
        }

        private async void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentMode == "Running")
            {
                await StopClickingAsync();
                return;
            }

            currentMode = "Running";
            UpdateModeVisuals();

            if (SettingsContainer.Visibility == Visibility.Visible)
                SettingsButton_Click(SettingsButton, e);

            pointManager.SetRuntimeMode(true);

            await StartClickingAsync().ConfigureAwait(false);
        }

        private void LoopButton_Click(object? sender, RoutedEventArgs e)
        {
            if (sender == LoopButton)
            {
                looped = !looped;
                SaveSetting("Looped", looped.ToString());
            }


            if (looped)
            {
                LoopButton.Background = (System.Windows.Media.Brush)FindResource("ControlFillColorSecondaryBrush");
            }
            else
            {
                LoopButton.Background = Brushes.Transparent;
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            currentMode = (currentMode != "Add") ? "Add" : "Normal";
            UpdateModeVisuals();
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            currentMode = (currentMode != "Remove") ? "Remove" : "Normal";

            UpdateModeVisuals();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            pointManager.ClearAll();
            currentMode = "Normal";
            UpdateModeVisuals();
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (pointManager.Points.Count <= 0)
                return;

            currentMode = "Normal";
            UpdateModeVisuals();

            try
            {
                var saveDialog = new SaveFileDialog
                {
                    FileName = "",
                    DefaultExt = FileType,
                    Title = "Save Points File",
                    Filter = $"Point Auto Clicker File (*{FileType})|*{FileType}",
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                var data = new AppFileData
                {
                    Points = pointManager.Points.ToList()
                };

                await AppFileOperations.SaveToFileAsync(saveDialog.FileName, data);
            }
            catch
            {

            }
        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            currentMode = "Normal";
            UpdateModeVisuals();

            try
            {
                var openDialog = new Microsoft.Win32.OpenFileDialog
                {
                    DefaultExt = FileType,
                    Title = "Save Points File",
                    Filter = $"Point Auto Clicker File (*{FileType})|*{FileType}",
                };

                if (openDialog.ShowDialog() != true)
                    return;

                var data = await LoadFromFileAsync(openDialog.FileName);

                LoadPoints(data);
            }
            catch
            {

            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            currentMode = "Normal";
            UpdateModeVisuals();
            bool isSettingsOpen = !(SettingsContainer.Visibility == Visibility.Visible);

            PointsContainer.Visibility = isSettingsOpen ? Visibility.Collapsed : Visibility.Visible;
            SettingsContainer.Visibility = isSettingsOpen ? Visibility.Visible : Visibility.Collapsed;

            if (isSettingsOpen)
            {
                SettingsButton.Background = (System.Windows.Media.Brush)FindResource("ControlFillColorSecondaryBrush");
            }
            else
            {
                SettingsButton.Background = Brushes.Transparent;
            }
        }

        private void ClickTypeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string buttonTag = selectedItem.Tag?.ToString() ?? "System";
                
                clickType = buttonTag;
                SaveSetting("ClickType", buttonTag);
                currentClickType = MouseHandler.GetClickTypeFromString(buttonTag);
            }
        }

        private void MouseButtonSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string buttonTag = selectedItem.Tag?.ToString() ?? "System";
                button = buttonTag;

                SaveSetting("Button", buttonTag);
                currentButton = MouseHandler.GetMouseButtonFromString(buttonTag);
            }
        }

        private void LoopCountTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(LoopCountTextBox.Text, out int newLoopCount))
            {
                if (newLoopCount >= 0)
                {
                    loopCount = newLoopCount;
                    LoopCountTextBox.Text = newLoopCount.ToString();
                    SaveSetting("LoopCount", newLoopCount.ToString());
                }
                else
                {
                    LoopCountTextBox.Text = loopCount.ToString();
                }
            }
            else
            {
                LoopCountTextBox.Text = loopCount.ToString();
            }
        }

        private void HotkeyRegisterGrid_MouseEnter(object sender, MouseEventArgs e)
        {
            canChangeHotkey = true;
            HotkeyChangeHint.Visibility = Visibility.Collapsed;
            CurrentHotkeyButtons.Visibility = Visibility.Visible;
        }

        private void HotkeyRegisterGrid_MouseLeave(object sender, MouseEventArgs e)
        {
            canChangeHotkey = false;
            HotkeyChangeHint.Visibility = Visibility.Visible;
            CurrentHotkeyButtons.Visibility = Visibility.Collapsed;
        }

        private void MenuStyleSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                menuStyle = selectedItem.Tag?.ToString() ?? "Left";

                SetMenuDock();
                SaveSetting("MenuStyle", menuStyle);
            }
        }

        private void AlwaysOnTopSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                bool value = Convert.ToBoolean(selectedItem.Tag);

                alwaysOnTop = value;
                Topmost = value;

                SaveSetting("AlwaysOnTop", value.ToString());
            }
        }

        private void AppThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string themeTag = selectedItem.Tag?.ToString() ?? "System";
                theme = themeTag;
                this.ThemeMode = themeTag switch
                {
                    "Light" => ThemeMode.Light,
                    "Dark" => ThemeMode.Dark,
                    _ => ThemeMode.System
                };

                SaveSetting("Theme", themeTag);
            }
        }

        private void AppLanguageSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox)
                return;

            if (comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string languageTag = selectedItem.Tag?.ToString() ?? "System";

                theme = languageTag;
                ApplyLanguage(languageTag);
                SaveSetting("Language", languageTag);
            }
        }

        private async void CheckForUpdatesButton_Click(object sender, RoutedEventArgs e)
        {
            if (updateAvailable)
            {
                LaunchLink("SupportLink");
                return;
            }

            try
            {
                CheckForUpdatesButtonText.Text = string.Empty;
                CheckForUpdatesButton.IsEnabled = false;
                CheckForUpdatesProgress.Visibility = Visibility.Visible;

                CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "CurrentlyCheckingForUpdates");

                var (status, message) = await SafeCheckForUpdatesAsync();

                CheckForUpdatesButton.IsEnabled = true;
                CheckForUpdatesProgress.Visibility = Visibility.Collapsed;

                switch (status)
                {
                    case UpdateStatus.UpdateAvailable:
                        updateAvailable = true;
                        CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "UpdateAvailable");
                        CheckForUpdatesButtonText.SetResourceReference(TextBlock.TextProperty, "SettingsCheckForUpdateButtonUpdate");
                        break;

                    case UpdateStatus.UpToDate:
                        updateAvailable = false;
                        CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "UpdateUpToDate");
                        CheckForUpdatesButtonText.SetResourceReference(TextBlock.TextProperty, "SettingsCheckForUpdateButtonCheck");
                        break;

                    case UpdateStatus.CheckFailed:
                    default:
                        updateAvailable = false;
                        CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "UpdateError");
                        CheckForUpdatesButtonText.SetResourceReference(TextBlock.TextProperty, "SettingsCheckForUpdateButtonCheck");
                        break;
                }
            }
            catch
            {
                CheckForUpdatesButton.IsEnabled = true;
                CheckForUpdatesProgress.Visibility = Visibility.Collapsed;

                CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "UpdateError");
                CheckForUpdatesButtonText.SetResourceReference(TextBlock.TextProperty, "SettingsCheckForUpdateButtonCheck");
            }
        }
        private void GithubHyperlink_Click(object sender, RoutedEventArgs e) => LaunchLink("GithubLink");

        private void YoutubeHyperlink_Click(object sender, RoutedEventArgs e) => LaunchLink("YoutubeLink");

        private void DiscordHyperlink_Click(object sender, RoutedEventArgs e) => LaunchLink("SupportLink");
        #endregion

        #region Generic Events Methods
        private void ComboBox_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
        {
            e.Handled = true;
        }

        private void IntegerBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox intBox) return;
            if ((string?)intBox.Tag != "DurationTextBox") return;

            if (intBox == GlobalDurationTextBox)
            {
                if (int.TryParse(GlobalDurationTextBox.Text, out int newDuration))
                {
                    if (newDuration >= 10)
                    {
                        duration = newDuration;
                        SaveSetting("Duration", newDuration.ToString());
                        GlobalDurationTextBox.Text = newDuration.ToString();
                    }
                    else
                    {
                        GlobalDurationTextBox.Text = duration.ToString();
                    }
                }
                else
                {
                    GlobalDurationTextBox.Text = duration.ToString();
                }

                return;
            }

            if (int.TryParse(intBox.Text, out int newValue))
            {
                if (newValue >= 10)
                {
                    intBox.Text = newValue.ToString();
                }
                else
                {
                    intBox.Text = "10";
                }
            }
            else
            {
                intBox.Text = "10";
            }
        }

        private void IntegerBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || e.Key != Key.Return) return;

            ((UIElement)e.OriginalSource).MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));

            e.Handled = true;
        }

        private void IntegerBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }
        #endregion

        #region Exposed Methods
        public void LaunchLink(string linkName)
        {
            if (TryFindResource(linkName) is string link && !string.IsNullOrWhiteSpace(link))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = link,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Support link not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void LoadPoints(AppFileData data)
        {
            pointManager.ClearAll();
            foreach (var point in data.Points)
            {
                pointManager.AddPoint(point.Position, point.Button, point.ClickType, point.Duration);
            }
            if (SettingsContainer.Visibility == Visibility.Visible) SettingsButton_Click(SettingsButton, null!);
        }
        #endregion

        #region General Methods
        private async Task StartClickingAsync()
        {
            lock (runLock)
            {
                if (isRunning)
                    return;

                isRunning = true;
            }

            pointManager.SetRuntimeMode(true);
            currentMode = "Running";
            UpdateModeVisuals();

            runningTask = Task.Run(async () =>
            {
                try
                {
                    bool shouldLoop = looped;
                    int loops = loopCount;
                    if (!shouldLoop)
                        loops = 1;

                    int currentLoop = 0;
                    var orderedPoints = pointManager.Points.OrderBy(p => p.Order).ToList();

                    if (orderedPoints.Count == 0)
                    {
                        while (isRunning)
                        {
                            MouseHandler.SimulateClick(currentButton, currentClickType);
                            await Task.Delay(duration);
                        }
                        return;
                    }

                    do
                    {
                        foreach (var point in orderedPoints)
                        {
                            if (!isRunning)
                                return;

                            MouseHandler.MoveMouse(point.Position);
                            MouseHandler.SimulateClick(point.Button, point.ClickType);

                            await Task.Delay(point.Duration);
                        }

                        currentLoop++;
                        if (shouldLoop && loops > 0 && currentLoop >= loops)
                            break;
                    }
                    while (isRunning && (shouldLoop && (loops == 0 || currentLoop < loops)));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Clicker failed: {ex.Message}");
                }
                finally
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        lock (runLock)
                        {
                            isRunning = false;
                            runningTask = null;
                        }

                        pointManager.SetRuntimeMode(false);
                        currentMode = "Normal";
                        UpdateModeVisuals();
                    });
                }
            });
        }

        private async Task StopClickingAsync()
        {
            lock (runLock)
            {
                if (!isRunning)
                    return;

                isRunning = false;
            }

            try
            {
                if (runningTask != null)
                    await runningTask;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"StopClickingAsync error: {ex.Message}");
            }

            runningTask = null;
        }

        private void UpdateModeVisuals()
        {
            var highlightBrush = (System.Windows.Media.Brush)FindResource("ControlFillColorSecondaryBrush");
            var transparentBrush = System.Windows.Media.Brushes.Transparent;

            AddButton.Background = currentMode == "Add" ? highlightBrush : transparentBrush;
            RemoveButton.Background = currentMode == "Remove" ? highlightBrush : transparentBrush;
            ToggleButton.Background = currentMode == "Running" ? highlightBrush : transparentBrush;
        }

        private void SaveSetting(string key, string value)
        {
            if (appSettings == null || configuration == null) return;

            try
            {
                if (appSettings[key] != null)
                    appSettings[key].Value = value;
                else
                    appSettings.Add(key, value);

                configuration.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Config Save Error] {ex.Message}");
            }
        }

        private T GetSetting<T>(string key, T defaultValue)
        {
            if (appSettings == null || configuration == null)
                return defaultValue;

            try
            {
                if (appSettings[key] != null)
                {
                    string rawValue = appSettings[key].Value ?? string.Empty;

                    // Enums
                    if (typeof(T).IsEnum && Enum.TryParse(typeof(T), rawValue, out var enumValue))
                        return (T)enumValue;

                    // Primitives
                    if (typeof(T) == typeof(bool) && bool.TryParse(rawValue, out var boolValue))
                        return (T)(object)boolValue;

                    if (typeof(T) == typeof(int) && int.TryParse(rawValue, out var intValue))
                        return (T)(object)intValue;

                    if (typeof(T) == typeof(double) && double.TryParse(rawValue, out var doubleValue))
                        return (T)(object)doubleValue;

                    if (typeof(T) == typeof(string))
                        return (T)(object)rawValue;

                    if (typeof(T) == typeof(HashSet<Key>))
                    {
                        var keys = rawValue
                            .Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                            .Select(k => Enum.TryParse<Key>(k, out var key) ? key : Key.None)
                            .Where(k => k != Key.None)
                            .ToHashSet();

                        return (T)(object)keys;
                    }

                    if (typeof(T) == typeof(List<Key>))
                    {
                        var keys = rawValue
                            .Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                            .Select(k => Enum.TryParse<Key>(k, out var key) ? key : Key.None)
                            .Where(k => k != Key.None)
                            .ToList();

                        return (T)(object)keys;
                    }

                    if (typeof(T) == typeof(Point))
                    {
                        var parts = rawValue.Split(',');
                        if (parts.Length == 2 &&
                            int.TryParse(parts[0], out int x) &&
                            int.TryParse(parts[1], out int y))
                            return (T)(object)new Point(x, y);
                    }

                    if (typeof(T) == typeof(Size))
                    {
                        var parts = rawValue.Split(',');
                        if (parts.Length == 2 &&
                            double.TryParse(parts[0], out double w) &&
                            double.TryParse(parts[1], out double h))
                            return (T)(object)new Size(w, h);
                    }
                }

                SetSetting(key, defaultValue);
                return defaultValue;
            }
            catch
            {
                SetSetting(key, defaultValue);
                return defaultValue;
            }
        }

        private void SetSetting<T>(string key, T value)
        {
            if (appSettings == null || configuration == null)
                return;

            try
            {
                string stringValue = value switch
                {
                    List<Key> keyList => string.Join("+", keyList.Select(k => k.ToString())),
                    Point p => $"{p.X},{p.Y}",
                    Size s => $"{s.Width},{s.Height}",
                    _ => value?.ToString() ?? string.Empty
                };

                if (appSettings[key] != null)
                    appSettings[key].Value = stringValue;
                else
                    appSettings.Add(key, stringValue);

                configuration.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Config Save Error] {ex.Message}");
            }
        }

        private string GetFormattedToggleKey()
        {
            return string.Join("+", toggleKey);
        }

        private static string FormatKeyName(Key key)
        {
            if (key >= Key.D0 && key <= Key.D9)
                return ((int)key - (int)Key.D0).ToString();

            if (key >= Key.NumPad0 && key <= Key.NumPad9)
                return "Num" + ((int)key - (int)Key.NumPad0);

            return key switch
            {
                Key.LeftCtrl or Key.RightCtrl => "Ctrl",
                Key.LeftShift or Key.RightShift => "Shift",
                Key.LeftAlt or Key.RightAlt => "Alt",
                Key.LWin or Key.RWin => "Win",

                >= Key.F1 and <= Key.F24 => key.ToString().ToUpper(),

                Key.Return => "Enter",
                Key.Back => "Backspace",
                Key.Delete => "Del",
                Key.Insert => "Ins",
                Key.Tab => "Tab",
                Key.Space => "Space",
                Key.Escape => "Esc",
                Key.Capital => "CapsLock",

                Key.Home => "Home",
                Key.End => "End",
                Key.PageUp => "PgUp",
                Key.PageDown => "PgDn",
                Key.Up => "↑",
                Key.Down => "↓",
                Key.Left => "←",
                Key.Right => "→",

                Key.OemPlus => "Plus",
                Key.OemMinus => "Minus",
                Key.OemComma => ",",
                Key.OemPeriod => ".",
                Key.Oem1 => ";",
                Key.Oem2 => "/",
                Key.Oem3 => "`",
                Key.Oem4 => "[",
                Key.Oem5 => "\\",
                Key.Oem6 => "]",
                Key.Oem7 => "'",

                Key.Multiply => "Multiply",
                Key.Divide => "Divide",
                Key.Subtract => "Subtract",
                Key.Add => "Add",
                Key.Decimal => "Decimal",

                _ => key.ToString()
            };
        }

        private void SetMenuDock()
        {
            if (menuStyle == "Left")
            {
                Grid.SetRow(MenuGrid, 1);
                Grid.SetColumn(MenuGrid, 0);

                ActionsPanel.Orientation = Orientation.Vertical;
                MiscPanel.Orientation = Orientation.Vertical;

                MenuGrid.VerticalAlignment = VerticalAlignment.Stretch;
                MenuGrid.HorizontalAlignment = HorizontalAlignment.Left;

                ActionsPanel.HorizontalAlignment = HorizontalAlignment.Center;
                ActionsPanel.VerticalAlignment = VerticalAlignment.Top;

                MiscPanel.HorizontalAlignment = HorizontalAlignment.Center;
                MiscPanel.VerticalAlignment = VerticalAlignment.Bottom;
            }
            else
            {
                Grid.SetRow(MenuGrid, 0);
                Grid.SetColumn(MenuGrid, 1);

                ActionsPanel.Orientation = Orientation.Horizontal;
                MiscPanel.Orientation = Orientation.Horizontal;

                MenuGrid.HorizontalAlignment = HorizontalAlignment.Stretch;
                MenuGrid.VerticalAlignment = VerticalAlignment.Top;

                ActionsPanel.HorizontalAlignment = HorizontalAlignment.Left;
                ActionsPanel.VerticalAlignment = VerticalAlignment.Center;

                MiscPanel.HorizontalAlignment = HorizontalAlignment.Right;
                MiscPanel.VerticalAlignment = VerticalAlignment.Center;
            }
        }

        private void ApplyLanguage(string language)
        {
            try
            {
                string targetLanguage = language;

                if (language.Equals("System", StringComparison.OrdinalIgnoreCase))
                {
                    string systemLang = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;

                    switch (systemLang)
                    {
                        case "ar":
                            targetLanguage = "Arabic";
                            break;
                        case "en":
                            targetLanguage = "English";
                            break;
                        default:
                            targetLanguage = "English";
                            break;
                    }
                }

                string dictPath = $"pack://application:,,,/PointAC;component/Localization/{targetLanguage}.xaml";
                var newDict = new ResourceDictionary { Source = new Uri(dictPath, UriKind.Absolute) };

                var oldDict = Application.Current.Resources.MergedDictionaries
                    .FirstOrDefault(d => d.Source != null && d.Source.OriginalString.Contains("/Localization/"));
                if (oldDict != null)
                    Application.Current.Resources.MergedDictionaries.Remove(oldDict);

                Application.Current.Resources.MergedDictionaries.Add(newDict);

                if (targetLanguage.Equals("Arabic", StringComparison.OrdinalIgnoreCase))
                {
                    this.FlowDirection = FlowDirection.RightToLeft;
                }
                else
                {
                    this.FlowDirection = FlowDirection.LeftToRight;
                }
            }
            catch
            {
                try
                {
                    var fallbackDict = new ResourceDictionary
                    {
                        Source = new Uri("pack://application:,,,/PointAC;component/Localization/English.xaml", UriKind.Absolute)
                    };

                    Application.Current.Resources.MergedDictionaries.Add(fallbackDict);
                    this.FlowDirection = FlowDirection.LeftToRight;
                }
                catch { }
            }
        }

        private void UpdateHoverState(PointEntry? nearest, double minDist, int hoverRadius)
        {
            foreach (var p in pointManager.Points)
            {
                bool shouldHover = (p == nearest && minDist <= hoverRadius);
                if (p.IsHovered != shouldHover)
                {
                    p.IsHovered = shouldHover;
                    if (shouldHover)
                    {
                        var container = PointsList.ItemContainerGenerator.ContainerFromItem(p) as FrameworkElement;
                        container?.BringIntoView();
                    }
                }
            }
        }

        private bool IsShortcutPressed(KeyEventArgs e, HashSet<Key> shortcut)
        {
            bool ctrl = shortcut.Contains(Key.LeftCtrl) || shortcut.Contains(Key.RightCtrl);
            bool shift = shortcut.Contains(Key.LeftShift) || shortcut.Contains(Key.RightShift);
            bool alt = shortcut.Contains(Key.LeftAlt) || shortcut.Contains(Key.RightAlt);

            bool ctrlPressed = (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl));
            bool shiftPressed = (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift));
            bool altPressed = (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt));

            var mainKeys = shortcut.Except(new[] {
                Key.LeftCtrl, Key.RightCtrl,
                Key.LeftShift, Key.RightShift,
                Key.LeftAlt, Key.RightAlt
            });

            return mainKeys.Contains(e.Key)
                && (!ctrl || ctrlPressed)
                && (!shift || shiftPressed)
                && (!alt || altPressed);
        }

        private async Task<(UpdateStatus, string)> SafeCheckForUpdatesAsync()
        {
            try
            {
                var currentVersion = Version.Parse(AppVersion);
                string versionUrl = "https://raw.githubusercontent.com/xh4l1l/Versions/refs/heads/main/PointAC";

                var (status, latestVersion) = await UpdateChecker.CheckForUpdatesAsync(versionUrl, currentVersion);

                string message = status switch
                {
                    UpdateStatus.UpdateAvailable => (string)FindResource("UpdateAvailable"),
                    UpdateStatus.UpToDate => (string)FindResource("UpdateUpToDate"),
                    UpdateStatus.CheckFailed => (string)FindResource("UpdateError"),
                    _ => "Unknown status"
                };

                return (status, message);
            }
            catch
            {
                return (UpdateStatus.CheckFailed, (string)FindResource("UpdateError"));
            }
        }


        private bool IsClickInsideAppWindow(System.Drawing.Point screenPoint)
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;

            if (GetWindowRect(hwnd, out RECT rect))
            {
                bool inside = screenPoint.X >= rect.Left && screenPoint.X <= rect.Right &&
                              screenPoint.Y >= rect.Top && screenPoint.Y <= rect.Bottom;
                return inside;
            }

            return false;
        }

        private System.Drawing.Rectangle GetElementScreenBounds(FrameworkElement element)
        {
            if (!element.IsLoaded)
                return System.Drawing.Rectangle.Empty;

            var transform = element.TransformToAncestor(this);
            var topLeft = transform.Transform(new System.Windows.Point(0, 0));
            var bottomRight = transform.Transform(new System.Windows.Point(element.ActualWidth, element.ActualHeight));

            var screenTopLeft = PointToScreen(topLeft);
            var screenBottomRight = PointToScreen(bottomRight);

            return System.Drawing.Rectangle.FromLTRB(
                (int)screenTopLeft.X,
                (int)screenTopLeft.Y,
                (int)screenBottomRight.X,
                (int)screenBottomRight.Y
            );
        }

        private static double Distance(Point a, Point b)
        {
            int dx = a.X - b.X;
            int dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
        #endregion

        #region Overriden Methods
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            isRunning = false;
            currentMode = "Closed";
            D2DOverlay.Instance?.Dispose();
            MouseHandler.StopListeningToAll();
            KeyboardHandler.StopListeningToAll();
            Application.Current.Shutdown();
        }
        #endregion

        #region Win32 Interop
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion

    }
}