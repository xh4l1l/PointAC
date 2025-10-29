using System.IO;
using System.Data;
using PointAC.Core;
using System.Windows;
using Microsoft.Win32;
using PointAC.Services;
using PointAC.Management;
using System.Diagnostics;
using System.Configuration;
using System.Windows.Input;
using System.Windows.Media;
using PointAC.Miscellaneous;
using System.Windows.Controls;
using Point = System.Drawing.Point;
using static PointAC.Management.IOManager;
using WindowState = System.Windows.WindowState;
using Orientation = System.Windows.Controls.Orientation;
using Configuration = System.Configuration.Configuration;
using ClickType = PointAC.Services.MouseService.ClickType;
using MouseButton = PointAC.Services.MouseService.MouseButton;
using UpdateStatus = PointAC.Services.UpdateService.UpdateStatus;

namespace PointAC
{
    public partial class MainWindow : Window
    {
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
        private string backdrop = "Mica";
        private string language = "System";
        private string clickType = "Single";
        private string targetImage = "System"; // Partially Implemented, UI Required.
        private string backgroundImage = "None"; // Will be implemented later.
        private HashSet<Key> toggleShortcut = new() { Key.F6 };
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

            configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            appSettings = configuration.AppSettings.Settings;

            theme = GetSetting("Theme", theme);
            button = GetSetting("Button", button);
            looped = GetSetting("Looped", looped);
            backdrop = GetSetting("Backdrop", backdrop);
            language = GetSetting("Language", language);
            duration = GetSetting("Duration", duration);
            menuStyle = GetSetting("MenuStyle", menuStyle);
            loopCount = GetSetting("LoopCount", loopCount);
            clickType = GetSetting("ClickType", clickType);
            alwaysOnTop = GetSetting("AlwaysOnTop", alwaysOnTop);
            windowState = GetSetting("WindowState", WindowState.Normal);
            var size = GetSetting("WindowSize", new Size(Width, Height));
            toggleShortcut = GetSetting("ToggleShortcut", toggleShortcut);
            var location = GetSetting("WindowLocation", new Point((int)Left, (int)Top));

            windowTop = location.X;
            windowLeft = location.Y;
            windowWidth = size.Width;
            windowHeight = size.Height;

            Top = windowTop;
            Left = windowLeft;
            Width = windowWidth;
            Height = windowHeight;
            WindowState = windowState;
            UIManager.ApplyLanguage(this, language);
            UIManager.ApplyTheme(this, theme, backdrop);
            currentButton = MouseService.GetMouseButtonFromString(button);
            currentClickType = MouseService.GetClickTypeFromString(clickType);

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

            MouseService.ListenToMouseClick(screenPoint =>
            {
                if (RandomUtilities.IsClickInsideAppWindow(this, screenPoint))
                {
                    var addButtonBounds = RandomUtilities.GetElementScreenBounds(this, AddButton);
                    if (addButtonBounds.Contains(screenPoint))
                        return false;

                    return false;
                }

                if (currentMode == "Add")
                {
                    pointManager.AddPoint(targetImage, screenPoint, currentButton, currentClickType, duration);
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

            MouseService.ListenToMouseMove(screenPoint =>
            {
                if (pointManager.IsRuntimeMode || currentMode == "Add")
                    return;

                if (Math.Abs(screenPoint.X - lastMousePosition.X) < 2 &&
                    Math.Abs(screenPoint.Y - lastMousePosition.Y) < 2)
                    return;

                lastMousePosition = screenPoint;

                var nearest = pointManager.Points
                    .OrderBy(p => RandomUtilities.Distance(p.Position, screenPoint))
                    .FirstOrDefault();

                double minDist = nearest == null ? double.MaxValue : RandomUtilities.Distance(nearest.Position, screenPoint);
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

            KeyboardService.ListenToKeyDown(key =>
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
                        .Select(k => RandomUtilities.FormatKeyName(k))
                        .ToList();

                    Dispatcher.Invoke(() =>
                    {
                        CurrentHotkeyButtons.Text = RandomUtilities.GetFormattedShortcut(currentlyPressedKeys);
                    });

                    toggleShortcut = [.. modifiers.Concat(normalKeys)];
                    SaveSetting("ToggleShortcut", string.Join("+", toggleShortcut.Select(k => k.ToString())));
                    return;
                }

                if (RandomUtilities.IsShortcutPressed(new KeyEventArgs(
                    Keyboard.PrimaryDevice,
                    PresentationSource.FromVisual(Application.Current.MainWindow),
                    0, key), toggleShortcut))
                {
                    ToggleButton_Click(null!, null!);
                }
            });

            KeyboardService.ListenToKeyUp(key =>
            {
                currentlyPressedKeys.Remove(key);
            });
        }

        #region Event Methods
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoopCountTextBox.Text = loopCount.ToString();
            GlobalDurationTextBox.Text = duration.ToString();
            CurrentHotkeyButtons.Text = RandomUtilities.GetFormattedShortcut(toggleShortcut);

            LoopButton_Click(null!, null!);
            CheckForUpdatesButton_Click(null!, null!);
            SelectComboItemByTag(AppThemeSelector, theme);
            SelectComboItemByTag(MouseButtonSelector, button);
            SelectComboItemByTag(ClickTypeSelector, clickType);
            SelectComboItemByTag(AppBackdropSelector, backdrop);
            SelectComboItemByTag(AppLanguageSelector, language);
            SelectComboItemByTag(MenuStyleSelector, menuStyle, SetMenuDock);
            SelectComboItemByTag(AlwaysOnTopSelector, alwaysOnTop.ToString().ToLower(), () => Topmost = alwaysOnTop);
        }

        private void Window_Closed(object sender, EventArgs e) => pointManager.ClearAll();

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && Path.GetExtension(files[0]).Equals(FileType, StringComparison.OrdinalIgnoreCase))
                    e.Effects = DragDropEffects.Copy;
            }
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            currentMode = "Normal";
            UpdateModeVisuals();

            try
            {
                if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 0) return;

                string filePath = files[0];
                if (!Path.GetExtension(filePath).Equals(FileType, StringComparison.OrdinalIgnoreCase)) return;

                var data = await LoadFromFileAsync(filePath);
                LoadPoints(data);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Window_Drop] {ex.Message}");
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized) return;
            windowState = WindowState;
            SaveSetting("WindowState", windowState.ToString());
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState != WindowState.Normal) return;
            SaveSetting("WindowSize", $"{Math.Round(windowWidth)},{Math.Round(windowHeight)}");
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            if (WindowState != WindowState.Normal) return;

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

            LoopButton.Background = looped
                ? (Brush)FindResource("ControlFillColorSecondaryBrush")
                : Brushes.Transparent;
        }

        private void AddButton_Click(object sender, RoutedEventArgs e) => ToggleMode("Add");
        private void RemoveButton_Click(object sender, RoutedEventArgs e) => ToggleMode("Remove");
        private void ClearButton_Click(object sender, RoutedEventArgs e) { pointManager.ClearAll(); ToggleMode("Normal"); }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (pointManager.Points.Count <= 0) return;
            ToggleMode("Normal");

            try
            {
                var dlg = new SaveFileDialog
                {
                    DefaultExt = FileType,
                    Title = "Save Points File",
                    Filter = $"Point Auto Clicker File (*{FileType})|*{FileType}",
                };

                if (dlg.ShowDialog() == true)
                    await IOManager.SaveToFileAsync(dlg.FileName, new AppFileData { Points = pointManager.Points.ToList() });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SaveButton] {ex.Message}");
            }
        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleMode("Normal");

            try
            {
                var dlg = new OpenFileDialog
                {
                    DefaultExt = FileType,
                    Title = "Load Points File",
                    Filter = $"Point Auto Clicker File (*{FileType})|*{FileType}",
                };

                if (dlg.ShowDialog() == true)
                    LoadPoints(await LoadFromFileAsync(dlg.FileName));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadButton] {ex.Message}");
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleMode("Normal");
            bool open = SettingsContainer.Visibility != Visibility.Visible;

            PointsContainer.Visibility = open ? Visibility.Collapsed : Visibility.Visible;
            SettingsContainer.Visibility = open ? Visibility.Visible : Visibility.Collapsed;

            SettingsButton.Background = open
                ? (Brush)FindResource("ControlFillColorSecondaryBrush")
                : Brushes.Transparent;
        }

        private void LoopCountTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(LoopCountTextBox.Text, out int newCount) && newCount >= 0)
            {
                loopCount = newCount;
                SaveSetting("LoopCount", newCount.ToString());
            }

            LoopCountTextBox.Text = loopCount.ToString();
        }

        private void Selector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox || comboBox.SelectedItem is not ComboBoxItem item)
                return;

            string tag = item.Tag?.ToString() ?? "System";

            switch (comboBox.Name)
            {
                case nameof(ClickTypeSelector):
                    clickType = tag;
                    currentClickType = MouseService.GetClickTypeFromString(tag);
                    SaveSetting("ClickType", tag);
                    break;

                case nameof(MouseButtonSelector):
                    button = tag;
                    currentButton = MouseService.GetMouseButtonFromString(tag);
                    SaveSetting("Button", tag);
                    break;

                case nameof(AppThemeSelector):
                    theme = tag;
                    SaveSetting("Theme", tag);
                    UIManager.ApplyTheme(this, theme, backdrop);
                    break;

                case nameof(AppBackdropSelector):
                    backdrop = tag;
                    SaveSetting("Backdrop", tag);
                    UIManager.ApplyTheme(this, theme, backdrop);
                    break;

                case nameof(AppLanguageSelector):
                    language = tag;
                    SaveSetting("Language", tag);
                    UIManager.ApplyLanguage(this, language);
                    break;

                case nameof(AlwaysOnTopSelector):
                    alwaysOnTop = Convert.ToBoolean(tag);
                    Topmost = alwaysOnTop;
                    SaveSetting("AlwaysOnTop", alwaysOnTop.ToString());
                    break;

                case nameof(MenuStyleSelector):
                    menuStyle = tag;
                    SetMenuDock();
                    SaveSetting("MenuStyle", menuStyle);
                    break;
            }

            if (SettingsScrollViewer != null) RandomUtilities.ScrollIntoViewIfNotVisible(comboBox, SettingsScrollViewer);
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

        private void CheckForUpdatesButton_Click(object sender, RoutedEventArgs e) => _ = CheckForUpdatesAsync();

        private async Task CheckForUpdatesAsync()
        {
            if (updateAvailable)
            {
                LaunchLink("SupportLink");
                return;
            }

            try
            {
                CheckForUpdatesButtonText.Text = "";
                CheckForUpdatesButton.IsEnabled = false;
                CheckForUpdatesProgress.Visibility = Visibility.Visible;
                CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, "CurrentlyCheckingForUpdates");

                Random random = new();
                await Task.Delay(random.Next(1000, 3000)); // Intentional delay to let the user see the progress bar...

                var (status, _) = await UpdateService.CheckForUpdatesAsync(
                    (string)FindResource("VersionLink"),
                    Version.Parse((string)FindResource("AppVersion")));

                ReflectUpdateStatus(status);
            }
            catch
            {
                ReflectUpdateStatus(UpdateStatus.CheckFailed);
            }
        }

        private void GithubHyperlink_Click(object s, RoutedEventArgs e) => LaunchLink("GithubLink");
        private void YoutubeHyperlink_Click(object s, RoutedEventArgs e) => LaunchLink("YoutubeLink");
        private void DiscordHyperlink_Click(object s, RoutedEventArgs e) => LaunchLink("SupportLink");
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
            else { }
        }

        private void ToggleMode(string mode)
        {
            currentMode = (currentMode != mode) ? mode : "Normal";
            UpdateModeVisuals();
        }

        private void SelectComboItemByTag(ComboBox combo, string tag, Action? after = null)
        {
            try
            {
                combo.SelectionChanged -= Selector_SelectionChanged;

                var item = combo.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(i => i.Tag?.ToString()?.Equals(tag, StringComparison.OrdinalIgnoreCase) == true);

                if (item != null)
                {
                    combo.SelectedItem = item;
                    after?.Invoke();
                }
            }
            finally
            {
                combo.SelectionChanged += Selector_SelectionChanged;
            }
        }

        private void ReflectUpdateStatus(UpdateStatus status)
        {
            CheckForUpdatesButton.IsEnabled = true;
            CheckForUpdatesProgress.Visibility = Visibility.Collapsed;

            string descKey, btnKey;
            updateAvailable = status == UpdateStatus.UpdateAvailable;

            (descKey, btnKey) = status switch
            {
                UpdateStatus.UpdateAvailable => ("UpdateAvailable", "SettingsCheckForUpdateButtonUpdate"),
                UpdateStatus.UpToDate => ("UpdateUpToDate", "SettingsCheckForUpdateButtonCheck"),
                _ => ("UpdateError", "SettingsCheckForUpdateButtonCheck")
            };

            CheckForUpdatesDescription.SetResourceReference(TextBlock.TextProperty, descKey);
            CheckForUpdatesButtonText.SetResourceReference(TextBlock.TextProperty, btnKey);
        }

        public void LoadPoints(AppFileData data)
        {
            pointManager.ClearAll();
            foreach (var point in data.Points)
            {
                pointManager.AddPoint(targetImage, point.Position, point.Button, point.ClickType, point.Duration);
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
                            MouseService.SimulateClick(currentButton, currentClickType);
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

                            MouseService.MoveMouse(point.Position);
                            MouseService.SimulateClick(point.Button, point.ClickType);

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
            catch { }

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

        private void SetMenuDock()
        {
            if (menuStyle == "Left")
            {
                Grid.SetRow(MenuGrid, 1);
                Grid.SetColumn(MenuGrid, 0);

                Thickness margin = new Thickness
                {
                    Top = 2,
                    Bottom = 2,
                    Left = 4,
                    Right = 16
                };

                PointsList.Margin = margin;
                SettingsList.Margin = margin;

                MiscPanel.Orientation = Orientation.Vertical;
                ActionsPanel.Orientation = Orientation.Vertical;
                MiscPanel.VerticalAlignment = VerticalAlignment.Bottom;
                MenuGrid.VerticalAlignment = VerticalAlignment.Stretch;
                ActionsPanel.VerticalAlignment = VerticalAlignment.Top;
                MenuGrid.HorizontalAlignment = HorizontalAlignment.Left;
                MiscPanel.HorizontalAlignment = HorizontalAlignment.Center;
                ActionsPanel.HorizontalAlignment = HorizontalAlignment.Center;
            }
            else
            {
                Grid.SetRow(MenuGrid, 0);
                Grid.SetColumn(MenuGrid, 1);

                Thickness margin = new Thickness
                {
                    Top = 4,
                    Bottom = 4,
                    Left = 16,
                    Right = 16
                };

                PointsList.Margin = margin;
                SettingsList.Margin = margin;

                MiscPanel.Orientation = Orientation.Horizontal;
                ActionsPanel.Orientation = Orientation.Horizontal;
                MenuGrid.VerticalAlignment = VerticalAlignment.Top;
                MiscPanel.VerticalAlignment = VerticalAlignment.Center;
                MiscPanel.HorizontalAlignment = HorizontalAlignment.Right;
                ActionsPanel.VerticalAlignment = VerticalAlignment.Center;
                MenuGrid.HorizontalAlignment = HorizontalAlignment.Stretch;
                ActionsPanel.HorizontalAlignment = HorizontalAlignment.Left;
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
        #endregion

        #region Overridden Methods
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            isRunning = false;
            currentMode = "Closed";
            PointsOverlay.Instance?.Dispose();
            MouseService.StopListeningToAll();
            KeyboardService.StopListeningToAll();
            Application.Current.Shutdown();
        }
        #endregion

        #region Save and Load Methods
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

                    if (typeof(T).IsEnum && Enum.TryParse(typeof(T), rawValue, out var enumValue))
                        return (T)enumValue;

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
            catch { }
        }
        #endregion
    }
}