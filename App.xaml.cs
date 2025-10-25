using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace PointAC
{
    public partial class App : Application
    {
        private MainWindow? mainWindow;
        private static Mutex? appMutex;
        private const string PipeName = "PointAutoClickerPipe";
        private const string MutexName = "PointAutoClickerMutex";


        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            Directory.SetCurrentDirectory(AppContext.BaseDirectory);

            bool isNewInstance;
            appMutex = new Mutex(true, MutexName, out isNewInstance);

            if (!isNewInstance)
            {
                if (e.Args.Length > 0)
                    await SendArgsToExistingInstanceAsync(e.Args);
                Environment.Exit(0);
                return;
            }

            RunWithArgs(e.Args);
        }

        private void RunWithArgs(string[] args)
        {
            mainWindow = new MainWindow();
            this.MainWindow = mainWindow;
            mainWindow.Show();

            if (args.Length > 0)
                _ = LoadStartupFilesAsync(args);

            StartPipeServer();
        }

        private async Task LoadStartupFilesAsync(string[] args)
        {
            foreach (var arg in args)
            {
                string resolved = NormalizePath(arg);

                if (File.Exists(resolved))
                {
                    try
                    {
                        var data = await AppFileOperations.LoadFromFileAsync(resolved);
                        mainWindow?.LoadPoints(data);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to load file:\n\n{ex.Message}",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"File not found:\n{resolved}",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private async Task SendArgsToExistingInstanceAsync(string[] args)
        {
            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    await client.ConnectAsync(1000);
                    using (var writer = new StreamWriter(client) { AutoFlush = true })
                    {
                        foreach (var arg in args)
                            await writer.WriteLineAsync(arg);
                    }
                }
            }
            catch { }
        }

        private void StartPipeServer()
        {
            Thread pipeThread = new(() =>
            {
                while (true)
                {
                    try
                    {
                        using var server = new NamedPipeServerStream(PipeName, PipeDirection.In);
                        server.WaitForConnection();

                        using var reader = new StreamReader(server);
                        string? line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            string resolved = NormalizePath(line);

                            Application.Current.Dispatcher.Invoke(async () =>
                            {
                                if (File.Exists(resolved))
                                {
                                    try
                                    {
                                        var data = await AppFileOperations.LoadFromFileAsync(resolved);
                                        mainWindow?.LoadPoints(data);
                                        mainWindow?.Activate();
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"Failed to load file:\n\n{ex.Message}",
                                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                            });
                        }
                    }
                    catch { }
                }
            })
            {
                IsBackground = true,
                Name = "PointAutoClickerPipeThread"
            };

            pipeThread.Start();
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            path = path.Trim().Trim('"').Trim();

            if (!Path.IsPathRooted(path))
                path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, path));

            return path;
        }
    }
}