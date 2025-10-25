using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using SharpDX;
using SharpDX.Direct2D1;
using SharpDX.DXGI;
using SharpDX.Mathematics.Interop;
using SharpDX.WIC;
using Factory = SharpDX.Direct2D1.Factory;
using PixelFormat = SharpDX.Direct2D1.PixelFormat;

namespace PointAC
{
    public class D2DOverlay : IDisposable
    {
        private static D2DOverlay? _instance;
        private static IntPtr _winEventHook;
        private static WinEventDelegate? _winEventProc;
        private readonly object _lock = new object();
        public static D2DOverlay Instance => _instance ??= new D2DOverlay();

        private delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        private static readonly WndProc _wndProc = DefWindowProc;

        private IntPtr _hwnd;
        private Factory? _factory;
        private WindowRenderTarget? _target;
        private ImagingFactory? _wicFactory;
        private readonly Dictionary<Guid, (SharpDX.Direct2D1.Bitmap Image, RawVector2 Pos, RawVector2 Size)> _images = new();

        private bool _running;
        private Thread? _renderThread;

        public event Action? RendererRecreated;

        private D2DOverlay()
        {
            SetProcessDPIAware();

            InitializeOverlayWindow();
            InitializeD2D();
            StartRenderLoop();


            AppDomain.CurrentDomain.ProcessExit += (_, _) => Dispose();
        }

        private void InitializeOverlayWindow()
        {
            const string CLASS_NAME = "D2DOverlayWindow";
            var wndClass = new WNDCLASSEX
            {
                cbSize = Marshal.SizeOf<WNDCLASSEX>(),
                lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
                hInstance = Process.GetCurrentProcess().Handle,
                lpszClassName = CLASS_NAME,
                hbrBackground = IntPtr.Zero
            };

            RegisterClassEx(ref wndClass);

            _hwnd = CreateWindowEx(
                WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
                CLASS_NAME,
                "Direct2D Overlay",
                WS_POPUP,
                0, 0,
                GetSystemMetrics(0), GetSystemMetrics(1),
                IntPtr.Zero,
                IntPtr.Zero,
                Process.GetCurrentProcess().Handle,
                IntPtr.Zero
            );

            int exStyle = GetWindowLong(_hwnd, GWL_EXSTYLE);
            exStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT | WS_EX_LAYERED;
            SetWindowLong(_hwnd, GWL_EXSTYLE, exStyle);

            if (IntPtr.Size == 8)
                SetWindowLongPtr64(_hwnd, GWL_HWNDPARENT, GetDesktopWindow());
            else
                SetWindowLong(_hwnd, GWL_HWNDPARENT, GetDesktopWindow().ToInt32());

            const uint LWA_ALPHA = 0x2;
            const uint LWA_COLORKEY = 0x1;
            SetLayeredWindowAttributes(_hwnd, 0, 255, LWA_ALPHA | LWA_COLORKEY);

            ShowWindow(_hwnd, 4);

            SetWindowPos(_hwnd, new IntPtr(-2), 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            SetWindowPos(_hwnd, new IntPtr(-1), 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            _winEventProc = (hWinEventHook, eventType, hwnd, idObject, idChild, thread, time) =>
            {
                if (_hwnd != IntPtr.Zero)
                {
                    SetWindowPos(_hwnd, new IntPtr(-1), 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                }
            };

            _winEventHook = SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_MINIMIZEEND,
            IntPtr.Zero, _winEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);

            AppDomain.CurrentDomain.ProcessExit += (_, _) =>
            {
                if (_winEventHook != IntPtr.Zero)
                {
                    UnhookWinEvent(_winEventHook);
                    _winEventHook = IntPtr.Zero;
                }
                try { DestroyWindow(_hwnd); } catch { }
            };
        }

        private void CleanupD2D()
        {
            try
            {
                _running = false;

                lock (_lock)
                {
                    foreach (var (_, data) in _images)
                    {
                        try
                        {
                            data.Image?.Dispose();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[D2DOverlay] Failed to dispose image: {ex.Message}");
                        }
                    }

                    _images.Clear();

                    try
                    {
                        _target?.Dispose();
                        _target = null;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[D2DOverlay] Failed to dispose target: {ex.Message}");
                    }

                    try
                    {
                        _factory?.Dispose();
                        _factory = null;

                        _wicFactory?.Dispose();
                        _wicFactory = null;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[D2DOverlay] Factory cleanup error: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine("[D2DOverlay] CleanupD2D completed.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[D2DOverlay] CleanupD2D failed: {ex.Message}");
            }
        }

        private void InitializeD2D()
        {
            _factory = new Factory();
            _wicFactory = new ImagingFactory();

            RenderTargetProperties renderProps = new RenderTargetProperties(
                SharpDX.Direct2D1.RenderTargetType.Default,
                new PixelFormat(Format.B8G8R8A8_UNorm, SharpDX.Direct2D1.AlphaMode.Premultiplied),
                96.0f,
                96.0f,
                SharpDX.Direct2D1.RenderTargetUsage.None,
                SharpDX.Direct2D1.FeatureLevel.Level_DEFAULT
            );

            var hwndProps = new HwndRenderTargetProperties
            {
                Hwnd = _hwnd,
                PixelSize = new Size2(GetSystemMetrics(0), GetSystemMetrics(1)),
                PresentOptions = SharpDX.Direct2D1.PresentOptions.None
            };

            _target = new WindowRenderTarget(_factory, renderProps, hwndProps)
            {
                AntialiasMode = AntialiasMode.PerPrimitive
            };
        }

        private void StartRenderLoop()
        {
            _running = true;
            _renderThread = new Thread(() =>
            {
                while (_running)
                {
                    try
                    {
                        if (_target == null)
                            continue;

                        lock (_lock)
                        {
                            _target.BeginDraw();
                            _target.Clear(new SharpDX.Mathematics.Interop.RawColor4(0, 0, 0, 0));

                            foreach (var (_, data) in _images)
                            {
                                var rect = new SharpDX.Mathematics.Interop.RawRectangleF(
                                    data.Pos.X,
                                    data.Pos.Y,
                                    data.Pos.X + data.Size.X,
                                    data.Pos.Y + data.Size.Y
                                );

                                try
                                {
                                    _target.DrawBitmap(data.Image, rect, 1.0f, SharpDX.Direct2D1.BitmapInterpolationMode.Linear);
                                }
                                catch
                                {

                                }
                            }

                            _target.EndDraw();
                        }
                    }
                    catch (SharpDX.SharpDXException ex) when (ex.ResultCode == SharpDX.Direct2D1.ResultCode.RecreateTarget)
                    {
                        System.Diagnostics.Debug.WriteLine("[D2DOverlay] RecreateTarget triggered — cleaning up and reinitializing D2D...");
                        CleanupD2D();
                        Thread.Sleep(250);
                        InitializeD2D();

                        RendererRecreated?.Invoke();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[D2DOverlay] Render loop error: {ex.Message}");
                    }

                    Thread.Sleep(32); // ~30 FPS
                }
            })
            {
                IsBackground = true
            };
            _renderThread.Start();
        }

        public Guid Add(string imagePath, System.Drawing.Point pos, System.Drawing.Size size)
        {
            using var decoder = new SharpDX.WIC.BitmapDecoder(
                _wicFactory, imagePath, DecodeOptions.CacheOnDemand
            );
            using var frame = decoder.GetFrame(0);
            using var converter = new FormatConverter(_wicFactory);
            converter.Initialize(frame, SharpDX.WIC.PixelFormat.Format32bppPBGRA);

            var bmp = SharpDX.Direct2D1.Bitmap.FromWicBitmap(_target, converter);
            var id = Guid.NewGuid();
            _images[id] = (bmp, new RawVector2(pos.X, pos.Y), new RawVector2(size.Width, size.Height));
            return id;
        }

        public void Remove(Guid handle)
        {
            if (_images.TryGetValue(handle, out var img))
            {
                img.Image.Dispose();
                _images.Remove(handle);
            }
        }

        public void Clear()
        {
            foreach (var item in _images.Values)
                item.Image.Dispose();
            _images.Clear();
        }

        public void Dispose()
        {
            _running = false;
            _renderThread?.Join();

            Clear();

            _target?.Dispose();
            _wicFactory?.Dispose();
            _factory?.Dispose();

            if (_hwnd != IntPtr.Zero)
                DestroyWindow(_hwnd);
        }

        #region Win32 Interop
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint EVENT_SYSTEM_MINIMIZESTART = 0x0016;
        private const uint EVENT_SYSTEM_MINIMIZEEND = 0x0017;

        private const int WS_POPUP = unchecked((int)0x80000000);
        private const int WS_EX_LAYERED = 0x80000;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const int WS_EX_TOPMOST = 0x00000008;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int LWA_ALPHA = 0x2;

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOREDRAW = 0x0008;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint SWP_HIDEWINDOW = 0x0080;
        private const uint SWP_NOCOPYBITS = 0x0100;
        private const uint SWP_NOOWNERZORDER = 0x0200;
        private const uint SWP_NOSENDCHANGING = 0x0400;

        private const int GWL_EXSTYLE = -20;
        private const int GWL_HWNDPARENT = -8;

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int index);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern ushort RegisterClassEx([In] ref WNDCLASSEX lpwcx);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CreateWindowEx(
            int exStyle, string lpClassName, string lpWindowName, int dwStyle,
            int x, int y, int width, int height,
            IntPtr hwndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam
        );

        [DllImport("user32.dll")]
        private static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct WNDCLASSEX
        {
            public int cbSize;
            public int style;
            public IntPtr lpfnWndProc;
            public int cbClsExtra;
            public int cbWndExtra;
            public IntPtr hInstance;
            public IntPtr hIcon;
            public IntPtr hCursor;
            public IntPtr hbrBackground;
            public string lpszMenuName;
            public string lpszClassName;
            public IntPtr hIconSm;
        }
        #endregion
    }
}