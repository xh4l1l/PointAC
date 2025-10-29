using SharpDX;
using SharpDX.Direct2D1;
using SharpDX.DXGI;
using SharpDX.Mathematics.Interop;
using SharpDX.WIC;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Resources;
using AlphaMode = SharpDX.Direct2D1.AlphaMode;
using BitmapInterpolationMode = SharpDX.Direct2D1.BitmapInterpolationMode;
using Factory = SharpDX.Direct2D1.Factory;
using PixelFormat = SharpDX.Direct2D1.PixelFormat;
using ResultCode = SharpDX.Direct2D1.ResultCode;

namespace PointAC.Core
{
    public class PointsOverlay : IDisposable
    {
        private static PointsOverlay? _instance;
        private static IntPtr _winEventHook;
        private static WinEventDelegate? _winEventProc;
        private readonly object _lock = new();

        public static PointsOverlay Instance => _instance ??= new PointsOverlay();

        private delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        private static readonly WndProc _wndProc = DefWindowProc;

        private IntPtr _hwnd;
        private Factory? _factory;
        private WindowRenderTarget? _target;
        private ImagingFactory? _wicFactory;
        private bool _disposed;

        private readonly Dictionary<Guid, (SharpDX.Direct2D1.Bitmap Image, RawVector2 Pos, RawVector2 Size)> _images = new();

        public event Action? RendererRecreated;

        private PointsOverlay()
        {
            SetProcessDPIAware();
            InitializeOverlayWindow();
            InitializeD2D();

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
                "Points Overlay",
                WS_POPUP,
                0, 0,
                GetSystemMetrics(0), GetSystemMetrics(1),
                GetDesktopWindow(),
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

            SetWindowPos(_hwnd, new IntPtr(-1), 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            _winEventProc = (_, _, _, _, _, _, _) =>
            {
                if (_disposed || _hwnd == IntPtr.Zero) return;
                SetWindowPos(_hwnd, new IntPtr(-1), 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            };

            _winEventHook = SetWinEventHook(
                EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_MINIMIZEEND,
                IntPtr.Zero, _winEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
        }

        private void InitializeD2D()
        {
            _factory = new Factory();
            _wicFactory = new ImagingFactory();

            RenderTargetProperties renderProps = new(
                RenderTargetType.Default,
                new PixelFormat(Format.B8G8R8A8_UNorm, AlphaMode.Premultiplied),
                96.0f, 96.0f,
                RenderTargetUsage.None,
                FeatureLevel.Level_DEFAULT
            );

            var hwndProps = new HwndRenderTargetProperties
            {
                Hwnd = _hwnd,
                PixelSize = new Size2(GetSystemMetrics(0), GetSystemMetrics(1)),
                PresentOptions = PresentOptions.None
            };

            _target = new WindowRenderTarget(_factory, renderProps, hwndProps)
            {
                AntialiasMode = AntialiasMode.PerPrimitive
            };
        }

        private void CleanupD2D()
        {
            lock (_lock)
            {
                foreach (var (_, data) in _images)
                    try { data.Image.Dispose(); } catch { }

                _images.Clear();

                _target?.Dispose();
                _wicFactory?.Dispose();
                _factory?.Dispose();

                _target = null;
                _wicFactory = null;
                _factory = null;
            }
        }

        public void Render()
        {
            if (_disposed || _target == null)
                return;

            lock (_lock)
            {
                try
                {
                    _target.BeginDraw();
                    _target.Clear(new RawColor4(0, 0, 0, 0));

                    foreach (var (_, data) in _images)
                    {
                        if (data.Image == null || data.Image.IsDisposed) continue;

                        var rect = new RawRectangleF(
                            data.Pos.X,
                            data.Pos.Y,
                            data.Pos.X + data.Size.X,
                            data.Pos.Y + data.Size.Y);

                        _target.DrawBitmap(data.Image, rect, 1.0f, BitmapInterpolationMode.Linear);
                    }

                    _target.EndDraw();
                }
                catch (SharpDXException ex) when (ex.ResultCode == ResultCode.RecreateTarget)
                {
                    Debug.WriteLine("[D2DOverlay] RecreateTarget triggered — reinitializing D2D...");
                    CleanupD2D();
                    Thread.Sleep(250);
                    InitializeD2D();
                    RendererRecreated?.Invoke();
                    Render();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[D2DOverlay] Render error: {ex.Message}");
                }
            }
        }

        public Guid Add(string imagePath, System.Drawing.Point pos, System.Drawing.Size size)
        {
            if (_wicFactory == null || _target == null)
                throw new InvalidOperationException("Renderer not initialized.");

            BitmapDecoder decoder;

            if (imagePath.StartsWith("pack://", StringComparison.OrdinalIgnoreCase))
            {
                Uri uri = new Uri(imagePath, UriKind.Absolute);
                StreamResourceInfo sri = System.Windows.Application.GetResourceStream(uri)
                    ?? throw new FileNotFoundException("Resource not found: " + imagePath);
                decoder = new BitmapDecoder(_wicFactory, sri.Stream, DecodeOptions.CacheOnDemand);
            }
            else
            {
                decoder = new BitmapDecoder(_wicFactory, imagePath, DecodeOptions.CacheOnDemand);
            }

            using (decoder)
            using (var frame = decoder.GetFrame(0))
            using (var converter = new FormatConverter(_wicFactory))
            {
                converter.Initialize(frame, SharpDX.WIC.PixelFormat.Format32bppPBGRA);

                var bmp = SharpDX.Direct2D1.Bitmap.FromWicBitmap(_target, converter);
                var id = Guid.NewGuid();

                lock (_lock)
                {
                    _images[id] = (bmp, new RawVector2(pos.X, pos.Y), new RawVector2(size.Width, size.Height));
                }

                Render();
                return id;
            }
        }

        public void Remove(Guid handle)
        {
            lock (_lock)
            {
                if (_images.TryGetValue(handle, out var img))
                {
                    try { img.Image.Dispose(); } catch { }
                    _images.Remove(handle);
                }
            }
            Render();
        }

        public void Clear()
        {
            lock (_lock)
            {
                foreach (var item in _images.Values)
                    try { item.Image.Dispose(); } catch { }

                _images.Clear();
            }
            Render();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_winEventHook != IntPtr.Zero)
                {
                    UnhookWinEvent(_winEventHook);
                    _winEventHook = IntPtr.Zero;
                }
            }
            catch { }

            try { Clear(); } catch { }
            try { CleanupD2D(); } catch { }

            if (_hwnd != IntPtr.Zero)
            {
                try { DestroyWindow(_hwnd); } catch { }
                _hwnd = IntPtr.Zero;
            }
        }

        #region Win32 Interop
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint EVENT_SYSTEM_MINIMIZEEND = 0x0017;

        private const int WS_POPUP = unchecked((int)0x80000000);
        private const int WS_EX_LAYERED = 0x80000;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const int WS_EX_TOPMOST = 0x00000008;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int LWA_ALPHA = 0x2;
        private const int LWA_COLORKEY = 0x1;

        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

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