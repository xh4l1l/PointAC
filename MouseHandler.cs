using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Point = System.Drawing.Point;

namespace PointAC
{
    public static class MouseHandler
    {
        private static Action<Point>? _onMove;
        private static Func<Point, bool>? _onClick;
        private static IntPtr _hookId = IntPtr.Zero;
        private static LowLevelMouseProc _proc = HookCallback;

        private const int WH_MOUSE_LL = 14;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_LBUTTONDOWN = 0x0201;

        #region Hook Listening
        public static void ListenToMouseClick(Func<Point, bool> onClick)
        {
            _onClick = onClick;
            _hookId = SetHook(_proc);
        }

        public static void StopListeningToMouseClick()
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
            _onClick = null;
        }

        public static void ListenToMouseMove(Action<Point> onMove)
        {
            _onMove = onMove;
            if (_hookId == IntPtr.Zero)
                _hookId = SetHook(_proc);
        }

        public static void StopListeningToMouseMove()
        {
            _onMove = null;

            if (_onClick == null && _hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        public static void StopListeningToAll()
        {
            _onClick = null;
            _onMove = null;

            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }
        #endregion

        #region Hook Callback
        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                var pt = new Point(hookStruct.pt.x, hookStruct.pt.y);

                if (wParam == (IntPtr)WM_LBUTTONDOWN && _onClick != null)
                {
                    bool swallow = _onClick(pt);
                    if (swallow)
                        return (IntPtr)1;
                }

                if (wParam == (IntPtr)WM_MOUSEMOVE && _onMove != null)
                {
                    _onMove(pt);
                }
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }
        #endregion

        #region Controlling Mouse Methods

        /// <summary>Moves the mouse cursor to the given screen position.</summary>
        public static void MoveMouse(Point position)
        {
            SetCursorPos(position.X, position.Y);
        }

        /// <summary>Simulates a left mouse click at the current cursor position.</summary>
        public static void SimulateLeftClick()
        {
            mouse_event(MouseEventFlags.LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseEventFlags.LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        /// <summary>Simulates a right mouse click at the current cursor position.</summary>
        public static void SimulateRightClick()
        {
            mouse_event(MouseEventFlags.RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseEventFlags.RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }

        /// <summary>Simulates a middle (scroll wheel) click at the current cursor position.</summary>
        public static void SimulateMiddleClick()
        {
            mouse_event(MouseEventFlags.MIDDLEDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseEventFlags.MIDDLEUP, 0, 0, 0, UIntPtr.Zero);
        }

        /// <summary>Simulates a click using the user's primary mouse button (respects left/right-handed swap).</summary>
        public static void SimulatePrimaryClick()
        {
            var primary = GetPrimaryMouseButton();
            switch (primary)
            {
                case MouseButton.Left:
                    SimulateLeftClick();
                    break;
                case MouseButton.Right:
                    SimulateRightClick();
                    break;
                default:
                    SimulateLeftClick();
                    break;
            }
        }

        /// <summary>
        /// Simulates a single click on the specified button.
        /// </summary>
        public static void SimulateSingleClick(MouseButton button)
        {
            switch (button)
            {
                case MouseButton.Left:
                    SimulateLeftClick();
                    break;
                case MouseButton.Right:
                    SimulateRightClick();
                    break;
                case MouseButton.Middle:
                    SimulateMiddleClick();
                    break;
                case MouseButton.Default:
                    SimulatePrimaryClick();
                    break;
            }
        }

        /// <summary>
        /// Simulates a double click using the system double-click interval.
        /// </summary>
        public static async void SimulateDoubleClick(MouseButton button)
        {
            SimulateSingleClick(button);
            SimulateSingleClick(button);
        }

        /// <summary>Convenience wrappers for specific double-clicks.</summary>
        public static void SimulateLeftDoubleClick() => SimulateDoubleClick(MouseButton.Left);
        public static void SimulateRightDoubleClick() => SimulateDoubleClick(MouseButton.Right);
        public static void SimulateMiddleDoubleClick() => SimulateDoubleClick(MouseButton.Middle);
        public static void SimulatePrimaryDoubleClick() => SimulateDoubleClick(MouseButton.Default);

        /// <summary>
        /// Simulates a click on the specified mouse button (single or double).
        /// </summary>
        public static void SimulateClick(MouseButton button, ClickType type)
        {
            switch (type)
            {
                case ClickType.Single:
                    SimulateSingleClick(button);
                    break;

                case ClickType.Double:
                    SimulateDoubleClick(button);
                    break;
            }
        }

        /// <summary>Returns the user's primary mouse button.</summary>
        public static MouseButton GetPrimaryMouseButton()
        {
            bool swapped = GetSystemMetrics(SM_SWAPBUTTON) != 0;
            return swapped ? MouseButton.Right : MouseButton.Left;
        }
        #endregion

        #region Type Converters Methods
        public static MouseButton GetMouseButtonFromString(string mouseButton) => mouseButton switch
        {
            "Left" => MouseButton.Left,
            "Right" => MouseButton.Right,
            "Middle" => MouseButton.Middle,
            _ => MouseButton.Default
        };

        public static string GetStringFromMouseButton(MouseButton mouseButton) => mouseButton switch
        {
            MouseButton.Left => "Left",
            MouseButton.Right => "Right",
            MouseButton.Middle => "Middle",
            _ => "Default"
        };

        public static ClickType GetClickTypeFromString(string clickType) => clickType switch
        {
            "Single" => ClickType.Single,
            "Double" => ClickType.Double,
            _ => ClickType.Single
        };

        public static string GetStringFromClickType(ClickType clickType) => clickType switch
        {
            ClickType.Single => "Single",
            ClickType.Double => "Double",
            _ => "Single"
        };
        #endregion

        #region WinAPI
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x, y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData, flags, time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void mouse_event(MouseEventFlags dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        private static extern uint GetDoubleClickTime();

        private const int SM_SWAPBUTTON = 23;

        [Flags]
        private enum MouseEventFlags : uint
        {
            LEFTDOWN = 0x0002,
            LEFTUP = 0x0004,
            RIGHTDOWN = 0x0008,
            RIGHTUP = 0x0010,
            MIDDLEDOWN = 0x0020,
            MIDDLEUP = 0x0040
        }

        public enum MouseButton
        {
            Left,
            Right,
            Middle,
            Default
        }

        public enum ClickType { Single, Double }
        #endregion
    }
}