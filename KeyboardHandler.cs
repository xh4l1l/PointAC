using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace PointAC
{
    public static class KeyboardHandler
    {
        private static Action<Key>? _onKeyDown;
        private static Action<Key>? _onKeyUp;
        private static Action<Key>? _onKeyPress;

        private static IntPtr _hookId = IntPtr.Zero;
        private static LowLevelKeyboardProc _proc = HookCallback;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        #region Public API

        public static void ListenToKeyDown(Action<Key> onKeyDown)
        {
            _onKeyDown = onKeyDown;
            EnsureHook();
        }

        public static void ListenToKeyUp(Action<Key> onKeyUp)
        {
            _onKeyUp = onKeyUp;
            EnsureHook();
        }

        public static void ListenToKeyPress(Action<Key> onKeyPress)
        {
            _onKeyPress = onKeyPress;
            EnsureHook();
        }

        public static void StopListeningToKeyDown()
        {
            _onKeyDown = null;

            if (_onKeyUp == null && _onKeyPress == null && _hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        public static void StopListeningToKeyUp()
        {
            _onKeyUp = null;

            if (_onKeyDown == null && _onKeyPress == null && _hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        public static void StopListeningToKeyPress()
        {
            _onKeyPress = null;

            if (_onKeyDown == null && _onKeyUp == null && _hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }
        public static void StopListeningToAll()
        {
            _onKeyDown = null;
            _onKeyUp = null;
            _onKeyPress = null;

            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
        }

        #endregion

        #region Internals

        private static void EnsureHook()
        {
            if (_hookId == IntPtr.Zero)
            {
                using var curProcess = Process.GetCurrentProcess();
                using var curModule = curProcess.MainModule!;
                _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Key key = KeyInterop.KeyFromVirtualKey(vkCode);

                if (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN)
                {
                    _onKeyDown?.Invoke(key);
                    _onKeyPress?.Invoke(key);
                }
                else if (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP)
                {
                    _onKeyUp?.Invoke(key);
                }
            }

            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        #endregion

        #region WinAPI

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        #endregion
    }
}