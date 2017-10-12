using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;
using System.IO;
using System.Reflection;
using WindowsInput;
using WindowsInput.Native;
using Elp;
using System.Text;
using IWshRuntimeLibrary;

class InterceptKeys
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private static LowLevelKeyboardProc _proc = HookCallback;
    private static IntPtr _hookID = IntPtr.Zero;

    // appGuid - something unique to this program
    private static string appGuid = "a8d89asud89asud98123di23hd923jfiosdklfsiodf0";
    private static InputSimulator mInputSimulator;

    private static string thisLocation = Assembly.GetEntryAssembly().Location;
    private static string thisDirectory = Path.GetDirectoryName(thisLocation) + "\\";
    private static string thisProcessname = Process.GetCurrentProcess().ProcessName;
    private static string thisFilePath = Path.Combine(thisDirectory, thisProcessname + ".exe");

    private static string fileName = Process.GetCurrentProcess().ProcessName + ".exe";
    private static string shortcutFileName = Application.ProductName + ".lnk";
    private static string installDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Elp\\";

    public static string myStartUpDirPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\";


    private static bool isInstalled()
    {
        return (new DirectoryInfo(installDir).Exists
            && new FileInfo(installDir + fileName).Exists
            && new FileInfo(installDir + "WindowsInput.xml").Exists
            && new FileInfo(installDir + "WindowsInput.dll").Exists);
    }

    private static void removeFiles()
    {
        try {
            System.IO.File.Delete(installDir + fileName);
            System.IO.File.Delete(installDir + "WindowsInput.xml");
            System.IO.File.Delete(installDir + "WindowsInput.dll");
            System.IO.File.Delete(myStartUpDirPath + shortcutFileName);
        } catch (Exception) {
            MessageBox.Show("Возникла ошибка при удалении файлов");
        }
    }

    private static bool install()
    {
        try { new DirectoryInfo(installDir).Create(); } catch (Exception) { }
        try {
            System.IO.File.Copy(thisDirectory + fileName, installDir + fileName, true);
            System.IO.File.Copy(thisDirectory + "WindowsInput.xml", installDir + "WindowsInput.xml", true);
            System.IO.File.Copy(thisDirectory + "WindowsInput.dll", installDir + "WindowsInput.dll", true);
            CreateStartupFolderShortcut(myStartUpDirPath + "\\" + shortcutFileName, installDir + fileName);
            return true;
        } catch (Exception) {
            MessageBox.Show("Возникла ошибка при копировании файлов");
            return false;
        }
    }

    public static void CreateStartupFolderShortcut(string pShortcutPath, string pTargetPath)
    {
        WshShellClass wshShell = new WshShellClass();
        IWshShortcut shortcut;

        // Create the shortcut
        shortcut = (IWshShortcut)wshShell.CreateShortcut(pShortcutPath);

        shortcut.TargetPath = pTargetPath;     
        shortcut.WorkingDirectory = Path.GetDirectoryName(pTargetPath);
        shortcut.Description = "Elp I Autostart";
        // shortcut.IconLocation = Application.StartupPath + @"\App.ico";
        shortcut.Save();
    }

    public static void Main()
    {
        // Allow only one instance
        using (Mutex mutex = new Mutex(false, "Global\\" + appGuid)) {
            if (!mutex.WaitOne(0, false)) {
                MessageBox.Show("Программа уже работает!");
                return;
            }

            mInputSimulator = new InputSimulator();

            if (!isInstalled()) {
                // New installation
                install();

                MessageBox.Show("Программа запущена и успешно установлена в автозапуск!", "Elp I", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _hookID = SetHook(_proc);
                Application.Run();
            } else {
                // Already installed

                // Check if it should be updated

                // Old version
                var versionInfo = FileVersionInfo.GetVersionInfo(installDir + fileName);
                int exVersionMajor = versionInfo.FileMajorPart;
                int exVersionMinor = versionInfo.FileMinorPart;

                // New version
                versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                int myVersionMajor = versionInfo.FileMajorPart;
                int myVersionMinor = versionInfo.FileMinorPart;

                if (myVersionMajor > exVersionMajor || (myVersionMajor == exVersionMajor && myVersionMinor > exVersionMinor)) {
                    // Must update
                    foreach (var process in Process.GetProcessesByName("elp")) {
                        if (process.Id != Process.GetCurrentProcess().Id)
                            process.Kill();
                    }

                    // Update to new version
                    removeFiles();
                    if (install()) {
                        MessageBox.Show("Программа успешно обновлена", "Elp I", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }


                if (new FileInfo(myStartUpDirPath + shortcutFileName).Exists) {
                    // Shortcut has already been installed
                }

                _hookID = SetHook(_proc);
                Application.Run();
            }
            UnhookWindowsHookEx(_hookID);
        }
    }
    
    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule) {
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
        }
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN) {
            int vkCode = Marshal.ReadInt32(lParam);
            CultureInfo ci = GetCurrentKeyboardLayout();
            if (ci.TwoLetterISOLanguageName == "ru") {
                var isShiftKeyDown = mInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.SHIFT);

                if ((Keys)vkCode == Keys.Oem5 && !isShiftKeyDown) {
                    // Determines if the shift key is currently down

                    mInputSimulator.Keyboard.TextEntry("Ӏ");
                    return (IntPtr)1;
                }
            }
        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, IntPtr proccess);
    [DllImport("user32.dll")] static extern IntPtr GetKeyboardLayout(uint thread);
    public static CultureInfo GetCurrentKeyboardLayout()
    {
        try {
            IntPtr foregroundWindow = GetForegroundWindow();
            uint foregroundProcess = GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
            int keyboardLayout = GetKeyboardLayout(foregroundProcess).ToInt32() & 0xFFFF;
            return new CultureInfo(keyboardLayout);
        } catch (Exception _) {
            return new CultureInfo(1033); // Assume English if something went wrong.
        }
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
}