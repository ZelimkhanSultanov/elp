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
    public static string thisFilePath = Path.Combine(thisDirectory, thisProcessname + ".exe");

    public static string installDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Elp\\";

    [DllImport("shell32.dll")]
    static extern bool SHGetSpecialFolderPath(IntPtr hwndOwner, [Out] StringBuilder lpszPath, int nFolder, bool fCreate);
    const int CSIDL_COMMON_STARTMENU = 0x16;  // All Users\Start Menu
    public static string fileName = Process.GetCurrentProcess().ProcessName + ".exe";
    //    public static string aStartUpDirPath = "";
    public static string cStartUpDirPath = "";
    private static string realPath = "";

    private static void copyFiles()
    {
        try {
            System.IO.File.Copy(thisDirectory + fileName, installDir + fileName, true);
            System.IO.File.Copy(thisDirectory + "WindowsInput.xml", installDir + "WindowsInput.xml", true);
            System.IO.File.Copy(thisDirectory + "WindowsInput.dll", installDir + "WindowsInput.dll", true);
            CreateStartupFolderShortcut();
        } catch (Exception ex) {
            MessageBox.Show("Возникла ошибка при копировании файлов");
            MessageBox.Show(ex.Message);
        }
    }

    public static void CreateStartupFolderShortcut()
    {
        WshShellClass wshShell = new WshShellClass();
        IWshRuntimeLibrary.IWshShortcut shortcut;
        string startUpFolderPath =
          Environment.GetFolderPath(Environment.SpecialFolder.Startup);

        // Create the shortcut
        shortcut =
          (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(
            startUpFolderPath + "\\" +
            Application.ProductName + ".lnk");

        shortcut.TargetPath = Application.ExecutablePath;
        shortcut.WorkingDirectory = Application.StartupPath;
        shortcut.Description = "Elp I Autostart";
        // shortcut.IconLocation = Application.StartupPath + @"\App.ico";
        shortcut.Save();
    }

    public static void Main()
    {
        DirectoryInfo di = new System.IO.DirectoryInfo(installDir);
        if (!di.Exists) di.Create();


        /* StringBuilder path = new StringBuilder(260);
         SHGetSpecialFolderPath(IntPtr.Zero, path, CSIDL_COMMON_STARTMENU, false);
       //  aStartUpDirPath = path.ToString() + "\\";
       */
        cStartUpDirPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\";
        


        mInputSimulator = new InputSimulator();

        // Is it should be updated?
        // Look if there is any existing copy of the application

        //  if (new FileInfo(aStartUpDirPath + fileName).Exists) {
        //       realPath = aStartUpDirPath + fileName;
        //   } else 

        if (new FileInfo(cStartUpDirPath + fileName).Exists) {
            realPath = cStartUpDirPath + fileName;
        }

        if (!realPath.Equals("")) {
            var versionInfo = FileVersionInfo.GetVersionInfo(realPath);
            int exVersionMajor = versionInfo.FileMajorPart;
            int exVersionMinor = versionInfo.FileMinorPart;

            versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            int myVersionMajor = versionInfo.FileMajorPart;
            int myVersionMinor = versionInfo.FileMinorPart;

            if (myVersionMajor > exVersionMajor || (myVersionMajor == exVersionMajor && myVersionMinor > exVersionMinor)) {
                // Must update
                foreach (var process in Process.GetProcessesByName("elp")) {
                    if (process.Id != Process.GetCurrentProcess().Id)
                        process.Kill();
                }

                System.IO.File.Delete(realPath);
                copyFiles();
            }
        }


        using (Mutex mutex = new Mutex(false, "Global\\" + appGuid)) {
            if (!mutex.WaitOne(0, false)) {
                MessageBox.Show("Программа уже работает!");
                return;
            }
            // It's the first time
            if (realPath.Equals("")) {
                copyFiles();
                MessageBox.Show("Программа запущена и успешно установлена в автозапуск для текущего пользователя!");
                _hookID = SetHook(_proc);
                Application.Run();
            } else {
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