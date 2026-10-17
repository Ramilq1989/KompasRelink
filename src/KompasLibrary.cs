// ActiveX-библиотека КОМПАС-3D: пункт меню «Замена путей ссылок…» открывает окно KompasRelink.
// Собирается вместе с KompasRelink.cs в KompasRelinkLib.dll (см. build.bat).
// Регистрация: KompasRelink.exe --install (для текущего пользователя, без прав администратора).
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace KompasRelink
{
    [ComVisible(true)]
    [Guid("6B1F3C52-9D7E-4A8B-B3E1-2C5D8F0A7E41")]
    [ProgId("KompasRelink.Library")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class KompasLibrary
    {
        public const string ClassGuid = "{6B1F3C52-9D7E-4A8B-B3E1-2C5D8F0A7E41}";
        public const string ProgIdName = "KompasRelink.Library";

        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "KompasRelink — замена путей";
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string DisplayLibraryName()
        {
            return "KompasRelink — замена путей";
        }

        // TRUE: в ExternalRunCommand придёт IApplication (API7), а не KompasObject
        public bool IsOnApplication7()
        {
            return true;
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            switch (number)
            {
                case 1: itemType = 1; command = 1; return "Замена путей ссылок…";
                default: itemType = 3; command = -1; return "";
            }
        }

        public void ExternalRunCommand([In] short command, [In] short mode,
                                       [In, MarshalAs(UnmanagedType.IDispatch)] object kompas)
        {
            try
            {
                object app = kompas;
                try { ((dynamic)kompas).ApplicationName(true); }
                catch { app = ((dynamic)kompas).ksGetApplication7(); }   // старый вариант: пришёл KompasObject
                Application.EnableVisualStyles();
                using (var f = new MainForm(app))
                {
                    f.StartPosition = FormStartPosition.CenterScreen;
                    // библиотека работает внутри процесса КОМПАСа: его главное окно — владелец формы
                    f.ShowDialog(new MainWindow(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("KompasRelink: " + e.Message, "KompasRelink", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        class MainWindow : IWin32Window
        {
            readonly IntPtr h;
            public MainWindow(IntPtr hwnd) { h = hwnd; }
            public IntPtr Handle { get { return h; } }
        }

        // regasm /codebase тоже работает: добавляет признак библиотеки КОМПАС
        [ComRegisterFunction]
        public static void RegasmRegister(Type t)
        {
            using (var k = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + ClassGuid, true))
                if (k != null) k.CreateSubKey("Kompas_Library").Close();
        }

        [ComUnregisterFunction]
        public static void RegasmUnregister(Type t)
        {
            try { Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + ClassGuid + @"\Kompas_Library"); } catch { }
        }

        public static Icon LoadIcon()
        {
            try
            {
                using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("KompasRelink.ico"))
                    return s != null ? new Icon(s) : null;
            }
            catch { return null; }
        }
    }

    // Регистрация DLL как COM-библиотеки КОМПАС для текущего пользователя (HKCU, без администратора)
    public static class LibraryInstaller
    {
        const string Root = @"Software\Classes\";

        public static void Register(string dll)
        {
            dll = Path.GetFullPath(dll);
            var asm = AssemblyName.GetAssemblyName(dll);
            string codeBase = "file:///" + dll.Replace('\\', '/');     // как пишет regasm, без %-кодирования
            string g = KompasLibrary.ClassGuid;
            using (var k = Registry.CurrentUser.CreateSubKey(Root + @"CLSID\" + g))
            {
                k.SetValue(null, KompasLibrary.ProgIdName);
                using (var s = k.CreateSubKey("InprocServer32"))
                {
                    s.SetValue(null, "mscoree.dll");
                    s.SetValue("ThreadingModel", "Both");
                    s.SetValue("Class", typeof(KompasLibrary).FullName);
                    s.SetValue("Assembly", asm.FullName);
                    s.SetValue("RuntimeVersion", "v4.0.30319");
                    s.SetValue("CodeBase", codeBase);
                    using (var v = s.CreateSubKey(asm.Version.ToString()))
                    {
                        v.SetValue("Class", typeof(KompasLibrary).FullName);
                        v.SetValue("Assembly", asm.FullName);
                        v.SetValue("RuntimeVersion", "v4.0.30319");
                        v.SetValue("CodeBase", codeBase);
                    }
                }
                k.CreateSubKey("ProgId").SetValue(null, KompasLibrary.ProgIdName);
                k.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}").Close();
                k.CreateSubKey("Kompas_Library").Close();
            }
            using (var k = Registry.CurrentUser.CreateSubKey(Root + KompasLibrary.ProgIdName))
            {
                k.SetValue(null, KompasLibrary.ProgIdName);
                k.CreateSubKey("CLSID").SetValue(null, KompasLibrary.ClassGuid);
            }
        }

        public static void Unregister()
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(Root + @"CLSID\" + KompasLibrary.ClassGuid); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(Root + KompasLibrary.ProgIdName); } catch { }
        }
    }
}
