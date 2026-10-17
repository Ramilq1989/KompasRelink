// KompasRelink — массовая замена путей к внешним файлам в документах КОМПАС-3D (v20+).
// Сборка: build.bat (csc .NET Framework 4, C# 5). Без interop-сборок: позднее связывание COM.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("KompasRelink")]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace KompasRelink
{
    // Нужные члены IKompasDocument1 (dispinterface, kAPI7.tlb). DISPID взяты из библиотеки типов.
    [ComImport, Guid("58890FE8-E671-4561-994A-600DD29032E4"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IKompasDocument1
    {
        [DispId(10007)] object this[bool allFiles] { get; }            // ExternalFilesNames
        [DispId(10008)] bool SaveAsEx(string pathName, int saveMode);
        [DispId(10022)] int OpenVersion { get; }
        [DispId(10023)] bool ReplaceExternalFilesNames(bool allFiles, string oldFileName, string newFileName);
    }

    public class Options
    {
        public List<string> Roots = new List<string>();
        public List<KeyValuePair<string, string>> Map = new List<KeyValuePair<string, string>>();
        public bool Apply;
        public bool Backup = true;
        public string SaveVersion = "keep";   // keep | current | 17..23
        public string[] Exts = { ".a3d", ".m3d", ".cdw", ".spw", ".frw", ".kdw" };
    }

    public class Relinker
    {
        // ksSaveDocumentVersionEnum: основная версия КОМПАСа -> код для SaveAsEx
        static readonly Dictionary<int, int> SDV = new Dictionary<int, int> {
            { 17, 21 }, { 18, 23 }, { 19, 25 }, { 20, 26 }, { 21, 27 }, { 22, 28 }, { 23, 29 } };
        public static int[] SaveVersions { get { var k = new List<int>(SDV.Keys); k.Sort(); k.Reverse(); return k.ToArray(); } }

        readonly Options o;
        readonly Action<string> log;
        readonly Action<int, int> progress;
        public volatile bool Cancel;

        public Relinker(Options o, Action<string> log, Action<int, int> progress)
        {
            this.o = o; this.log = log; this.progress = progress;
        }

        public static string Norm(string p)
        {
            p = p.Trim().Trim('"').Replace('/', '\\');
            if (p.Length >= 2 && char.IsLetter(p[0]) && p[1] == '\\') p = p.Substring(0, 1) + ":" + p.Substring(1);
            if (p.Length == 1 && char.IsLetter(p[0])) p += ":";
            return p.TrimEnd('\\');
        }

        public static List<KeyValuePair<string, string>> ParseMap(IEnumerable<string> lines)
        {
            var r = new List<KeyValuePair<string, string>>();
            foreach (var raw in lines)
            {
                var l = raw.Trim();
                if (l.Length == 0 || l.StartsWith("#")) continue;
                int i = l.IndexOf('=');
                if (i <= 0) throw new ArgumentException("Строка замены должна быть вида СТАРЫЙ=НОВЫЙ: " + l);
                r.Add(new KeyValuePair<string, string>(Norm(l.Substring(0, i)), Norm(l.Substring(i + 1))));
            }
            r.Sort((a, b) => b.Key.Length.CompareTo(a.Key.Length));   // длинные префиксы первыми
            return r;
        }

        public string Remap(string path)
        {
            var p = path.Replace('/', '\\');
            foreach (var kv in o.Map)
            {
                if (p.Equals(kv.Key, StringComparison.OrdinalIgnoreCase) ||
                    p.StartsWith(kv.Key + "\\", StringComparison.OrdinalIgnoreCase))
                    return kv.Value + p.Substring(kv.Key.Length);
            }
            return null;
        }

        IEnumerable<string> FindFiles(string root)
        {
            var stack = new Stack<string>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                var d = stack.Pop();
                string[] files = new string[0], dirs = new string[0];
                try { files = Directory.GetFiles(d); dirs = Directory.GetDirectories(d); }
                catch (Exception e) { log("  нет доступа: " + d + " (" + e.Message + ")"); }
                foreach (var f in files)
                    foreach (var x in o.Exts)
                        if (f.EndsWith(x, StringComparison.OrdinalIgnoreCase)) { yield return f; break; }
                for (int i = dirs.Length - 1; i >= 0; i--)
                    if (!Path.GetFileName(dirs[i]).StartsWith("_relink_backup", StringComparison.OrdinalIgnoreCase))
                        stack.Push(dirs[i]);
            }
        }

        static string Csv(string s) { return "\"" + (s ?? "").Replace("\"", "\"\"") + "\""; }

        public object ExternalApp;          // IApplication, когда работаем внутри КОМПАСа (библиотека)
        public Action Pump;                 // обработка сообщений окна между файлами (режим библиотеки)

        public void Run()
        {
            var files = new List<string>();
            foreach (var r in o.Roots)
            {
                if (File.Exists(r)) files.Add(Path.GetFullPath(r));
                else if (Directory.Exists(r)) { log("Поиск файлов: " + r); files.AddRange(FindFiles(Path.GetFullPath(r))); }
                else log("Не найдено: " + r);
            }
            if (files.Count == 0) { log("Файлы КОМПАС не найдены."); return; }

            string baseDir = Directory.Exists(o.Roots[0]) ? Path.GetFullPath(o.Roots[0]) : Path.GetDirectoryName(Path.GetFullPath(o.Roots[0]));
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backup = Path.Combine(baseDir, "_relink_backup_" + stamp);
            string report = Path.Combine(baseDir, "relink_report_" + stamp + ".csv");

            log("Замены:");
            foreach (var kv in o.Map) log("  " + kv.Key + "  ->  " + kv.Value);
            log("Файлов найдено: " + files.Count + "   режим: " + (o.Apply ? "ЗАПИСЬ" : "ПРОБНЫЙ (ничего не сохраняется)"));

            dynamic app; bool started = false;
            if (ExternalApp != null) app = ExternalApp;
            else try { app = Marshal.GetActiveObject("Kompas.Application.7"); }
            catch { log("Запуск КОМПАС-3D…"); app = Activator.CreateInstance(Type.GetTypeFromProgID("Kompas.Application.7", true)); started = true; }
            string appName = app.ApplicationName(true);
            var m = Regex.Match(appName ?? "", @"v(\d+)");
            int appMajor = m.Success ? int.Parse(m.Groups[1].Value) : 0;
            log("КОМПАС: " + appName + "   сохранение: " + o.SaveVersion);
            log("");

            int oldHide = app.HideMessage;
            app.HideMessage = 2;     // ksHideMessageNo — на вопросы «найти файл?» отвечать «Нет»
            var rows = new StringBuilder();
            rows.AppendLine("Файл;Старая ссылка;Новая ссылка;Новый файл существует;Статус");
            int nChanged = 0, nErr = 0;
            var t0 = DateTime.Now;
            try
            {
                for (int i = 0; i < files.Count && !Cancel; i++)
                {
                    string path = files[i];
                    string tag = "[" + (i + 1) + "/" + files.Count + "] " + path;
                    progress(i, files.Count);
                    if (Pump != null) Pump();
                    dynamic doc = null;
                    try
                    {
                        doc = app.Documents.Open(path, false, !o.Apply);
                        if (doc == null) throw new Exception("КОМПАС не открыл файл");
                        var d1 = (IKompasDocument1)doc;
                        var links = d1[false] as Array;
                        var todo = new List<KeyValuePair<string, string>>();
                        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);   // в чертеже — по ссылке на каждый вид
                        if (links != null)
                            foreach (object l in links)
                            {
                                string s = l as string; if (s == null || !seen.Add(s)) continue;
                                string n = Remap(s);
                                if (n != null && n != s) todo.Add(new KeyValuePair<string, string>(s, n));
                            }
                        if (todo.Count == 0) { doc.Close(0); doc = null; continue; }

                        log(tag);
                        bool okAll = true;
                        foreach (var kv in todo)
                        {
                            bool exists = File.Exists(kv.Value);
                            string status = "будет заменено";
                            if (o.Apply)
                            {
                                bool ok = d1.ReplaceExternalFilesNames(false, kv.Key, kv.Value);
                                okAll &= ok;
                                status = ok ? "заменено" : "ОШИБКА замены";
                            }
                            log("    " + kv.Key + "\r\n      -> " + kv.Value + "   [" + status + (exists ? "" : ", НОВОГО ФАЙЛА НЕТ") + "]");
                            rows.AppendLine(Csv(path) + ";" + Csv(kv.Key) + ";" + Csv(kv.Value) + ";" + (exists ? "да" : "НЕТ") + ";" + status);
                        }
                        if (o.Apply)
                        {
                            int fv = d1.OpenVersion >> 24;
                            int target = 0;   // 0 — обычное сохранение в текущей версии
                            if (o.SaveVersion == "keep") target = fv >= 17 ? fv : 0;
                            else if (o.SaveVersion != "current") target = int.Parse(o.SaveVersion);
                            if (target != 0 && appMajor != 0 && target >= appMajor) target = 0;
                            if (target != 0 && !SDV.ContainsKey(target))
                                throw new Exception("файл версии v" + target + ": сохранение в неё не поддерживается");
                            if (o.Backup)
                            {
                                string rel = path.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase)
                                    ? path.Substring(baseDir.Length).TrimStart('\\') : Path.GetFileName(path);
                                string dst = Path.Combine(backup, rel);
                                Directory.CreateDirectory(Path.GetDirectoryName(dst));
                                File.Copy(path, dst, true);
                            }
                            var mtime = File.GetLastWriteTimeUtc(path);
                            if (target == 0) doc.Save();
                            else if (!d1.SaveAsEx(path, SDV[target])) throw new Exception("не удалось сохранить в формате v" + target);
                            if (File.GetLastWriteTimeUtc(path) == mtime)
                                throw new Exception("не удалось сохранить (файл занят или только для чтения?)");
                            string ver = "v" + (target != 0 ? target : appMajor);
                            log(okAll ? "    сохранено (" + ver + ")" : "    сохранено (" + ver + "), но не все ссылки заменились");
                        }
                        doc.Close(0); doc = null;
                        nChanged++;
                    }
                    catch (Exception e)
                    {
                        nErr++;
                        log(tag + "\r\n    ОШИБКА: " + e.Message);
                        rows.AppendLine(Csv(path) + ";;;;" + Csv("ОШИБКА: " + e.Message));
                        if (doc != null) try { doc.Close(0); } catch { }
                    }
                }
            }
            finally
            {
                try { app.HideMessage = oldHide; } catch { }
                if (started) try { app.Quit(); } catch { }
                progress(files.Count, files.Count);
            }
            File.WriteAllText(report, rows.ToString(), new UTF8Encoding(true));
            log("");
            if (Cancel) log("ОСТАНОВЛЕНО пользователем.");
            log("Файлов со ссылками на замену: " + nChanged + ", ошибок: " + nErr +
                "   время: " + (DateTime.Now - t0).ToString(@"hh\:mm\:ss"));
            log("Отчёт: " + report);
            if (o.Apply && o.Backup && nChanged > 0) log("Резервные копии: " + backup);
            if (!o.Apply) log("Это был пробный прогон. Для записи нажмите «Заменить».");
        }
    }

    public class MainForm : Form
    {
        TextBox tbRoot = new TextBox(), tbMap = new TextBox(), tbLog = new TextBox();
        ComboBox cbVer = new ComboBox();
        CheckBox chBackup = new CheckBox();
        Button btBrowse = new Button(), btCheck = new Button(), btApply = new Button(), btStop = new Button();
        ProgressBar pb = new ProgressBar();
        Label lbProg = new Label();
        Relinker current;
        readonly object kompasApp;          // не null — форма открыта из библиотеки внутри КОМПАСа

        public MainForm() : this(null) { }

        public MainForm(object kompasApp)
        {
            this.kompasApp = kompasApp;
            Text = "KompasRelink — замена путей в сборках КОМПАС-3D";
            Font = new Font("Segoe UI", 9f);
            if (kompasApp != null) { var ic = KompasLibrary.LoadIcon(); if (ic != null) Icon = ic; }
            else try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            ClientSize = new Size(860, 640);
            MinimumSize = new Size(640, 480);

            var l1 = new Label { Text = "Папка (обрабатываются все подпапки):", Left = 12, Top = 12, AutoSize = true };
            tbRoot.SetBounds(12, 32, 740, 24); tbRoot.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            btBrowse.Text = "Обзор…"; btBrowse.SetBounds(760, 31, 88, 26); btBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btBrowse.Click += delegate {
                using (var d = new FolderBrowserDialog { SelectedPath = tbRoot.Text })
                    if (d.ShowDialog(this) == DialogResult.OK) tbRoot.Text = d.SelectedPath;
            };

            var l2 = new Label { Text = "Замены, по одной в строке:  СТАРЫЙ=НОВЫЙ   (слэши / и \\ равнозначны, регистр не важен)", Left = 12, Top = 66, AutoSize = true };
            tbMap.Multiline = true; tbMap.ScrollBars = ScrollBars.Vertical; tbMap.Font = new Font("Consolas", 10f);
            tbMap.SetBounds(12, 86, 836, 72); tbMap.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            tbMap.Text = "//metall=G:";

            var l3 = new Label { Text = "Сохранять в версии:", Left = 12, Top = 172, AutoSize = true };
            cbVer.DropDownStyle = ComboBoxStyle.DropDownList; cbVer.SetBounds(140, 168, 230, 24);
            cbVer.Items.Add("как был файл (рекомендуется)");
            cbVer.Items.Add("текущей версии КОМПАСа");
            foreach (int v in Relinker.SaveVersions) cbVer.Items.Add("v" + v);
            cbVer.SelectedIndex = 0;
            chBackup.Text = "Резервные копии"; chBackup.Checked = true; chBackup.SetBounds(390, 169, 150, 24);

            btCheck.Text = "Проверить"; btCheck.SetBounds(560, 166, 90, 28); btCheck.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btApply.Text = "Заменить"; btApply.SetBounds(656, 166, 90, 28); btApply.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btStop.Text = "Стоп"; btStop.SetBounds(758, 166, 90, 28); btStop.Anchor = AnchorStyles.Top | AnchorStyles.Right; btStop.Enabled = false;
            btCheck.Click += delegate { Start(false); };
            btApply.Click += delegate {
                if (MessageBox.Show(this, "Записать изменения в файлы?\r\nЗакройте эти сборки в КОМПАСе перед заменой.",
                    Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK) Start(true);
            };
            btStop.Click += delegate { if (current != null) { current.Cancel = true; Log("Остановка после текущего файла…"); } };

            pb.SetBounds(12, 204, 700, 18); pb.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            lbProg.SetBounds(720, 204, 128, 18); lbProg.Anchor = AnchorStyles.Top | AnchorStyles.Right; lbProg.TextAlign = ContentAlignment.MiddleRight;

            tbLog.Multiline = true; tbLog.ReadOnly = true; tbLog.ScrollBars = ScrollBars.Both; tbLog.WordWrap = false;
            tbLog.Font = new Font("Consolas", 9f); tbLog.BackColor = SystemColors.Window;
            tbLog.SetBounds(12, 230, 836, 398); tbLog.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;

            Controls.AddRange(new Control[] { l1, tbRoot, btBrowse, l2, tbMap, l3, cbVer, chBackup, btCheck, btApply, btStop, pb, lbProg, tbLog });
            LoadSettings();
            FormClosing += delegate { SaveSettings(); };
        }

        static string SettingsPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KompasRelink.txt"); } }

        void LoadSettings()
        {
            try
            {
                var lines = File.ReadAllLines(SettingsPath, Encoding.UTF8);
                if (lines.Length > 0) tbRoot.Text = lines[0];
                if (lines.Length > 1) tbMap.Text = string.Join("\r\n", lines, 1, lines.Length - 1);
            }
            catch { }
        }

        void SaveSettings()
        {
            try { File.WriteAllText(SettingsPath, tbRoot.Text + "\r\n" + tbMap.Text, Encoding.UTF8); } catch { }
        }

        void Log(string s)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(Log), s); return; }
            tbLog.AppendText(s + "\r\n");
        }

        void Progress(int i, int n)
        {
            if (InvokeRequired) { BeginInvoke(new Action<int, int>(Progress), i, n); return; }
            pb.Maximum = Math.Max(n, 1); pb.Value = Math.Min(i, pb.Maximum);
            lbProg.Text = i + " / " + n;
        }

        void Start(bool apply)
        {
            var o = new Options { Apply = apply, Backup = chBackup.Checked };
            if (tbRoot.Text.Trim().Length == 0) { MessageBox.Show(this, "Укажите папку.", Text); return; }
            o.Roots.Add(tbRoot.Text.Trim());
            try { o.Map = Relinker.ParseMap(tbMap.Lines); }
            catch (Exception e) { MessageBox.Show(this, e.Message, Text); return; }
            if (o.Map.Count == 0) { MessageBox.Show(this, "Укажите хотя бы одну замену СТАРЫЙ=НОВЫЙ.", Text); return; }
            o.SaveVersion = cbVer.SelectedIndex == 0 ? "keep" : cbVer.SelectedIndex == 1 ? "current" : cbVer.Text.Substring(1);
            SaveSettings();

            tbLog.Clear();
            SetBusy(true);
            current = new Relinker(o, Log, Progress);
            if (kompasApp != null)
            {
                // внутри КОМПАСа: работаем в потоке окна, между файлами обрабатываем сообщения (кнопка «Стоп»)
                current.ExternalApp = kompasApp;
                current.Pump = Application.DoEvents;
                try { current.Run(); }
                catch (Exception e) { Log("ОШИБКА: " + e.Message); }
                finally { SetBusy(false); }
                return;
            }
            var r = current;
            var th = new Thread(delegate () {
                try { r.Run(); }
                catch (Exception e) { Log("ОШИБКА: " + e.Message); }
                finally { BeginInvoke(new Action(delegate { SetBusy(false); })); }
            });
            th.SetApartmentState(ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
        }

        void SetBusy(bool b)
        {
            btCheck.Enabled = btApply.Enabled = btBrowse.Enabled = tbRoot.Enabled = tbMap.Enabled = cbVer.Enabled = chBackup.Enabled = !b;
            btStop.Enabled = b;
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            // Консольный режим: KompasRelink.exe <папка> --map СТАРЫЙ=НОВЫЙ [--map …] [--apply] [--no-backup] [--save-version keep|current|23]
            if (args.Length > 0) return Cli(args);
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
            return 0;
        }

        [DllImport("kernel32.dll")] static extern bool AttachConsole(int pid);

        // Регистрация KompasRelinkLib.dll (рядом с exe) как библиотеки КОМПАС для текущего пользователя
        static int Install(bool install)
        {
            string dll = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "KompasRelinkLib.dll");
            try
            {
                if (!install) { LibraryInstaller.Unregister(); Say("Регистрация библиотеки снята."); return 0; }
                if (!File.Exists(dll)) { Say("Не найден " + dll); return 1; }
                LibraryInstaller.Register(dll);
                Say("Библиотека зарегистрирована:\r\n" + dll +
                    "\r\n\r\nВ КОМПАСе: Приложения → Конфигуратор приложений → Добавить приложение → выберите KompasRelinkLib.dll.");
                return 0;
            }
            catch (Exception e) { Say("ОШИБКА: " + e.Message); return 1; }
        }

        static void Say(string s)
        {
            Console.WriteLine(s);
            if (Console.IsOutputRedirected) return;
            MessageBox.Show(s, "KompasRelink");
        }

        static int Cli(string[] args)
        {
            AttachConsole(-1);
            Console.OutputEncoding = Encoding.UTF8;
            if (args[0] == "--install" || args[0] == "--uninstall") return Install(args[0] == "--install");
            var o = new Options();
            var maps = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (a == "--map" && i + 1 < args.Length) maps.Add(args[++i]);
                else if (a == "--apply") o.Apply = true;
                else if (a == "--no-backup") o.Backup = false;
                else if (a == "--save-version" && i + 1 < args.Length) o.SaveVersion = args[++i].ToLowerInvariant();
                else o.Roots.Add(a);
            }
            if (o.Roots.Count == 0 || maps.Count == 0)
            {
                Console.WriteLine("KompasRelink.exe <папка> --map СТАРЫЙ=НОВЫЙ [--apply] [--no-backup] [--save-version keep|current|23]");
                return 2;
            }
            o.Map = Relinker.ParseMap(maps);
            new Relinker(o, Console.WriteLine, delegate { }).Run();
            return 0;
        }
    }
}
