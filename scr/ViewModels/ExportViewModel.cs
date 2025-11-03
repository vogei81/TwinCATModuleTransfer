using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using TCatSysManagerLib;
using TwinCATModuleTransfer.Services;
using TwinCATModuleTransfer.Utils;

namespace TwinCATModuleTransfer.ViewModels
{
    public class ExportViewModel : BaseViewModel
    {
        public ObservableCollection<SelectionTreeItem> Tree { get; private set; }
            = new ObservableCollection<SelectionTreeItem>();

        public ObservableCollection<string> RecentSolutions { get; private set; }
            = new ObservableCollection<string>();

        public string SolutionPath
        {
            get { return _solutionPath; }
            set { _solutionPath = value; Raise(); }
        }
        private string _solutionPath;

        public string ExportFolder
        {
            get { return _exportFolder; }
            set { _exportFolder = value; Raise(); }
        }
        private string _exportFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TwinCAT_Module_Export");

        public string BaseTreePath { get { return _baseTreePath; } set { _baseTreePath = value; Raise(); } }
        private string _baseTreePath = @"TIID^Device 1 (EtherCAT)^-KF002 (EK1101)^-KF005 (EK1122)";

        public bool SaveMappings { get { return _saveMappings; } set { _saveMappings = value; Raise(); } }
        private bool _saveMappings = true;

        public bool IncludeNcAxes { get { return _includeNcAxes; } set { _includeNcAxes = value; Raise(); } }
        private bool _includeNcAxes = true;

        public SelectionTreeItem SelectedNode { get; set; }

        public DteHost DteHost { get { return _dteHost; } set { _dteHost = value; Raise(); } }
        private DteHost _dteHost = DteHost.XaeShell;

        public DteHost[] DteHosts { get; } = new[] { DteHost.XaeShell, DteHost.VisualStudio2017 };

        // ---------- Status & Busy ----------
        private string _status = "";
        public string Status { get { return _status; } set { _status = value; Raise(); } }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; Raise(); } }

        private string _busyText;
        public string BusyText { get { return _busyText; } set { _busyText = value; Raise(); } }

        // merkt sich, welche Solution über den langlebigen Service offen ist
        private string _openedSolutionPath;

        public ICommand BrowseSolutionCommand
        {
            get
            {
                return new Relay(() =>
                {
                    var startDir = !string.IsNullOrWhiteSpace(SolutionPath)
                                   ? (File.Exists(SolutionPath) ? Path.GetDirectoryName(SolutionPath) : SolutionPath)
                                   : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    var f = Dialogs.OpenFile("Solution (*.sln)|*.sln", startDir, "TwinCAT/VS Solution auswählen");
                    if (!string.IsNullOrEmpty(f))
                    {
                        SolutionPath = f;
                        AddRecentSolution(f);
                        SetStatus("Solution gesetzt: " + f);
                    }
                    else
                    {
                        SetStatus("Keine Solution gewählt.");
                    }
                });
            }
        }

        public ICommand BrowseExportFolderCommand
        {
            get
            {
                return new Relay(() =>
                {
                    var f = Dialogs.SelectFolder("Exportordner wählen", ExportFolder);
                    if (!string.IsNullOrEmpty(f))
                    {
                        ExportFolder = f;
                        SetStatus("Exportordner: " + f);
                    }
                });
            }
        }

        public ICommand LoadTreeCommand
        {
            get
            {
                return new RelayAsync(async () =>
                {
                    if (string.IsNullOrWhiteSpace(SolutionPath) || !File.Exists(SolutionPath))
                    {
                        SetStatus("Bitte zuerst eine gültige Solution (.sln) wählen.");
                        return;
                    }

                    // Busy sofort anzeigen
                    BusyText = "Solution wird geladen …";
                    IsBusy = true;
                    await YieldToUi(); // UI rendern lassen, Overlay wird sichtbar

                    try
                    {
                        Tree.Clear();
                        Directory.CreateDirectory(ExportFolder);

                        var tc = TcAutomationService.GetOrCreate(DteHost, true);
                        if (tc.SystemManager == null || !IsSameSolutionOpen(SolutionPath))
                        {
                            SetStatus("Öffne Solution …");
                            tc.OpenSolution(SolutionPath);
                            _openedSolutionPath = SolutionPath;
                            SetStatus("Solution geöffnet.");
                        }

                        ITcSmTreeItem baseItem = null;

                        BusyText = "Basisknoten prüfen …";
                        await YieldToUi();

                        // 1) eingestellter Basispfad
                        try { baseItem = tc.Lookup(BaseTreePath); }
                        catch (COMException) { baseItem = null; }

                        // 2) Auto-Detect unter TIID
                        if (baseItem == null)
                        {
                            SetStatus("Suche Basisknoten (Auto‑Detect) …");
                            BusyText = "Suche I/O-Basisknoten …";
                            await YieldToUi();

                            var detected = TryAutoDetectIoModule(tc);
                            if (!string.IsNullOrEmpty(detected))
                            {
                                BaseTreePath = detected;
                                baseItem = tc.Lookup(detected);
                                SetStatus("Basisknoten gefunden: " + detected);
                            }
                            else
                            {
                                SetStatus("Basisknoten nicht gefunden. Prüfe Pfad/Hardwarebaum.");
                            }
                        }

                        if (baseItem != null)
                        {
                            BusyText = "Baumaufbau I/O …";
                            await YieldToUi();

                            Tree.Add(ProjectTreeService.BuildTreeFrom(baseItem, baseItem.Name, true, 10));
                        }

                        // NC/Motion
                        if (IncludeNcAxes)
                        {
                            BusyText = "Suche NC/Motion …";
                            await YieldToUi();

                            var added = 0;

                            foreach (var candidate in new[] { "TMC", "TM", "TIRC", "TMC^Axes", "Motion", "NC" })
                            {
                                try
                                {
                                    var root = tc.Lookup(candidate);
                                    if (root != null && !Tree.Any(x => x.Path == root.PathName))
                                    {
                                        Tree.Add(ProjectTreeService.BuildTreeFrom(root, "[NC/Motion] " + root.Name, true, 8));
                                        added++;
                                    }
                                }
                                catch { }
                            }

                            var discovered = NcDiscoveryService.DiscoverNcRoots(tc.SystemManager, 8);
                            foreach (var r in discovered)
                            {
                                if (!Tree.Any(x => x.Path == r.Path))
                                {
                                    Tree.Add(r);
                                    added++;
                                }
                            }

                            SetStatus(added > 0 ? "NC/Motion eingefügt." : "Hinweis: NC/Motion nicht gefunden.");
                        }

                        if (Tree.Count == 0)
                            SetStatus("Kein auswählbarer Knoten gefunden.");
                        else
                            SetStatus("Baum geladen.");
                    }
                    catch (Exception ex)
                    {
                        SetStatus("FEHLER beim Laden: " + ex.Message);
                    }
                    finally
                    {
                        IsBusy = false;
                        BusyText = null;
                    }
                });
            }
        }

        public ICommand CheckAllCommand
        {
            get { return new Relay(() => { foreach (var r in Tree) ProjectTreeService.CheckAll(r, true); SetStatus("Alle Knoten markiert."); }); }
        }

        public ICommand ExportCommand
        {
            get
            {
                return new RelayAsync(async () =>
                {
                    if (string.IsNullOrWhiteSpace(SolutionPath) || !File.Exists(SolutionPath))
                    {
                        SetStatus("Keine gültige Solution gewählt.");
                        return;
                    }
                    if (Tree.Count == 0)
                    {
                        SetStatus("Kein Baum geladen. Bitte 'Baum laden' ausführen.");
                        return;
                    }

                    BusyText = "Export läuft …";
                    IsBusy = true;
                    await YieldToUi();

                    try
                    {
                        var tc = TcAutomationService.GetOrCreate(DteHost, true);
                        if (tc.SystemManager == null || !IsSameSolutionOpen(SolutionPath))
                        {
                            SetStatus("Öffne Solution …");
                            tc.OpenSolution(SolutionPath);
                            _openedSolutionPath = SolutionPath;
                        }

                        int idx = 0;
                        int exported = 0;

                        foreach (var root in Tree)
                        {
                            foreach (var node in ProjectTreeService.EnumerateChecked(root))
                            {
                                var item = tc.Lookup(node.Path);
                                var safeName = Sanitize(item.Name) + "_" + (++idx);
                                var xmlFile = Path.Combine(ExportFolder, safeName + ".xml");
                                var tceFile = Path.Combine(ExportFolder, safeName + ".childexport");

                                var xml = tc.ProduceItemXml(item);
                                File.WriteAllText(xmlFile, xml);
                                SetStatus("XML exportiert: " + safeName + ".xml");

                                try { tc.ExportChild(item, tceFile); SetStatus("Child exportiert: " + safeName + ".childexport"); } catch { }
                                exported++;
                            }
                        }

                        if (SaveMappings)
                        {
                            SetStatus("Erzeuge Mappings …");
                            var mapping = tc.ProduceMappingInfo();
                            File.WriteAllText(Path.Combine(ExportFolder, "_Mappings.xml"), mapping);
                            SetStatus("Mappings gespeichert: _Mappings.xml");
                        }

                        SetStatus(string.Format("Export abgeschlossen. {0} Elemente exportiert.", exported));
                    }
                    catch (Exception ex)
                    {
                        SetStatus("Exportfehler: " + ex.Message);
                    }
                    finally
                    {
                        IsBusy = false;
                        BusyText = null;
                    }
                });
            }
        }

        // ---------- Helpers ----------
        private async Task YieldToUi()
        {
            try
            {
                var app = Application.Current;
                if (app != null && app.Dispatcher != null)
                {
                    await app.Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                }
                else
                {
                    await Task.Delay(50);
                }
            }
            catch
            {
                // not fatal
            }
        }

        private void SetStatus(string message)
        {
            var ts = DateTime.Now.ToString("HH:mm:ss");
            if (string.IsNullOrEmpty(Status)) Status = "[" + ts + "] " + message;
            else Status += Environment.NewLine + "[" + ts + "] " + message;
            Raise("Status");
        }

        private bool IsSameSolutionOpen(string path)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(_openedSolutionPath)
                    && string.Equals(Path.GetFullPath(_openedSolutionPath),
                                     Path.GetFullPath(path),
                                     StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        public void AddRecentSolution(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            for (int i = RecentSolutions.Count - 1; i >= 0; i--)
                if (string.Equals(RecentSolutions[i], path, StringComparison.OrdinalIgnoreCase))
                    RecentSolutions.RemoveAt(i);

            RecentSolutions.Insert(0, path);
            while (RecentSolutions.Count > 8) RecentSolutions.RemoveAt(RecentSolutions.Count - 1);
        }

        private string TryAutoDetectIoModule(TcAutomationService tc)
        {
            ITcSmTreeItem ioRoot = null;
            try { ioRoot = tc.Lookup("TIID"); } catch { return null; }
            if (ioRoot == null) return null;

            var q = new System.Collections.Generic.Queue<ITcSmTreeItem>();
            q.Enqueue(ioRoot);

            string fallbackEk1101 = null;

            while (q.Count > 0)
            {
                var n = q.Dequeue();
                string nm;
                try { nm = n.Name ?? string.Empty; }
                catch { nm = string.Empty; }

                if (nm.Contains("EK1122"))
                    return n.PathName;
                if (nm.Contains("EK1101") && fallbackEk1101 == null)
                    fallbackEk1101 = n.PathName;

                int cc = 0;
                try { cc = n.ChildCount; } catch (COMException) { cc = 0; }
                for (int i = 1; i <= cc; i++)
                {
                    ITcSmTreeItem c = null;
                    try { c = n.Child[i]; } catch { }
                    if (c != null) q.Enqueue(c);
                }
            }
            return fallbackEk1101;
        }

        private static string Sanitize(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name;
        }

        public class Relay : ICommand
        {
            private readonly Action _a;
            public Relay(Action a) { _a = a; }
            public bool CanExecute(object parameter) { return true; }
            public void Execute(object parameter) { _a(); }
            public event EventHandler CanExecuteChanged { add { } remove { } }
        }

        public class RelayAsync : ICommand
        {
            private readonly Func<Task> _run;
            public RelayAsync(Func<Task> run) { _run = run; }
            public bool CanExecute(object parameter) { return true; }
            public async void Execute(object parameter) { await _run(); }
            public event EventHandler CanExecuteChanged { add { } remove { } }
        }
    }
}