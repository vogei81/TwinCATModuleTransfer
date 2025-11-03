using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using TwinCATModuleTransfer.Models;
using TwinCATModuleTransfer.Services;
using TwinCATModuleTransfer.Utils;

namespace TwinCATModuleTransfer.ViewModels
{
    public class ImportViewModel : BaseViewModel
    {
        public string SolutionPath { get { return _solutionPath; } set { _solutionPath = value; Raise(); } }
        private string _solutionPath;

        public string ModuleFolder { get { return _moduleFolder; } set { _moduleFolder = value; Raise(); } }
        private string _moduleFolder;

        public string OldPlcBasePath { get { return _old; } set { _old = value; Raise(); } }
        private string _old;

        public string NewPlcBasePath { get { return _new; } set { _new = value; Raise(); } }
        private string _new;

        public string Log { get { return _log; } set { _log = value; Raise(); } }
        private string _log;

        public DteHost DteHost { get { return _dteHost; } set { _dteHost = value; Raise(); } }
        private DteHost _dteHost = DteHost.XaeShell;

        public DteHost[] DteHosts { get; } = new[] { DteHost.XaeShell, DteHost.VisualStudio2017 };

        public ObservableCollection<SelectionTreeItem> Tree { get; private set; } =
            new ObservableCollection<SelectionTreeItem>();

        public SelectionTreeItem SelectedTargetParent { get; set; }

        // ---------- Status & Busy ----------
        private string _status = "";
        public string Status { get { return _status; } set { _status = value; Raise(); } }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; Raise(); } }

        private string _busyText;
        public string BusyText { get { return _busyText; } set { _busyText = value; Raise(); } }

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
                    if (f != null)
                    {
                        SolutionPath = f;
                        SetStatus("Solution gesetzt: " + f);
                    }
                });
            }
        }

        public ICommand BrowseModuleFolderCommand
        {
            get
            {
                return new Relay(() =>
                {
                    var f = Dialogs.SelectFolder("Export-/Modulordner wählen", ModuleFolder);
                    if (f != null)
                    {
                        ModuleFolder = f;
                        SetStatus("Modulordner: " + f);
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
                        SetStatus("Bitte gültige Solution (.sln) wählen.");
                        return;
                    }

                    BusyText = "Solution wird geladen …";
                    IsBusy = true;
                    await YieldToUi();

                    var sb = new StringBuilder();
                    Tree.Clear();

                    try
                    {
                        var tc = TcAutomationService.GetOrCreate(DteHost, true);

                        if (tc.SystemManager == null || !IsSameSolutionOpen(SolutionPath))
                        {
                            SetStatus("Öffne Solution …");
                            tc.OpenSolution(SolutionPath);
                            _openedSolutionPath = SolutionPath;
                            SetStatus("Solution geöffnet.");
                        }

                        foreach (var top in new[] { "TIID", "TIPC", "TIRC", "TM", "TMC", "NC", "Motion" })
                        {
                            try
                            {
                                BusyText = "Lade Knoten: " + top + " …";
                                await YieldToUi();

                                var item = tc.Lookup(top);
                                if (item != null)
                                {
                                    Tree.Add(Services.ProjectTreeService.BuildTreeFrom(item, item.Name, true, 6));
                                    SetStatus("Knoten geladen: " + top);
                                }
                            }
                            catch
                            {
                                // ignore
                            }
                        }

                        if (Tree.Count == 0) SetStatus("Keine Top-Level-Knoten gefunden.");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("FEHLER beim Laden des Baums: " + ex.Message);
                        SetStatus("FEHLER beim Laden: " + ex.Message);
                    }
                    finally
                    {
                        Log = sb.ToString();
                        IsBusy = false;
                        BusyText = null;
                    }
                });
            }
        }

        public ICommand ImportCommand
        {
            get
            {
                return new RelayAsync(async () =>
                {
                    if (SelectedTargetParent == null)
                    {
                        SetStatus("Bitte Ziel‑Parent in der Baumansicht wählen.");
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(ModuleFolder) || !Directory.Exists(ModuleFolder))
                    {
                        SetStatus("Bitte gültigen Modul-/Exportordner wählen.");
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(SolutionPath) || !File.Exists(SolutionPath))
                    {
                        SetStatus("Bitte gültige Solution (.sln) wählen.");
                        return;
                    }

                    BusyText = "Import läuft …";
                    IsBusy = true;
                    await YieldToUi();

                    var sb = new StringBuilder();

                    try
                    {
                        var tc = TcAutomationService.GetOrCreate(DteHost, true);

                        if (tc.SystemManager == null || !IsSameSolutionOpen(SolutionPath))
                        {
                            SetStatus("Öffne Solution …");
                            tc.OpenSolution(SolutionPath);
                            _openedSolutionPath = SolutionPath;
                            SetStatus("Solution geöffnet.");
                        }

                        var target = tc.Lookup(SelectedTargetParent.Path);
                        if (target == null)
                            throw new InvalidOperationException("Ziel‑Parent nicht gefunden: " + SelectedTargetParent.Path);

                        // 1) Module importieren
                        var files = Directory.GetFiles(ModuleFolder, "*.childexport").OrderBy(f => f)
                                   .Concat(Directory.GetFiles(ModuleFolder, "*.xml").OrderBy(f => f))
                                   .ToArray();

                        int imported = 0;
                        foreach (var f in files)
                        {
                            var name = Path.GetFileName(f);
                            if (name.Equals("_Mappings.xml", StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (f.EndsWith(".childexport", StringComparison.OrdinalIgnoreCase))
                            {
                                tc.ImportChildInto(target, f);
                                SetStatus("ImportChild: " + name);
                                imported++;
                            }
                            else
                            {
                                var xml = File.ReadAllText(f);
                                tc.ConsumeItemXml(target, xml);
                                SetStatus("ConsumeXml: " + name);
                                imported++;
                            }
                        }

                        if (imported == 0) SetStatus("Hinweis: Keine Modul-Dateien im Ordner gefunden.");

                        // 2) Mappings
                        var mappingPath = Path.Combine(ModuleFolder, "_Mappings.xml");
                        if (File.Exists(mappingPath))
                        {
                            SetStatus("Verlinkungen anwenden …");
                            var oldXml = File.ReadAllText(mappingPath);
                            var opts = new MappingRewriteOptions
                            {
                                OldPlcBasePath = OldPlcBasePath,
                                NewPlcBasePath = NewPlcBasePath
                            };
                            var newXml = Services.MappingService.RewriteMappingXml(oldXml, opts);
                            tc.ReplaceAllMappings(newXml);
                            SetStatus("Mappings angewendet.");
                        }
                        else
                        {
                            SetStatus("Hinweis: _Mappings.xml nicht gefunden – keine Verlinkungen angewendet.");
                        }

                        // 3) Aktivieren
                        SetStatus("Aktiviere Konfiguration / TwinCAT Neustart …");
                        tc.ActivateAndRestart();
                        SetStatus("Aktivierung abgeschlossen.");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("FEHLER beim Import: " + ex.Message);
                        SetStatus("FEHLER beim Import: " + ex.Message);
                    }
                    finally
                    {
                        Log = sb.ToString();
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
            catch { }
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