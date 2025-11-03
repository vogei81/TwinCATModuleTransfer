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
using static TwinCATModuleTransfer.Services.PackagingService;

namespace TwinCATModuleTransfer.ViewModels
{
    public class ImportViewModel : BaseViewModel
    {
        public string SolutionPath { get { return _solutionPath; } set { _solutionPath = value; Raise(); } }
        private string _solutionPath;

        public string PackagePath { get { return _packagePath; } set { _packagePath = value; Raise(); } }
        private string _packagePath;

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

        // Busy & Status
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

        public ICommand BrowsePackageCommand
        {
            get
            {
                return new Relay(() =>
                {
                    var startDir = !string.IsNullOrWhiteSpace(PackagePath)
                       ? (File.Exists(PackagePath) ? Path.GetDirectoryName(PackagePath) : PackagePath)
                       : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    var f = Dialogs.OpenFile("TwinCAT Modul-Paket (*.tcmodpkg)|*.tcmodpkg", startDir, "Paket auswählen");
                    if (f != null)
                    {
                        PackagePath = f;
                        SetStatus("Paket gewählt: " + f);
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
                    if (string.IsNullOrWhiteSpace(PackagePath) || !File.Exists(PackagePath))
                    {
                        SetStatus("Bitte gültige Paketdatei (*.tcmodpkg) wählen.");
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

                        // Paket entpacken
                        SetStatus("Paket wird geöffnet …");
                        var tempWork = Path.Combine(Path.GetTempPath(), "TCModImport_" + Guid.NewGuid().ToString("N"));
                        var pkg = OpenPackage(PackagePath, tempWork);

                        // Ziel-Parent
                        var target = tc.Lookup(SelectedTargetParent.Path);
                        if (target == null)
                            throw new InvalidOperationException("Ziel‑Parent nicht gefunden: " + SelectedTargetParent.Path);

                        // In Manifest-Reihenfolge importieren
                        int imported = 0;
                        foreach (var it in pkg.Manifest.Items.OrderBy(i => i.Index))
                        {
                            // Child zuerst (wenn vorhanden), dann XML konsumieren
                            if (!string.IsNullOrEmpty(it.ChildFile))
                            {
                                var childAbs = pkg.ResolveItemPath(it.ChildFile);
                                if (File.Exists(childAbs))
                                {
                                    tc.ImportChildInto(target, childAbs);
                                    SetStatus("ImportChild: " + it.ChildFile);
                                }
                            }

                            var xmlAbs = pkg.ResolveItemPath(it.XmlFile);
                            if (File.Exists(xmlAbs))
                            {
                                var xml = File.ReadAllText(xmlAbs, Encoding.UTF8);
                                tc.ConsumeItemXml(target, xml);
                                SetStatus("ConsumeXml: " + it.XmlFile);
                            }

                            imported++;
                        }

                        if (imported == 0) SetStatus("Hinweis: Keine Items im Paket.");

                        // Mappings anwenden (optional)
                        if (pkg.Manifest.ContainsMappings && !string.IsNullOrWhiteSpace(pkg.MappingsPath) && File.Exists(pkg.MappingsPath))
                        {
                            SetStatus("Verlinkungen anwenden …");
                            var oldXml = File.ReadAllText(pkg.MappingsPath, Encoding.UTF8);
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
                            SetStatus("Hinweis: Keine Mappings im Paket.");
                        }

                        // Aktivieren
                        SetStatus("Aktiviere Konfiguration / TwinCAT Neustart …");
                        tc.ActivateAndRestart();
                        SetStatus("Aktivierung abgeschlossen.");

                        // Temp löschen
                        try { Directory.Delete(tempWork, true); } catch { }
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

        // Helpers
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
