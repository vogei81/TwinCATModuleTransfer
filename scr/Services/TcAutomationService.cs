using System;
using System.IO;
using EnvDTE;
using TCatSysManagerLib;

namespace TwinCATModuleTransfer.Services
{
    public enum DteHost { XaeShell, VisualStudio2017 }

    /// <summary>
    /// Verwaltet eine langlebige DTE/ITcSysManager-Instanz:
    /// - Startet XAE/VS im Hintergrund (unsichtbar).
    /// - Hält Solution offen, bis Export abgeschlossen oder App beendet ist.
    /// - Stellt Lookup/Export/Mapping-Methoden bereit.
    /// </summary>
    public class TcAutomationService : IDisposable
    {
        // ---------- Singleton-ähnliche Verwaltung ----------
        private static TcAutomationService _current;
        public static TcAutomationService Current { get { return _current; } }

        /// <summary>
        /// Liefert eine existierende Instanz oder erstellt eine neue.
        /// </summary>
        public static TcAutomationService GetOrCreate(DteHost host, bool hideWindow)
        {
            if (_current == null)
            {
                _current = new TcAutomationService();
                _current.Launch(host, hideWindow);
            }
            return _current;
        }

        /// <summary>
        /// Schließt die aktuelle Instanz (Solution schließen, DTE quitten).
        /// </summary>
        public static void ShutdownCurrent()
        {
            if (_current != null)
            {
                _current.Shutdown();
                _current = null;
            }
        }

        // ---------- Instanz ----------
        private DTE _dte;
        private ITcSysManager _sysMan;
        private Solution _solution;
        private DteHost _host;
        private bool _launched;

        private TcAutomationService() { }

        public void Launch(DteHost host, bool hideWindow)
        {
            if (_launched) return;

            Utils.ComMessageFilter.Register();
            _host = host;

            var progId = (host == DteHost.XaeShell) ? "TcXaeShell.DTE.15.0" : "VisualStudio.DTE.15.0";
            var t = Type.GetTypeFromProgID(progId, true);
            _dte = (DTE)Activator.CreateInstance(t);

            // "Unsichtbar" & leise
            _dte.SuppressUI = true;
            try { _dte.MainWindow.Visible = !hideWindow; } catch { }
            try { _dte.MainWindow.WindowState = vsWindowState.vsWindowStateMinimize; } catch { }

            _launched = true;
        }

        public void OpenSolution(string slnPath)
        {
            if (!_launched) throw new InvalidOperationException("TcAutomationService nicht gestartet.");

            _solution = _dte.Solution;
            _solution.Open(slnPath);

            // Erstes Projekt der Solution nehmen (TwinCAT-Projekt)
            var project = _solution.Projects.Item(1);
            _sysMan = (ITcSysManager)project.Object;
        }

        public ITcSysManager SystemManager { get { return _sysMan; } }

        public ITcSmTreeItem Lookup(string path) { return _sysMan.LookupTreeItem(path); }

        public string ProduceItemXml(ITcSmTreeItem item) { return item.ProduceXml(false); } // bool-Flag statt int

        public void ConsumeItemXml(ITcSmTreeItem item, string xml) { item.ConsumeXml(xml); }

        public void ExportChild(ITcSmTreeItem child, string filePath)
        {
            var parent = child.Parent;
            parent.ExportChild(child.Name, filePath);
        }

        public void ImportChildInto(ITcSmTreeItem parent, string filePath)
        {
            parent.ImportChild(filePath);
        }

        public string ProduceMappingInfo()
        {
            var sysMan3 = _sysMan as ITcSysManager3;
            if (sysMan3 != null) return sysMan3.ProduceMappingInfo();

            dynamic dyn = _sysMan; // Runtime-Bindung für unterschiedliche Typlib-Versionen
            return dyn.ProduceMappingInfo();
        }

        public void ReplaceAllMappings(string newMappingXml)
        {
            try
            {
                var sysMan3 = _sysMan as ITcSysManager3;
                if (sysMan3 != null) sysMan3.ClearMappingInfo();

                dynamic dyn = _sysMan;   // „ConsumeMappingInfo“ je nach Interfaceversion
                dyn.ConsumeMappingInfo(newMappingXml);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("ConsumeMappingInfo/Mapping-API nicht verfügbar. Prüfe TwinCAT XAE/Type Library.", ex);
            }
        }

        public void ActivateAndRestart()
        {
            _sysMan.ActivateConfiguration();
            _sysMan.StartRestartTwinCAT();
        }

        /// <summary>
        /// Schließt Solution & DTE bewusst (z. B. am Programmende).
        /// </summary>
        public void Shutdown()
        {
            try { if (_solution != null) _solution.Close(true); } catch { }
            try { if (_dte != null) _dte.Quit(); } catch { }
            _solution = null;
            _dte = null;
            _sysMan = null;
            _launched = false;
            Utils.ComMessageFilter.Revoke();
        }

        public void Dispose()
        {
            // hier NICHT automatisch schließen – Lebensdauer wird extern gesteuert (ShutdownCurrent)
        }
    }
}