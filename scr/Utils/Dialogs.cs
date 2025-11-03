using System;
using System.IO;
using System.Windows;

// Alias-Namen, damit es keine Konflikte zwischen WPF und WinForms gibt:
using Forms = System.Windows.Forms;            // FolderBrowserDialog, DialogResult
using WpfApp = System.Windows.Application;     // WPF Application.Current.MainWindow

namespace TwinCATModuleTransfer.Utils
{
    public static class Dialogs
    {
        /// <summary>
        /// Öffnet einen Datei-Dialog (z. B. für .sln-Dateien).
        /// </summary>
        /// <param name="filter">z. B. "Solution (*.sln)|*.sln"</param>
        /// <param name="initialDir">Startverzeichnis oder bereits ausgewählter Pfad</param>
        /// <param name="title">Titelzeile des Dialogs</param>
        /// <returns>Vollständiger Dateipfad oder null</returns>
        public static string OpenFile(string filter, string initialDir = null, string title = "Solution auswählen")
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = string.IsNullOrWhiteSpace(filter) ? "Alle Dateien (*.*)|*.*" : filter,
                Title = string.IsNullOrWhiteSpace(title) ? "Datei auswählen" : title,
                Multiselect = false,
                CheckFileExists = true
            };

            // InitialDirectory (sofern sinnvoll)
            try
            {
                if (!string.IsNullOrWhiteSpace(initialDir))
                {
                    string dir = null;

                    if (Directory.Exists(initialDir))
                        dir = initialDir;
                    else if (File.Exists(initialDir))
                        dir = Path.GetDirectoryName(initialDir);

                    if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                        ofd.InitialDirectory = dir;
                }
            }
            catch
            {
                // Ignorieren – Dialog öffnet trotzdem
            }

            // Owner setzen (WPF), damit der Dialog im Vordergrund bleibt
            var owner = (WpfApp.Current != null) ? WpfApp.Current.MainWindow : null;
            bool? result = (owner != null) ? ofd.ShowDialog(owner) : ofd.ShowDialog();

            return result == true ? ofd.FileName : null;
        }

        /// <summary>
        /// Öffnet einen Ordner-Dialog (WinForms) für Export/Import-Ziele.
        /// </summary>
        /// <param name="description">Beschreibung im Dialog</param>
        /// <param name="initialDir">Vorwahl des Startordners</param>
        /// <returns>Ausgewählter Ordnerpfad oder null</returns>
        public static string SelectFolder(string description = "Ordner auswählen", string initialDir = null)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = string.IsNullOrWhiteSpace(description) ? "Ordner auswählen" : description;

                if (!string.IsNullOrWhiteSpace(initialDir) && Directory.Exists(initialDir))
                    dlg.SelectedPath = initialDir;

                var result = dlg.ShowDialog();
                return result == Forms.DialogResult.OK ? dlg.SelectedPath : null;
            }
        }
    }
}