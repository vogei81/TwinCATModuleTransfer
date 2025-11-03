using System;
using System.Windows;

namespace TwinCATModuleTransfer
{
    public partial class App : Application
    {
        [STAThread]
        public static void Main()
        {
            var app = new App();
            app.InitializeComponent(); // keine StartupUri gesetzt
            app.Run(new MainWindow()); // root Window hostet UserControlsprotected override void OnExit(ExitEventArgs e)
        }

            protected override void OnExit(ExitEventArgs e)
            {
                base.OnExit(e);
                TwinCATModuleTransfer.Services.TcAutomationService.ShutdownCurrent();
            }
    }
}