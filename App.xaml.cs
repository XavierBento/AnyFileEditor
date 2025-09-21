// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : App.xaml.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Application entry points and lifecycle events for the WPF app.
// ============================================================================
// File: AnyFileEditor_fixed_all/App.xaml.cs
// Purpose: Code-behind for a XAML view.
// Context: Event handlers, UI wiring
// Notes: Keep behavior unchanged; annotate intent/why.
using System;

namespace TxtOrganizer
{
    // Fully-qualify WPF types to avoid clash with WinForms.Application
    /// <summary>App — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public partial class App : System.Windows.Application
    {
        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                System.Windows.MessageBox.Show(e.Exception.ToString(), "Unhandled UI exception");
                e.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                System.Windows.MessageBox.Show(ex?.ToString() ?? "Unknown crash", "Unhandled exception");
            };
        }

        protected override void OnStartup(System.Windows.StartupEventArgs e)
        {
            base.OnStartup(e);
            var w = new MainWindow();
            this.MainWindow = w;
            w.Show();
        }
    }
}
