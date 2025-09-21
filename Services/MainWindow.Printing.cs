// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.Printing.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Print pipeline for FlowDocument (A4 pagination).
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.Printing.cs
// Purpose: C# implementation file in the editor application.
// Context: May interact with ThemeManager, Tabs, or file I/O
// Notes: Do not alter public API or behavior.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Input;
using ControlzEx.Theming;
using System.Printing;
using System.IO.Compression;
// ---- Disambiguation aliases (avoid CS0104) ----

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace TxtOrganizer
{
    /// <summary>MainWindow — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public partial class MainWindow
    {
        private void Print_Click(object sender, RoutedEventArgs e)
        {
            var doc = PrepareForA4(CloneDocument());
            var pd = new SWC.PrintDialog();
            if (pd.ShowDialog() == true)
            {
                pd.PrintDocument(((IDocumentPaginatorSource)doc).DocumentPaginator, "AnyFile Editor");
                UpdateStatus("Printed.");
            }
        }
    }
}
