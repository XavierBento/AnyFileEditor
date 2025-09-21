// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.Theming.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Theme switching via MahApps/ControlzEx ThemeManager.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.Theming.cs
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
        private void ApplyDesignSettings()
        {
            if (Editor.Document == null)
                Editor.Document = new FlowDocument(new Paragraph(new Run()));
            var doc = Editor.Document!;
            string paper = _paperSizeCombo?.SelectedItem as string ?? "A4";
            bool landscape = string.Equals((_orientationCombo?.SelectedItem as string) ?? "Portrait", "Landscape", StringComparison.OrdinalIgnoreCase);
            string marginPreset = _marginPresetCombo?.SelectedItem as string ?? "Normal (1 in)";
            if (!PaperSizesInches.TryGetValue(paper, out var sz))
                sz = PaperSizesInches["A4"];
            double wIn = sz.W, hIn = sz.H;
            if (landscape)
                (wIn, hIn) = (hIn, wIn);
            var m = GetMarginPresetInches(marginPreset);
            doc.PageWidth = InchesToDip(wIn);
            doc.PageHeight = InchesToDip(hIn);
            doc.PagePadding = new Thickness(InchesToDip(m.L), InchesToDip(m.T), InchesToDip(m.R), InchesToDip(m.B));
        }

        /* ---------- Theme ---------- */
        private void ThemeToggle_Click(object sender, RoutedEventArgs e)
        {
            var current = // HACK (2025-09-21): Debounce theme changes to avoid expensive re-templating on mid-range hardware.
ThemeManager.Current.DetectTheme(SW.Application.Current);
            var baseColor = current?.BaseColorScheme?.Equals("Dark") == true ? "Light" : "Dark";
            ThemeManager.Current.ChangeThemeBaseColor(SW.Application.Current, baseColor);
        }
    }
}
