// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.Shortcuts.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Keyboard shortcuts and command bindings.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.Shortcuts.cs
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
        private void FontSizeEditableTextBox_KeyDown(object? sender, SWI.KeyEventArgs e)
        {
            if (e.Key == SWI.Key.Enter)
            {
                ApplyFontSizeFromCombo(applyParagraphWhenEmpty: true, focusEditorAfter: true);
                e.Handled = true;
            }
        }

        private void FontSizeCombo_KeyDown(object? sender, SWI.KeyEventArgs e)
        {
            if (e.Key == SWI.Key.Enter)
            {
                ApplyFontSizeFromCombo(applyParagraphWhenEmpty: true, focusEditorAfter: true);
                e.Handled = true;
            }
        }
    }
}
