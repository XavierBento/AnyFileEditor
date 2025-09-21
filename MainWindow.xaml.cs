// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : MainWindow.xaml.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Main application window markup (MetroWindow) with tabbed editor.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.xaml.cs
// Purpose: Interaction logic and event handlers for the main window shell.
// Context: Tabbed editor management, Theme switching via MahApps/ControlzEx, File I/O (DOCX/RTF/ODT/TXT), Printing
// Notes: Avoid logic changes; comments only., Keep tab state and UI in sync.
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

// Invariant: the backing tab list and the visual TabControl must remain index-aligned;
// selection, persistence, and close/add operations assume parity.



namespace TxtOrganizer
{
    /// <summary>MainWindow — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public partial class MainWindow : MahApps.Metro.Controls.MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += (_, __) =>
            {
                EnsureDesignTab();
                EnsureEditorTabStrip();
                HideCurrentFileIndicator();
            };
            InitFontFamilyCombo();
            InitFontSizeCombo();
            UpdateStatus("Choose a working folder to begin.");
        }

        // Represents a font option in the font-family dropdown
        private sealed class FontChoice
        {
            public string Name { get; set; } = string.Empty;
            public System.Windows.Media.FontFamily? Family { get; set; }
            // Special sentinel used by the UI to show the "Add font…" entry
            public bool IsAddItem { get; set; }

            /// <summary>ToString — see remarks for intent and side effects.</summary>
/// <remarks>Non-functional docs only; behavior unchanged.</remarks>

            public override string ToString() => Name;
        }

        // Entry used for the file dropdown
        private sealed class FileEntry
        {
            public string Name { get; }
            public string FullPath { get; }

            /// <summary>FileEntry — see remarks for intent and side effects.</summary>
/// <param name="name">See implementation for usage and contracts.</param>
/// <param name="fullPath">See implementation for usage and contracts.</param>
/// <remarks>Non-functional docs only; behavior unchanged.</remarks>

            public FileEntry(string name, string fullPath)
            {
                Name = name;
                FullPath = fullPath;
            }

            /// <summary>ToString — see remarks for intent and side effects.</summary>
/// <remarks>Non-functional docs only; behavior unchanged.</remarks>

            public override string ToString() => Name;
        }

        private string? _rootFolder;
        private string? _currentFilePath;
        private string? _docxPreviewFolder; // temp directory for html + assets
        // Font-size UX
        private SWC.TextBox? _fontSizeEditableTextBox;
        private double? _lastSizeAppliedFromCombo;
        // Font-family UX
        private readonly List<FontChoice> _fontChoices = new();
        private static string DefaultDocsFolder => Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        // ---------- Design tab state (paper size / orientation / margins) ----------
        private SWC.ComboBox? _paperSizeCombo;
        private SWC.ComboBox? _orientationCombo;
        private SWC.ComboBox? _marginPresetCombo;
        // inches for common paper sizes (width, height)
        private static readonly Dictionary<string, (double W, double H)> PaperSizesInches = new()
        {
            {
                "A4",
                (8.27, 11.69)
            },
            {
                "A3",
                (11.69, 16.54)
            },
            {
                "Letter",
                (8.5, 11.0)
            },
            {
                "Legal",
                (8.5, 14.0)
            },
            {
                "A5",
                (5.83, 8.27)
            },
        };
        // ---------- File editor tabs (multi-document) ----------
        private SWC.TabControl? _editorTabs;
        private SWC.TabItem? _plusTab;
        private sealed class DocTab
        {
            public FlowDocument Document { get; set; } = new FlowDocument(new Paragraph(new Run()));
            public string? FullPath { get; set; }

            /// <summary>ToString — see remarks for intent and side effects.</summary>
/// <remarks>Non-functional docs only; behavior unchanged.</remarks>

            public override string ToString() => System.IO.Path.GetFileName(FullPath ?? "Untitled");
        }
    }
}
