// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/GlobalUsings.Aliases.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Project-wide alias usings to avoid type ambiguities.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/GlobalUsings.Aliases.cs
// Purpose: C# implementation file in the editor application.
// Context: May interact with ThemeManager, Tabs, or file I/O
// Notes: Do not alter public API or behavior.
// Centralized alias directives for the whole project.
// Requires C# 10+ (global using) – OK on net8.0-windows.

global using WF     = System.Windows.Forms;    // FolderBrowserDialog, ColorDialog
global using MWin32 = Microsoft.Win32;         // WPF file dialogs
global using W      = System.Windows;          // MessageBox, Clipboard, etc.  <-- added to fix CS0246
global using SW     = System.Windows;          // (kept for existing code paths)
global using SWC    = System.Windows.Controls; // Controls
global using SWD    = System.Windows.Documents;// FlowDocument, Paragraph, Run, Table
global using SMB    = System.Windows.Media;    // Brushes, Color, Typeface
global using SWI    = System.Windows.Input;    // Key, RoutedCommand, ICommand
global using WP     = DocumentFormat.OpenXml.Wordprocessing;
