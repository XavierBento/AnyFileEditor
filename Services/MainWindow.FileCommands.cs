// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.FileCommands.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: File open/save/export commands for the editor.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.FileCommands.cs
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
        // ========= Tab persistence helpers =========
        private static FlowDocument CloneFlowDocument(FlowDocument source)
        {
            var clone = new FlowDocument();
            var srcRange = new TextRange(source.ContentStart, source.ContentEnd);
            using var ms = new MemoryStream();
            srcRange.Save(ms, System.Windows.DataFormats.XamlPackage);
            ms.Position = 0;
            var dstRange = new TextRange(clone.ContentStart, clone.ContentEnd);
            dstRange.Load(ms, System.Windows.DataFormats.XamlPackage);
            return clone;
        }

        private void SaveEditorDocIntoTab(SWC.TabItem? tabItem)
        {
            if (tabItem == null || Editor?.Document == null)
                return;
            if (tabItem.Tag is DocTab meta)
            {
                meta.Document = CloneFlowDocument(Editor.Document);
            }
        }

        private void BindLoadedDocToActiveTab(string filePath, FlowDocument doc)
        {
            if (_editorTabs == null)
                return;
            if (ReferenceEquals(_editorTabs.SelectedItem, _plusTab) || _editorTabs.SelectedItem == null)
            {
                var newTab = CreateBlankTab();
                int idx = Math.Max(0, _editorTabs.Items.Count - 1);
                _editorTabs.Items.Insert(idx, newTab);
                _editorTabs.SelectedItem = newTab;
            }

            var ti = GetActiveTab();
            if (ti == null)
                return;
            if (ti.Tag is not DocTab meta)
            {
                meta = new DocTab();
                ti.Tag = meta;
            }

            meta.Document = doc;
            meta.FullPath = filePath;
            ti.Header = System.IO.Path.GetFileName(filePath);
            Editor.Document = doc;
            _currentFilePath = filePath;
        }

        private string GetInitialDir() => string.IsNullOrEmpty(_rootFolder) ? DefaultDocsFolder : _rootFolder!;
        private void OpenFileInNewTab()
        {
            var ofd = new MWin32.OpenFileDialog
            {
                Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Word Document (*.docx)|*.docx|All Files|*.*",
                InitialDirectory = GetInitialDir()
            };
            if (ofd.ShowDialog() != true)
                return;
            if (_editorTabs != null)
            {
                var newTab = CreateBlankTab();
                int insertIndex = Math.Max(0, _editorTabs.Items.Count - 1);
                _editorTabs.Items.Insert(insertIndex, newTab);
                _editorTabs.SelectedItem = newTab;
            }

            string path = ofd.FileName;
            if (path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
                LoadRtfFile(path);
            else if (path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                LoadDocxFile(path);
            else
                LoadTxtFile(path);
        }

        /* ---------- Loaders ---------- */
        private void LoadRtfFile(string path)
        {
            // Load into an offscreen WPF RichTextBox to avoid mutating the shared editor doc
            var rtb = new SWC.RichTextBox();
            using (var fs = File.OpenRead(path))
            {
                var r = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                r.Load(fs, System.Windows.DataFormats.Rtf);
            }

            var flow = CloneFlowDocument(rtb.Document);
            ShowEditor();
            BindLoadedDocToActiveTab(path, flow);
            CurrentFileLabel.Text = Path.GetFileName(path);
            UpdateStatus($"Loaded {Path.GetFileName(path)}");
        }

        private void LoadTxtFile(string path)
        {
            string text = File.ReadAllText(path, Encoding.UTF8);
            var flow = new FlowDocument(new Paragraph(new Run(text)));
            ShowEditor();
            BindLoadedDocToActiveTab(path, flow);
            CurrentFileLabel.Text = Path.GetFileName(path);
            UpdateStatus($"Loaded {Path.GetFileName(path)} (plain text)");
        }

        /* ---------- Save As… (multi-format) ---------- */
        private void SaveAsMulti()
        {
            if (string.IsNullOrEmpty(_rootFolder))
            {
                SW.MessageBox.Show("Choose a working folder first.");
                return;
            }

            string baseName = string.IsNullOrEmpty(_currentFilePath) ? "Document" : Path.GetFileNameWithoutExtension(_currentFilePath);
            string currentExt = string.IsNullOrEmpty(_currentFilePath) ? ".rtf" : Path.GetExtension(_currentFilePath).ToLowerInvariant();
            if (currentExt is not (".rtf" or ".txt" or ".docx" or ".odt"))
                currentExt = ".rtf";
            var sfd = new MWin32.SaveFileDialog
            {
                Title = "Save As",
                Filter = "Rich Text (*.rtf)|*.rtf|" + "Plain Text (*.txt)|*.txt|" + "Word Document (*.docx)|*.docx|" + "OpenDocument Text (*.odt)|*.odt",
                DefaultExt = currentExt,
                AddExtension = true,
                InitialDirectory = GetInitialDir(),
                FileName = baseName + currentExt
            };
            if (sfd.ShowDialog() != true)
                return;
            string path = sfd.FileName;
            string ext = Path.GetExtension(path).ToLowerInvariant();
            try
            {
                switch (ext)
                {
                    case ".rtf":
                        SaveCurrentAsRtf(path);
                        UpdateStatus($"Saved {Path.GetFileName(path)} (RTF)");
                        break;
                    case ".txt":
                        SaveCurrentAsTxt(path);
                        UpdateStatus($"Saved {Path.GetFileName(path)} (TXT)");
                        break;
                    case ".docx":
                        SaveCurrentAsDocx(path);
                        UpdateStatus($"Saved {Path.GetFileName(path)} (DOCX)");
                        break;
                    case ".odt":
                        SaveCurrentAsOdt(path);
                        UpdateStatus($"Saved {Path.GetFileName(path)} (ODT)");
                        break;
                    default:
                        SaveCurrentAsRtf(path);
                        UpdateStatus($"Saved {Path.GetFileName(path)} (RTF)");
                        break;
                }

                _currentFilePath = path;
                CurrentFileLabel.Text = Path.GetFileName(path);
                if (GetActiveTab() is SWC.TabItem ati && GetMeta(ati) is DocTab ameta)
                {
                    ameta.FullPath = path;
                    ati.Header = System.IO.Path.GetFileName(path);
                }

                RefreshFileList();
            }
            catch (Exception ex)
            {
                SW.MessageBox.Show($"Save failed:\n{ex.Message}", "AnyFile Editor", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveCurrentAsRtf(string path)
        {
            ShowEditor();
            EnsureInitialParagraph();
            var range = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd);
            using (var fs = File.Create(path))
            {
                range.Save(fs, SW.DataFormats.Rtf);
            }
        }

        private void SaveCurrentAsTxt(string path)
        {
            ShowEditor();
            EnsureInitialParagraph();
            var range = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd);
            File.WriteAllText(path, range.Text, Encoding.UTF8);
        }

        private void SaveCurrentAsDocx(string path)
        {
            ShowEditor();
            EnsureInitialParagraph();
            SaveDocxFromFlowDocument(Editor.Document, path);
        }

        private void SaveCurrentAsOdt(string path)
        {
            // Minimal support for now: write RTF content with .odt name, or plain text as fallback.
            try
            {
                var tempRtf = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".rtf");
                SaveCurrentAsRtf(tempRtf);
                File.Copy(tempRtf, path, overwrite: true);
            }
            catch
            {
                SaveCurrentAsTxt(path);
            }
        }

        private void SaveRtf_Click(object sender, RoutedEventArgs e) => SaveAsMulti();
        private void ExportTxt_Click(object sender, RoutedEventArgs e) => ExportAsTxt();
        private void ExportAsTxt()
        {
            if (string.IsNullOrEmpty(_rootFolder))
            {
                SW.MessageBox.Show("Choose a working folder first.");
                return;
            }

            var suggested = string.IsNullOrEmpty(_currentFilePath) ? "Untitled.txt" : Path.ChangeExtension(Path.GetFileName(_currentFilePath), ".txt");
            var sfd = new MWin32.SaveFileDialog
            {
                Filter = "Plain Text (*.txt)|*.txt",
                InitialDirectory = GetInitialDir(),
                FileName = suggested
            };
            if (sfd.ShowDialog() != true)
                return;
            ShowEditor();
            EnsureInitialParagraph();
            var range = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd);
            File.WriteAllText(sfd.FileName, range.Text, Encoding.UTF8);
            RefreshFileList();
            UpdateStatus($"Exported {Path.GetFileName(sfd.FileName)} (plain text)");
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            string? path = (FilesCombo.SelectedItem as FileEntry)?.FullPath;
            if (string.IsNullOrEmpty(path))
            {
                if (string.IsNullOrEmpty(_rootFolder))
                {
                    SW.MessageBox.Show("Choose a working folder first.");
                    return;
                }

                var ofd = new MWin32.OpenFileDialog
                {
                    Filter = "Rich Text (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Word Document (*.docx)|*.docx|All Files|*.*",
                    InitialDirectory = GetInitialDir()
                };
                if (ofd.ShowDialog() == true)
                    path = ofd.FileName;
            }

            if (string.IsNullOrEmpty(path))
                return;
            if (path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
                LoadRtfFile(path);
            else if (path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                LoadDocxFile(path);
            else
                LoadTxtFile(path);
            RefreshFileList();
        }

        private void FilesCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FilesCombo.SelectedItem is not FileEntry fe)
                return;
            var path = fe.FullPath;
            if (path.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
                LoadRtfFile(path);
            else if (path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                LoadDocxFile(path);
            else
                LoadTxtFile(path);
        }

        private void NewFile_Click(object sender, RoutedEventArgs e)
        {
            ShowEditor();
            Editor.Document = new FlowDocument();
            EnsureInitialParagraph();
            _currentFilePath = null;
            CleanupDocxPreview();
            CurrentFileLabel.Text = "(unsaved)";
            UpdateStatus("New document.");
            Editor.Focus();
        }

        private void OpenPalette_Click(object sender, RoutedEventArgs e)
        {
            if (ColorPaletteBtn?.ContextMenu is { } cm)
            {
                cm.PlacementTarget = ColorPaletteBtn;
                cm.IsOpen = true;
            }
        }

        /* ---------- Print + PDF ---------- */
        private FlowDocument CloneDocument()
        {
            var src = new TextRange(Editor.Document.ContentStart, Editor.Document.ContentEnd);
            using var ms = new MemoryStream();
            src.Save(ms, SW.DataFormats.XamlPackage);
            ms.Position = 0;
            var clone = new FlowDocument();
            var dst = new TextRange(clone.ContentStart, clone.ContentEnd);
            dst.Load(ms, SW.DataFormats.XamlPackage);
            return clone;
        }

        private void ExportPdf_Click(object sender, RoutedEventArgs e)
        {
            var doc = PrepareForA4(CloneDocument());
            var pd = new SWC.PrintDialog();
            try
            {
                var server = new LocalPrintServer();
                var qs = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
                var pdfQ = qs.FirstOrDefault(q => q.Name.IndexOf("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) >= 0);
                if (pdfQ != null)
                    pd.PrintQueue = pdfQ;
            }
            catch
            { /* ignore */
            }

            if (pd.ShowDialog() == true)
            {
                pd.PrintDocument(((IDocumentPaginatorSource)doc).DocumentPaginator, "Export to PDF");
                UpdateStatus("Exported to PDF.");
            }
        }

        private static void SaveDocxFromFlowDocument(FlowDocument flowDoc, string filePath)
        {
            using var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            var main = wordDoc.AddMainDocumentPart();
            main.Document = new WP.Document(new WP.Body());
            var body = main.Document.Body!;
            foreach (var block in flowDoc.Blocks)
            {
                switch (block)
                {
                    case Paragraph p:
                        AppendParagraphFromWpf(body, p);
                        break;
                    case Section sec:
                        foreach (var b in sec.Blocks)
                        {
                            if (b is Paragraph bp)
                                AppendParagraphFromWpf(body, bp);
                            else if (b is SWD.Table inner)
                                AppendTableFromWpf(body, inner);
                            else
                            {
                                var tr = new TextRange(b.ContentStart, b.ContentEnd).Text ?? "";
                                body.Append(new WP.Paragraph(new WP.Run(new WP.Text(tr) { Space = SpaceProcessingModeValues.Preserve })));
                            }
                        }

                        break;
                    case SWD.Table t:
                        AppendTableFromWpf(body, t);
                        break;
                    default:
                        var text = new TextRange(block.ContentStart, block.ContentEnd).Text ?? "";
                        body.Append(new WP.Paragraph(new WP.Run(new WP.Text(text) { Space = SpaceProcessingModeValues.Preserve })));
                        break;
                }
            }

            main.Document.Save();
        }

        private void ExportDocx_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_rootFolder))
            {
                SW.MessageBox.Show("Choose a working folder first.");
                return;
            }

            var suggested = string.IsNullOrEmpty(_currentFilePath) ? "Document.docx" : Path.ChangeExtension(Path.GetFileName(_currentFilePath), ".docx");
            var sfd = new MWin32.SaveFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                InitialDirectory = GetInitialDir(),
                FileName = suggested
            };
            if (sfd.ShowDialog() == true)
            {
                SaveDocxFromFlowDocument(Editor.Document, sfd.FileName);
                UpdateStatus($"Exported {Path.GetFileName(sfd.FileName)} (.docx, formatting preserved)");
            }
        }

        private void LoadDocxFile(string path)
        {
            CleanupDocxPreview();
            try
            {
                using var doc = WordprocessingDocument.Open(path, false);
                EnsureRequiredDocxParts(doc);
                var flow = BuildFlowFromDocx(doc);
                ShowEditor();
                BindLoadedDocToActiveTab(path, flow);
                UpdateStatus($"Loaded {Path.GetFileName(path)}");
            }
            catch (Exception ex)
            {
                SW.MessageBox.Show($"Failed to open DOCX: {ex.Message}", "DOCX preview error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    

        private static void SaveRtfFromFlowDocument(FlowDocument doc, string filePath)
        {
            using var fs = File.Create(filePath);
            var range = new TextRange(doc.ContentStart, doc.ContentEnd);
            range.Save(fs, System.Windows.DataFormats.Rtf);
        }

        private static void SaveTxtFromFlowDocument(FlowDocument doc, string filePath)
        {
            var text = new TextRange(doc.ContentStart, doc.ContentEnd).Text ?? string.Empty;
            File.WriteAllText(filePath, text);
        }

        private static void SaveOdtFromFlowDocument(FlowDocument doc, string filePath)
        {
            // Reuse existing ODT pipeline by exporting Editor.Document like before, 
            // but with provided doc (serialize to DOCX-like then convert to ODT zip)
            // Minimal approach: generate a basic ODT using existing code path (already present for SaveCurrentAsOdt)
            // For now, fall back to TXT inside .odt if specialized path not available
            var tempTxt = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".txt");
            try
            {
                SaveTxtFromFlowDocument(doc, tempTxt);
                File.Copy(tempTxt, filePath, overwrite: true);
            }
            finally
            {
                try { File.Delete(tempTxt); } catch { }
            }
        }
    }
}
