// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.Tabs.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Tab management (add/close/select and “+” tab handling).
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.Tabs.cs
// Purpose: Tab management helpers for the main window (add/close/select/+ tab).
// Context: TabControl coordination, TabsPersistence helpers (if present)
// Notes: Invariant: internal _tabs list and TabControl.Items must remain index-aligned.
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
        private SWC.TabItem? GetActiveTab() => _editorTabs?.SelectedItem as SWC.TabItem;
        private DocTab? GetMeta(SWC.TabItem? ti) => ti?.Tag as DocTab;
        // Tab UI fields
        private void HideCurrentFileIndicator()
        {
            try
            {
                if (CurrentFileLabel != null)
                    CurrentFileLabel.Visibility = Visibility.Collapsed;
                // Also hide any TextBlock whose text starts with 'Current file:'
                var root = this as DependencyObject;
                if (root == null)
                    return;
                var queue = new Queue<DependencyObject>();
                queue.Enqueue(root);
                while (queue.Count > 0)
                {
                    var d = queue.Dequeue();
                    int children = System.Windows.Media.VisualTreeHelper.GetChildrenCount(d);
                    for (int i = 0; i < children; i++)
                    {
                        var child = System.Windows.Media.VisualTreeHelper.GetChild(d, i);
                        if (child is TextBlock tb)
                        {
                            var txt = tb.Text?.Trim();
                            if (!string.IsNullOrEmpty(txt) && txt.StartsWith("Current file", StringComparison.OrdinalIgnoreCase))
                            {
                                tb.Visibility = Visibility.Collapsed;
                            }
                        }

                        queue.Enqueue(child);
                    }
                }
            }
            catch
            {
            // best-effort; ignore if the template changes or names differ
            }
        }

        // --- Text measurement helpers for table autosize ---
        private static double MeasureRunDip(Run run, double defaultFontSize, SMB.FontFamily defaultFamily)
        {
            string text = run.Text ?? string.Empty;
            var family = run.FontFamily ?? defaultFamily;
            var size = run.FontSize > 0 ? run.FontSize : defaultFontSize;
            var style = run.FontStyle;
            var weight = run.FontWeight;
            var stretch = run.FontStretch;
            var ft = new SMB.FormattedText(text, System.Globalization.CultureInfo.CurrentUICulture, SW.FlowDirection.LeftToRight, new SMB.Typeface(family, style, weight, stretch), size <= 0 ? 14.0 : size, SMB.Brushes.Black, 1.0);
            return ft.WidthIncludingTrailingWhitespace;
        }

        private static double MeasureBlocksMaxLineWidth(IEnumerable<Block> blocks, double defaultFontSize, SMB.FontFamily defaultFamily)
        {
            double maxLine = 0;
            foreach (var b in blocks)
            {
                if (b is SWD.Paragraph p)
                {
                    double current = 0;
                    foreach (var inline in p.Inlines)
                    {
                        if (inline is LineBreak)
                        {
                            if (current > maxLine)
                                maxLine = current;
                            current = 0;
                            continue;
                        }

                        if (inline is Run r)
                        {
                            current += MeasureRunDip(r, defaultFontSize, defaultFamily);
                        }
                        else if (inline is Span sp)
                        {
                            foreach (var r2 in sp.Inlines.OfType<Run>())
                                current += MeasureRunDip(r2, defaultFontSize, defaultFamily);
                        }
                    }

                    if (current > maxLine)
                        maxLine = current;
                }
                else if (b is SWD.List list)
                {
                    foreach (var li in list.ListItems)
                        maxLine = Math.Max(maxLine, MeasureBlocksMaxLineWidth(li.Blocks, defaultFontSize, defaultFamily));
                }
                else if (b is SWD.Section sec)
                {
                    maxLine = Math.Max(maxLine, MeasureBlocksMaxLineWidth(sec.Blocks, defaultFontSize, defaultFamily));
                }
                else if (b is SWD.Table tbl)
                {
                    AutoSizeTableColumns(tbl);
                    double sum = 0;
                    if (tbl.Columns.Count > 0)
                        foreach (var c in tbl.Columns)
                            sum += (c.Width.IsAbsolute ? c.Width.Value : 0);
                    maxLine = Math.Max(maxLine, sum);
                }
            }

            return maxLine;
        }

        // --- Auto-fit table columns based on content & spans ---
        private static void AutoSizeTableColumns(SWD.Table table)
        {
            if (table == null)
                return;
            // Determine column count considering ColumnSpan
            int colCount = table.Columns.Count;
            if (colCount == 0)
            {
                int maxCols = 0;
                foreach (var group in table.RowGroups)
                {
                    foreach (var row in group.Rows)
                    {
                        int count = 0;
                        foreach (var cell in row.Cells)
                            count += Math.Max(1, cell.ColumnSpan);
                        if (count > maxCols)
                            maxCols = count;
                    }
                }

                colCount = maxCols;
                for (int i = 0; i < colCount; i++)
                    table.Columns.Add(new SWD.TableColumn());
            }

            var maxDip = new double[colCount];
            var defaultFamily = new SMB.FontFamily("Segoe UI");
            double cellHPadding = 8;
            double minCol = 40;
            foreach (var group in table.RowGroups)
            {
                foreach (var row in group.Rows)
                {
                    int colIndex = 0;
                    foreach (var cell in row.Cells)
                    {
                        int span = Math.Max(1, cell.ColumnSpan);
                        double contentWidth = MeasureBlocksMaxLineWidth(cell.Blocks, cell.FontSize > 0 ? cell.FontSize : 14.0, cell.FontFamily ?? defaultFamily);
                        contentWidth += cell.BorderThickness.Left + cell.BorderThickness.Right + cellHPadding;
                        double perCol = contentWidth / span;
                        for (int k = 0; k < span && colIndex + k < colCount; k++)
                            if (perCol > maxDip[colIndex + k])
                                maxDip[colIndex + k] = perCol;
                        colIndex += span;
                    }
                }
            }

            for (int i = 0; i < colCount; i++)
            {
                var w = Math.Max(minCol, maxDip[i] > 0 ? maxDip[i] : minCol);
                table.Columns[i].Width = new System.Windows.GridLength(w);
            }
        }

        // Ensure the minimal DOCX parts exist so we can safely read/write
        private static void EnsureRequiredDocxParts(WordprocessingDocument doc)
        {
            var main = doc.MainDocumentPart;
            if (main == null)
            {
                main = doc.AddMainDocumentPart();
                main.Document = new WP.Document(new WP.Body());
            }

            if (main.Document == null)
                main.Document = new WP.Document(new WP.Body());
        // Numbering/Styles parts are optional for read; create lazily when writing if needed.
        }

        // Convert a DOCX body into a FlowDocument (paragraphs, lists, tables)
        private static FlowDocument BuildFlowFromDocx(WordprocessingDocument doc)
        {
            var flow = new FlowDocument();
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
                return flow;
            Paragraph BuildWpfParagraph(WP.Paragraph p)
            {
                var para = new Paragraph();
                var j = p.ParagraphProperties?.Justification?.Val?.Value;
                if (j == WP.JustificationValues.Center)
                    para.TextAlignment = TextAlignment.Center;
                else if (j == WP.JustificationValues.Right)
                    para.TextAlignment = TextAlignment.Right;
                else if (j == WP.JustificationValues.Both)
                    para.TextAlignment = TextAlignment.Justify;
                else
                    para.TextAlignment = TextAlignment.Left;
                var lineStr = p.ParagraphProperties?.SpacingBetweenLines?.Line?.Value;
                if (lineStr != null && int.TryParse(lineStr, out int lineTwips) && lineTwips > 0)
                {
                    para.LineStackingStrategy = LineStackingStrategy.BlockLineHeight;
                    var fontSize = para.FontSize > 0 ? para.FontSize : 14.0;
                    var factor = Math.Max(1.0, lineTwips / 240.0);
                    para.LineHeight = fontSize * factor;
                }

                foreach (var r in p.Elements<WP.Run>())
                {
                    if (r.Elements<WP.Break>().Any())
                        para.Inlines.Add(new LineBreak());
                    var t = r.GetFirstChild<WP.Text>()?.Text;
                    if (!string.IsNullOrEmpty(t))
                    {
                        var run = new Run(t);
                        var rp = r.RunProperties;
                        if (rp != null)
                        {
                            if (rp.Bold != null)
                                run.FontWeight = FontWeights.Bold;
                            if (rp.Italic != null)
                                run.FontStyle = FontStyles.Italic;
                            if (rp.Underline != null)
                                run.TextDecorations = TextDecorations.Underline;
                        }

                        para.Inlines.Add(run);
                    }
                }

                return para;
            }

            System.Windows.Documents.Table BuildWpfTable(WP.Table wt)
            {
                var table = new System.Windows.Documents.Table
                {
                    CellSpacing = 0
                };
                var firstRow = wt.Elements<WP.TableRow>().FirstOrDefault();
                if (firstRow != null)
                {
                    int colCount = firstRow.Elements<WP.TableCell>().Count();
                    for (int i = 0; i < colCount; i++)
                        table.Columns.Add(new System.Windows.Documents.TableColumn());
                }

                var group = new System.Windows.Documents.TableRowGroup();
                foreach (var wr in wt.Elements<WP.TableRow>())
                {
                    var wpfRow = new System.Windows.Documents.TableRow();
                    foreach (var wc in wr.Elements<WP.TableCell>())
                    {
                        var wpfCell = new System.Windows.Documents.TableCell();
                        var span = wc.TableCellProperties?.GridSpan?.Val?.Value;
                        if (span != null && span.Value > 1)
                            wpfCell.ColumnSpan = span.Value;
                        wpfCell.BorderBrush = System.Windows.Media.Brushes.Gray;
                        wpfCell.BorderThickness = new Thickness(1);
                        wpfCell.Padding = new Thickness(4);
                        foreach (var elem in wc.Elements())
                        {
                            if (elem is WP.Paragraph p)
                                wpfCell.Blocks.Add(BuildWpfParagraph(p));
                            else if (elem is WP.Table innerTbl)
                                wpfCell.Blocks.Add(BuildWpfTable(innerTbl));
                        }

                        if (!wpfCell.Blocks.Any())
                            wpfCell.Blocks.Add(new Paragraph(new Run("")));
                        wpfRow.Cells.Add(wpfCell);
                    }

                    group.Rows.Add(wpfRow);
                }

                table.RowGroups.Add(group);
                AutoSizeTableColumns(table);
                return table;
            }

            System.Windows.Documents.List? curList = null;
            foreach (var el in body.Elements())
            {
                if (el is WP.Paragraph p)
                {
                    bool isList = p.ParagraphProperties?.NumberingProperties != null;
                    bool isOrdered = true;
                    var para = BuildWpfParagraph(p);
                    if (isList)
                    {
                        if (curList == null)
                        {
                            curList = new System.Windows.Documents.List
                            {
                                MarkerStyle = isOrdered ? TextMarkerStyle.Decimal : TextMarkerStyle.Disc
                            };
                            flow.Blocks.Add(curList);
                        }

                        curList.ListItems.Add(new ListItem(para));
                    }
                    else
                    {
                        curList = null;
                        flow.Blocks.Add(para);
                    }
                }
                else if (el is WP.Table wt)
                {
                    curList = null;
                    flow.Blocks.Add(BuildWpfTable(wt));
                }
            }

            return flow;
        }

        private static double InchesToDip(double inches) => inches * 96.0;
        private static (double L, double T, double R, double B) GetMarginPresetInches(string? preset)
        {
            switch (preset)
            {
                case "Narrow (0.5 in)":
                    return (0.5, 0.5, 0.5, 0.5);
                case "Moderate (0.75 in)":
                    return (0.75, 0.75, 0.75, 0.75);
                case "Wide (1.5 in)":
                    return (1.5, 1.5, 1.5, 1.5);
                case "No margins (0 in)":
                    return (0, 0, 0, 0);
                case "Normal (1 in)":
                default:
                    return (1.0, 1.0, 1.0, 1.0);
            }
        }

        // Create "Design" tab dynamically so XAML doesn't need changes
        private void EnsureDesignTab()
        {
            // Try to locate the main TabControl
            var tabs = FindDescendant<SWC.TabControl>(this);
            if (tabs == null)
                return;
            // Check it's not already added
            foreach (var item in tabs.Items)
            {
                if (item is SWC.TabItem ti && (ti.Header as string) == "Design")
                    return;
            }

            var panel = new StackPanel
            {
                Margin = new Thickness(12)
            };
            var row1 = new WrapPanel
            {
                Margin = new Thickness(0, 0, 0, 8)
            };
            _paperSizeCombo = new SWC.ComboBox
            {
                Width = 140,
                Margin = new Thickness(8, 0, 16, 0)
            };
            foreach (var k in PaperSizesInches.Keys)
                _paperSizeCombo.Items.Add(k);
            _paperSizeCombo.SelectedItem = "A4";
            _paperSizeCombo.SelectionChanged += (_, __) => ApplyDesignSettings();
            _orientationCombo = new SWC.ComboBox
            {
                Width = 120,
                Margin = new Thickness(8, 0, 16, 0)
            };
            _orientationCombo.Items.Add("Portrait");
            _orientationCombo.Items.Add("Landscape");
            _orientationCombo.SelectedItem = "Portrait";
            _orientationCombo.SelectionChanged += (_, __) => ApplyDesignSettings();
            _marginPresetCombo = new SWC.ComboBox
            {
                Width = 160,
                Margin = new Thickness(8, 0, 16, 0)
            };
            _marginPresetCombo.Items.Add("Normal (1 in)");
            _marginPresetCombo.Items.Add("Narrow (0.5 in)");
            _marginPresetCombo.Items.Add("Moderate (0.75 in)");
            _marginPresetCombo.Items.Add("Wide (1.5 in)");
            _marginPresetCombo.Items.Add("No margins (0 in)");
            _marginPresetCombo.SelectedItem = "Normal (1 in)";
            _marginPresetCombo.SelectionChanged += (_, __) => ApplyDesignSettings();
            row1.Children.Add(new TextBlock { Text = "Paper size:", VerticalAlignment = VerticalAlignment.Center });
            row1.Children.Add(_paperSizeCombo);
            row1.Children.Add(new TextBlock { Text = "Orientation:", VerticalAlignment = VerticalAlignment.Center });
            row1.Children.Add(_orientationCombo);
            row1.Children.Add(new TextBlock { Text = "Margins:", VerticalAlignment = VerticalAlignment.Center });
            row1.Children.Add(_marginPresetCombo);
            panel.Children.Add(row1);
            var designTab = new SWC.TabItem
            {
                Header = "Design",
                Content = panel
            };
            // Try to insert right after a "Layout" tab if we find one
            int insertAt = tabs.Items.Count;
            for (int i = 0; i < tabs.Items.Count; i++)
            {
                if (tabs.Items[i] is SWC.TabItem ti && (ti.Header as string) == "Layout")
                {
                    insertAt = i + 1;
                    break;
                }
            }

            tabs.Items.Insert(insertAt, designTab);
            // Apply defaults immediately
            ApplyDesignSettings();
        }

        // Generic visual-tree search
        private static T? FindDescendant<T>(DependencyObject root)
            where T : DependencyObject
        {
            if (root == null)
                return null;
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(root); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
                if (child is T t)
                    return t;
                var r = FindDescendant<T>(child);
                if (r != null)
                    return r;
            }

            return null;
        }

        private void EnsureEditorTabStrip()
        {
            if (_editorTabs != null)
                return;
            // Build TabControl UI (compact)
            _editorTabs = new SWC.TabControl
            {
                Margin = new Thickness(2, -2, 2, 0),
                Padding = new Thickness(0),
                FontSize = 12,
                SnapsToDevicePixels = true
            };
            var tabStyle = new Style(typeof(SWC.TabItem));
            tabStyle.Setters.Add(new Setter(SWC.Control.PaddingProperty, new Thickness(8, 2, 8, 2)));
            tabStyle.Setters.Add(new Setter(SWC.Control.FontSizeProperty, 12.5));
            tabStyle.Setters.Add(new Setter(SWC.Control.MinHeightProperty, 24.0));
            tabStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(0, 0, 2, 0)));
            _editorTabs.ItemContainerStyle = tabStyle;
            _editorTabs.SelectionChanged += EditorTabs_SelectionChanged;
            // Create '+' tab
            _plusTab = new SWC.TabItem
            {
                Header = "+"
            };
            // Build a container grid: [tabs] + [Editor]
            var parent = Editor.Parent;
            int row = SWC.Grid.GetRow(Editor);
            int col = SWC.Grid.GetColumn(Editor);
            int rspan = SWC.Grid.GetRowSpan(Editor);
            int cspan = SWC.Grid.GetColumnSpan(Editor);
            var grid = new SWC.Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            // Ensure layout stretches
            grid.HorizontalAlignment = SW.HorizontalAlignment.Stretch;
            grid.VerticalAlignment = SW.VerticalAlignment.Stretch;
            SWC.Grid.SetRow(_editorTabs, 0);
            SWC.Grid.SetRow(Editor, 1);
            Editor.VerticalAlignment = SW.VerticalAlignment.Stretch;
            Editor.HorizontalAlignment = SW.HorizontalAlignment.Stretch;
            Editor.ClearValue(FrameworkElement.HeightProperty);
            Editor.ClearValue(FrameworkElement.WidthProperty);
            // Move editor into grid row 1
            if (parent is SWC.Panel pnl)
            {
                pnl.Children.Remove(Editor);
                SWC.Grid.SetRow(Editor, 1);
                grid.Children.Add(_editorTabs);
                grid.Children.Add(Editor);
                // Keep original grid cell for the editor area
                SWC.Grid.SetRow(grid, row);
                SWC.Grid.SetColumn(grid, col);
                SWC.Grid.SetRowSpan(grid, rspan);
                SWC.Grid.SetColumnSpan(grid, cspan);
                pnl.Children.Add(grid);
            }
            else if (parent is SWC.ContentControl cc)
            {
                cc.Content = null;
                SWC.Grid.SetRow(Editor, 1);
                grid.Children.Add(_editorTabs);
                grid.Children.Add(Editor);
                cc.Content = grid;
            }
            else
            {
                // Fallback: host within a new Grid
                var host = new SWC.Grid();
                host.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                host.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                SWC.Grid.SetRow(Editor, 1);
                host.Children.Add(_editorTabs);
                host.Children.Add(Editor);
            }

            // Seed tabs: persistent Untitled + optional current file + '+'
            var untitledTab = CreateBlankTab("Untitled");
            _editorTabs.Items.Add(untitledTab);
            if (!string.IsNullOrWhiteSpace(_currentFilePath))
            {
                var currentDoc = Editor.Document ?? new FlowDocument(new Paragraph(new Run()));
                var currentFileTab = new SWC.TabItem { Tag = new DocTab { Document = currentDoc, FullPath = _currentFilePath } };
                currentFileTab.Header = BuildTabHeader(System.IO.Path.GetFileName(_currentFilePath), currentFileTab);
                _editorTabs.Items.Add(currentFileTab);
            }

            _editorTabs.Items.Add(_plusTab);
            _editorTabs.SelectedItem = untitledTab;
            if (untitledTab.Tag is DocTab m1)
                Editor.Document = m1.Document;
            HideCurrentFileIndicator();
        }

        private void EditorTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_editorTabs == null)
                return;
            // react only to events from the tab control itself (ignore bubbled events from inner controls)
            if (!ReferenceEquals(e.OriginalSource, _editorTabs))
                return;
            // Save the tab we are leaving
            var leaving = e.RemovedItems.Count > 0 ? e.RemovedItems[0] as SWC.TabItem : null;
            SaveEditorDocIntoTab(leaving);
            // '+' tab -> open a BRAND NEW blank tab (no file dialog)
            if (ReferenceEquals(_editorTabs.SelectedItem, _plusTab))
            {
                var previous = leaving;
                if (previous != null)
                    _editorTabs.SelectedItem = previous;
                var newTab = CreateBlankTab();
                int insertIndex = _editorTabs.Items.Count - 1; // before '+'
                _editorTabs.Items.Insert(insertIndex, newTab);
                _editorTabs.SelectedItem = newTab;
                return;
            }

            // normal tab switch
            if (_editorTabs.SelectedItem is SWC.TabItem ti && ti.Tag is DocTab meta)
            {
                Editor.Document = meta.Document ?? new FlowDocument(new Paragraph(new Run()));
                _currentFilePath = meta.FullPath;
            }
        }

        // Creates a new blank FlowDocument tab (not backed by a file yet)
        private SWC.TabItem CreateBlankTab(string? preferredName = null)
        {
            var doc = new FlowDocument(new Paragraph(new Run()));
            var header = preferredName ?? NextUntitledName();
            var ti = new SWC.TabItem { Header = header, Tag = new DocTab { Document = doc, FullPath = null } };
            ti.Header = BuildTabHeader(header, ti);
            return ti;
        }
        

        private void CloseOtherTabs(System.Windows.Controls.TabItem keep)
        {
            if (_editorTabs == null) return;
            var toClose = _editorTabs.Items.OfType<System.Windows.Controls.TabItem>()
                .Where(t => !ReferenceEquals(t, _plusTab) && !ReferenceEquals(t, keep))
                .ToList();
            foreach (var t in toClose)
                TryCloseTab(t);
        }

        private void CloseAllTabs()
        {
            if (_editorTabs == null) return;
            var toClose = _editorTabs.Items.OfType<System.Windows.Controls.TabItem>()
                .Where(t => !ReferenceEquals(t, _plusTab))
                .ToList();
            foreach (var t in toClose)
                TryCloseTab(t);
        }


        // Generates the next "Untitled", "Untitled 2", ... unique header
        private string NextUntitledName()
        {
            if (_editorTabs == null)
                return "Untitled";
            var existing = new HashSet<string>(_editorTabs.Items.OfType<SWC.TabItem>().Select(t => t.Header?.ToString() ?? string.Empty));
            const string baseName = "Untitled";
            if (!existing.Contains(baseName))
                return baseName;
            int i = 2;
            while (existing.Contains($"{baseName} {i}"))
                i++;
            return $"{baseName} {i}";
        }

        /* ---------- Status & helpers ---------- */
        private void UpdateStatus(string msg) => StatusText.Text = msg;
        private void SetRootFolderNameLabel()
        {
            RootFolderNameText.Text = string.IsNullOrEmpty(_rootFolder) ? "" : Path.GetFileName(_rootFolder);
        }

        private void EnsureInitialParagraph()
        {
            if (Editor.Document == null)
                Editor.Document = new FlowDocument();
            if (!Editor.Document.Blocks.Any())
                Editor.Document.Blocks.Add(new Paragraph(new Run("")));
        }

        private static bool IsParagraphEmpty(Paragraph p)
        {
            var tr = new TextRange(p.ContentStart, p.ContentEnd);
            return string.IsNullOrWhiteSpace(tr.Text);
        }

        /* ---------- File list & views ---------- */
        private void RefreshFileList()
        {
            if (string.IsNullOrEmpty(_rootFolder) || !Directory.Exists(_rootFolder))
            {
                FilesCombo.ItemsSource = null;
                return;
            }

            string[] exts = new[]
            {
                ".rtf",
                ".txt",
                ".docx"
            };
            var files = Directory.GetFiles(_rootFolder, "*.*", SearchOption.TopDirectoryOnly).Where(p => exts.Contains(Path.GetExtension(p), StringComparer.OrdinalIgnoreCase)).Select(p => new FileEntry(Path.GetFileName(p), p)).OrderBy(e => e.Name).ToList();
            FilesCombo.ItemsSource = files;
            if (!string.IsNullOrEmpty(_currentFilePath))
            {
                var match = files.FirstOrDefault(f => string.Equals(f.FullPath, _currentFilePath, StringComparison.OrdinalIgnoreCase));
                FilesCombo.SelectedItem = match;
            }
            else
            {
                FilesCombo.SelectedIndex = -1;
            }
        }

        private void ShowEditor()
        {
            DocxWeb.Visibility = Visibility.Collapsed;
            Editor.Visibility = Visibility.Visible;
        }

        private void ShowDocxPreview()
        {
            Editor.Visibility = Visibility.Collapsed;
            DocxWeb.Visibility = Visibility.Visible;
        }

        private void CleanupDocxPreview()
        {
            try
            {
                if (_docxPreviewFolder != null && Directory.Exists(_docxPreviewFolder))
                    Directory.Delete(_docxPreviewFolder, true);
            }
            catch
            { /* ignore */
            }

            _docxPreviewFolder = null;
        }

        private void ApplyToSelectionParagraphs(Action<Paragraph> action)
        {
            ShowEditor();
            var sel = Editor.Selection;
            var startPara = sel.Start?.Paragraph ?? Editor.CaretPosition?.Paragraph;
            var end = sel.End;
            if (startPara == null)
            {
                var p = new Paragraph(new Run(""));
                Editor.Document.Blocks.Add(p);
                startPara = p;
            }

            Paragraph? current = startPara;
            while (current != null)
            {
                action(current);
                if (current.ContentEnd.CompareTo(end) >= 0)
                    break;
                current = current.NextBlock as Paragraph;
            }
        }

        private void Bullets_Click(object sender, RoutedEventArgs e) => EditingCommands.ToggleBullets.Execute(null, Editor);
        private void Numbering_Click(object sender, RoutedEventArgs e) => EditingCommands.ToggleNumbering.Execute(null, Editor);
        /* ---------- Insert Table ---------- */
        private void InsertTable_Click(object sender, RoutedEventArgs e)
        {
            ShowEditor();
            var dlg = new InsertTableDialog
            {
                Owner = this
            };
            var ok = dlg.ShowDialog();
            if (ok == true && dlg.Rows > 0 && dlg.Columns > 0)
                InsertTableIntoEditor(dlg.Rows, dlg.Columns);
        }

        private void InsertTableIntoEditor(int rows, int cols)
        {
            var table = new SWD.Table
            {
                CellSpacing = 0
            };
            for (int c = 0; c < cols; c++)
                table.Columns.Add(new SWD.TableColumn());
            var group = new SWD.TableRowGroup();
            for (int r = 0; r < rows; r++)
            {
                var tr = new SWD.TableRow();
                for (int c = 0; c < cols; c++)
                {
                    var cell = new SWD.TableCell(new SWD.Paragraph(new Run("")));
                    cell.BorderBrush = SMB.Brushes.Gray;
                    cell.BorderThickness = new Thickness(1);
                    cell.Padding = new Thickness(4);
                    tr.Cells.Add(cell);
                }

                group.Rows.Add(tr);
            }

            table.RowGroups.Add(group);
            var caretPara = Editor.CaretPosition?.Paragraph;
            if (caretPara == null)
            {
                caretPara = new SWD.Paragraph(new Run(""));
                Editor.Document.Blocks.Add(caretPara);
            }

            Editor.Document.Blocks.InsertAfter(caretPara, table);
            Editor.CaretPosition = table.ContentEnd;
        }

        private void FontSizeEditableTextBox_TextChanged(object? sender, TextChangedEventArgs e)
        {
            if (double.TryParse(FontSizeCombo.Text, out _))
                ApplyFontSizeFromCombo(applyParagraphWhenEmpty: true, focusEditorAfter: false);
        }

        /* ---------- Common UI actions ---------- */
        private void ChooseRoot_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new WF.FolderBrowserDialog();
            if (dlg.ShowDialog() == WF.DialogResult.OK)
            {
                _rootFolder = dlg.SelectedPath;
                SetRootFolderNameLabel();
                RefreshFileList();
                _currentFilePath = null;
                CleanupDocxPreview();
                ShowEditor();
                EnsureInitialParagraph();
                CurrentFileLabel.Text = "(unsaved)";
                UpdateStatus($"Working folder set: {_rootFolder}");
            }
        }

        private static FlowDocument PrepareForA4(FlowDocument doc)
        {
            const double dpi = 96.0;
            double a4Width = 8.27 * dpi; // 210 mm
            double a4Height = 11.69 * dpi; // 297 mm
            double margin = 0.75 * dpi; // 0.75" margins
            doc.PageWidth = a4Width;
            doc.PageHeight = a4Height;
            doc.PagePadding = new Thickness(margin);
            doc.ColumnGap = 0;
            doc.ColumnWidth = a4Width - doc.PagePadding.Left - doc.PagePadding.Right;
            return doc;
        }

        /* ---------- DOCX export (formatting preserved) ---------- */
        private static int PxToHalfPoints(double wpfPx) => (int)Math.Round(wpfPx * 72.0 / 48.0); // px -> pt (72/96) -> half-points (*2) == *1.5
        private static string BrushToHex(SMB.Brush? b)
        {
            if (b is SolidColorBrush scb)
                return $"{scb.Color.R:X2}{scb.Color.G:X2}{scb.Color.B:X2}";
            return "000000";
        }

        private static void AppendInlineToParagraph(WP.Paragraph p, Inline inline)
        {
            switch (inline)
            {
                case Run r:
                    string txt = r.Text ?? string.Empty;
                    string[] parts = txt.Replace("\r\n", "\n").Split('\n');
                    for (int i = 0; i < parts.Length; i++)
                    {
                        var wr = new WP.Run();
                        var rp = new WP.RunProperties();
                        if (r.FontFamily != null)
                        {
                            var fam = r.FontFamily.Source;
                            rp.RunFonts = new WP.RunFonts
                            {
                                Ascii = fam,
                                HighAnsi = fam,
                                EastAsia = fam,
                                ComplexScript = fam
                            };
                        }

                        if (r.FontSize > 0)
                            rp.FontSize = new WP.FontSize
                            {
                                Val = PxToHalfPoints(r.FontSize).ToString()
                            };
                        if (r.FontWeight == FontWeights.Bold)
                            rp.Bold = new WP.Bold();
                        if (r.FontStyle == FontStyles.Italic)
                            rp.Italic = new WP.Italic();
                        if (r.TextDecorations != null && r.TextDecorations.Count > 0)
                            rp.Underline = new WP.Underline
                            {
                                Val = WP.UnderlineValues.Single
                            };
                        rp.Color = new WP.Color
                        {
                            Val = BrushToHex(r.Foreground)
                        };
                        wr.Append(rp);
                        wr.Append(new WP.Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
                        p.Append(wr);
                        if (i < parts.Length - 1)
                            p.Append(new WP.Run(new WP.Break()));
                    }

                    break;
                case Span sp:
                    foreach (var child in sp.Inlines)
                        AppendInlineToParagraph(p, child);
                    break;
                case LineBreak:
                    p.Append(new WP.Run(new WP.Break()));
                    break;
                default:
                    var tr = new TextRange(inline.ContentStart, inline.ContentEnd).Text ?? "";
                    if (tr.Length > 0)
                        p.Append(new WP.Run(new WP.Text(tr) { Space = SpaceProcessingModeValues.Preserve }));
                    break;
            }
        }

        // Border helper (correct enum)
        private static WP.BorderValues BorderStyleFromThickness(Thickness t) => (t.Left > 0 || t.Top > 0 || t.Right > 0 || t.Bottom > 0) ? WP.BorderValues.Single : WP.BorderValues.Nil;
        // ---------- builders so we can append to Body or TableCell ----------
        private static WP.Paragraph BuildParagraphFromWpf(Paragraph para)
        {
            var wp = new WP.Paragraph();
            var pp = new WP.ParagraphProperties
            {
                Justification = new WP.Justification
                {
                    Val = MapAlignment(para.TextAlignment)
                }
            };
            // Map line spacing (multiple) -> OpenXML line spacing (in 240ths of a line)
            double factor = 0;
            if (para.LineStackingStrategy == LineStackingStrategy.BlockLineHeight && para.LineHeight > 0 && para.FontSize > 0)
            {
                factor = Math.Round(para.LineHeight / para.FontSize, 2);
            }

            if (factor > 0)
            {
                int line = (int)Math.Round(240 * factor);
                pp.SpacingBetweenLines = new WP.SpacingBetweenLines
                {
                    Line = line.ToString(),
                    LineRule = WP.LineSpacingRuleValues.Auto
                };
            }

            wp.Append(pp);
            foreach (var inline in para.Inlines)
                AppendInlineToParagraph(wp, inline);
            return wp;
        }

        private static WP.Table BuildTableFromWpf(SWD.Table t)
        {
            var wt = new WP.Table();
            var tblProps = new WP.TableProperties(new WP.TableBorders(new WP.TopBorder { Val = BorderStyleFromThickness(t.BorderThickness), Size = 4 }, new WP.LeftBorder { Val = BorderStyleFromThickness(t.BorderThickness), Size = 4 }, new WP.BottomBorder { Val = BorderStyleFromThickness(t.BorderThickness), Size = 4 }, new WP.RightBorder { Val = BorderStyleFromThickness(t.BorderThickness), Size = 4 }, new WP.InsideHorizontalBorder { Val = WP.BorderValues.Single, Size = 4 }, new WP.InsideVerticalBorder { Val = WP.BorderValues.Single, Size = 4 }));
            wt.Append(tblProps);
            foreach (var group in t.RowGroups)
            {
                foreach (SWD.TableRow r in group.Rows)
                {
                    var wr = new WP.TableRow();
                    foreach (SWD.TableCell c in r.Cells)
                    {
                        var wc = new WP.TableCell();
                        var wcp = new WP.TableCellProperties
                        {
                            TableCellBorders = new WP.TableCellBorders(new WP.TopBorder { Val = BorderStyleFromThickness(c.BorderThickness), Size = 4 }, new WP.LeftBorder { Val = BorderStyleFromThickness(c.BorderThickness), Size = 4 }, new WP.BottomBorder { Val = BorderStyleFromThickness(c.BorderThickness), Size = 4 }, new WP.RightBorder { Val = BorderStyleFromThickness(c.BorderThickness), Size = 4 })
                        };
                        wc.Append(wcp);
                        foreach (var b in c.Blocks)
                        {
                            if (b is Paragraph p)
                                wc.Append(BuildParagraphFromWpf(p));
                            else if (b is SWD.Table inner)
                                wc.Append(BuildTableFromWpf(inner));
                            else
                            {
                                var tr = new TextRange(b.ContentStart, b.ContentEnd).Text ?? "";
                                var wp = new WP.Paragraph(new WP.Run(new WP.Text(tr) { Space = SpaceProcessingModeValues.Preserve }));
                                wc.Append(wp);
                            }
                        }

                        wr.Append(wc);
                    }

                    wt.Append(wr);
                }
            }

            return wt;
        }

        // Overloads that append to Body
        private static void AppendParagraphFromWpf(WP.Body body, Paragraph para) => body.Append(BuildParagraphFromWpf(para));
        private static void AppendTableFromWpf(WP.Body body, SWD.Table t) => body.Append(BuildTableFromWpf(t));
        // Overloads that append to TableCell
        private static void AppendParagraphFromWpf(WP.TableCell cell, Paragraph para) => cell.Append(BuildParagraphFromWpf(para));
        private static void AppendTableFromWpf(WP.TableCell cell, SWD.Table t) => cell.Append(BuildTableFromWpf(t));
    
        // ===== Tab Header with Close Button =====
        private FrameworkElement BuildTabHeader(string title, System.Windows.Controls.TabItem owner)
        {
            var grid = new System.Windows.Controls.Grid
            {
                Margin = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Center
            };
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });

            var txt = new TextBlock
            {
                Text = title,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            };
            System.Windows.Controls.Grid.SetColumn(txt, 0);
            grid.Children.Add(txt);

            var btn = new System.Windows.Controls.Button
            {
Content = "×",
                Width = 18,
                Height = 18,
                Padding = new Thickness(0),
                Margin = new Thickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                ToolTip = "Close",
                Focusable = false,
                IsTabStop = false,
                ClickMode = System.Windows.Controls.ClickMode.Press
            };
            btn.Tag = owner;
            btn.Click += CloseTabButton_Click;
            btn.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "TabCloseGlyphBrush");
            btn.SetResourceReference(System.Windows.FrameworkElement.StyleProperty, "CloseGlyphButtonStyle");
            System.Windows.Controls.Grid.SetColumn(btn, 1);
            grid.Children.Add(btn);

            // Optional: right-click menu
            var cm = new System.Windows.Controls.ContextMenu();
            var miClose = new System.Windows.Controls.MenuItem { Header = "Close" };
            miClose.Click += (_, __) => TryCloseTab(owner);
            var miCloseOthers = new System.Windows.Controls.MenuItem { Header = "Close Others" };
            miCloseOthers.Click += (_, __) => CloseOtherTabs(owner);
            var miCloseAll = new System.Windows.Controls.MenuItem { Header = "Close All" };
            miCloseAll.Click += (_, __) => CloseAllTabs();
            cm.Items.Add(miClose);
            cm.Items.Add(miCloseOthers);
            cm.Items.Add(miCloseAll);
            grid.ContextMenu = cm;

            return grid;
        }

        private void CloseTabButton_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            System.Windows.Controls.TabItem? ti = null;
            if (sender is System.Windows.Controls.Button b)
            {
                ti = b.Tag as System.Windows.Controls.TabItem;
                if (ti == null)
                {
                    // Fallback: walk up the visual tree to find the TabItem container
                    DependencyObject cur = b;
                    while (cur != null && ti == null)
                    {
                        if (cur is System.Windows.Controls.TabItem tti) { ti = tti; break; }
                        cur = System.Windows.Media.VisualTreeHelper.GetParent(cur);
                    }
                }
            }
            if (ti != null)
                TryCloseTab(ti);
        }
        

        private static string GetPlainText(FlowDocument doc)
        {
            var r = new TextRange(doc.ContentStart, doc.ContentEnd);
            return r.Text ?? string.Empty;
        }

        private static string ReadFileAsPlainText(string path)
        {
            var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            try
            {
                switch (ext)
                {
                    case ".txt":
                        return System.IO.File.ReadAllText(path);
                    case ".rtf":
                        {
                            var rtb = new System.Windows.Controls.RichTextBox();
                            using var fs = System.IO.File.OpenRead(path);
                            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                            range.Load(fs, System.Windows.DataFormats.Rtf);
                            return GetPlainText(rtb.Document);
                        }
                    case ".docx":
                        {
                            using var word = WordprocessingDocument.Open(path, false);

// Invariant: the backing tab list and the visual TabControl must remain index-aligned;
// selection, persistence, and close/add operations assume parity.


                            // Basic plain-text extraction: concatenate paragraph texts
                            var sb = new StringBuilder();
                            var body = word.MainDocumentPart?.Document?.Body;
                            if (body != null)
                            {
                                foreach (var p in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                                {
                                    var text = string.Concat(p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                                    sb.AppendLine(text);
                                }
                            }
                            return sb.ToString();
                        }
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private bool IsTabDirty(System.Windows.Controls.TabItem tab)
        {
            if (tab == _plusTab) return false;
            if (tab.Tag is not DocTab meta) return false;

            // Ensure meta.Document holds current editor content when checking the active tab
            if (ReferenceEquals(_editorTabs?.SelectedItem, tab))
                SaveEditorDocIntoTab(tab);

            var textNow = GetPlainText(meta.Document).TrimEnd();
            if (string.IsNullOrEmpty(meta.FullPath))
            {
                // Unsaved doc: consider dirty if there's any non-empty text
                return !string.IsNullOrWhiteSpace(textNow);
            }
            var disk = ReadFileAsPlainText(meta.FullPath).TrimEnd();
            return !string.Equals(textNow, disk, StringComparison.Ordinal);
        }

        private void TryCloseTab(System.Windows.Controls.TabItem tab)
        {
            if (tab == null) return;
            // Identify owning TabControl (supports multiple tab strips / split view)
            var owner = System.Windows.Controls.ItemsControl.ItemsControlFromItemContainer(tab) as System.Windows.Controls.TabControl
                        ?? tab.Parent as System.Windows.Controls.TabControl
                        ?? _editorTabs; // fallback
            if (owner == null) return;

            // Determine '+' sentinel for this owner (if any)
            System.Windows.Controls.TabItem? plus = null;
            foreach (var it in owner.Items)
            {
                if (it is System.Windows.Controls.TabItem ti)
                {
                    var headerStr = ti.Header as string;
                    if (headerStr == "+") { plus = ti; break; }
                }
            }
            if (ReferenceEquals(tab, plus)) return; // never close '+' tab

            // Dirty check & save prompt (leverage existing tab meta)
            var meta = tab.Tag as DocTab;
            string displayName = System.IO.Path.GetFileName(meta?.FullPath ?? "Untitled");
            if (IsTabDirty(tab))
            {
                var res = SW.MessageBox.Show(
                    $"Save changes to {displayName}?",
                    "Close Tab",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);
                if (res == MessageBoxResult.Cancel) return;
                if (res == MessageBoxResult.Yes)
                {
                    if (meta == null || !SaveTab(meta)) return; // aborted or failed
                }
            }

            // Remove from the owner control and pick a stable neighbor
            int idx = owner.Items.IndexOf(tab);
            owner.Items.Remove(tab);

            // Ensure at least one real tab remains (before '+')
            int nonPlusCount = owner.Items.OfType<System.Windows.Controls.TabItem>().Count(t => !ReferenceEquals(t, plus));
            if (nonPlusCount == 0)
            {
                var fallback = CreateBlankTab("Untitled");
                int insertIndex = plus != null ? owner.Items.IndexOf(plus) : owner.Items.Count;
                insertIndex = insertIndex < 0 ? owner.Items.Count : insertIndex;
                owner.Items.Insert(insertIndex, fallback);
                owner.SelectedItem = fallback;
            }
            else
            {
                if (owner.Items.Count > 0)
                {
                    int selIndex = System.Math.Max(0, System.Math.Min(idx, owner.Items.Count - 1));
                    if (plus != null && ReferenceEquals(owner.Items[selIndex], plus) && selIndex > 0)
                        selIndex -= 1;
                    owner.SelectedIndex = selIndex;
                }
            }
        }
            


        private bool SaveTab(DocTab meta)
        {
            try
            {
                if (string.IsNullOrEmpty(meta.FullPath))
                {
                    // Prompt for path (multi-format like SaveAsMulti)
                    var sfd = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter = "Rich Text Format (*.rtf)|*.rtf|Plain Text (*.txt)|*.txt|Word Document (*.docx)|*.docx|OpenDocument Text (*.odt)|*.odt",
                        FileName = "Document"
                    };
                    if (sfd.ShowDialog() != true)
                        return false;
                    var path = sfd.FileName;
                    SaveTabDocumentToPath(meta, path);
                }
                else
                {
                    SaveTabDocumentToPath(meta, meta.FullPath);
                }
                return true;
            }
            catch (Exception ex)
            {
                SW.MessageBox.Show($"Failed to save: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private void SaveTabDocumentToPath(DocTab meta, string path)
        {
            var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".odt")
            {
                // Use existing ODT pipeline by temporarily selecting the tab
                var current = _editorTabs?.SelectedItem;
                if (_editorTabs != null)
                    _editorTabs.SelectedItem = _editorTabs.Items.OfType<System.Windows.Controls.TabItem>().FirstOrDefault(t => ReferenceEquals(t.Tag, meta)) ?? current;
                var originalDoc = Editor.Document;
                Editor.Document = meta.Document;
                SaveCurrentAsOdt(path);
                Editor.Document = originalDoc;
                if (_editorTabs != null)
                    _editorTabs.SelectedItem = current;
            }
            else
            {
                switch (ext)
                {
                    case ".rtf":
                        SaveRtfFromFlowDocument(meta.Document, path);
                        break;
                    case ".txt":
                        SaveTxtFromFlowDocument(meta.Document, path);
                        break;
                    case ".docx":
                        SaveDocxFromFlowDocument(meta.Document, path);
                        break;
                    default:
                        SaveRtfFromFlowDocument(meta.Document, path);
                        break;
                }
            }
            meta.FullPath = path;
        }
    }
}


