// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.Formatting.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Text formatting commands and editor operations.
// ============================================================================
// TODO (2025-09-21): Replace manual DOCX spacing workaround when upstream issue is fixed.
// File: AnyFileEditor_fixed_all/MainWindow.Formatting.cs
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
        private TextPointer EnsureInsertionPosition()
        {
            EnsureInitialParagraph();
            var caret = Editor.CaretPosition;
            if (caret.Paragraph == null)
            {
                var para = Editor.Document.Blocks.LastBlock as Paragraph ?? new Paragraph(new Run(""));
                if (para.Parent == null)
                    Editor.Document.Blocks.Add(para);
                caret = para.ContentEnd;
            }

            if (!caret.IsAtInsertionPosition)
                caret = caret.GetInsertionPosition(LogicalDirection.Forward);
            return caret;
        }

        /* ---------- Color helpers ---------- */
        private void ApplyTextColor(SMB.Color c)
        {
            var brush = new SolidColorBrush(c);
            Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        }

        /* ---------- Text formatting ---------- */
        private void AlignLeft_Click(object sender, RoutedEventArgs e) => ApplyToSelectionParagraphs(p => p.TextAlignment = TextAlignment.Left);
        private void AlignCenter_Click(object sender, RoutedEventArgs e) => ApplyToSelectionParagraphs(p => p.TextAlignment = TextAlignment.Center);
        private void AlignRight_Click(object sender, RoutedEventArgs e) => ApplyToSelectionParagraphs(p => p.TextAlignment = TextAlignment.Right);
        private void Bold_Click(object sender, RoutedEventArgs e)
        {
            var current = Editor.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            var next = (current is FontWeight fw && fw == FontWeights.Bold) ? FontWeights.Normal : FontWeights.Bold;
            Editor.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, next);
        }

        private void Italic_Click(object sender, RoutedEventArgs e)
        {
            var current = Editor.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            var next = (current is SW.FontStyle fs && fs == SW.FontStyles.Italic) ? SW.FontStyles.Normal : SW.FontStyles.Italic;
            Editor.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, next);
        }

        private void Underline_Click(object sender, RoutedEventArgs e)
        {
            var current = Editor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            if (current is TextDecorationCollection decs && decs == TextDecorations.Underline)
                Editor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, null);
            else
                Editor.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
        }

        private void LineSpacingCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LineSpacingCombo.SelectedItem is ComboBoxItem item && double.TryParse(item.Content.ToString(), out double factor))
            {
                ApplyToSelectionParagraphs(p =>
                {
                    p.LineStackingStrategy = LineStackingStrategy.BlockLineHeight;
                    var size = p.FontSize > 0 ? p.FontSize : 14.0;
                    p.LineHeight = size * factor;
                });
            }
        }

        /* ---------- Font SIZE dropdown ---------- */
        private void InitFontSizeCombo()
        {
            int[] first = new[]
            {
                2,
                4,
                6,
                8,
                10
            };
            foreach (var v in first)
                FontSizeCombo.Items.Add(v);
            for (int i = 0; i <= 15; i++)
                FontSizeCombo.Items.Add(11 + i * 2);
            FontSizeCombo.SelectionChanged += FontSizeCombo_SelectionChanged;
            FontSizeCombo.LostKeyboardFocus += FontSizeCombo_LostKeyboardFocus;
            FontSizeCombo.Loaded += (_, __) =>
            {
                FontSizeCombo.ApplyTemplate();
                _fontSizeEditableTextBox = FontSizeCombo.Template.FindName("PART_EditableTextBox", FontSizeCombo) as SWC.TextBox;
                if (_fontSizeEditableTextBox != null)
                {
                    _fontSizeEditableTextBox.TextChanged += FontSizeEditableTextBox_TextChanged;
                    _fontSizeEditableTextBox.KeyDown += FontSizeEditableTextBox_KeyDown;
                }
            };
            FontSizeCombo.AddHandler(SWI.Keyboard.KeyDownEvent, new SWI.KeyEventHandler(FontSizeCombo_KeyDown), true);
        }

        private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontSizeCombo.IsDropDownOpen)
            {
                ApplyFontSizeFromCombo(applyParagraphWhenEmpty: true, focusEditorAfter: true);
            }
        }

        private void FontSizeCombo_LostKeyboardFocus(object sender, SWI.KeyboardFocusChangedEventArgs e)
        {
            ApplyFontSizeFromCombo(applyParagraphWhenEmpty: true, focusEditorAfter: false);
        }

        private void ApplyFontSizeFromCombo(bool applyParagraphWhenEmpty, bool focusEditorAfter)
        {
            if (!double.TryParse(FontSizeCombo.Text, out double size) || size <= 0)
                return;
            if (_lastSizeAppliedFromCombo.HasValue && Math.Abs(_lastSizeAppliedFromCombo.Value - size) < 0.01)
            {
                if (focusEditorAfter)
                    Editor.Focus();
                return;
            }

            EnsureInitialParagraph();
            var sel = Editor.Selection;
            if (sel.IsEmpty)
            {
                var tp = EnsureInsertionPosition();
                var para = tp.Paragraph;
                if (applyParagraphWhenEmpty && para != null && IsParagraphEmpty(para))
                    para.FontSize = size;
                sel.Select(tp, tp);
            }

            sel.ApplyPropertyValue(TextElement.FontSizeProperty, size);
            _lastSizeAppliedFromCombo = size;
            if (focusEditorAfter)
                Editor.Focus();
        }

        /* ---------- Font FAMILY dropdown ---------- */
        private void InitFontFamilyCombo()
        {
            _fontChoices.Clear();
            foreach (var ff in Fonts.SystemFontFamilies)
            {
                var name = string.IsNullOrWhiteSpace(ff.Source) ? ff.FamilyNames.Values.FirstOrDefault() ?? "Font" : ff.Source;
                _fontChoices.Add(new FontChoice { Name = name, Family = ff });
            }

            _fontChoices.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.CurrentCultureIgnoreCase));
            _fontChoices.Add(new FontChoice { Name = "[ Add font… ]", IsAddItem = true });
            FontFamilyCombo.ItemsSource = _fontChoices;
            try
            {
                var current = Editor?.Selection?.GetPropertyValue(TextElement.FontFamilyProperty);
                if (current is SMB.FontFamily curFam)
                {
                    var match = _fontChoices.FirstOrDefault(fc => !fc.IsAddItem && string.Equals(fc.Name, curFam.Source, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                        FontFamilyCombo.SelectedItem = match;
                }
            }
            catch
            {
            }

            FontFamilyCombo.SelectionChanged += FontFamilyCombo_SelectionChanged;
        }

        private void FontFamilyCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontFamilyCombo.SelectedItem is not FontChoice choice)
                return;
            if (choice.IsAddItem)
            {
                AddFontFromFile();
                FontFamilyCombo.SelectedItem = null;
                return;
            }

            if (choice.Family != null)
                ApplyFontFamily(choice.Family, applyParagraphWhenEmpty: true, focusEditorAfter: true);
        }

        private void AddFontFromFile()
        {
            var ofd = new MWin32.OpenFileDialog
            {
                Title = "Add font",
                Filter = "Font files (*.ttf;*.otf)|*.ttf;*.otf",
                Multiselect = false
            };
            if (ofd.ShowDialog() != true)
                return;
            var path = ofd.FileName;
            try
            {
                var glyph = new SMB.GlyphTypeface(new Uri(path, UriKind.Absolute));
                var familyName = glyph.FamilyNames.Values.FirstOrDefault() ?? Path.GetFileNameWithoutExtension(path);
                var folderUri = new Uri("file:///" + Path.GetDirectoryName(path)!.Replace('\\', '/') + "/");
                var ff = new SMB.FontFamily(folderUri, "./#" + familyName);
                var addIndex = _fontChoices.FindIndex(fc => fc.IsAddItem);
                var newChoice = new FontChoice
                {
                    Name = familyName,
                    Family = ff
                };
                if (addIndex < 0)
                    _fontChoices.Add(newChoice);
                else
                    _fontChoices.Insert(addIndex, newChoice);
                FontFamilyCombo.ItemsSource = null;
                FontFamilyCombo.ItemsSource = _fontChoices;
                FontFamilyCombo.SelectedItem = newChoice;
                ApplyFontFamily(ff, applyParagraphWhenEmpty: true, focusEditorAfter: true);
            }
            catch (Exception ex)
            {
                SW.MessageBox.Show($"Unable to load font:\n{ex.Message}", "AnyFile Editor", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ApplyFontFamily(SMB.FontFamily family, bool applyParagraphWhenEmpty, bool focusEditorAfter)
        {
            EnsureInitialParagraph();
            var sel = Editor.Selection;
            if (sel.IsEmpty)
            {
                var tp = EnsureInsertionPosition();
                var para = tp.Paragraph;
                if (applyParagraphWhenEmpty && para != null && IsParagraphEmpty(para))
                    para.FontFamily = family;
                sel.Select(tp, tp);
            }

            sel.ApplyPropertyValue(TextElement.FontFamilyProperty, family);
            if (focusEditorAfter)
                Editor.Focus();
        }

        private void PaletteSwatch_Click(object sender, RoutedEventArgs e)
        {
            if (sender is SWC.Button b && b.Background is SolidColorBrush sb)
            {
                ApplyTextColor(sb.Color);
            }

            if (ColorPaletteBtn?.ContextMenu is { } cm)
                cm.IsOpen = false;
        }

        private void ColorPicker_Click(object sender, RoutedEventArgs e)
        {
            var cd = new WF.ColorDialog();
            if (cd.ShowDialog() == WF.DialogResult.OK)
            {
                var c = SMB.Color.FromArgb(cd.Color.A, cd.Color.R, cd.Color.G, cd.Color.B);
                ApplyTextColor(c);
            }

            if (ColorPaletteBtn?.ContextMenu is { } cm)
                cm.IsOpen = false;
        }

        private static WP.JustificationValues MapAlignment(TextAlignment a) => a switch
        {
            TextAlignment.Center => WP.JustificationValues.Center,
            TextAlignment.Right => WP.JustificationValues.Right,
            TextAlignment.Justify => WP.JustificationValues.Both,
            _ => WP.JustificationValues.Left
        };
    }
}
