// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/InsertTableDialog.xaml.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Dialog view for inserting a table into the document.
// ============================================================================
// File: AnyFileEditor_fixed_all/InsertTableDialog.xaml.cs
// Purpose: Code-behind for a XAML view.
// Context: Event handlers, UI wiring
// Notes: Keep behavior unchanged; annotate intent/why.

namespace TxtOrganizer
{
    /// <summary>InsertTableDialog — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public partial class InsertTableDialog : MahApps.Metro.Controls.MetroWindow
    {
        public int Rows    => (int)(RowsUpDown.Value ?? 0);
        public int Columns => (int)(ColsUpDown.Value ?? 0);

        public InsertTableDialog()
        {
            InitializeComponent();
        }

        private void Ok_Click(object sender, SW.RoutedEventArgs e)
        {
            if (Rows <= 0 || Columns <= 0)
            {
                SW.MessageBox.Show("Rows and columns must be at least 1.");
                return;
            }
            DialogResult = true;
        }
    }
}

