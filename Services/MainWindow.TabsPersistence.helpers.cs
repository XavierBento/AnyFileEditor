// ============================================================================
// Project   : AnyFile Editor (TxtOrganizer)
// File      : Services/MainWindow.TabsPersistence.helpers.cs
// Author    : Xavier Bento
// Version   : v1.0
// Created   : 2025-09-21
// Description: Helper methods for persisting/restoring tab state.
// ============================================================================
// File: AnyFileEditor_fixed_all/MainWindow.TabsPersistence.helpers.cs
// Purpose: Tab management helpers for the main window (add/close/select/+ tab).
// Context: TabControl coordination, TabsPersistence helpers (if present)
// Notes: Invariant: internal _tabs list and TabControl.Items must remain index-aligned.
using System.Linq;
using System.Windows.Documents;

// Invariant: the backing tab list and the visual TabControl must remain index-aligned;
// selection, persistence, and close/add operations assume parity.



namespace TxtOrganizer
{
    /// <summary>MainWindow — role and responsibilities within the AnyFile Editor app.</summary>
/// <remarks>Documented without behavior changes on 2025-09-21.</remarks>
    public partial class MainWindow
    {
        private (SWC.TabItem? item, DocTab? meta) GetActiveDocTab()
        {
            if (_editorTabs?.SelectedItem is SWC.TabItem ti && ti.Tag is DocTab meta) return (ti, meta);
            return (null, null);
        }

        private void SaveActiveDocToTab(SWC.TabItem? tabFromEvent = null)
        {
            var ti = tabFromEvent ?? (_editorTabs?.SelectedItem as SWC.TabItem);
            if (ti?.Tag is DocTab meta && Editor != null && Editor.Document != null)
                meta.Document = Editor.Document;
        }
    }
}

