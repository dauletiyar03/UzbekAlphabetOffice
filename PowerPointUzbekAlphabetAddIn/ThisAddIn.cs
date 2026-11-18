using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PowerPointUzbekAlphabetAddIn
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() { return new UzbekAlphabetRibbon(this); }
        internal void ConvertSelection()
        {
            PowerPoint.Selection selection = Application.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionText) return;
            PowerPoint.TextRange range = selection.TextRange;
            if (range == null) return;
            string original = range.Text;
            string converted = UzbekAlphabetConverter.ConvertText(original);
            if (converted != original) range.Text = converted;
        }
        private void InternalStartup() { }
    }
}
