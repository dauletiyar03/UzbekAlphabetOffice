using System;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelUzbekAlphabetAddIn
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() { return new UzbekAlphabetRibbon(this); }
        internal void ConvertSelection()
        {
            Excel.Range selected = Application.Selection as Excel.Range;
            if (selected == null) return;
            foreach (Excel.Range cell in selected.Cells)
            {
                if ((cell.HasFormula is bool hasFormula && hasFormula) || cell.Value2 == null) continue;
                string value = cell.Value2 as string;
                if (value == null) continue;
                string converted = UzbekAlphabetConverter.ConvertText(value);
                if (converted != value) cell.Value2 = converted;
            }
        }
        private void InternalStartup() { }
    }
}
