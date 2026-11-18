using System.Runtime.InteropServices;
using Microsoft.Office.Core;
namespace ExcelUzbekAlphabetAddIn
{
    [ComVisible(true)] public sealed class UzbekAlphabetRibbon : IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        public UzbekAlphabetRibbon(ThisAddIn addIn) { _addIn = addIn; }
        public string GetCustomUI(string ribbonId) { return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='UzbekAlphabetTab' label='O‘zbek alifbosi'><group id='UzbekAlphabetGroup' label='Yangi alifbo'><button id='ConvertSelection' label='Выделенное в новый алфавит' size='large' imageMso='ReplaceDialog' screentip='Преобразовать выделенные ячейки' onAction='ConvertSelection'/></group></tab></tabs></ribbon></customUI>"; }
        public void ConvertSelection(IRibbonControl control) { _addIn.ConvertSelection(); }
    }
}
