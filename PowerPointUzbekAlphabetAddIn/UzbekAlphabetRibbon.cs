using System.Runtime.InteropServices;
using Microsoft.Office.Core;
namespace PowerPointUzbekAlphabetAddIn
{
    [ComVisible(true)] public sealed class UzbekAlphabetRibbon : IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        public UzbekAlphabetRibbon(ThisAddIn addIn) { _addIn = addIn; }
        public string GetCustomUI(string ribbonId) { return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='UzbekAlphabetTab' label='O‘zbek alifbosi'><group id='UzbekAlphabetGroup' label='Yangi alifbo'><button id='ConvertSelection' label='Выделенный текст в новый алфавит' size='large' imageMso='ReplaceDialog' screentip='Преобразовать выделенный текст' onAction='ConvertSelection'/></group></tab></tabs></ribbon></customUI>"; }
        public void ConvertSelection(IRibbonControl control) { _addIn.ConvertSelection(); }
    }
}
