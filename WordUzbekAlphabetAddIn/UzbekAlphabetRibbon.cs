using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace WordUzbekAlphabetAddIn
{
    [ComVisible(true)]
    public sealed class UzbekAlphabetRibbon : IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        private IRibbonUI _ribbon;

        public UzbekAlphabetRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='UzbekAlphabetTab' label='O‘zbek alifbosi'>
        <group id='UzbekAlphabetGroup' label='Yangi alifbo'>
          <toggleButton id='AlphabetMode'
                        label='Avtomatik almashtirish'
                        size='large'
                        imageMso='SpellingAndGrammar'
                        screentip='Yangi alifbo'
                        supertip='Ch, Sh, O‘ va G‘ yozilishini yangi alifboga darhol almashtiradi.'
                        getPressed='GetPressed'
                        onAction='ToggleConversion'/>
          <button id='ConvertSelection'
                  label='Преобразовать выделение'
                  imageMso='ReplaceDialog'
                  screentip='Преобразовать выделенный текст'
                  onAction='ConvertSelection'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public bool GetPressed(IRibbonControl control)
        {
            return _addIn.ConversionEnabled;
        }

        public void ToggleConversion(IRibbonControl control, bool pressed)
        {
            _addIn.SetConversionEnabled(pressed);
            _ribbon?.InvalidateControl("AlphabetMode");
        }

        public void ConvertSelection(IRibbonControl control)
        {
            _addIn.ConvertSelection();
        }
    }
}
