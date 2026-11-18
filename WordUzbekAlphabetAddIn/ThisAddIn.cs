using System;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordUzbekAlphabetAddIn
{
    public partial class ThisAddIn
    {
        private bool _isReplacing;
        private System.Windows.Forms.Timer _conversionTimer;
        internal bool ConversionEnabled { get; private set; } = true;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Word does not reliably raise WindowSelectionChange for every typed
            // character. Poll on Word's UI thread so pairs are still recognized.
            _conversionTimer = new System.Windows.Forms.Timer { Interval = 100 };
            _conversionTimer.Tick += OnConversionTimerTick;
            _conversionTimer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_conversionTimer == null) return;
            _conversionTimer.Stop();
            _conversionTimer.Tick -= OnConversionTimerTick;
            _conversionTimer.Dispose();
            _conversionTimer = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new UzbekAlphabetRibbon(this);
        }

        internal void SetConversionEnabled(bool enabled)
        {
            ConversionEnabled = enabled;
        }

        private void OnConversionTimerTick(object sender, EventArgs e)
        {
            if (!ConversionEnabled || _isReplacing || Application.Documents.Count == 0)
                return;

            Word.Selection selection = Application.Selection;
            if (selection == null)
                return;

            // Only convert text typed at the insertion point in the main story.
            if (selection.Type != Word.WdSelectionType.wdSelectionIP ||
                selection.StoryType != Word.WdStoryType.wdMainTextStory)
                return;

            Word.Range cursor = selection.Range;
            if (cursor.Start < 2)
                return;

            Word.Range lastTwoCharacters = cursor.Duplicate;
            lastTwoCharacters.SetRange(cursor.Start - 2, cursor.Start);
            string replacement = UzbekAlphabetConverter.ConvertPair(lastTwoCharacters.Text);

            if (replacement == null)
                return;

            _isReplacing = true;
            try
            {
                int replacementStart = lastTwoCharacters.Start;
                lastTwoCharacters.Text = replacement;

                // Word can move the insertion point while a range is replaced.
                // Put it immediately after the single replacement character.
                Word.Range updatedCursor = Application.ActiveDocument.Range(
                    replacementStart + replacement.Length,
                    replacementStart + replacement.Length);
                updatedCursor.Select();
            }
            finally
            {
                _isReplacing = false;
            }
        }

        internal void ConvertSelection()
        {
            Word.Selection selection = Application.Selection;
            if (selection == null || selection.Type == Word.WdSelectionType.wdSelectionIP)
                return;

            Word.Range selectedRange = selection.Range.Duplicate;
            int originalStart = selectedRange.Start;
            string originalText = selectedRange.Text;
            int replacements = 0;

            _isReplacing = true;
            try
            {
                // Work from right to left so replacing a pair does not shift the
                // positions of pairs that are still to be processed.
                for (int i = originalText.Length - 2; i >= 0; i--)
                {
                    string replacement = UzbekAlphabetConverter.ConvertPair(originalText.Substring(i, 2));
                    if (replacement == null) continue;

                    Word.Range pairRange = selectedRange.Duplicate;
                    pairRange.SetRange(originalStart + i, originalStart + i + 2);
                    pairRange.Text = replacement;
                    replacements++;
                }

                if (replacements > 0)
                {
                    int newEnd = originalStart + originalText.Length - replacements;
                    Application.ActiveDocument.Range(originalStart, newEnd).Select();
                }
            }
            finally
            {
                _isReplacing = false;
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
