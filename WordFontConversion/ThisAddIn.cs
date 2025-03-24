using System;
using System.Collections.Generic;
using System.Linq;
using WordFontConversion.NepaliFont;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Text;

namespace WordFontConversion
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Application ID for the font conversion plugin.
        /// </summary>
        public String fontAppId = null;

        /// <summary>
        /// Font service flag.
        /// </summary>
        public bool fontServiceEnabled = false;

        /// <summary>
        /// MS word app, name alias.
        /// </summary>
        public Word.Application fontConversionApp = null;

        public Word.Selection currentSelection = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeFontApplication();
            Globals.ThisAddIn.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            ((Word.ApplicationEvents4_Event)Application).NewDocument += Application_NewDocument;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            InitializeFontApplication();
        }

        private void Application_NewDocument(Word.Document Doc)
        {
            InitializeFontApplication();
        }

        /// <summary>
        /// Perform initialization of font application environment.
        /// </summary>
        private void InitializeFontApplication()
        {
            this.fontAppId = COMUtility.GenerateAppId();
            if (this.fontConversionApp == null) this.fontConversionApp = Application;

            Globals.ThisAddIn.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            Globals.ThisAddIn.Application.WindowSelectionChange +=
               new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Globals.ThisAddIn.Application.WindowBeforeRightClick +=
                new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            
        }

        private void Application_WindowSelectionChange(Word.Selection currentSelection)
        {
        }

        /// <summary>
        /// Perform operations after right click happened in the application.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <param name="cancel"></param>
        private void Application_WindowBeforeRightClick(Word.Selection currentSelection, ref bool cancel)
        {
            Boolean isEligible = Globals.ThisAddIn.fontServiceEnabled && COMUtility.IsNepaliTextSelected(currentSelection);
            if (!isEligible)
            {
                RemoveNepaliFontContextMenu(currentSelection);
                return;
            }

            this.currentSelection = currentSelection;
            List<String> targetFonts = ConversionMenuHelper.FindTargetedFonts(currentSelection);
            if ((Boolean)currentSelection.Information[Word.WdInformation.wdWithInTable]) AddNepaliFontContextMenuInTable(targetFonts);
            else AddNepaliFontContextMenu(targetFonts);        
        }

        /// <summary>
        /// Add context menu for the given font selection.
        /// </summary>
        /// <param name="targetFonts"></param>
        private void AddNepaliFontContextMenu(List<String> targetFonts)
        {
            RemoveNepaliFontContextMenu(this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TEXT]);           
            CommandBar fontCommandBar = this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TEXT];
            CommandBarPopup parentCommandBarControl = CreateSubMenu(fontCommandBar);
            GenerateFontOptions(parentCommandBarControl, targetFonts);
        }

        /// <summary>
        /// Build and display context menu for the current font conversion in Nepali.
        /// </summary>
        private void AddNepaliFontContextMenuInTable(List<String> targetFonts)
        {            
            RemoveNepaliFontContextMenu(this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TABLES]);
            CommandBar fontCommandBar = this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TABLES];
            CommandBarPopup parentCommandBarControl = CreateSubMenu(fontCommandBar);
            GenerateFontOptions(parentCommandBarControl, targetFonts);
        }

        /// <summary>
        /// Creates the sub-menu for Nepali font conversion after right click.
        /// </summary>
        /// <param name="fontCommandBar"></param>
        /// <returns></returns>
        private CommandBarPopup CreateSubMenu(CommandBar fontCommandBar)
        {
            bool isFound = false;
            CommandBarPopup parentCommandBarControl = null;
            foreach (var commandBarPopup in fontCommandBar.Controls.OfType<CommandBarPopup>())
            {
                if (commandBarPopup.Tag.Equals(ConversionMenuHelper.COMMAND_BAR_POPUP_TAG))
                {
                    isFound = true;
                    parentCommandBarControl = commandBarPopup;
                    break;
                }
            }
            if (!isFound)
            {
                parentCommandBarControl = (CommandBarPopup)fontCommandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
                parentCommandBarControl.Caption = ConversionMenuHelper.NEPALI_FONT_CHECK_COMMAND;
                parentCommandBarControl.Tag = ConversionMenuHelper.COMMAND_BAR_POPUP_TAG;
                parentCommandBarControl.Visible = true;
            }
            return parentCommandBarControl;
        }

        /// <summary>
        /// Generates the suggestion menu for the current Nepali word.
        /// </summary>
        /// <param name="suggestionOptions"></param>
        /// <param name="currentWord"></param>
        private void GenerateFontOptions(CommandBarPopup fontOptions, List<String> targetFonts)
        {
            targetFonts.ForEach(fontName => BuildContextMenuItem(fontOptions, fontName));
        }

        /// <summary>
        /// Builds the menu item for the font name.
        /// </summary>
        /// <param name="fontOptions"></param>
        /// <param name="fontName"></param>
        private void BuildContextMenuItem(CommandBarPopup fontOptions, String fontName)
        {
            var commandBarButton = (CommandBarButton)fontOptions
                .Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, false);
            commandBarButton.Click += FontConversionEventHandler;
            commandBarButton.Caption = fontName;
            commandBarButton.FaceId = 340;
            commandBarButton.Tag = COMUtility.BuildFontNameTag(fontName);
            commandBarButton.BeginGroup = false;
            commandBarButton = null;
        }

        /// <summary>
        /// Perform replacement of text after font conversion.
        /// </summary>
        /// <param name="fontNameControl"></param>
        /// <param name="CancelDefault"></param>
        private void FontConversionEventHandler(CommandBarButton fontNameControl, ref bool CancelDefault)
        {
            /* *
             * Initialize current range before processing.
             * */           
            Word.Range workingRange = this.currentSelection.Range.Duplicate;
   
            /* *
             * Perform conversion as well as render the converted text here. If the system is unable to make converstion, 
             * it leaves such section to the original one.
             * */
            String targetFont = COMUtility.ServerFontNameOf(fontNameControl.Caption);
            List<SayakConversion> successConversions = CollectConversionRanges(workingRange)
                .Select(conversionRange => PerformFontConversion(conversionRange, targetFont))
                .Where(converstion => !converstion.IsEmpty())
                .ToList();

            //Debug.WriteLine("Total: " + successConversions.Count);
            //string s = string.Join(";", successConversions.Find(result => result.IsFailedStatus()).serverResponse.Select(x => x.Key + "=" + x.Value).ToArray());
            //Debug.WriteLine(s);

            /* *
             * Check if there are some messages from the server part to show the client like,
             * subscription expired or payment not happened yet.
             * In such case, there will be message shown to redirect the Sayak web site.
             * */
            if (successConversions.Any(result => result.IsFailedStatus())) {
                SayakConversion sayakConversion = successConversions.Find(conversion => conversion.IsFailedStatus());
                Tuple<String, String> remoteURLMessage = sayakConversion.BuildRemoteURLMessage();
                DialogResult result = MessageBox.Show(remoteURLMessage.Item2, FontPluginLiterals.SAYAK_SERVICE_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (result == DialogResult.Yes) System.Diagnostics.Process.Start(remoteURLMessage.Item1);
            }
            workingRange.Select();
        }

        /// <summary>
        /// Convert the given ranage of the document with the specified target font.
        /// </summary>
        /// <param name="conversionRange"></param>
        /// <param name="targetedFont"></param>
        /// <returns>conversionResult</returns>
        private SayakConversion PerformFontConversion(Word.Range conversionRange, String targetedFont)
        {            
            Tuple<Boolean, String> successSourceFont = COMUtility.SourceFontOfServer(conversionRange.Font.Name, conversionRange);
            SayakConversion sayakConversion;
            if (successSourceFont.Item1){
                sayakConversion = FontService.SayakConversionOf(successSourceFont.Item2, targetedFont, conversionRange.Text);
                if (sayakConversion.IsSuccessStatus()) WriteSuccessResult(targetedFont, sayakConversion, conversionRange);
            }
            else sayakConversion = SayakConversion.Empty();
            return sayakConversion;
        }

        /// <summary>
        /// Collect conversion ranges in terms of group of words.
        /// </summary>
        /// <returns>listOfWordRanges</returns>
        private List<Word.Range> CollectConversionRanges(Word.Range workingRange)
        {            
            /* *
             * Step 1. Check if the current selection has no input.
             * */
            if (workingRange.Words.Count < 1) return new List<Word.Range>();

            /*Debug.WriteLine("\n" + workingRange.Words.Count);
            foreach(Word.Range word in workingRange.Words)
            {
                Debug.WriteLine("Word: {" + word.Text + "} FontName: " + word.Font.Name + "\tIs Nepali Unicode: " + COMUtility.IsNepaliUnicodeText(word, true));            
            }*/

            /* *
             * Perform grouping based on the the sequece of same font ranges.
             * */
            List<Tuple<int, int, string>> groupedRanges = new List<Tuple<int, int, string>>() { WordRangeTuple(workingRange.Words.First) };
            foreach(Word.Range word in workingRange.Words)
            {
                var lastRange = groupedRanges.Last();
                if (String.IsNullOrWhiteSpace(word.Text) || lastRange.Item3 == word.Font.Name)
                {                    
                    var newRange = new Tuple<int, int, string>(lastRange.Item1, word.End, lastRange.Item3);
                    groupedRanges.RemoveAt(groupedRanges.Count - 1);
                    groupedRanges.Add(newRange);
                }
                else
                {
                    groupedRanges.Add(new Tuple<int, int, string>(word.Start, word.End, word.Font.Name));
                }
            }

            /* *
             * Convert the sequence of font ranges to the word range with reference to the
             * current active document, and finally returns the result.
             * */
            return groupedRanges.Select(groupedRange => TupleToWordRange(groupedRange)).ToList();
        }

        /// <summary>
        /// Perform mapping of word range to the tuple representation for grouping.
        /// </summary>
        /// <param name="wordRange"></param>
        /// <returns>tupleOfIndixesWithFontName</returns>
        private Tuple<int, int, string> WordRangeTuple(Word.Range wordRange)
        {
            int start = wordRange.Start;
            int end = wordRange.End;
            string fontName = wordRange.Font.Name;
            return new Tuple<int, int, string>(start, end, fontName);
        }

        /// <summary>
        /// Perform mapping of tuple to the word range represenation to make inverse operation back
        /// to word range.
        /// </summary>
        /// <param name="rangeTuple"></param>
        /// <returns>wordRangeOfIndexes</returns>
        private Word.Range TupleToWordRange(Tuple<int, int, string> rangeTuple)
        {
            object start = rangeTuple.Item1;
            object end = rangeTuple.Item2;

            // #Prevent to merge with next formatting.
            if (String.IsNullOrWhiteSpace(this.currentSelection.Document.Words.Last.Characters.Last.Text))
                end = rangeTuple.Item2 - 1;

            var newRange = this.currentSelection.Document.Range(start, end);
            return newRange;
        }

        /// <summary>
        /// Write conversion result to the document.
        /// </summary>
        /// <param name="targetedFont"></param>
        /// <param name="sayakConversion"></param>
        private void WriteSuccessResult(String targetedFont, SayakConversion sayakConversion, Word.Range conversionRange)
        {
            String conversionResult = sayakConversion.ConverionResult();
            String renderFont = COMUtility.LocalFontNameOf(targetedFont);

            /**
             * This seciton is necessary for the TTF rendering to avoid the common characters that
             * are available both UTF and TTF encoded Nepali font and the render is begin with 
             * TTF font.
             * */
            if (COMUtility.IsTTFKey(targetedFont)) conversionRange.InsertBefore(FontPluginLiterals.TTF_PREFIX);

            /**
             * Replace the converted font data.
             * */
            object startLocation = conversionRange.Start;
            object endLocation = conversionRange.End;

            Word.Range currentRange = conversionRange.Document.Range(ref startLocation, ref endLocation);
            currentRange.Text = conversionResult;
            currentRange.Font.Name = renderFont;
        }

        /// <summary>
        /// Write conversion result to the document.
        /// </summary>
        /// <param name="targetedFont"></param>
        /// <param name="sayakConversion"></param>
        private void WriteSuccessResult(String targetedFont, SayakConversion sayakConversion)
        {
            WriteSuccessResult(targetedFont, sayakConversion, this.currentSelection.Range);                   
        }

        /// <summary>
        /// Removes the context menu for Nepali Font Conversion System.
        /// </summary>
        /// <param name="spellingCommandBar"></param>
        private void RemoveNepaliFontContextMenu(Word.Selection currentSelection)
        {
            if ((bool)currentSelection.Information[Word.WdInformation.wdWithInTable])
                RemoveNepaliFontContextMenu(this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TABLES]);
            else
                RemoveNepaliFontContextMenu(this.fontConversionApp.CommandBars[ConversionMenuHelper.COMMAND_BAR_TEXT]);
        }

        /// <summary>
        /// Removes the context menu for Nepali Spelling Check.
        /// </summary>
        /// <param name="fontCommandBar"></param>
        private void RemoveNepaliFontContextMenu(CommandBar fontCommandBar)
        {
            foreach (var commandBarPopup in fontCommandBar.Controls.OfType<CommandBarPopup>())
            {
                if (commandBarPopup.Tag.Equals(ConversionMenuHelper.COMMAND_BAR_POPUP_TAG))
                {
                    commandBarPopup.Delete();
                }
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
