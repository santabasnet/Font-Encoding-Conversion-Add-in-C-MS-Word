using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordFontConversion.NepaliFont
{
    class ConversionMenuHelper
    {
        /// <summary>
        /// Constant representation of tag to Nepali Spelling Check command.
        /// </summary>
        public static readonly String COMMAND_BAR_POPUP_TAG = "nepali_font_conversion";

        /// <summary>
        /// Represents Nepali spelling check option literal.
        /// </summary>
        public static readonly String NEPALI_FONT_CHECK_COMMAND = "नेपाली फन्ट रूपान्तरण";

        /// <summary>
        /// Text option for command bar.
        /// </summary>
        public static readonly String COMMAND_BAR_TEXT = "Text";

        /// <summary>
        /// Text option for command bar in tables.
        /// </summary>
        public static readonly String COMMAND_BAR_TABLES = "Table Text";

        /// <summary>
        /// Returns the supported list of fonts for the current selection. If there is multiple fonts selected, 
        /// it returns all the supported targeted fonts.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns>listOfSupportedFonts.</returns>
        public static List<String> FindTargetedFonts(Word.Selection currentSelection)
        {
            if (String.IsNullOrWhiteSpace(currentSelection.Font.Name)) return COMUtility.AllSupportedFonts();

            String currentFont = COMUtility.UnicodeNameOf(currentSelection);
            return COMUtility.AllSupportedFonts().Where(fontName => fontName != currentFont).ToList();
        }
    }
}
