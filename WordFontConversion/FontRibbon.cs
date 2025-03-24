using System;
using Microsoft.Office.Tools.Ribbon;
using WordFontConversion.NepaliFont;
using System.Windows.Forms;
using System.Net;
using System.Diagnostics;

namespace WordFontConversion
{
    public partial class FontRibbon
    {
        private void FontRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void fontConverterButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (fontConverterButton.Checked)
            {
                SetNepaliFontConversion();
            }
            else
                DisableNepaliFontConversion();
        }

        /// <summary>
        /// Validates if the current environment is suitable for Nepali font conversion.
        /// </summary>
        /// <returns></returns>
        private Boolean IsValidFontSettings()
        {
            // Check remote font conversion server.
            if (FontService.IsNotRemoteServerAvailable())
            {
                {
                    MessageBox.Show(ResponseLiterals.MESSAGE_SERVICE_NOT_AVAIABLE, ResponseLiterals.HEADING_SERICE_NOT_AVAILABLE, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    fontConverterButton.Checked = false;
                    return false;
                }
            }
            else return true;
        }

        private void SetNepaliFontConversion()
        {
            if (!IsValidFontSettings())
            {
                fontConverterButton.Checked = false;
                return;
            }
            this.nepaliFont.Label = FontPluginLiterals.FONT_ACTION_ON;  
            Globals.ThisAddIn.fontServiceEnabled = true;
        }

        private void DisableNepaliFontConversion()
        {
            this.nepaliFont.Label = FontPluginLiterals.FONT_ACTION_OFF;
            Globals.ThisAddIn.fontServiceEnabled = false;
        }
    }
}
