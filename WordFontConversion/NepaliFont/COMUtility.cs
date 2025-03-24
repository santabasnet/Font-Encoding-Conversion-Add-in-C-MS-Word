using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Word;
using System.Text;
using System.Security.Cryptography;
using System.Diagnostics;

namespace WordFontConversion.NepaliFont
{
    class COMUtility
    {

        /// <summary>
        /// Represents the selection type, and is used while making a right click.
        /// </summary>
        private static readonly String SELECTION_TYPE = "Selection";
        private static readonly String PREETI_KEY = "PREETI";
        private static readonly String KANTIPUR_KEY = "KANTIPUR";
        private static readonly String HIMALB_KEY = "HIMLAB";
        private static readonly String AAKRITI_KEY = "AAKRITI";
        private static readonly String AALEKH_KEY = "AALEKH";
        private static readonly String GANESS_KEY = "GANESS";
        private static readonly String NAVJEEVAN_KEY = "NAVJEEVAN";
        private static readonly String PCSNEPALI_KEY = "PCSNEPALI";
        private static readonly String SHANGRILA_KEY = "SHANGRILA";
        private static readonly String SHREENATH_KEY = "SHREENATH";
        private static readonly String SUMOD_KEY = "SUMOD";
        private static readonly String UNICODE_KEY = "UNICODE";

        private static readonly String UNICODE_NAME = "Mangal";

        /// <summary>
        /// List of Nepali Punctuations.
        /// </summary>
        private static readonly List<Char> NEPALI_PUNCTUATIONS = new List<Char>() {
            ',', ';', ':', '?', '!', '"', '—', '-'
        };

        /// <summary>
        /// Available fonts that are to be converted through font conversion services.
        /// </summary>
        private static readonly Dictionary<String, String> NEPALI_FONTS = new Dictionary<String, String>() {
            { PREETI_KEY, "प्रीति" },
            { KANTIPUR_KEY, "कान्तिपूर" },
            { HIMALB_KEY, "हिमाली" },
            { AAKRITI_KEY, "आकृति" },
            { AALEKH_KEY, "आलेख" },
            { GANESS_KEY, "गणेश" },
            { NAVJEEVAN_KEY, "नवजीवन" },
            { PCSNEPALI_KEY, "पीसीएस नेपाली" },
            { SHANGRILA_KEY, "साङ्ग्रिला" },
            { SHREENATH_KEY, "श्रीनाथ" },
            { SUMOD_KEY, "सूमोद" },
            { UNICODE_KEY, "यूनिकोड" }
        };

        /// <summary>
        /// Available fonts that have different name in the local system.
        /// </summary>
        private static readonly Dictionary<String, String> LOCAL_NEPALI_FONTS = new Dictionary<String, String>() {
            { PREETI_KEY, "Preeti" },
            { KANTIPUR_KEY, "Kantipur" },
            { HIMALB_KEY, "Himalb" },
            { AAKRITI_KEY, "Aakriti" },
            { AALEKH_KEY, "Aalekh" },
            { GANESS_KEY, "Ganess" },
            { NAVJEEVAN_KEY, "Navjeevan" },
            { PCSNEPALI_KEY, "PCS NEPALI" },
            { SHANGRILA_KEY, "Shangrila Numeric" },
            { SHREENATH_KEY, "Shreenath Bold" },
            { SUMOD_KEY, "Sumod Acharya" },
            { UNICODE_KEY, UNICODE_NAME }
        };

        /// <summary>
        /// Initialize the local font name once.
        /// </summary>
        /// <returns>localFontNames</returns>
        private static List<String> initializeLocalFontNames()
        {
            var fonts = LOCAL_NEPALI_FONTS.Values.ToList();
            fonts.Remove(UNICODE_NAME);
            fonts.Add(UNICODE_KEY);
            return fonts;
        }

        /// <summary>
        /// Initialize the TTF font names once.
        /// </summary>
        /// <returns></returns>
        private static List<String> initializeTTFFontNames()
        {
            var fontKeys = LOCAL_NEPALI_FONTS.Keys.ToList();
            fontKeys.Remove(UNICODE_KEY);
            return fontKeys;
        }

        /// <summary>
        /// Local nepal fonts.
        /// </summary>
        private static readonly List<String> LOCAL_NEPALI_FONT_NAMES = initializeLocalFontNames();

        /// <summary>
        /// All supported TTF fonts.
        /// </summary>
        private static readonly List<String> TTF_FONTS_KEYS = initializeTTFFontNames();

        /// <summary>
        /// Returns the unicode name for the given selection, because the context menu only supports the unicode 
        /// characters.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static String UnicodeNameOf(Word.Selection currentSelection)
        {
            String fontName = currentSelection.Font.Name;
            if (LOCAL_NEPALI_FONT_NAMES.Contains(fontName)) return UnicodeNameOf(fontName);
            else if (IsNepaliEncodedText(currentSelection)) return RequestLiterals.UNICODE_NEPALI; 
            else return FontPluginLiterals.EMPTY;
        }

        /// <summary>
        /// Varifies if the given key, ie. font key is of ttf type or not.
        /// </summary>
        /// <param name="key"></param>
        /// <returns>true if it is ttf key.</returns>
        public static Boolean IsTTFKey(String key)
        {
            return TTF_FONTS_KEYS.Contains(key);
        }

        /// <summary>
        /// Returns the given name of the font in Devanagari representation, 
        /// used for building context menu.
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns></returns>
        private static String UnicodeNameOf(String fontName)
        {
            var key = LOCAL_NEPALI_FONTS.FirstOrDefault(fontEntry => fontEntry.Value == fontName).Key;
            return NEPALI_FONTS[key];
        }

        /// <summary>
        /// Returns the list of Nepali Font Names, that are used to render during right click.
        /// </summary>
        /// <returns></returns>
        public static List<String> AllSupportedFonts()
        {
            return NEPALI_FONTS.Values.ToList();
        }

        /// <summary>
        /// It returns the name of the given unicode font name which is used to send to the
        /// server during font conversion.
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static String ServerFontNameOf(String fontName)
        {
            return NEPALI_FONTS.FirstOrDefault(fontEntry => fontEntry.Value == fontName).Key;
        }

        /// <summary>
        /// Local font name after font conversion.
        /// </summary>
        /// <param name="constantName"></param>
        /// <returns>localFontName</returns>
        public static String LocalFontNameOf(String constantName)
        {
            String localFontName;
            return LOCAL_NEPALI_FONTS.TryGetValue(constantName, out localFontName) ? localFontName : constantName;            
        }

        /// <summary>
        /// Source font detection before sending to the server.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="text"></param>
        /// <returns>tuple of success and the font name.</returns>
        public static Tuple<Boolean, String> SourceFontOfServer(String fontName, Word.Range conversionRange)
        {
            if (String.IsNullOrWhiteSpace(fontName)) return new Tuple<Boolean, String>(false, fontName);                                  
            else if (HasLocalName(fontName)) return new Tuple<Boolean, String>(true, SourceFontFromLocalName(fontName));
            else if (IsNepaliUnicodeText(conversionRange, true)) return new Tuple<Boolean, String>(true, RequestLiterals.UNICODE);
            else return new Tuple<Boolean, String>(false, fontName);
        }

        /// <summary>
        /// Verifies whether the content font is of local given name or not.
        /// </summary>
        /// <param name="contentFont"></param>
        /// <returns>true if the content font is of local name defined or not.</returns>
        private static Boolean HasLocalName(String contentFont)
        {
            return LOCAL_NEPALI_FONT_NAMES.Contains(contentFont);
        }

        /// <summary>
        /// Returns the conversion source font of the given font name in terms of given
        /// local name.
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns>sourceNameInLocalName</returns>
        private static String SourceFontFromLocalName(String fontName)
        {
            return LOCAL_NEPALI_FONTS.FirstOrDefault(fontEntry => fontEntry.Value == fontName).Key;
        }

        /// <summary>
        /// It identifies if the current selection has Nepali fonts and they are supported by the font
        /// conversion API.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static Boolean IsNepaliTextSelected(Word.Selection currentSelection)
        {
            return IsTextSelected(currentSelection) && IsNepaliEncodedText(currentSelection);
        }

        /// <summary>
        /// It generates the application ID based on the installed machine.
        /// </summary>
        /// <returns></returns>
        public static String GenerateAppId()
        {
            String appId = "fontconversion-9998";//GenerateId();
            //String appId = "fontconversion-8989";
            //return encryptId(appId);
            return appId;
        }

        /// <summary>
        /// Generates the unique ID for Client APP identity.
        /// </summary>
        /// <returns>appId</returns>
        private static String GenerateId()
        {
            List<String> macAddress = MacAddress();
            List<String> osInformation = OSInformation();
            List<String> generatedTokens = new List<String>()
            {
                ServiceName(),
                osInformation.First(),
                macAddress.First(),                
                macAddress.Last(),
                osInformation.Last()
            };
            return String.Join(FontPluginLiterals.TOKEN_SEPARATOR, generatedTokens);
        }

        /// <summary>
        /// Encryption with AES, with hashed sha256 (From Sachin).
        /// </summary>
        /// <param name="text"></param>
        /// <param name="key"></param>
        /// <returns>encryptedKey</returns>
        private static String encryptId(String text, String key = "h!jje2@1")
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.PKCS7;

            SHA256 sha256 = SHA256.Create();
            byte[] passwordHashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(key));
            byte[] passwordIvBytes = Encoding.UTF8.GetBytes(key);
            byte[] keyBytes = new byte[0x20];
            byte[] iVBytes = new byte[0x10];

            int lenIvBytes = passwordIvBytes.Length;
            int lenKeyBytes = passwordHashBytes.Length;
            Array.Copy(passwordHashBytes, keyBytes, lenKeyBytes);
            Array.Copy(passwordIvBytes, iVBytes, lenIvBytes);

            rijndaelCipher.Key = keyBytes;
            rijndaelCipher.IV = iVBytes;
            ICryptoTransform transform = rijndaelCipher.CreateEncryptor();
            byte[] plainText = Encoding.UTF8.GetBytes(text);

            return Convert.ToBase64String(transform.TransformFinalBlock(plainText, 0, plainText.Length)).Replace('+', '-').Replace('/', '_');
        }

        /// <summary>
        /// It determines the all the sections that are excluded to check the nepali spelling errors. 
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static Boolean IsIgnoredObjectSelected(Selection currentSelection)
        {
            if ((bool)currentSelection.Information[WdInformation.wdInFootnoteEndnotePane]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInCommentPane]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInClipboard]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInBibliography]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInCitation]) return true;
            if ((bool)currentSelection.Information[WdInformation.wdInContentControl]) return true;
            return false;
        }

        /// <summary>
        /// Identify the given selection whether it is in Nepali or not.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns>true if Nepali text is selected.</returns>
        public static Boolean IsNepaliEncodedText(Word.Selection currentSelection)
        {           
            List<Word.Range> wordRanges = ExtractWordRanges(currentSelection);
            return wordRanges.Any(wordRange => LOCAL_NEPALI_FONT_NAMES.Contains(wordRange.Font.Name) || IsNepaliUnicodeText(wordRange));
        }

        /// <summary>
        /// Maps the current selection to the list of word ranges.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns>listOfWordRanges</returns>
        private static List<Word.Range> ExtractWordRanges(Word.Selection currentSelection)
        {
            List<Word.Range> wordRanges = new List<Word.Range>();
            if (currentSelection == null) return wordRanges;

            foreach(Word.Range word in currentSelection.Range.Words) if(!String.IsNullOrEmpty(word.Font.Name)) wordRanges.Add(word);            
            return wordRanges;
        }

        /// <summary>
        /// Identifies if the current selection has all the Nepali Unicode Characters.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static Boolean IsNepaliUnicodeText(Word.Selection currentSelection, Boolean ignorePunctuation = false)
        {
            return IsDevanagariEncodedText(currentSelection.Text.Trim(), ignorePunctuation);
        }

        /// <summary>
        /// Identifies if the current range has all the Nepali Unicode Characters.
        /// </summary>
        /// <param name="currentSelection"></param>
        /// <returns></returns>
        public static Boolean IsNepaliUnicodeText(Word.Range currentRange, Boolean ignorePunctuation = false)
        {            
            return IsDevanagariEncodedText(currentRange.Text.Trim(), ignorePunctuation);
        }

        /// <summary>
        /// Verifies the Devanagari unicode range for the given text.
        /// </summary>
        /// <param name="currentText"></param>
        /// <returns>true if the given text contains all the characters in Devanagari.</returns>
        private static Boolean IsDevanagariEncodedText(String currentText, Boolean ignorePunctuation = false)
        {
            if (!ignorePunctuation) return currentText.Length == CountDevanagariCharacters(currentText);           

            int devanagariStats = CountDevanagariCharactersWithPunctuations(currentText);
            Double score = (Double)devanagariStats / (Double)currentText.Length;
            return score >= 0.94;
        }

        private static int CountDevanagariCharacters(String currentText)
        {
            return currentText.ToCharArray().Where(ch => ((ch >= 0x900 && ch <= 0x97f))).Count();
        }

        private static int CountDevanagariCharactersWithPunctuations(String currentText)
        {
            return currentText.ToCharArray().Where(ch => ((ch >= 0x900 && ch <= 0x97f) || COMUtility.IsNepaliPunctuation(ch) || Char.IsWhiteSpace(ch))).Count();
        }

        /// <summary>
        /// Generates list of Hex codes of 6 digits from the first machine MAC address.
        /// </summary>
        /// <returns>listOfMACAddressOfHexCodes</returns>
        private static List<String> MacAddress()
        {
            String macAddress = NetworkInterface
                .GetAllNetworkInterfaces()
                .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .Select(nic => nic.GetPhysicalAddress().ToString())
                .FirstOrDefault();

            return (from Match token in Regex.Matches(macAddress, @"\S{6}") select token.Value).ToList();
        }

        /// <summary>
        /// Generates the list of Hex codes of 6 digits from the current version of Installed OS.
        /// </summary>
        /// <returns>listOfOSInformationInHexCodes</returns>
        private static List<String> OSInformation()
        {
            String osInformation = String.Join(FontPluginLiterals.EMPTY, Environment.OSVersion.VersionString.ToCharArray().Select(letter => ((int)letter).ToString("X2")));
            return (from Match token in Regex.Matches(osInformation, @"\S{6}") select token.Value).ToList();
        }

        /// <summary>
        /// Generates the string represention of the Hex digits from the installed office version number.
        /// </summary>
        /// <returns>wordVersionInHexCode</returns>
        private static String WordVersion()
        {
            String wordVersion = Globals.ThisAddIn.Application.Version.ToString();
            return ((int)(Double.Parse(wordVersion))).ToString("X2");
        }

        /// <summary>
        /// Extracts service name from the settings.
        /// </summary>
        /// <returns>serviceName</returns>
        private static String ServiceName()
        {
            return FontSettings.Default.serviceName;
        }

        /// <summary>
        /// Verifies if the current selection is of text form or not.
        /// </summary>
        /// <param name="Sel"></param>
        /// <returns>true if text content selected.s</returns>
        private static Boolean IsTextSelected(Word.Selection Sel)
        {
            return Information.TypeName(Sel) == COMUtility.SELECTION_TYPE && !String.IsNullOrWhiteSpace(Sel.Text);
        }

        /// <summary>
        /// Verifies whether the given character lies among the Nepali puncturations or not.
        /// </summary>
        /// <param name="ch"></param>
        /// <returns>true if the given character is Nepali puncturation.</returns>
        public static Boolean IsNepaliPunctuation(Char ch)
        {
            return NEPALI_PUNCTUATIONS.Contains(ch);
        }

        /// <summary>
        /// Generates UUID for the font option during context menu generation.
        /// </summary>
        /// <param name="fontName"></param>
        /// <returns>fontNameTag</returns>
        public static String BuildFontNameTag(String fontName)
        {
            return String.Join("", new List<String>() { fontName, "|", System.Guid.NewGuid().ToString()});
        }


    }
}
