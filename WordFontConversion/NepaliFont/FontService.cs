using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;


namespace WordFontConversion.NepaliFont
{
    class FontService
    {
        /// <summary>
        ///  Verifies the availability of remote server by sending a single word to convert Nepali font.
        /// </summary>
        /// <returns>true if the remote server sends the status OK.</returns>
        public static Boolean IsRemoteServerAvailable()
        {
            RemoteFontData remoteFontData = RemoteFontData.ofUnicodeConversion(RequestLiterals.PREETI, RequestLiterals.DEMO_TEXT);
            var fontRequestResult = FontsRemoteIO.SendMultipartPostRequest(remoteFontData.ToString());
            return fontRequestResult.Item1 == HttpStatusCode.OK;
        }

        /// <summary>
        /// Negation of the availability of the remote server.
        /// </summary>
        /// <returns>true if the remote server does not sends the status OK.</returns>
        public static Boolean IsNotRemoteServerAvailable()
        {
            return !IsRemoteServerAvailable();
        }

        /// <summary>
        /// Perform font conversion using remote service.
        /// </summary>
        /// <param name="sourceFont"></param>
        /// <param name="destinationFont"></param>
        /// <param name="text"></param>
        /// <returns>sayakConversion instance as the conversion result.</returns>
        public static SayakConversion SayakConversionOf(String sourceFont, String destinationFont, String text)
        {
            RemoteFontData remoteFontData = RemoteFontData.ofConversion(sourceFont, destinationFont, text);
            String fontRequestResult = FontsRemoteIO.SendMultipartPostRequest(remoteFontData.ToString()).Item2;
            return SayakConversion.Of(fontRequestResult);
        }

    }

    public static class HttpContentExtension
    {
        public static async Task<string> ReadAsStringUTF8Async(this HttpContent content)
        {
            return await content.ReadAsStringAsync(Encoding.UTF8);
        }

        public static async Task<string> ReadAsStringAsync(this HttpContent content, Encoding encoding)
        {
            using (var reader = new StreamReader((await content.ReadAsStreamAsync()), encoding))
            {
                return reader.ReadToEnd();
            }
        }
    }

    class FontsRemoteIO
    {
        /// <summary>
        /// Default timeout for remote connection.
        /// </summary>
        private static readonly int DEFAULT_TIMEOUT = 5000;

        /// <summary>
        /// Spelling client to make remote request.
        /// </summary>
        private static readonly HttpClient fontClient = initClient();

        /// <summary>
        /// Remote server url for spelling check.
        /// </summary>
        private static readonly String SERVER_URL = FontSettings.Default.remoteURL;

        /// <summary>
        /// Sends asynchronuus post reqeust to the spelling server.
        /// </summary>
        /// <param name="jsonBody"></param>
        /// <returns></returns>
        public static Tuple<HttpStatusCode, String> SendMultipartPostRequest(String jsonBody)
        {
            /*try
            {
                var multipartData = new MultipartFormDataContent();
                var jsonContent = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                multipartData.Add(jsonContent, String.Format("\"{0}\"", RequestLiterals.DATA));
                using (HttpResponseMessage fontResponse = fontClient.PostAsync(SERVER_URL, multipartData).Result)
                {
                    if (fontResponse.IsSuccessStatusCode)
                    {
                        var result = fontResponse.Content.ReadAsStringUTF8Async().Result;
                        return new Tuple<HttpStatusCode, string>(fontResponse.StatusCode, result);
                    }
                    else
                        return new Tuple<HttpStatusCode, String>(HttpStatusCode.BadRequest, ResponseLiterals.FailedResponse(ResponseLiterals.FAILED_TO_READ_FROM_SERVER));
                }
            }
            catch(Exception e)
            {
                return new Tuple<HttpStatusCode, String>(HttpStatusCode.BadRequest, ResponseLiterals.FailedResponse(e.Message));
            }*/

            MultiPartFormUpload multiPartFormUpload = new MultiPartFormUpload();
            var headers = new NameValueCollection() { };
            var files = new List<FileInfo>() { };
            try
            {
                MultiPartFormUpload.UploadResponse response = multiPartFormUpload.Upload(SERVER_URL, headers, jsonContent(jsonBody), files);
                var code = response.HttpStatusCode;
                var body = response.ResponseBody;
                return new Tuple<HttpStatusCode, string>(code, body);
            }
            catch (Exception ex)
            {
                return new Tuple<HttpStatusCode, String>(HttpStatusCode.BadRequest, ResponseLiterals.FailedResponse(ex.Message));
            }

        }

        /// <summary>
        /// Builds name value collection for the body content.
        /// </summary>
        /// <param name="jsonBody"></param>
        /// <returns>bodyContent</returns>
        private static NameValueCollection jsonContent(String jsonBody)
        {
            NameValueCollection bodyContent = new NameValueCollection() { };
            bodyContent.Add(RequestLiterals.DATA, jsonBody);
            return bodyContent;
        }

        /// <summary>
        /// Constructus service new URL using seetting variable and generated AppID.
        /// </summary>
        /// <returns>serviceRenewRequestURL</returns>
        public static String BuildServiceRenewURL()
        {
            return $"{FontSettings.Default.serviceRenewURL}?appId={COMUtility.GenerateAppId()}";
        }

        /// <summary>
        /// Constructs service payment URL using setting variable and generated AppID.
        /// </summary>
        /// <returns>servicePaymentRequestURL</returns>
        public static String BuildServicePaymentURL()
        {
            return $"{FontSettings.Default.servicePaymentURL}?appId={COMUtility.GenerateAppId()}";
        }

        /// <summary>
        /// Constructs default contact URL.
        /// </summary>
        /// <returns>contactURL</returns>
        public static String BuildContactURL()
        {
            return "https://hijje.com/#/user/contact";
        }

        /// <summary>
        /// Initialize the http client object.
        /// </summary>
        /// <returns></returns>
        private static HttpClient initClient()
        {
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json; charset=utf-8");
            httpClient.Timeout = TimeSpan.FromMilliseconds(DEFAULT_TIMEOUT);
            return httpClient;
        }
    }

    class FontData
    {
        public String sourceFont { get; set; }
        public String destinationFont { get; set; }
        public String text { get; set; }

        public FontData(String sourceFont, String destinationFont, String text)
        {
            this.sourceFont = sourceFont;
            this.destinationFont = destinationFont;
            this.text = text;
        }

        public Boolean IsEmpty()
        {
            return String.IsNullOrWhiteSpace(this.text) || String.IsNullOrWhiteSpace(this.sourceFont) || String.IsNullOrWhiteSpace(this.destinationFont);
        }

        public Boolean IsValid()
        {
            return !IsEmpty();
        }

        override
        public String ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static FontData of(String sourceFont, String destinationFont, String text)
        {
            return new FontData(sourceFont, destinationFont, text);
        }

        public static FontData ofUnicodeConversion(String sourceFont, String text)
        {
            return FontData.of(sourceFont, RequestLiterals.UNICODE, text);
        }

        public static FontData ofTTFConversion(String destinationFont, String text)
        {
            return FontData.of(RequestLiterals.UNICODE, destinationFont, text);
        }
    }

    class LanguageParameters
    {
        public String language { get; set; } = FontPluginLiterals.DEFAULT_LANGUAGE;
        public List<FontData> words { get; set; } = FontPluginLiterals.EMPTY_FONT_DATA;

        public LanguageParameters(String language, List<FontData> words)
        {
            this.language = language;
            this.words = words;
        }

        override
        public String ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static LanguageParameters of(List<FontData> words)
        {
            return new LanguageParameters(FontPluginLiterals.DEFAULT_LANGUAGE, words);
        }
    }

    class RemoteFontData
    {
        public String requestMethod { get; set; }
        public LanguageParameters languageParams { get; set; }
        public String token { get; set; } = FontPluginLiterals.EMPTY;
        public String wordPluginId { get; set; }
        
        public RemoteFontData(String requestMethod, LanguageParameters languageParams, String wordPluginId, String token = "")
        {
            this.requestMethod = requestMethod;
            this.languageParams = languageParams;
            this.wordPluginId = wordPluginId;
            this.token = token;
        }

        override
        public String ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static RemoteFontData ofConversion(String sourceFont, String destinationFont, String text)
        {
            return new RemoteFontData(
                    RequestLiterals.ACTION_FONT_CONVERSION,
                    LanguageParameters.of(new List<FontData>() { FontData.of(sourceFont, destinationFont, text) }),
                    Globals.ThisAddIn.fontAppId
                );
        }

        public static RemoteFontData ofUnicodeConversion(String sourceFont, String text)
        {
            return RemoteFontData.ofConversion(sourceFont, RequestLiterals.UNICODE, text);
        }

        public static RemoteFontData ofTTFConversion(String destinationFont, String text)
        {
            return RemoteFontData.ofConversion(RequestLiterals.UNICODE, destinationFont, text);
        }
    }

    class SayakConversion
    {
        public Dictionary<String, Object> serverResponse { get; set; }

        public SayakConversion()
        {
            this.serverResponse = new Dictionary<String, Object>();
        }

        public SayakConversion(String serverResponse)
        {
            this.serverResponse = JsonConvert.DeserializeObject<Dictionary<String, Object>>(serverResponse);
        }

        public Boolean IsEmpty()
        {
            return this.serverResponse.Count < 1;
        }

        public Boolean IsFailedStatus()
        {
            return !IsSuccessStatus();
        }

        public Boolean IsSuccessStatus()
        {            
            string status = !this.serverResponse.ContainsKey(ResponseLiterals.STATUS) ? FontPluginLiterals.EMPTY : this.serverResponse[ResponseLiterals.STATUS].ToString();
            return status == ResponseLiterals.SUCCESS;
        }

        public Boolean IsLicenseExpired()
        {
            return this.serverResponse[ResponseLiterals.MESSAGE_ID].ToString() == ResponseLiterals.REDIRECT_TO_SERVICE_RENEW;
        }

        public Boolean IsTrialOver()
        {
            return this.serverResponse[ResponseLiterals.MESSAGE_ID].ToString() == ResponseLiterals.REDIRECT_TO_SERVICE_PAYMENT;
        }

        public String ConverionResult()
        {
            return !this.serverResponse.ContainsKey(ResponseLiterals.STATUS) ? FontPluginLiterals.EMPTY : this.serverResponse[ResponseLiterals.RESULT].ToString();
        }

        public String MessageId()
        {
            return this.serverResponse[ResponseLiterals.MESSAGE_ID].ToString();
        }

        public static SayakConversion Of(String serverResponse)
        {
            return new SayakConversion(serverResponse);
        }

        public static SayakConversion Empty()
        {
            return new SayakConversion();
        }

        /// <summary>
        /// Constructs remote message to display for the font service user.
        /// </summary>
        /// <param name="sayakConversion"></param>
        /// <returns>tupleOfURLMessage</returns>
        public Tuple<String, String> BuildRemoteURLMessage()
        {
            String remoteURL;
            String remoteMessage;
            if (this.IsLicenseExpired())
            {
                remoteURL = FontsRemoteIO.BuildServiceRenewURL();
                remoteMessage = $"Nepali Font Service licence expired!\r\n Goto Sayak Renew Page : {remoteURL}";
            }
            else if (this.IsTrialOver())
            {
                remoteURL = FontsRemoteIO.BuildServicePaymentURL();
                remoteMessage = $"Nepali Font Service trial is over!\r\n Goto Sayak Payment Page : {remoteURL}";
            }
            else
            {
                remoteURL = FontsRemoteIO.BuildContactURL();
                remoteMessage = $"Unable to make font converion. Goto Sayak contact page : {remoteURL}";
            }
            return Tuple.Create(remoteURL, remoteMessage);
        }

    }


    class ResponseLiterals
    {
        /// <summary>
        /// Message literals.
        /// </summary>
        public static readonly string TRUE = "true";
        public static readonly string FALSE = "false";
        public static readonly string DEFAULT_STATUS = FALSE;
        public static readonly string REDIRECT_TO_SERVICE_PAYMENT = "REDIRECT_TO_SERVICE_PAYMENT";
        public static readonly string REDIRECT_TO_SERVICE_RENEW = "REDIRECT_TO_SERVICE_RENEW";
        public static readonly string SERVICE_NOT_AVAILABLE = "SERVICE_NOT_AVAILABLE";
        public static readonly string MESSAGE_TRIAL_OVER = "Trial Period Over!, We request you to register and make payment.";
        public static readonly string MESSAGE_SERVICE_EXPIRED = "Sorry, Your subscribed service duration is expired! We request you extend the font service.";
        public static readonly string MESSAGE_SERVICE_NOT_AVAIABLE = "Sorry, The remote font serive is currently unavailable, please try after sometime. Exiting now ...";
        public static readonly string FAILED_TO_READ_FROM_SERVER = "Sorry, unable to read the data from server.";
        public static readonly string HEADING_SERICE_NOT_AVAILABLE = "Service Not Available !";
        public static readonly string SAYAK_RESPONSE_TYPE = "SayakResponse";
        public static readonly string SEMANTRO_CONTEXT = "http://semantro.com";
        public static readonly string SEMANTRO_SAYAK_RESPONSE_ID = "http://semantro.com/SayakResponse";
        public static readonly string ID_PLACE_HOLDER = "@";

        /// <summary>
        /// Variable literals.
        /// </summary>
        public static readonly string STATUS = "status";
        public static readonly string MESSAGE = "message";
        public static readonly string MESSAGE_ID = "messageId";
        public static readonly string TYPE = "type";
        public static readonly string CONTEXT = "context";
        public static readonly string ID = "id";
        public static readonly string SUCCESS = "success";
        public static readonly string FAIL = "fail";
        public static readonly string RESULT = "result";

        public static String FailedResponse(String message)
        {
            var failedResponse = new Dictionary<String, String>()
            {
                { "status", "fail" },
                { "result", message },
                { "messageId", "रूपानतरण असफल भयो ।" }
            };

            return JsonConvert.SerializeObject(failedResponse, Formatting.Indented);
        }
    }

    class RequestLiterals
    {
        public static readonly string DATA = "data";
        public static readonly string ACTION_FONT_CONVERSION = "fontconversion";


        /// <summary>
        /// Font Data literals.
        /// </summary>
        public static readonly string UNICODE = "Unicode";
        public static readonly string UNICODE_NEPALI = "यूनिकोड";
        public static readonly string PREETI = "Preeti";
        public static readonly string DEMO_TEXT = "g]kfnL";

    }

    class FontPluginLiterals
    {
        public static readonly string EMPTY = "";
        public static readonly string TOKEN_SEPARATOR = "-";
        public static readonly string TTF_PREFIX = "a";
        public static readonly List<FontData> EMPTY_FONT_DATA = new List<FontData>();
        public static readonly string DEFAULT_LANGUAGE = "ne";
        public static readonly string FONT_ACTION_ON = "Font Action: ON";
        public static readonly string FONT_ACTION_OFF = "Font Action: OFF";
        public static readonly string SAYAK_SERVICE_NAME = "Sayak Font Service";
        public static readonly string UNSUPPORTED_FONT = "Unsupported Font";
        public static readonly string MESSAGE_UNSUPPORTED_FONT = "Current font is not supported yet to make font conversion.";

    }
}
