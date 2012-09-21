#region License
/*
 **************************************************************
 *  Author: Rick Strahl 
 *          © West Wind Technologies, 2008
 *          http://www.west-wind.com/
 * 
 * Created: 09/08/2008
 *
 * Permission is hereby granted, free of charge, to any person
 * obtaining a copy of this software and associated documentation
 * files (the "Software"), to deal in the Software without
 * restriction, including without limitation the rights to use,
 * copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the
 * Software is furnished to do so, subject to the following
 * conditions:
 * 
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 * OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 **************************************************************  
*/
#endregion

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Globalization;


namespace Westwind.Utilities
{   
    /// <summary>
    /// String utility class that provides a host of string related operations
    /// </summary>
    public partial class StringUtils
    {

        /// <summary>
        /// Replaces and  and Quote characters to HTML safe equivalents.
        /// </summary>
        /// <param name="html">HTML to convert</param>
        /// <returns>Returns an HTML string of the converted text</returns>
        public static string FixHTMLForDisplay(string html)
        {
            html = html.Replace("<", "&lt;");
            html = html.Replace(">", "&gt;");
            html = html.Replace("\"", "&quot;");
            return html;
        }

        /// <summary>
        /// Strips HTML tags out of an HTML string and returns just the text.
        /// </summary>
        /// <param name="html">Html String</param>
        /// <returns></returns>
        public static string StripHtml(string html)
        {
            html = Regex.Replace(html, @"<(.|\n)*?>", string.Empty);
            html = html.Replace("\t", " ");
            html = html.Replace("\r\n", string.Empty);
            html = html.Replace("   ", " ");
            return html.Replace("  ", " ");
        }

        /// <summary>
        /// Fixes a plain text field for display as HTML by replacing carriage returns 
        /// with the appropriate br and p tags for breaks.
        /// </summary>
        /// <param name="String Text">Input string</param>
        /// <returns>Fixed up string</returns>
        public static string DisplayMemo(string htmlText)
        {
            htmlText = htmlText.Replace("\r\n", "\r");
            htmlText = htmlText.Replace("\n", "\r");
            //HtmlText = HtmlText.Replace("\r\r","<p>");
            htmlText = htmlText.Replace("\r", "<br />");
            return htmlText;
        }

        /// <summary>
        /// Method that handles handles display of text by breaking text.
        /// Unlike the non-encoded version it encodes any embedded HTML text
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string DisplayMemoEncoded(string text)
        {
            bool PreTag = false;
            if (text.IndexOf("<pre>") > -1)
            {
                text = text.Replace("<pre>", "__pre__");
                text = text.Replace("</pre>", "__/pre__");
                PreTag = true;
            }


            // *** fix up line breaks into <br><p>
            text = DisplayMemo(HtmlEncode(text)); //HttpUtility.HtmlEncode(Text));

            if (PreTag)
            {
                text = text.Replace("__pre__", "<pre>");
                text = text.Replace("__/pre__", "</pre>");
            }

            return text;
        }

        /// <summary>
        /// Expands links into HTML hyperlinks inside of text or HTML.
        /// </summary>
        /// <param name="text">The text to expand</param>
        /// <param name="target">Target frame where links are displayed</param>
        /// <param name="parseFormattedLinks">Allows parsing of links in the following format [text|www.site.com]</param>
        /// <returns></returns>
        public static string ExpandUrls(string text, string target, bool parseFormattedLinks)
        {
            if (target == null)
                target = string.Empty;

            ExpandUrlsParser Parser = new ExpandUrlsParser();
            Parser.Target = target;
            Parser.ParseFormattedLinks = parseFormattedLinks;
            return Parser.ExpandUrls(text);
        }
        /// <summary>
        /// Expands links into HTML hyperlinks inside of text or HTML.
        /// </summary>
        /// <param name="text">The text to expand</param>
        /// <param name="target">Target frame where links are displayed</param>
        public static string ExpandUrls(string text, string target)
        {
            return ExpandUrls(text, null, false);
        }
        /// <summary>
        /// Expands links into HTML hyperlinks inside of text or HTML.
        /// </summary>
        /// <param name="text">The text to expand</param>
        public static string ExpandUrls(string text)
        {
            return ExpandUrls(text, null, false);
        }

        /// <summary>
        /// Create an Href HTML link
        /// </summary>
        /// <param name="text"></param>
        /// <param name="url"></param>
        /// <param name="target"></param>
        /// <param name="additionalMarkup"></param>
        /// <returns></returns>
        public static string Href(string text, string url, string target, string additionalMarkup)
        {
            return "<a href=\"" + url + "\" " +
                (string.IsNullOrEmpty(target) ? string.Empty : "target=\"" + target + "\" ") +
                (string.IsNullOrEmpty(additionalMarkup) ? string.Empty : additionalMarkup) +
                ">" + text + "</a>";
        }

        /// <summary>
        /// Created an Href HTML link
        /// </summary>
        /// <param name="text"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string Href(string text, string url)
        {
            return Href(text, url, null, null);
        }

        /// <summary>
        /// Creates an HREF HTML Link
        /// </summary>
        /// <param name="url"></param>
        public static string Href(string url)
        {
            return Href(url, url, null, null);
        }

        /// <summary>
        /// Extracts a string from between a pair of delimiters. Only the first 
        /// instance is found.
        /// </summary>
        /// <param name="source">Input String to work on</param>
        /// <param name="StartDelim">Beginning delimiter</param>
        /// <param name="endDelim">ending delimiter</param>
        /// <param name="CaseInsensitive">Determines whether the search for delimiters is case sensitive</param>
        /// <returns>Extracted string or ""</returns>
        public static string ExtractString(string source, string beginDelim,
                                           string endDelim, bool caseSensitive,
                                           bool allowMissingEndDelimiter)
        {
            int at1, at2;

            if (string.IsNullOrEmpty(source))
                return string.Empty;

            if (caseSensitive)
            {
                at1 = source.IndexOf(beginDelim);
                if (at1 == -1)
                    return string.Empty;

                at2 = source.IndexOf(endDelim, at1 + beginDelim.Length);
            }
            else
            {
                //string Lower = source.ToLower();
                at1 = source.IndexOf(beginDelim, 0, source.Length, StringComparison.OrdinalIgnoreCase);
                if (at1 == -1)
                    return string.Empty;

                at2 = source.IndexOf(endDelim, at1 + beginDelim.Length, StringComparison.OrdinalIgnoreCase);
            }

            if (allowMissingEndDelimiter && at2 == -1)
                return source.Substring(at1 + beginDelim.Length);

            if (at1 > -1 && at2 > 1)
                return source.Substring(at1 + beginDelim.Length, at2 - at1 - beginDelim.Length);

            return string.Empty;
        }

        /// <summary>
        /// Extracts a string from between a pair of delimiters. Only the first
        /// instance is found.
        /// <seealso>Class wwUtils</seealso>
        /// </summary>
        /// <param name="source">
        /// Input String to work on
        /// </param>
        /// <param name="beginDelim"></param>
        /// <param name="endDelim">
        /// ending delimiter
        /// </param>
        /// <param name="CaseInSensitive"></param>
        /// <returns>String</returns>
        public static string ExtractString(string source, string beginDelim, string endDelim, bool caseSensitive)
        {
            return ExtractString(source, beginDelim, endDelim, caseSensitive, false);
        }

        /// <summary>
        /// Extracts a string from between a pair of delimiters. Only the first 
        /// instance is found. Search is case insensitive.
        /// </summary>
        /// <param name="source">
        /// Input String to work on
        /// </param>
        /// <param name="StartDelim">
        /// Beginning delimiter
        /// </param>
        /// <param name="endDelim">
        /// ending delimiter
        /// </param>
        /// <returns>Extracted string or string.Empty</returns>
        public static string ExtractString(string source, string beginDelim, string endDelim)
        {
            return ExtractString(source, beginDelim, endDelim, false, false);
        }


        /// <summary>
        /// String replace function that support
        /// </summary>
        /// <param name="origString">Original input string</param>
        /// <param name="findString">The string that is to be replaced</param>
        /// <param name="replaceWith">The replacement string</param>
        /// <param name="instance">Instance of the FindString that is to be found. if Instance = -1 all are replaced</param>
        /// <param name="caseInsensitive">Case insensitivity flag</param>
        /// <returns>updated string or original string if no matches</returns>
        public static string ReplaceStringInstance(string origString, string findString,
                                                   string replaceWith, int instance,
                                                   bool caseInsensitive)
        {
            if (instance == -1)
                return ReplaceString(origString, findString, replaceWith, caseInsensitive);
            if (instance == 0)
                return origString;

            int at1 = 0;
            for (int x = 0; x < instance; x++)
            {

                if (caseInsensitive)
                    at1 = origString.IndexOf(findString, at1, origString.Length - at1, StringComparison.OrdinalIgnoreCase);
                else
                    at1 = origString.IndexOf(findString, at1);

                if (at1 == -1)
                    return origString;

                if (x < instance - 1)
                    at1 += findString.Length;
            }

            return origString.Substring(0, at1) + replaceWith + origString.Substring(at1 + findString.Length);
        }

        /// <summary>
        /// Replaces a substring within a string with another substring with optional case sensitivity turned off.
        /// </summary>
        /// <param name="origString">String to do replacements on</param>
        /// <param name="findString">The string to find</param>
        /// <param name="replaceString">The string to replace found string wiht</param>
        /// <param name="caseInsensitive">If true case insensitive search is performed</param>
        /// <returns>updated string or original string if no matches</returns>
        public static string ReplaceString(string origString, string findString,
                                           string replaceString, bool caseInsensitive)
        {
            int at1 = 0;
            while (true)
            {
                if (caseInsensitive)
                    at1 = origString.IndexOf(findString, at1, origString.Length - at1, StringComparison.OrdinalIgnoreCase);
                else
                    at1 = origString.IndexOf(findString, at1);

                if (at1 == -1)
                    break;

                origString = origString.Substring(0, at1) + replaceString + origString.Substring(at1 + findString.Length);

                at1 += replaceString.Length;
            }

            return origString;
        }


        /// <summary>
        /// Determines whether a string is empty (null or zero length)
        /// </summary>
        /// <param name="text">Input string</param>
        /// <returns>true or false</returns>
        public static bool Empty(string text)
        {
            if (text == null || text.Trim().Length == 0)
                return true;

            return false;
        }

        /// <summary>
        /// Determines wheter a string is empty (null or zero length)
        /// </summary>
        /// <param name="text">Input string (in object format)</param>
        /// <returns>true or false/returns>
        public static bool Empty(object text)
        {
            return Empty(text as string);
        }

        /// <summary>
        /// Return a string in proper Case format
        /// </summary>
        /// <param name="Input"></param>
        /// <returns></returns>
        public static string ProperCase(string Input)
        {
            return Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Input);
        }

        /// <summary>
        /// Terminates a string with the given end string/character, but only if the
        /// value specified doesn't already exist and the string is not empty.
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        public static string TerminateString(string value, string terminator)
        {
            if (string.IsNullOrEmpty(value) || value.EndsWith(terminator))
                return value;

            return value + terminator;
        }

        /// <summary>
        /// Returns an abstract of the provided text by returning up to Length characters
        /// of a text string. If the text is truncated a ... is appended.
        /// </summary>
        /// <param name="text">Text to abstract</param>
        /// <param name="length">Number of characters to abstract to</param>
        /// <returns>string</returns>
        public static string TextAbstract(string text, int length)
        {
            if (text.Length <= length)
                return text;

            text = text.Substring(0, length);

            text = text.Substring(0, text.LastIndexOf(" "));
            return text + "...";
        }

        /// <summary>
        /// Creates an Abstract from an HTML document. Strips the 
        /// HTML into plain text, then creates an abstract.
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public static string HtmlAbstract(string html, int length)
        {
            return TextAbstract(StripHtml(html), length);
        }

        /// <summary>
        /// Simple Logging method that allows quickly writing a string to a file
        /// </summary>
        /// <param name="output"></param>
        /// <param name="filename"></param>
        public static void LogString(string output, string filename)
        {
            StreamWriter Writer = new StreamWriter(filename, true);
            Writer.WriteLine(DateTime.Now.ToString() + " - " + output);
            Writer.Close();
        }

        /// <summary>
        /// Creates short string id based on a GUID hashcode.
        /// Not guaranteed to be unique across machines, but unlikely
        /// to duplicate in medium volume situations.
        /// </summary>
        /// <returns></returns>
        public static string NewStringId()
        {
            return Guid.NewGuid().ToString().GetHashCode().ToString("x");
        }


        /// <summary>
        /// Parses an string into an integer. If the value can't be parsed
        /// a default value is returned instead
        /// </summary>
        /// <param name="input"></param>
        /// <param name="defaultValue"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static int ParseInt(string input, int defaultValue, IFormatProvider numberFormat)
        {
            int val = defaultValue;
            int.TryParse(input,NumberStyles.Any, numberFormat,out val);
            return val;
        }

        /// <summary>
        /// Parses an string into an integer. If the value can't be parsed
        /// a default value is returned instead
        /// </summary>
        /// <param name="input"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static int ParseInt(string input, int defaultValue)
        {
            return ParseInt(input, defaultValue,CultureInfo.CurrentCulture.NumberFormat);
        }

        /// <summary>
        /// Parses an string into an decimal. If the value can't be parsed
        /// a default value is returned instead
        /// </summary>
        /// <param name="input"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseDecimal(string input, decimal defaultValue, IFormatProvider numberFormat)
        {
            decimal val = defaultValue;
            decimal.TryParse(input, NumberStyles.Any ,numberFormat, out val);
            return val;
        }

        #region UrlEncoding and UrlDecoding without System.Web
        /// <summary>
        /// UrlEncodes a string without the requirement for System.Web
        /// </summary>
        /// <param name="String"></param>
        /// <returns></returns>
        public static string UrlEncode(string text)
        {
            
            StringBuilder sb = new StringBuilder(text.Length);

            int len = text.Length;
            for (int i = 0; i < len; i++)
            {

                int value = text[i];
                if (value == -1)
                    break;
                char charValue = (char)value;

                if (charValue >= 'a' && charValue <= 'z' ||
                    charValue >= 'A' && charValue <= 'Z' ||
                    charValue >= '0' && charValue <= '9')
                    sb.Append(charValue);
                else if (charValue == ' ')
                    sb.Append("+");
                else
                    sb.AppendFormat("%{0:X2}", value);
            }

            return sb.ToString();
        }

        /// <summary>
        /// UrlDecodes a string without requiring System.Web
        /// </summary>
        /// <param name="text">String to decode.</param>
        /// <returns>decoded string</returns>
        public static string UrlDecode(string text)
        {
            char temp = ' ';
            StringReader sr = new StringReader(text);
            StringBuilder sb = new StringBuilder(text.Length);

            while (true)
            {
                int lnVal = sr.Read();
                if (lnVal == -1)
                    break;
            
                char tchar = (char)lnVal;
                if (tchar == '+')
                    sb.Append(' ');
                else if (tchar == '%')
                {
                    // *** read the next 2 chars and parse into a char
                    temp = (char)Int32.Parse(((char)sr.Read()).ToString() + ((char)sr.Read()).ToString(),
                        System.Globalization.NumberStyles.HexNumber);
                    sb.Append(temp);
                }
                else
                    sb.Append(tchar);
            }
            sr.Close();

            return sb.ToString();
        }

        /// <summary>
        /// Retrieves a value by key from a UrlEncoded string.
        /// </summary>
        /// <param name="urlEncoded">UrlEncoded String</param>
        /// <param name="key">Key to retrieve value for</param>
        /// <returns>returns the value or "" if the key is not found or the value is blank</returns>
        public static string GetUrlEncodedKey(string urlEncoded, string key)
        {
            urlEncoded = "&" + urlEncoded + "&";

            int Index = urlEncoded.ToLower().IndexOf("&" + key.ToLower() + "=");
            if (Index < 0)
                return "";

            int lnStart = Index + 2 + key.Length;

            int Index2 = urlEncoded.IndexOf("&", lnStart);
            if (Index2 < 0)
                return "";

            return UrlDecode(urlEncoded.Substring(lnStart, Index2 - lnStart));
        }


        /// <summary>
        /// HTML-encodes a string and returns the encoded string.
        /// </summary>
        /// <param name="text">The text string to encode. </param>
        /// <returns>The HTML-encoded text.</returns>
        public static string HtmlEncode(string text)
        {
            if (text == null)
                return null;

            StringBuilder sb = new StringBuilder(text.Length);

            int len = text.Length;
            for (int i = 0; i < len; i++)
            {
                switch (text[i])
                {
                    case '&':
                        sb.Append("&amp;");
                        break;
                    case '>':
                        sb.Append("&gt;");
                        break;
                    case '<':
                        sb.Append("&lt;");
                        break;
                    case '"':
                        sb.Append("&quot;");
                        break;
                    default:
                        if (text[i] > 159)
                        {
                            sb.Append("&#");
                            sb.Append(((int)text[i]).ToString(CultureInfo.InvariantCulture));
                            sb.Append(";");
                        }
                        else
                            sb.Append(text[i]);
                        break;
                }
            }
            return sb.ToString();
        }

        #endregion
    }



    public class ExpandUrlsParser
    {
        public string Target = string.Empty;
        public bool ParseFormattedLinks = false;

        /// <summary>
        /// Expands links into HTML hyperlinks inside of text or HTML.
        /// </summary>
        /// <param name="text">The text to expand</param>    
        /// <returns></returns>
        public string ExpandUrls(string text)
        {
            MatchEvaluator matchEval = null;
            string pattern = null;
            string updated = null;


            // *** Expand embedded hyperlinks
            System.Text.RegularExpressions.RegexOptions options =
                                                                  RegexOptions.Multiline |
                                                                  RegexOptions.IgnoreCase;

            if (this.ParseFormattedLinks)
            {
                pattern = @"\[(.*?)\|(.*?)]";

                matchEval = new MatchEvaluator(this.ExpandFormattedLinks);
                updated = Regex.Replace(text, pattern, matchEval, options);
            }
            else
                updated = text;

            pattern = @"([""'=]|&quot;)?(http://|ftp://|https://|www\.|ftp\.[\w]+)([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])";

            matchEval = new MatchEvaluator(this.ExpandUrlsRegExEvaluator);
            updated = Regex.Replace(updated, pattern, matchEval, options);



            return updated;
        }

        /// <summary>
        /// Internal RegExEvaluator callback
        /// </summary>
        /// <param name="M"></param>
        /// <returns></returns>
        private string ExpandUrlsRegExEvaluator(System.Text.RegularExpressions.Match M)
        {
            string Href = M.Value; // M.Groups[0].Value;

            // *** if string starts within an HREF don't expand it
            if (Href.StartsWith("=") ||
                Href.StartsWith("'") ||
                Href.StartsWith("\"") ||
                Href.StartsWith("&quot;"))
                return Href;

            string Text = Href;

            if (Href.IndexOf("://") < 0)
            {
                if (Href.StartsWith("www."))
                    Href = "http://" + Href;
                else if (Href.StartsWith("ftp"))
                    Href = "ftp://" + Href;
                else if (Href.IndexOf("@") > -1)
                    Href = "mailto:" + Href;
            }

            string Targ = !string.IsNullOrEmpty(this.Target) ? " target='" + this.Target + "'" : "";

            return "<a href='" + Href + "'" + Targ +
                    ">" + Text + "</a>";
        }

        private string ExpandFormattedLinks(System.Text.RegularExpressions.Match M)
        {
            //string Href = M.Value; // M.Groups[0].Value;

            string Text = M.Groups[1].Value;
            string Href = M.Groups[2].Value;

            if (Href.IndexOf("://") < 0)
            {
                if (Href.StartsWith("www."))
                    Href = "http://" + Href;
                else if (Href.StartsWith("ftp"))
                    Href = "ftp://" + Href;
                else if (Href.IndexOf("@") > -1)
                    Href = "mailto:" + Href;
                else
                    Href = "http://" + Href;
            }

            string Targ = !string.IsNullOrEmpty(this.Target) ? " target='" + this.Target + "'" : "";

            return "<a href='" + Href + "'" + Targ +
                    ">" + Text + "</a>";
        }

    }
}