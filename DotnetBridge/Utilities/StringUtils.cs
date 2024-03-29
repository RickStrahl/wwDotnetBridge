#region License
/*
 **************************************************************
 *  Author: Rick Strahl 
 *          (c) West Wind Technologies, 2008
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
using System.Collections.Generic;
using System.Linq;

namespace Westwind.WebConnection
{
    /// <summary>
    /// String utility class that provides a host of string related operations
    /// </summary>
    public static class StringUtils
    {
        #region Basic String Tasks


        /// <summary>
        /// Trims a sub string from a string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="textToTrim"></param>
        /// <returns></returns>
        public static string TrimStart(string text, string textToTrim, bool caseInsensitive)
        {
            while (true)
            {
                string match = text.Substring(0, textToTrim.Length);

                if (match == textToTrim ||
                    (caseInsensitive && match.ToLower() == textToTrim.ToLower()))
                {
                    if (text.Length <= match.Length)
                        text = "";
                    else
                        text = text.Substring(textToTrim.Length);
                }
                else
                    break;
            }
            return text;
        }

        /// <summary>
        /// Replicates an input string n number of times
        /// </summary>
        /// <param name="input"></param>
        /// <param name="charCount"></param>
        /// <returns></returns>
        public static string Replicate(string input, int charCount)
        {
            return new StringBuilder().Insert(0, input, charCount).ToString();
        }

        /// <summary>
        /// Replicates a character n number of times and returns a string
        /// </summary>
        /// <param name="charCount"></param>
        /// <param name="character"></param>
        /// <returns></returns>
        public static string Replicate(char character, int charCount)
        {
            return new StringBuilder().Insert(0, character.ToString(), charCount).ToString();
        }

        /// <summary>
        /// Finds the nth index of strting in a string
        /// </summary>
        /// <param name="source"></param>
        /// <param name="matchString"></param>
        /// <param name="stringInstance"></param>
        /// <returns></returns>
        public static int IndexOfNth(string source, string matchString, int stringInstance, StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            if (string.IsNullOrEmpty(source))
                return -1;

            int lastPos = 0;
            int count = 0;
           
            while (count < stringInstance )
            {
                var len = source.Length - lastPos;
                lastPos = source.IndexOf(matchString, lastPos,len,stringComparison);
                if (lastPos == -1)
                    break;

                count++;
                if (count == stringInstance)
                    return lastPos;

                lastPos += matchString.Length;
            }
            return -1;
        }

        /// <summary>
        /// Returns the nth Index of a character in a string
        /// </summary>
        /// <param name="source"></param>
        /// <param name="matchChar"></param>
        /// <param name="charInstance"></param>
        /// <returns></returns>
        public static int IndexOfNth(this string source, char matchChar, int charInstance)        
        {
            if (string.IsNullOrEmpty(source))
                return -1;

            if (charInstance < 1)
                return -1;

            int count = 0;
            for (int i = 0; i < source.Length; i++)
            {
                if (source[i] == matchChar)
                {
                    count++;
                    if (count == charInstance)                 
                        return i;                 
                }
            }
            return -1;
        }



        /// <summary>
        /// Finds the nth index of strting in a string
        /// </summary>
        /// <param name="source"></param>
        /// <param name="matchString"></param>
        /// <param name="charInstance"></param>
        /// <returns></returns>
        public static int LastIndexOfNth(string source, string matchString, int charInstance, StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            if (string.IsNullOrEmpty(source))
                return -1;

            int lastPos = source.Length;
            int count = 0;

            while (count < charInstance)
            {                
                lastPos = source.LastIndexOf(matchString, lastPos, lastPos, stringComparison);
                if (lastPos == -1)
                    break;

                count++;
                if (count == charInstance)
                    return lastPos;                
            }
            return -1;
        }

        /// <summary>
        /// Finds the nth index of in a string from the end.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="matchChar"></param>
        /// <param name="charInstance"></param>
        /// <returns></returns>
        public static int LastIndexOfNth(string source, char matchChar, int charInstance)
        {
            if (string.IsNullOrEmpty(source))
                return -1;

            int count = 0;
            for (int i = source.Length-1 ; i > -1; i--)
            {
                if (source[i] == matchChar)
                {
                    count++;
                    if (count == charInstance)
                        return i;
                }
            }
            return -1;
        }
        #endregion

        #region String Casing

        /// <summary>
        /// Return a string in proper Case format
        /// </summary>
        /// <param name="Input"></param>
        /// <returns></returns>
        public static string ProperCase(string Input)
        {
            if (Input == null)
                return null;
            return Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Input);
        }

        /// <summary>
        /// Takes a phrase and turns it into CamelCase text.
        /// White Space, punctuation and separators are stripped
        /// </summary>
        /// <param name="phrase">Text to convert to CamelCase</param>
        public static string ToCamelCase(string phrase)
        {
            if (phrase == null)
                return string.Empty;

            StringBuilder sb = new StringBuilder(phrase.Length);

            // First letter is always upper case
            bool nextUpper = true;

            foreach (char ch in phrase)
            {
                if (char.IsWhiteSpace(ch) || char.IsPunctuation(ch) || char.IsSeparator(ch) || ch > 32 && ch < 48)
                {
                    nextUpper = true;
                    continue;
                }
                if (char.IsDigit(ch))
                {
                    sb.Append(ch);
                    nextUpper = true;
                    continue;
                }

                if (nextUpper)
                    sb.Append(char.ToUpper(ch));
                else
                    sb.Append(char.ToLower(ch));

                nextUpper = false;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Tries to create a phrase string from CamelCase text.
        /// Will place spaces before capitalized letters.
        /// 
        /// Note that this method may not work for round tripping 
        /// ToCamelCase calls, since ToCamelCase strips more characters
        /// than just spaces.
        /// </summary>
        /// <param name="camelCase"></param>
        /// <returns></returns>
        public static string FromCamelCase(string camelCase)
        {
            if (camelCase == null)
                throw new ArgumentException("Null is not allowed for StringUtils.FromCamelCase");

            StringBuilder sb = new StringBuilder(camelCase.Length + 10);
            bool first = true;
            char lastChar = '\0';

            foreach (char ch in camelCase)
            {
                if (!first &&
                    (char.IsUpper(ch) ||
                     char.IsDigit(ch) && !char.IsDigit(lastChar)))
                    sb.Append(' ');

                sb.Append(ch);
                first = false;
                lastChar = ch;
            }

            return sb.ToString(); ;
        }

        /// <summary>
        /// Terminates a string with the given end string/character, but only if the
        /// value specified doesn't already exist and the string is not empty.
        /// </summary>
        /// <param name="value">String to terminate</param>
        /// <param name="terminator">String to terminate the value string with</param>
        /// <returns></returns>
        public static string TerminateString(string value, string terminator)
        {
            if (string.IsNullOrEmpty(value))
                return terminator;

            if (value.EndsWith(terminator))
                return value;

            return value + terminator;
        }

        /// <summary>
        /// Trims a string to a specific number of max characters
        /// </summary>
        /// <param name="value"></param>
        /// <param name="charCount"></param>
        /// <returns></returns>
        public static string TrimTo(string value, int charCount)
        {
            if (value == null)
                return string.Empty;

            if (value.Length > charCount)
                return value.Substring(0, charCount);

            return value;
        }

        #endregion

        #region String Manipulation
        /// <summary>
        /// Extracts a string from between a pair of delimiters. Only the first 
        /// instance is found.
        /// </summary>
        /// <param name="source">Input String to work on</param>
        /// <param name="StartDelim">Beginning delimiter</param>
        /// <param name="endDelim">ending delimiter</param>
        /// <param name="CaseInsensitive">Determines whether the search for delimiters is case sensitive</param>
        /// <returns>Extracted string or ""</returns>
        public static string ExtractString(string source,
            string beginDelim,
            string endDelim,
            bool caseSensitive = false,
            bool allowMissingEndDelimiter = false,
            bool returnDelimiters = false)
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

            if (allowMissingEndDelimiter && at2 < 0)
            {
                if (!returnDelimiters)
                    return source.Substring(at1 + beginDelim.Length);
                else
                    return source.Substring(at1);
            }

            if (at1 > -1 && at2 > 1)
            {
                if (!returnDelimiters)
                    return source.Substring(at1 + beginDelim.Length, at2 - at1 - beginDelim.Length);

                return source.Substring(at1, at2 - at1 + endDelim.Length);
            }

            return string.Empty;
        }



        /// <summary>
        /// String replace function that supports replacing a specific instance with 
        /// case insensitivity
        /// </summary>
        /// <param name="origString">Original input string</param>
        /// <param name="findString">The string that is to be replaced</param>
        /// <param name="replaceWith">The replacement string</param>
        /// <param name="instance">Instance of the FindString that is to be found. 1 based. If Instance = -1 all are replaced</param>
        /// <param name="caseInsensitive">Case insensitivity flag</param>
        /// <returns>updated string or original string if no matches</returns>
        public static string ReplaceStringInstance(string origString, string findString,
            string replaceWith, int instance, bool caseInsensitive)
        {
            if (instance == -1)
                return ReplaceString(origString, findString, replaceWith, caseInsensitive);

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
        /// Returns an abstract of the provided text by returning up to Length characters
        /// of a text string. If the text is truncated a ... is appended.
        /// </summary>
        /// <param name="text">Text to abstract</param>
        /// <param name="length">Number of characters to abstract to</param>
        /// <returns>string</returns>
        public static string TextAbstract(string text, int length)
        {
            if (text == null)
                return string.Empty;

            if (text.Length <= length)
                return text;

            text = text.Substring(0, length);

            text = text.Substring(0, text.LastIndexOf(" "));
            return text + "...";
        }

        #endregion

        #region String Parsing
        /// <summary>
        /// Determines if a string is contained in a list of other strings
        /// </summary>
        /// <param name="s"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool Inlist(string s, params string[] list)
        {
            return list.Contains(s);
        }

        /// <summary>
        /// Parses a string into an array of lines broken
        /// by \r\n or \n
        /// </summary>
        /// <param name="s">String to check for lines</param>
        /// <param name="maxLines">Optional - max number of lines to return</param>
        /// <returns>array of strings, or null if the string passed was a null</returns>
        public static string[] GetLines(string s, int maxLines = 0)
        {
            if (s == null)
                return null;

            s = s.Replace("\r\n", "\n");

            if (maxLines < 1)
                return s.Split(new char[] { '\n' });

            return s.Split(new char[] { '\n' }).Take(maxLines).ToArray();
        }

        /// <summary>
        /// Returns a line count for a string
        /// </summary>
        /// <param name="s">string to count lines for</param>
        /// <returns></returns>
        public static int CountLines(string s)
        {
            if (string.IsNullOrEmpty(s))
                return 0;

            return s.Split('\n').Length;
        }

        /// <summary>
        /// Strips all non digit values from a string and only
        /// returns the numeric string.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string StripNonNumber(string input)
        {
            var chars = input.ToCharArray();
            StringBuilder sb = new StringBuilder();
            foreach (var chr in chars)
            {
                if (char.IsNumber(chr) || char.IsSeparator(chr))
                    sb.Append(chr);
            }

            return sb.ToString();
        }


        static Regex tokenizeRegex = new Regex("{{.*?}}");

        /// <summary>
        /// Tokenizes a string based on a start and end string. Replaces the values with a token
        /// value (#@#1#@# for example).
        /// 
        /// You can use Detokenize to get the original values back
        /// </summary>
        /// <param name="text"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="replaceDelimiter"></param>
        /// <returns></returns>
        public static List<string> TokenizeString(ref string text, string start, string end, string replaceDelimiter = "#@#")
        {
            var strings = new List<string>();
            var matches = tokenizeRegex.Matches(text);

            int i = 0;
            foreach (Match match in matches)
            {
                tokenizeRegex = new Regex(Regex.Escape(match.Value));
                text = tokenizeRegex.Replace(text, $"{replaceDelimiter}{i}{replaceDelimiter}", 1);
                strings.Add(match.Value);
                i++;
            }

            return strings;
        }


        /// <summary>
        /// Detokenizes a string tokenized with TokenizeString. Requires the collection created
        /// by detokenization
        /// </summary>
        /// <param name="text"></param>
        /// <param name="tokens"></param>
        /// <param name="replaceDelimiter"></param>
        /// <returns></returns>
        public static string DetokenizeString(string text, List<string> tokens, string replaceDelimiter = "#@#")
        {
            int i = 0;
            foreach (string token in tokens)
            {
                text = text.Replace($"{replaceDelimiter}{i}{replaceDelimiter}", token);
                i++;
            }
            return text;
        }

        /// <summary>
        /// Parses an string into an integer. If the value can't be parsed
        /// a default value is returned instead
        /// </summary>
        /// <param name="input">Input numeric string to be parsed</param>
        /// <param name="defaultValue">Optional default value if parsing fails</param>
        /// <param name="formatProvider">Optional NumberFormat provider. Defaults to current culture's number format</param>
        /// <returns></returns>
        public static int ParseInt(string input, int defaultValue=0, IFormatProvider numberFormat = null)
        {
            if (numberFormat == null)
                numberFormat = CultureInfo.CurrentCulture.NumberFormat;

            int val = defaultValue;
            if (!int.TryParse(input, NumberStyles.Any, numberFormat, out val))
                return defaultValue;
            return val;
        }



        /// <summary>
        /// Parses an string into an decimal. If the value can't be parsed
        /// a default value is returned instead
        /// </summary>
        /// <param name="input"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseDecimal(string input, decimal defaultValue = 0M, IFormatProvider numberFormat = null)
        {
            numberFormat = numberFormat ?? CultureInfo.CurrentCulture.NumberFormat;
            decimal val = defaultValue;
            if (!decimal.TryParse(input, NumberStyles.Any, numberFormat, out val))
                return defaultValue;
            return val;
        }

        #endregion

        #region String Ids
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
        /// Creates a new random string of upper, lower case letters and digits.
        /// Very useful for generating random data for storage in test data.
        /// </summary>
        /// <param name="size">The number of characters of the string to generate</param>
        /// <param name="includeNumbers"></param>
        /// <returns>randomized string</returns>
        public static string RandomString(int size, bool includeNumbers = false)
        {
            StringBuilder builder = new StringBuilder(size);
            char ch;
            int num;

            for (int i = 0; i < size; i++)
            {
                if (includeNumbers)
                    num = Convert.ToInt32(Math.Floor(62 * random.NextDouble()));
                else
                    num = Convert.ToInt32(Math.Floor(52 * random.NextDouble()));

                if (num < 26)
                    ch = Convert.ToChar(num + 65);
                // lower case
                else if (num > 25 && num < 52)
                    ch = Convert.ToChar(num - 26 + 97);
                // numbers
                else
                    ch = Convert.ToChar(num - 52 + 48);

                builder.Append(ch);
            }

            return builder.ToString();
        }
        private static Random random = new Random((int)DateTime.Now.Ticks);

        #endregion

        #region Encodings

        /// <summary>
        /// UrlEncodes a string without the requirement for System.Web
        /// </summary>
        /// <param name="String"></param>
        /// <returns></returns>
        // [Obsolete("Use System.Uri.EscapeDataString instead")]
        public static string UrlEncode(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            return Uri.EscapeDataString(text);
        }

        /// <summary>
        /// Encodes a few additional characters for use in paths
        /// Encodes: . #
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string UrlEncodePathSafe(string text)
        {
            string escaped = UrlEncode(text);
            return escaped.Replace(".", "%2E").Replace("#", "%23");
        }

        /// <summary>
        /// UrlDecodes a string without requiring System.Web
        /// </summary>
        /// <param name="text">String to decode.</param>
        /// <returns>decoded string</returns>
        public static string UrlDecode(string text)
        {
            // pre-process for + sign space formatting since System.Uri doesn't handle it
            // plus literals are encoded as %2b normally so this should be safe
            text = text.Replace("+", " ");
            string decoded = Uri.UnescapeDataString(text);
            return decoded;
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

            int Index = urlEncoded.IndexOf("&" + key + "=", StringComparison.OrdinalIgnoreCase);
            if (Index < 0)
                return string.Empty;

            int lnStart = Index + 2 + key.Length;

            int Index2 = urlEncoded.IndexOf("&", lnStart);
            if (Index2 < 0)
                return string.Empty;

            return UrlDecode(urlEncoded.Substring(lnStart, Index2 - lnStart));
        }

        /// <summary>
        /// Allows setting of a value in a UrlEncoded string. If the key doesn't exist
        /// a new one is set, if it exists it's replaced with the new value.
        /// </summary>
        /// <param name="urlEncoded">A UrlEncoded string of key value pairs</param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string SetUrlEncodedKey(string urlEncoded, string key, string value)
        {
            if (!urlEncoded.EndsWith("?") && !urlEncoded.EndsWith("&"))
                urlEncoded += "&";

            Match match = Regex.Match(urlEncoded, "[?|&]" + key + "=.*?&");

            if (match == null || string.IsNullOrEmpty(match.Value))
                urlEncoded = urlEncoded + key + "=" + UrlEncode(value) + "&";
            else
                urlEncoded = urlEncoded.Replace(match.Value, match.Value.Substring(0, 1) + key + "=" + UrlEncode(value) + "&");

            return urlEncoded.TrimEnd('&');
        }
        #endregion

        #region Binary Encoding

        /// <summary>
        /// Turns a BinHex string that contains raw byte values
        /// into a byte array
        /// </summary>
        /// <param name="hex">BinHex string (just two byte hex digits strung together)</param>
        /// <returns></returns>
        public static byte[] BinHexToBinary(string hex)
        {
            int offset = hex.StartsWith("0x") ? 2 : 0;
            if ((hex.Length % 2) != 0)
                throw new ArgumentException("Invalid Hex String Length...");

            byte[] ret = new byte[(hex.Length - offset) / 2];

            for (int i = 0; i < ret.Length; i++)
            {
                ret[i] = (byte)((ParseHexChar(hex[offset]) << 4)
                                | ParseHexChar(hex[offset + 1]));
                offset += 2;
            }
            return ret;
        }

        /// <summary>
        /// Converts a byte array into a BinHex string.
        /// BinHex is two digit hex byte values squished together
        /// into a string.
        /// </summary>
        /// <param name="data">Raw data to send</param>
        /// <returns>BinHex string or null if input is null</returns>
        public static string BinaryToBinHex(byte[] data)
        {
            if (data == null)
                return null;

            StringBuilder sb = new StringBuilder(data.Length * 2);
            foreach (byte val in data)
            {
                sb.AppendFormat("{0:x2}", val);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Converts a string into bytes for storage in any byte[] types
        /// buffer or stream format (like MemoryStream).
        /// </summary>
        /// <param name="text"></param>
        /// <param name="encoding">The character encoding to use. Defaults to Unicode</param>
        /// <returns></returns>
        public static byte[] StringToBytes(string text, Encoding encoding = null)
        {
            if (text == null)
                return null;

            if (encoding == null)
                encoding = Encoding.Unicode;

            return encoding.GetBytes(text);
        }

        /// <summary>
        /// Converts a byte array to a stringUtils
        /// </summary>
        /// <param name="buffer">raw string byte data</param>
        /// <param name="encoding">Character encoding to use. Defaults to Unicode</param>
        /// <returns></returns>
        public static string BytesToString(byte[] buffer, Encoding encoding = null)
        {
            if (buffer == null)
                return null;

            if (encoding == null)
                encoding = Encoding.Unicode;

            return encoding.GetString(buffer);
        }

        static int ParseHexChar(char c)
        {
            if (c >= '0' && c <= '9')
                return c - '0';
            if (c >= 'A' && c <= 'F')
                return c - 'A' + 10;
            if (c >= 'a' && c <= 'f')
                return c - 'a' + 10;

            throw new ArgumentException("Invalid Hex digit: " +  c);
        }

        static char[] base36CharArray = "0123456789abcdefghijklmnopqrstuvwxyz".ToCharArray();
        static string base36Chars = "0123456789abcdefghijklmnopqrstuvwxyz";

        /// <summary>
        /// Encodes an integer into a string by mapping to alpha and digits (36 chars)
        /// chars are embedded as lower case
        /// 
        /// Example: 4zx12ss
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string Base36Encode(long value)
        {
            string returnValue = "";
            bool isNegative = value < 0;
            if (isNegative)
                value = value * -1;

            do
            {
                returnValue = base36CharArray[value % base36CharArray.Length] + returnValue;
                value /= 36;
            } while (value != 0);

            return isNegative ? returnValue + "-" : returnValue;
        }

        /// <summary>
        /// Decodes a base36 encoded string to an integer
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static long Base36Decode(string input)
        {
            bool isNegative = false;
            if (input.EndsWith("-"))
            {
                isNegative = true;
                input = input.Substring(0, input.Length - 1);
            }

            char[] arrInput = input.ToCharArray();
            Array.Reverse(arrInput);
            long returnValue = 0;
            for (long i = 0; i < arrInput.Length; i++)
            {
                long valueindex = base36Chars.IndexOf(arrInput[i]);
                returnValue += Convert.ToInt64(valueindex * Math.Pow(36, i));
            }
            return isNegative ? returnValue * -1 : returnValue;
        }
        #endregion

        #region Miscellaneous    

        /// <summary>
        /// Strips any common white space from all lines of text that have the same
        /// common white space text. Effectively removes common code indentation from
        /// code blocks for example so you can get a left aligned code snippet.
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public static string NormalizeIndentation(string code)
        {
            // normalize tabs to 3 spaces
            string text = code.Replace("\t", "   ");

            string[] lines = text.Split(new string[3] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // keep track of the smallest indent
            int minPadding = 1000;

            foreach (var line in lines)
            {
                if (line.Length == 0)  // ignore blank lines
                    continue;

                int count = 0;
                foreach (char chr in line)
                {
                    if (chr == ' ' && count < minPadding)
                        count++;
                    else
                        break;
                }
                if (count == 0)
                    return code;

                minPadding = count;
            }

            string strip = new String(' ', minPadding);

            StringBuilder sb = new StringBuilder();
            foreach (var line in lines)
            {
                sb.AppendLine(StringUtils.ReplaceStringInstance(line, strip, "", 1, false));
            }

            return sb.ToString();
        }


        /// <summary>
        /// Simple Logging method that allows quickly writing a string to a file
        /// </summary>
        /// <param name="output"></param>
        /// <param name="filename"></param>
        /// <param name="encoding">if not specified used UTF-8</param>
        public static void LogString(string output, string filename, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;

            StreamWriter Writer = new StreamWriter(filename, true, encoding);
            Writer.WriteLine(DateTime.Now + " - " + output);
            Writer.Close();
        }

        /// <summary>
        /// Creates a Stream from a string. Internally creates
        /// a memory stream and returns that.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static Stream StringToStream(string text, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.Default;

            MemoryStream ms = new MemoryStream(text.Length * 2);
            byte[] data = encoding.GetBytes(text);
            ms.Write(data, 0, data.Length);
            ms.Position = 0;
            return ms;
        }

        /// <summary>
        /// Retrieves a value from an XML-like string
        /// </summary>
        /// <param name="propertyString"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetProperty(string propertyString, string key)
        {
            return StringUtils.ExtractString(propertyString, "<" + key + ">", "</" + key + ">");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyString"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string SetProperty(string propertyString, string key, string value)
        {
            string extract = StringUtils.ExtractString(propertyString, "<" + key + ">", "</" + key + ">");

            if (string.IsNullOrEmpty(value) && extract != string.Empty)
            {
                return propertyString.Replace(extract, "");
            }

            string xmlLine = "<" + key + ">" + value + "</" + key + ">";

            // replace existing
            if (extract != string.Empty)
                return propertyString.Replace(extract, xmlLine);

            // add new
            return propertyString + xmlLine + "\r\n";
        }

        #endregion




        static string HtmlSanitizeTagBlackList { get; } = "script|iframe|object|embed|form";

        static Regex _RegExScript = new Regex($@"(<({HtmlSanitizeTagBlackList})\b[^<]*(?:(?!<\/({HtmlSanitizeTagBlackList}))<[^<]*)*<\/({HtmlSanitizeTagBlackList})>)",
        RegexOptions.IgnoreCase | RegexOptions.Multiline);

        // strip javascript: and unicode representation of javascript:
        // href='javascript:alert(\"gotcha\")'
        // href='&#106;&#97;&#118;&#97;&#115;&#99;&#114;&#105;&#112;&#116;:alert(\"gotcha\");'
        static Regex _RegExJavaScriptHref = new Regex(
            @"<[^>]*?\s(href|src|dynsrc|lowsrc)=.{0,20}((javascript:)|(&#)).*?>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        static Regex _RegExOnEventAttributes = new Regex(
            @"<[^>]*?\s(on[^\s\\]{0,20}=([""].*?[""]|['].*?['])).*?(>|\/>)",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        /// <summary>
        /// Sanitizes HTML to some of the most of common XSS attacks
        /// </summary>
        /// <remarks>
        /// This provides rudimentary HTML sanitation catching the most obvious
        /// XSS script attack vectors. For mroe complete HTML Sanitation please look into
        /// a dedicated HTML Sanitizer.
        /// </remarks>
        /// <param name="html">input html</param>
        /// <param name="htmlTagBlacklist">A list of HTML tags that are stripped.</param>
        /// <returns>Sanitized HTML</returns>
        public static string SanitizeHtml(string html, string htmlTagBlacklist = "script|iframe|object|embed|form")
        {
            if (string.IsNullOrEmpty(html))
                return html;

            if (!string.IsNullOrEmpty(htmlTagBlacklist) || htmlTagBlacklist == HtmlSanitizeTagBlackList)
            {
                // Replace Script tags - reused expr is more efficient
                html = _RegExScript.Replace(html, string.Empty);
            }
            else
            {
                html = Regex.Replace(html,
                                        $@"(<({htmlTagBlacklist})\b[^<]*(?:(?!<\/({HtmlSanitizeTagBlackList}))<[^<]*)*<\/({htmlTagBlacklist})>)",
                                        "", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            }

            // Remove javascript: directives
            var matches = _RegExJavaScriptHref.Matches(html);
            foreach (Match match in matches)
            {
                if (match.Groups.Count > 2)
                {
                    var txt = match.Value.Replace(match.Groups[2].Value, "unsupported:");
                    html = html.Replace(match.Value, txt);
                }
            }

            // Remove onEvent handlers from elements
            matches = _RegExOnEventAttributes.Matches(html);
            foreach (Match match in matches)
            {
                var txt = match.Value;
                if (match.Groups.Count > 1)
                {
                    var onEvent = match.Groups[1].Value;
                    txt = txt.Replace(onEvent, string.Empty);
                    if (!string.IsNullOrEmpty(txt))
                        html = html.Replace(match.Value, txt);
                }
            }

            return html;
        }



        #region Obsolete
        /// <summary>
        /// Determines whether a string is empty (null or zero length)
        /// </summary>
        /// <param name="text">Input string</param>
        /// <returns>true or false</returns>
        [Obsolete("Use string.IsNullOrEmpty() instead")]
        public static bool Empty(string text)
        {
            return string.IsNullOrEmpty(text);
        }

        /// <summary>
        /// Determines wheter a string is empty (null or zero length)
        /// </summary>
        /// <param name="text">Input string (in object format)</param>
        /// <returns>true or false</returns>        
        [Obsolete("Use string.IsNullOrEmpty() instead")]
        public static bool Empty(object text)
        {
            return string.IsNullOrEmpty(text as string);
        }

        #endregion
    }
}
