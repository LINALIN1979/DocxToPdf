using System;

namespace DocxToPdf
{
    class Tools
    {
        public enum SizeEnum
        {
            /// <summary>
            /// Twentieths of a point.
            /// </summary>
            TwentiethsOfPoint, // Indentation

            /// <summary>
            /// 100th of a character, need to provide font size (point).
            /// </summary>
            HundredthsOfCharacter, // Indentation

            /// <summary>
            /// Half-point, used for font size.
            /// </summary>
            HalfPoint, // FontSize

            /// <summary>
            /// Eighths of a point, used for line border.
            /// </summary>
            LineBorder,

            /// <summary>
            /// 240th of a line, need to provide font size (point).
            /// </summary>
            TwoHundredFoutiesthOfLine,

            /// <summary>
            /// Size string (cm, mm, in, pt, pc, px), e.g. "12cm".
            /// </summary>
            String,
        }

        /// <summary>
        /// Convert different units to points.
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="type">Obj type.</param>
        /// <param name="fontSizePoint">Only used for sizeEnum.HundredthsOfCharacter.</param>
        /// <returns>Return points (72 points = 1 inch).</returns>
        public static float ConvertToPoint(object obj, SizeEnum type, float fontSizePoint)
        {
            float ret = 0f;

            if (type == SizeEnum.String)
            {
                String lsize = obj.ToString().Replace(" ", "").ToLower();
                if (lsize.IndexOf("cm") >= 0)
                    ret = (Convert.ToSingle(lsize.Replace("cm", "")) / 2.545f) * 72f;
                else if (lsize.IndexOf("mm") >= 0)
                    ret = (Convert.ToSingle(lsize.Replace("mm", "")) / 25.45f) * 72f;
                else if (lsize.IndexOf("in") >= 0)
                    ret = Convert.ToSingle(lsize.Replace("in", "")) * 72f;
                else if (lsize.IndexOf("pt") >= 0)
                    ret = Convert.ToSingle(lsize.Replace("pt", ""));
                else if (lsize.IndexOf("pc") >= 0)
                    ret = Convert.ToSingle(lsize.Replace("pc", "")) * 12f;
                else if (lsize.IndexOf("px") >= 0)
                    ret = Convert.ToSingle(lsize.Replace("px", "")); // assume 72 dpi
                else // default is px
                    ret = Convert.ToSingle(lsize);
            }
            else
            {
                float num = (float)Convert.ToSingle(obj);
                switch (type)
                {
                    case SizeEnum.TwentiethsOfPoint:
                        ret = (float)(num / 20);
                        break;
                    case SizeEnum.HundredthsOfCharacter:
                        ret = (float)(num / 100) * fontSizePoint;
                        break;
                    case SizeEnum.HalfPoint:
                        ret = (float)(num / 2);
                        break;
                    case SizeEnum.LineBorder:
                        ret = (float)(num / 8);
                        break;
                    case SizeEnum.TwoHundredFoutiesthOfLine:
                        ret = (float)(num / 240) * fontSizePoint;
                        break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Get color's brightness.
        /// </summary>
        /// <param name="color">Color string, e.g. FF0000.</param>
        /// <returns></returns>
        public static float RgbBrightness(String color)
        {
            if (color == "auto") // TODO: handle auto color
                color = "0";

            int rgb = Convert.ToInt32(color, 16);
            return RgbBrightness((rgb & 0xff0000) >> 16, (rgb & 0xff00) >> 8, (rgb & 0xff));
        }

        /// <summary>
        /// Get color's brightness.
        /// </summary>
        /// <param name="r">R</param>
        /// <param name="g">G</param>
        /// <param name="b">B</param>
        /// <returns></returns>
        public static float RgbBrightness(int r, int g, int b)
        {
            if (r > 255) r = 255;
            if (r < 0)   r = 0;
            if (g > 255) g = 255;
            if (g < 0)   g = 0;
            if (b > 255) b = 255;
            if (b < 0)   b = 0;

            // http://stackoverflow.com/questions/596216/formula-to-determine-brightness-of-rgb-color
            // http://www.codeproject.com/Articles/19045/Manipulating-colors-in-NET-Part-1
            //return (float)(0.2126f * r + 0.7152 * g + 0.0722 * b);

            // http://www.w3.org/TR/AERT#color-contrast
            return (float)(0.299f * r + 0.587 * g + 0.114 * b);
        }

        /// <summary>
        /// Convert percentage string (e.g. "50%" or "2500") to the value between 0~1.
        /// </summary>
        /// <param name="str">Percentage string, e.g. "50%" or "2500".</param>
        /// <returns>Return a float value between 0~1.</returns>
        public static float Percentage(string str)
        {
            try
            {
                str = str.Replace(" ", "");
                return (float)((str.EndsWith("%")) ?
                   (float)(Convert.ToSingle(str.Replace("%", "")) / 100) :
                   (float)(Convert.ToSingle(str) / 50) / 100);
            }
            catch (Exception) { return 0f; }
        }

        /// <summary>
        /// Roman numerals
        /// </summary>
        private static string[][] romanNumerals = new string[][]
        {
            new string[]{"", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"}, // ones
            new string[]{"", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"}, // tens
            new string[]{"", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"}, // hundreds
            new string[]{"", "M", "MM", "MMM"} // thousands
        };

        /// <summary>
        /// Taiwanese numerals
        /// </summary>
        private static string[][] taiwaneseNumerals = new string[][]
        {
            new string[]{"", "一", "二", "三", "四", "五", "六", "七", "八", "九"}, // ones
            new string[]{"零", "一十", "二十", "三十", "四十", "五十", "六十", "七十", "八十", "九十"}, // tens
            new string[]{"零", "一百", "二百", "三百", "四百", "五百", "六百", "七百", "八百", "九百"}, // hundreds
            new string[]{"零", "一千", "二千", "三千", "四千", "五千", "六千", "七千", "八千", "九千"}, // thousands
        };

        private static String IntToAnything(string[][] numerals, int number, bool ignoreZeroes)
        {
            // split integer string into array and reverse array
            char[] intArr = number.ToString().ToCharArray();
            Array.Reverse(intArr);
            String ret = "";
            int end = 0;

            if (ignoreZeroes)
            {
                int intArrMax = intArr.Length - 1;
                while (Convert.ToInt32(intArr[end].ToString()) == 0 && end < intArrMax)
                    end++;
            }

            // starting with the highest place (for 3046, it would be the thousands
            // place, or 3), get the roman numeral representation for that place
            // and add it to the final roman numeral string
            for (int i = intArr.Length - 1; i >= end; i--)
            {
                ret += numerals[i][Convert.ToInt32(intArr[i].ToString())];
            }

            return ret;
        }

        /// <summary>
        /// Convert integer to Roman numeral expression.
        /// </summary>
        /// <param name="number">Integer.</param>
        /// <param name="uppercase">True for uppder case, false for lower case.</param>
        /// <returns></returns>
        public static String IntToRoman(int number, bool uppercase)
        {
            return (uppercase) ? 
                IntToAnything(romanNumerals, number, false) : 
                IntToAnything(romanNumerals, number, false).ToLower();
        }

        /// <summary>
        /// Convert integer to Taiwanese numeral expression.
        /// </summary>
        /// <param name="number">Integer.</param>
        /// <returns></returns>
        public static String IntToTaiwanese(int number)
        {
            return IntToAnything(taiwaneseNumerals, number, true);
        }

        /// <summary>
        /// Convert VML string to System.Drawing.Image object.
        /// </summary>
        /// <param name="vml">VML string</param>
        /// <param name="width">Width in points</param>
        /// <param name="height">Height in points</param>
        /// <returns></returns>
        //public static Image ConvertVmlToImage(string vml, float width, float height)
        //{
        //    vml = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\"" +
        //          "      xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
        //          "      xmlns=\"http://www.w3.org/1999/xhtml\">" +
        //          "<head><STYLE>v\\:* {behavior:url(#default#VML);}</STYLE></head>" +
        //          "<body width=\"" + width + "pt\" height=\"" + height + "pt\">" + 
        //          vml + "</body></html>";

        //    Image img = null;
        //    using (MemoryStream vmlMs = new MemoryStream(Encoding.UTF8.GetBytes(vml)))
        //    {
        //        XPathDocument doc = new XPathDocument(vmlMs);

        //        XslCompiledTransform xsl = new XslCompiledTransform(true);
        //        xsl.Load("vml2svg.xsl", null, new XmlUrlResolver());

        //        StringWriter sw = new StringWriter();
        //        using (XmlTextWriter writer = new XmlTextWriter(sw))
        //            xsl.Transform(doc, null, writer);

        //        String svgStr = sw.ToString();
        //        if (svgStr.Length > 0)
        //        {
        //            // For debugging
        //            string svgFile = @"C:\Users\LINALIN\Desktop\test.svg";
        //            File.WriteAllText(svgFile, svgStr);

        //            SvgDocument svgDocument = SvgDocument.FromSvg<SvgDocument>(svgStr);
        //            //svgDocument.Ppi = 96;
        //            img = svgDocument.Draw();

        //            // For debugging
        //            string outFile = @"C:\Users\LINALIN\Desktop\test.png";
        //            img.Save(outFile, System.Drawing.Imaging.ImageFormat.Png);
        //        }
        //    }
        //    return img;
        //}
    }
}
