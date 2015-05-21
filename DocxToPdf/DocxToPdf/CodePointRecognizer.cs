using System;
using System.Collections.Generic;

namespace DocxToPdf
{
    /// <summary>
    /// Enumeration of font types.
    /// </summary>
    public enum FontTypeEnum
    {
        // Complex Script introduction:
        //   https://xmlgraphics.apache.org/fop/1.1/complexscripts.html
        //   http://jrgraphix.net/research/unicode_blocks.php

        ASCII,
        ComplexScript, // complex script, e.g. Thai, Arabic, Hebrew, and right-to-left languages
        EastAsian,
        HighANSI,
        UNKNOWN
    }

    /// <summary>
    /// Font type information, includes the name of Unicode block, and the font type should be used for this Unicode block.
    /// </summary>
    public class FontTypeInfo
    {
        /// <summary>
        /// Name of the Unicode block.
        /// </summary>
        public String Name = "";

        /// <summary>
        /// The font type (e.g. ComplexScript or EastAsian) of the Unicode block.
        /// </summary>
        public FontTypeEnum FontType = FontTypeEnum.UNKNOWN;

        /// <summary>
        /// If the vaule of the hint attribute is eastAsia then East Asian font is used, otherwise High ANSI font is used.
        /// </summary>
        public bool UseEastAsiaIfhintIsEastAsia = false;
    }

    internal class UnicodeBlockRange
    {
        public int Begin;
        public int End;
        public UnicodeBlockRange(int begin, int end)
        {
            this.Begin = begin;
            this.End = end;
        }
    }

    /// <summary>
    /// For internal usage, store the Unicode block related information.
    /// </summary>
    internal class UnicodeBlock : FontTypeInfo
    {
        private List<UnicodeBlockRange> Ranges = new List<UnicodeBlockRange>();

        public UnicodeBlock(String name, UnicodeBlockRange[] blocks, FontTypeEnum fontType, bool useEastAsiaIfhintIsEastAsia)
        {
            this.Name = name;
            this.Ranges.AddRange(blocks);
            this.FontType = fontType;
            this.UseEastAsiaIfhintIsEastAsia = useEastAsiaIfhintIsEastAsia;
        }

        /// <summary>
        /// Check the target Unicode value belongs to this code point or not.
        /// </summary>
        /// <param name="unicode">Target Unicode value.</param>
        /// <returns>Return true means the target Unicode value belongs to this code point, otherwise return false.</returns>
        public bool IsIn(int unicode)
        {
            return (this.Ranges.Find(r => (unicode >= r.Begin && unicode <= r.End)) != null) ? true : false;
        }
    }

    public class CodePointRecognizer
    {
        // https://social.msdn.microsoft.com/Forums/en-US/1bf1f185-ee49-4314-94e7-f4e1563b5c00/finding-which-font-is-to-be-used-to-displaying-a-character-from-pptx-xml?forum=os_binaryfile
        // Unicode character in a run, the font slot can be determined using the following two-step methodology:
        //   1. Use the table below to decide the classification of the content, based on its Unicode code point.
        //   2. If, after the first step, the character falls into East Asian classification and the value of the 
        //      hint attribute is eastAsia, then the character should use East Asian font slot
        //      1. Otherwise, if there is <w:cs/> or <w:rtl/> in this run, then the character should use Complex 
        //         Script font slot, regardless of its Unicode code point.
        //         1. Otherwise, the character is decided using the font slot that is corresponding to the 
        //            classification in the table above.
        // Once the font slot for the run has been determined using the above steps, the appropriate formatting 
        // elements (either complex script or non-complex script) will affect the content.

        private static List<UnicodeBlock> blocks = new List<UnicodeBlock>(new UnicodeBlock[] {
            new UnicodeBlock("Basic Latin", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0000, 0x007F), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Latin-1 Supplement", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x00A0, 0x00FF), }, FontTypeEnum.HighANSI, false), // TODO: exception not implemented
            new UnicodeBlock("Latin Extended-A", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0100, 0x017F), }, FontTypeEnum.HighANSI, false), // TODO: exception not implemented
            new UnicodeBlock("Latin Extended-B", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0180, 0x024F), }, FontTypeEnum.HighANSI, false), // TODO: exception not implemented
            new UnicodeBlock("IPA Extensions", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0250, 0x02AF), }, FontTypeEnum.HighANSI, false), // TODO: exception not implemented
            new UnicodeBlock("Spacing Modifier Letters", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x02B0, 0x02FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Combining Diacritical Marks", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0300, 0x036F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Greek", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0370, 0x03CF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Cyrillic", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0400, 0x04FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Hebrew", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0590, 0x05FF), }, FontTypeEnum.ASCII, false),

            new UnicodeBlock("Arabic", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0600, 0x06FF), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Syriac", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0700, 0x074F), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Arabic Supplement", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0750, 0x077F), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Thaana", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x0780, 0x07BF), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Hangul Jamo", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x1100, 0x11FF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Latin Extended Additional", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x1E00, 0x1EFF), }, FontTypeEnum.HighANSI, false), // TODO: exception not implemented
            new UnicodeBlock("Greek Extended", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x1F00, 0x1FFF), }, FontTypeEnum.HighANSI, false),
            new UnicodeBlock("General Punctuation", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2000, 0x206F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Superscripts and Subscripts", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2070, 0x209F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Currency Symbols", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x20A0, 0x20CF), }, FontTypeEnum.HighANSI, true),

            new UnicodeBlock("Combining Diacritical Marks for Symbols", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x20D0, 0x20FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Letter-like Symbols", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2100, 0x214F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Number Forms", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2150, 0x218F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Arrows", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2190, 0x21FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Mathematical Operators", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2200, 0x22FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Miscellaneous Technical", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2300, 0x23FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Control Pictures", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2400, 0x243F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Optical Character Recognition",
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2440, 0x245F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Enclosed Alphanumerics",
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2460, 0x24FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Box Drawing", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2500, 0x257F), }, FontTypeEnum.HighANSI, true),

            new UnicodeBlock("Block Elements", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2580, 0x259F), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Geometric Shapes", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x25A0, 0x25FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Miscellaneous Symbols", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2600, 0x26FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Dingbats", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2700, 0x27BF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("CJK Radicals Supplement", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2E80, 0x2EFF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Kangxi Radicals", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2F00, 0x2FDF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Ideographic Description Characters", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x2FF0, 0x2FFF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("CJK Symbols and Punctuation", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3000, 0x303F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Hiragana", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3040, 0x309F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Katakana", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x30A0, 0x30FF), }, FontTypeEnum.EastAsian, false),

            new UnicodeBlock("Bopomofo", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3100, 0x312F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Hangul Compatibility Jamo", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3130, 0x318F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Kanbun", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3190, 0x319F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Enclosed CJK Letters and Months", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3200, 0x32FF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("CJK Compatibility", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3300, 0x33FF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("CJK Unified Ideographs Extension A", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x3400, 0x4DBF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("CJK Unified Ideographs", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0x4E00, 0x9FAF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Yi Syllables", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xA000, 0xA48F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Yi Radicals", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xA490, 0xA4CF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Hangul Syllables", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xAC00, 0xD7AF), }, FontTypeEnum.EastAsian, false),

            new UnicodeBlock("High Surrogates", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xD800, 0xDB7F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("High Private Use Surrogates", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xDB80, 0xDBFF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Low Surrogates", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xDC00, 0xDFFF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Private Use Area", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xE000, 0xF8FF), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("CJK Compatibility Ideographs", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xF900, 0xFAFF), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Alphabetic Presentation Forms1", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFB00, 0xFB1C), }, FontTypeEnum.HighANSI, true),
            new UnicodeBlock("Alphabetic Presentation Forms2", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFB1D, 0xFB4F), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Arabic Presentation Forms-A", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFB50, 0xFDFF), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("CJK Compatibility Forms", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFE30, 0xFE4F), }, FontTypeEnum.EastAsian, false),
            new UnicodeBlock("Small Form Variants", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFE50, 0xFE6F), }, FontTypeEnum.EastAsian, false),

            new UnicodeBlock("Arabic Presentation Forms-B", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFE70, 0xFEFE), }, FontTypeEnum.ASCII, false),
            new UnicodeBlock("Halfwidth and Fullwidth Forms", 
	            new UnicodeBlockRange[] { new UnicodeBlockRange(0xFF00, 0xFFEF), }, FontTypeEnum.EastAsian, false),
        });

        /// <summary>
        /// Get font type information by a Unicode value. Font type information indicates the character should be display in ASCII/Complex Script/EastAsian/HighANSI.
        /// </summary>
        /// <param name="unicode">Unicode value of the character.</param>
        /// <returns>Return font type information.</returns>
        public static FontTypeInfo GetFontType(int unicode)
        {
            FontTypeInfo ret = new FontTypeInfo();

            UnicodeBlock unicodeBlock = blocks.Find(block => block.IsIn(unicode));
            if (unicodeBlock != null)
            {
                ret.Name = unicodeBlock.Name;
                ret.FontType = unicodeBlock.FontType;
                ret.UseEastAsiaIfhintIsEastAsia = unicodeBlock.UseEastAsiaIfhintIsEastAsia;
            }
            return ret;
        }
    }
}
