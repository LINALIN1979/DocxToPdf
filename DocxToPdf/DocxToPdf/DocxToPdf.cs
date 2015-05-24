using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using iTSPdf = iTextSharp.text.pdf;
using iTSText = iTextSharp.text;
using iTSTextColor = iTextSharp.text.BaseColor;
using Vml = DocumentFormat.OpenXml.Vml;
using Word = DocumentFormat.OpenXml.Wordprocessing;

// TODO: handle sectPr w:docGrid w:linePitch (how many lines per page)
// TODO: handle tblLook $17.7.6 (conditional formatting), it will define firstrow/firstcolum/...etc styles in styles.xml

namespace DocxToPdf
{
    #region Class FontFactory
    public class FontFactory
    {
        private static String fontFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        private class FontInfo
        {
            public String name;
            public String path;
        }
        private static List<FontInfo> fonts = new List<FontInfo>();

        /// <summary>
        /// Get font name from file, e.g. "Arial" from @"C:\windows\Fonts\arial.ttf".
        /// </summary>
        /// <param name="file">The full path of the font file.</param>
        /// <returns>The font name in English-United States.</returns>
        private static String getFontNameFromFile(String file)
        {
            PrivateFontCollection fontCollection = new PrivateFontCollection();
            fontCollection.AddFontFile(file);
            if (fontCollection.Families.Length > 0)
                return fontCollection.Families[0].GetName(0x409); // return font name in English-United States
            //return fontCollection.Families[0].GetName(0x404); // return font name in Chinese-Taiwan
            else
                return null;
        }

        /// <summary>
        /// The constructor of FontFactory.
        /// </summary>
        static FontFactory()
        {
            String[] files = Directory.GetFiles(FontFactory.fontFolderPath);
            foreach (String file in files)
            {
                String fontName = getFontNameFromFile(file);
                if (fontName != null)
                {
                    FontInfo font = new FontInfo();
                    font.path = file + (file.EndsWith(".ttc") ? ",0" : ""); // "xxx.ttc,0"
                    font.name = fontName;
                    FontFactory.fonts.Add(font);
                    //Tools.Output(String.Format("{0}: {1}", font.name, font.path));
                }
            }
        }

        /// <summary>
        /// Search for font file by font name. 
        /// </summary>
        /// <param name="name">Font name (e.g. Arial or 標楷體).</param>
        /// <returns>The full path of font file or null if can't find.</returns>
        public static String FuzzySearchByName(String name)
        {
            if (name == null)
                return null;

            String path = null;

            // translate font name from it's language to English-United States
            try
            {
                FontFamily fontFamily = new FontFamily(name);
                String enName = fontFamily.GetName(0x409);

                FontInfo fontInfo = FontFactory.fonts.FirstOrDefault(font => font.name.ToLower() == enName.ToLower());
                if (fontInfo == null)
                    fontInfo = FontFactory.fonts.FirstOrDefault(font =>
                        (font.name.IndexOf(enName, StringComparison.OrdinalIgnoreCase) >= 0) ? true : false);
                if (fontInfo != null)
                    path = fontInfo.path;
            }
            catch (ArgumentException e)
            {
                Debug.Write(String.Format("Failed to search FontInfo of \"{0}\", error = {1}", name, e));
            }
            return path;
        }

        /// <summary>
        /// Create iTextSharp.text.pdf.BaseFont by font name. This method doesn't configure font size, color or any other properties for BaseFont.
        /// </summary>
        /// <param name="fontName">Font name (e.g. Arial or 標楷體)</param>
        /// <returns>Return iTextSharp.text.pdf.BaseFont object or null if can't find font.</returns>
        public static iTSPdf.BaseFont CreatePdfBaseFontByFontName(String fontName)
        {
            if (fontName == null)
                return null;

            iTSPdf.BaseFont baseFont = null;
            String fontFilePath = FontFactory.FuzzySearchByName(fontName);
            if (fontFilePath != null)
            {
                baseFont = iTSPdf.BaseFont.CreateFont(
                    fontFilePath,
                    iTSPdf.BaseFont.IDENTITY_H,
                    iTSPdf.BaseFont.EMBEDDED);
            }
            return baseFont;
        }

        /// <summary>
        /// Create iTextSharp.text.Font by specifying the BaseFont, font size, and other font properties. 
        /// </summary>
        /// <param name="baseFont">iTextSharp.text.pdf.BaseFont.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="bold">Bold property.</param>
        /// <param name="italic">Italic property.</param>
        /// <param name="strike">Strike property.</param>
        /// <param name="color">Color.</param>
        /// <returns></returns>
        public static iTSText.Font CreateFont(iTSPdf.BaseFont baseFont,
            float fontSize,
            Word.Bold bold, Word.Italic italic, Word.Strike strike, Word.Color color)
        {
            iTSText.Font ret = null;

            if (baseFont == null)
                return ret;

            ret = new iTSText.Font(baseFont);

            int rgb = 0;
            if (color != null && color.Val != null)
            {
                if (color.Val.Value != "auto")
                    rgb = Convert.ToInt32(color.Val.Value, 16);
            }
            //ch.SetWordSpacing(8f);
            ret.SetStyle(
                (StyleHelper.GetToggleProperty(bold) ? iTSText.Font.BOLD : 0) |
                (StyleHelper.GetToggleProperty(italic) ? iTSText.Font.ITALIC : 0) |
                (StyleHelper.GetToggleProperty(strike) ? iTSText.Font.STRIKETHRU : 0)
                );
            if (rgb != 0) { ret.SetColor((rgb & 0xff0000) >> 16, (rgb & 0xff00) >> 8, (rgb & 0xff)); }
            if (fontSize > 0f) ret.Size = fontSize;

            Debug.Write(String.Format("{0}, {1}pt, {2}{3}{4}, RGB {5}",
                baseFont.FullFontName[0][3],
                (fontSize > 0f) ? fontSize : 12f,
                StyleHelper.GetToggleProperty(bold) ? "B" : "b",
                StyleHelper.GetToggleProperty(italic) ? "I" : "i",
                StyleHelper.GetToggleProperty(strike) ? "S" : "s",
                (color != null) ? color.Val.Value : "000000"
                ));
            return ret;
        }
    }
    #endregion

    public class DocxToPdf
    {
        private WordprocessingDocument doc = null;
        private StyleHelper stHelper = null;
        private iTSText.Document pdfDoc = null;

        /// <summary>
        /// Convert .docx to .pdf
        /// </summary>
        /// <param name="srcFile">Docx file path.</param>
        /// <param name="dstFile">Target output PDF file path.</param>
        public DocxToPdf(String srcFile, String dstFile)
        {
            // working with fonts
            // http://www.mikesdotnetting.com/article/81/itextsharp-working-with-fonts
            //   XP - @"C:\WINDOWS\Fonts"
            //string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) +
            //    @"\..\Fonts\kaiu.ttf";
            //iTSText.FontFactory.Register(fontPath);
            //iTSText.FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts");

            MemoryStream ms = new MemoryStream();
            using (this.doc = WordprocessingDocument.Open(srcFile, false))
            {
                this.stHelper = new StyleHelper(this.doc);

                Word.Body body = this.doc.MainDocumentPart.Document.Body;

                //// Remove empty page that shouldn't create new PDF page
                //IEnumerable<Word.Paragraph> pgs = doc.MainDocumentPart.Document.Body.Elements<Word.Paragraph>();
                //foreach (Word.Paragraph pg in pgs)
                //{
                //    bool noText = (pg.Descendants<Word.Text>().Count() == 0);
                //    bool hasBR = (pg.Descendants<Word.Break>().Count() > 0);
                //    bool hasLastRenderedPageBreak = (pg.Descendants<Word.LastRenderedPageBreak>().Count() > 0);

                //    if (noText & hasBR & hasLastRenderedPageBreak)
                //    {
                //        if (pg.Parent != null)
                //            pg.Remove();
                //    }
                //}

                int previousIndex = 0;
                int currentIndex = 0;
                for (Word.SectionProperties section = this.stHelper.CurrentSectPr; section != null; section = this.stHelper.NextSectPr)
                {
                    Word.PageSize size = StyleHelper.GetElement<Word.PageSize>(section);
                    Word.PageMargin margin = StyleHelper.GetElement<Word.PageMargin>(section);
                    iTSText.Rectangle pageSize = new iTSText.Rectangle(
                        Tools.ConvertToPoint(size.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, 0),
                        Tools.ConvertToPoint(size.Height.Value, Tools.SizeEnum.TwentiethsOfPoint, 0));
                    //// For PDF conversion, don't need to handle landscape property, it's mainly for printer
                    //bool landscape =
                    //    (size.Orient == null) ? false :
                    //    (size.Orient.Value == Word.PageOrientationValues.Landscape) ? true : false;

                    Console.WriteLine(String.Format("Page Settings:\n  {0}x{1}, left:{2}, right:{3}, top:{4}, bottom:{5}",
                        pageSize.Width, pageSize.Height,
                        Tools.ConvertToPoint(margin.Left.Value, Tools.SizeEnum.TwentiethsOfPoint, 0),
                        Tools.ConvertToPoint(margin.Right.Value, Tools.SizeEnum.TwentiethsOfPoint, 0),
                        Tools.ConvertToPoint(margin.Top.Value, Tools.SizeEnum.TwentiethsOfPoint, 0),
                        Tools.ConvertToPoint(margin.Bottom.Value, Tools.SizeEnum.TwentiethsOfPoint, 0)));
                    // TODO: footer = bottom edge of page to bottom edge of footer

                    float leftMargin = Tools.ConvertToPoint(margin.Left.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    float rightMargin = Tools.ConvertToPoint(margin.Right.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    float topMargin = Tools.ConvertToPoint(margin.Top.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    float bottomMargin = Tools.ConvertToPoint(margin.Bottom.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    float hdrMargin = (margin.Header != null && margin.Header.HasValue) ? 
                        Tools.ConvertToPoint(margin.Header.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f) : 0f;
                    float footrMargin = (margin.Footer != null && margin.Footer.HasValue) ?
                        Tools.ConvertToPoint(margin.Footer.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f) : 0f;

                    // top margin = greater one of margin.Top or (margin.Header + header height) (if has header)
                    List<iTSText.IElement> hdr = this.BuildHeader(section);
                    if (hdr.Count == 0) hdrMargin = 0f; // no header
                    else
                    {
                        float hdrHeight = this.calculateHeight(hdr, pageSize.Width - leftMargin - rightMargin);
                        if (hdrHeight > 0f) hdrMargin += hdrHeight;
                    }

                    // bottom margin = greater one of margin.Bottom or (margin.footer + footer height) (if has footer)
                    List<iTSText.IElement> footr = new List<iTSText.IElement>();
                    if (footr.Count == 0) footrMargin = 0f; // no footer
                    else
                    {
                        float footrHeight = -1f;
                        if (footrHeight > 0f) footrMargin += footrHeight;
                    }

                    // SetPageSize() and SetMargins can be used at any time in the document’s
                    // creation process, but be aware that the change will never affect 
                    // the current page, only the next page.
                    if (pdfDoc == null)
                    {
                        //doc.SetMargins(doc.LeftMargin, doc.RightMargin, posY + hdrHeight, doc.BottomMargin);
                        //writer.PageNumber;

                        pdfDoc = new iTSText.Document(pageSize, leftMargin, rightMargin,
                            (topMargin > hdrMargin) ? topMargin : hdrMargin, 
                            (bottomMargin > footrMargin) ? bottomMargin : footrMargin);
                        iTSPdf.PdfWriter writer = iTSPdf.PdfWriter.GetInstance(pdfDoc, ms);
                        writer.SetEncryption( // http://stackoverflow.com/questions/8419649/lock-pdf-against-editing-using-itextsharp
                            null, // null user password => users can open document WITHOUT password
                            System.Text.Encoding.UTF8.GetBytes("ppjj"), // owner password => required to modify document/permissions
                            iTSPdf.PdfWriter.ALLOW_PRINTING | iTSPdf.PdfWriter.ALLOW_COPY, // bitwise or => see iText API for permission parameter: http://api.itextpdf.com/itext/com/itextpdf/text/pdf/PdfWriter.html
                            iTSPdf.PdfWriter.ENCRYPTION_AES_128 // encryption level also in documentation referenced above
                            );
                        writer.ViewerPreferences = iTSPdf.PdfWriter.FitWindow
                                | iTSPdf.PdfWriter.PageLayoutOneColumn
                                | iTSPdf.PdfWriter.PageModeUseNone;
                        writer.PageEvent = new pdfPage(this); // header & footer

                        pdfDoc.Open();

//                        // silent printing when opening PDF file
//                        // http://www.sanjbee.com/content/?p=96
//                        // http://renjin.blogspot.tw/2010/10/printing-pdf-file-from-aspnet.html
//                        String silentPrint = @"
//                                                var pp = this.getPrintParams();
//                                                pp.interactive = pp.constants.interactionLevel.silent;
//                                                pp.pageHandling = pp.constants.handling.none;
//                                                var fv = pp.constants.flagValues;
//                                                pp.flags |= fv.setPageSize;
//                                                pp.flags |= (fv.suppressCenter | fv.suppressRotate);
//                                                this.print(pp);";
//                        writer.AddJavaScript(silentPrint);
                    }
                    else
                    {
                        pdfDoc.SetPageSize(pageSize);
                        pdfDoc.SetMargins(leftMargin, rightMargin,
                            (topMargin > hdrMargin) ? topMargin : hdrMargin,
                            (bottomMargin > footrMargin) ? bottomMargin : footrMargin);
                        pdfDoc.NewPage();
                    }

                    currentIndex = this.stHelper.CurrentSectPrIndexInBody;
                    if (currentIndex <= previousIndex)
                    { // nothing to be processed, go on to next elements
                        previousIndex++;
                        continue;
                    }

                    List<OpenXmlElement> openXMLList = body.Elements().ToList().GetRange(previousIndex, currentIndex - previousIndex);
                    bool previousIsParagraph = false;
                    for (int i = 0; i < openXMLList.Count; i++)
                    {
                        iTSText.IElement element = this.Dispatcher(openXMLList[i]);

                        //try {
                            if (element != null)
                            {
                                if (element.GetType() == typeof(iTSPdf.PdfPTable))
                                {
                                    iTSPdf.PdfPTable table = element as iTSPdf.PdfPTable;
                                    // http://stackoverflow.com/questions/1364435/itextsharp-splitlate-splitrows
                                    // SplitLate = true (default), the table will be split before the next row that does fit on the page.
                                    // SplitLate = false, the row that does not fully fit on the page will be split.
                                    // SplitRows = true (default), the row that does not fit on a page will be split.
                                    // SplitRows = false the row will be omitted.
                                    //  SplitLate && SplitRows: A row that does not fit on the page will be started on the next page and eventually split if it does not fit on that page either.
                                    //  SplitLate && !SplitRows: A row that does not fit on the page will be started on the next page and omitted if it does not fit on that page either.
                                    //  !SplitLate && SplitRows: A row that does not fit on the page will be split and continued on the next page and split again if it too large for the next page too.
                                    //  !SplitLate && !SplitRows: I'm a little unsure about this one. But from the sources it looks like it's the same as SplitLate && !SplitRows: A row that does not fit on the page will be started on the next page and omitted if it does not fit on that page either.
                                    table.SplitLate = false;
                                    table.SplitRows = true;

                                    if (previousIsParagraph)
                                        table.SpacingBefore = 6f; // magic: add default spacing before table

                                    //table.KeepRowsTogether(0);

                                    previousIsParagraph = false;
                                }
                                else if (element.GetType() == typeof(iTSText.Paragraph))
                                {
                                    //// if the last element of paragraph is to create a new 
                                    //// page AND reach the end of using CURRENT sectPr
                                    ////  ==>
                                    //// the new page shouldn't be created because the NEXT 
                                    //// sectPr will create a new page automatically
                                    //if (i == openXMLList.Count - 1)
                                    //{
                                    //    iTSText.Paragraph ph = element as iTSText.Paragraph;
                                    //    int lastIndex = ph.Count - 1;
                                    //    iTSText.Chunk last = ph.GetRange(lastIndex, 1)[0] as iTSText.Chunk;
                                    //    //iTSText.Chunk last = ph.ElementAt(lastIndex) as iTSText.Chunk; // iTextSharp 4.1.6 doesn't support ElementAt operation
                                    //    if (last != null && last.Attributes != null &&
                                    //        last.Attributes.Equals(iTSText.Chunk.NEXTPAGE.Attributes))
                                    //    {
                                    //        ph.RemoveAt(lastIndex);
                                    //    }
                                    //}

                                    previousIsParagraph = true;
                                }

                                pdfDoc.Add(element);
                            }
                        //}  catch (iTSText.DocumentException e) {
                        //    MessageBox.Show(e.StackTrace);
                        //}
                    }
                    previousIndex = currentIndex + 1;

                    // If reach the end of SectPr, call NextSectPr results in CurrentSectPr
                    // becomes null. When pdfDoc.close(), it calls OnEndPage() but no 
                    // header/footer will be drawn because CurrentSectPr is null. So add
                    // below check to avoid the problem.
                    if (this.stHelper.IsCurrentSectPrTheEnd)
                        break;
                }
                pdfDoc.Close();
            }

            if (ms.GetBuffer().Length > 0)
            { // output to file
                using (FileStream file = new FileStream(dstFile, FileMode.Create, FileAccess.Write))
                    file.Write(ms.GetBuffer(), 0, ms.GetBuffer().Length);
            }
            ms.Close();
        }

        /// <summary>
        /// Get font name from RunFonts, only from w:ascii(Theme)/w:hAnsi(Theme)/w:cs(Theme)/w:eastAsia(Theme) but not from w:hint.
        /// </summary>
        /// <param name="runFonts">Word.RunFonts object.</param>
        /// <param name="fontType">FontTypeInfo object.</param>
        /// <returns>Return font name if matches, otherwise return null.</returns>
        private String getFontNameFromRunFontsByFontType(Word.RunFonts runFonts, FontTypeInfo fontType)
        {
            if (runFonts == null)
                return null;

            String fontName = null;
            switch (fontType.FontType)
            {
                case FontTypeEnum.ASCII:
                    fontName =
                        (runFonts.AsciiTheme != null) ? runFonts.AsciiTheme.InnerText :
                        (runFonts.Ascii != null) ? runFonts.Ascii.Value : null;
                    break;
                case FontTypeEnum.ComplexScript:
                    fontName =
                        (runFonts.ComplexScriptTheme != null) ? runFonts.ComplexScriptTheme.InnerText :
                        (runFonts.ComplexScript != null) ? runFonts.ComplexScript.Value : null;
                    break;
                case FontTypeEnum.EastAsian:
                    fontName =
                        (runFonts.EastAsiaTheme != null) ? runFonts.EastAsiaTheme.InnerText :
                        (runFonts.EastAsia != null) ? runFonts.EastAsia.Value : null;
                    break;
                case FontTypeEnum.HighANSI:
                    fontName =
                        (runFonts.HighAnsiTheme != null) ? runFonts.HighAnsiTheme.InnerText :
                        (runFonts.HighAnsi != null) ? runFonts.HighAnsi.Value : null;
                    break;
                default:
                    break;
            }
            return fontName;
        }

        private String getLangFromLanguagesByFontType(Word.Languages langs, FontTypeInfo fontType)
        {
            if (langs == null)
                return null;

            String lang = null;
            switch (fontType.FontType)
            {
                case FontTypeEnum.EastAsian:
                    lang = (langs.EastAsia != null) ? langs.EastAsia.Value : null;
                    break;
                case FontTypeEnum.ComplexScript:
                    lang = (langs.Bidi != null) ? langs.Bidi.Value : null;
                    break;
                default:
                    lang = (langs.Val != null) ? langs.Val.Value : null;
                    break;
            }
            return lang;
        }

        /// <summary>
        /// Get iTSPdf.BaseFont by font name. The font name was extracted from RunFonts. 
        /// This method first try to create BaseFont by font name. If the font name is not normal font name and looks like eastAsia/cs/hAnsi/ascii or majorXX/minorXX, this method will try to search docDefaults and Theme.
        /// If all above procedures can't find font, it will try to use font type and w:lang to find the backup font.
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="language">Language complement, will be used in case can't find BaseFont by font name.</param>
        /// <returns></returns>
        private iTSPdf.BaseFont getBaseFontByFontName(String fontName, FontTypeInfo fontType, Word.Languages language)
        { // Use font name to find font file path
            iTSPdf.BaseFont baseFont = null;
            if (fontName != null)
            {
                baseFont = FontFactory.CreatePdfBaseFontByFontName(fontName);
                if (baseFont == null)
                {
                    // Cannot find font file path has two possibilities 
                    // 1. fontName = (eastAsia/cs/hAnsi/ascii), search docDefault
                    Word.RunFonts docDefaultRunFonts = this.stHelper.GetDocDefaults<Word.RunFonts>(StyleHelper.DocDefaultsType.Character);
                    if (docDefaultRunFonts != null)
                    {
                        if (fontName.IndexOf("eastAsia", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (docDefaultRunFonts.EastAsia != null)
                                baseFont = FontFactory.CreatePdfBaseFontByFontName(docDefaultRunFonts.EastAsia.Value);
                        }
                        else if (fontName.IndexOf("cs", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (docDefaultRunFonts.ComplexScript != null)
                                baseFont = FontFactory.CreatePdfBaseFontByFontName(docDefaultRunFonts.ComplexScript.Value);
                        }
                        else if (fontName.IndexOf("hAnsi", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (docDefaultRunFonts.HighAnsi != null)
                                baseFont = FontFactory.CreatePdfBaseFontByFontName(docDefaultRunFonts.HighAnsi.Value);
                        }
                        else if (fontName.IndexOf("ascii", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (docDefaultRunFonts.Ascii != null)
                                baseFont = FontFactory.CreatePdfBaseFontByFontName(docDefaultRunFonts.Ascii.Value);
                        }
                    }

                    // 2. fontName = majorXXX/minorXXX, search Theme
                    if (baseFont == null)
                    {
                        // http://blogs.msdn.com/b/officeinteroperability/archive/2013/04/22/office-open-xml-themes-schemes-and-fonts.aspx
                        // major fonts are mainly for styles as headings, whereas minor fonts are generally applied to body and paragraph text

                        baseFont = FontFactory.CreatePdfBaseFontByFontName(this.stHelper.GetThemeFontByType(fontName));
                    }
                }
            }

            // *************
            // Can't find rFont, search for w:lang
            if (baseFont == null && language != null)
            {
                String lang = this.getLangFromLanguagesByFontType(language, fontType);
                if (lang != null)
                {
                    // map lang (e.g. zh-TW) to script tag (e.g. Hant)
                    String scriptTag = LangScriptTag.GetScriptTagByLocale(lang);
                    if (scriptTag != null)
                    {
                        baseFont = FontFactory.CreatePdfBaseFontByFontName(this.stHelper.GetThemeFontByScriptTag(scriptTag));
                    }
                }
            }
            return baseFont;
        }

        /// <summary>
        /// Return iTextSharp.text.pdf.BaseFont base on text and fonts definitions of word file.
        /// The algorithm is as following,
        ///  1. Map text's code point to font type (Ascii/EastAsia/ComplexScript/HighAnsi)
        ///  2. Base on font type, go through Run:rPr > Paragraph:StyleId:rPr > Parapgraph:StyleId:pPr > default style > docDefault:rPr > Run:rPr:hint to find out the font name
        ///  3. According to the font name to create iTextSharp.text.pdf.BaseFont and return
        /// </summary>
        /// <param name="run">Run object where the text belongs to. It will be used for traversal.</param>
        /// <param name="text">The text to extract the code point from.</param>
        /// <returns>Return created iTextSharp.text.pdf.BaseFont if success, otherwise return null.</returns>
        private iTSPdf.BaseFont getBaseFont(Word.Run run, String text)
        {
            // 1. rFont
            // Run:rPr > Paragraph:StyleId:rPr > Paragraph:StyleId:pPr > default style > docDefault:rPr > Run:rPr:hint > Theme
            // 2. lang (if no rFont)
            // Run:rPr > Paragraph:StyleId:rPr > default style > docDefault:rPr
            String fontName = null;

            char[] chArray = text.ToCharArray();
            FontTypeInfo fontType = CodePointRecognizer.GetFontType(chArray[0]);

            // *************
            // rFont, search Run first
            Word.RunFonts runFonts = StyleHelper.GetDescendants<Word.RunFonts>(run);
            if (runFonts != null)
            {
                // Special case handling (seems w:hint only exists in Run)
                if (fontType.UseEastAsiaIfhintIsEastAsia &&
                    ((runFonts.Hint != null) ? (runFonts.Hint.Value == Word.FontTypeHintValues.EastAsia) : false))
                    fontType.FontType = FontTypeEnum.EastAsian;

                // Search Run's direct formatting
                fontName = this.getFontNameFromRunFontsByFontType(runFonts, fontType);

                // Search Run's rStyle
                if (fontName == null)
                {
                    if (run.RunProperties != null)
                    {
                        if (run.RunProperties.RunStyle != null)
                            fontName = this.getFontNameFromRunFontsByFontType(
                                (Word.RunFonts)this.stHelper.GetAppliedElement<Word.RunFonts>(
                                    this.stHelper.GetStyleById(run.RunProperties.RunStyle.Val)),
                                fontType);
                    }
                }
            }

            // No matched RunFonts from Run, search Paragraph pStyle
            if (fontName == null)
            {
                OpenXmlElement tmp = run.Parent;
                if ((tmp != null) && (tmp.GetType() == typeof(Word.Paragraph)))
                {
                    Word.Paragraph pg = tmp as Word.Paragraph;
                    if (pg.ParagraphProperties != null)
                    {
                        // Do no search in Paragraph.ParagraphProperties.rPr, that style
                        // is for paragraph glyph

                        if (fontName == null)
                        { // Search from Paragraph's pStyle
                            if (pg.ParagraphProperties.ParagraphStyleId != null)
                            {
                                fontName = this.getFontNameFromRunFontsByFontType(
                                    (Word.RunFonts)this.stHelper.GetAppliedElement<Word.RunFonts>(
                                        this.stHelper.GetStyleById(pg.ParagraphProperties.ParagraphStyleId.Val)),
                                    fontType);
                            }
                        }
                    }
                }
            }

            // Search in default styles
            if (fontName == null)
            {
                fontName = this.getFontNameFromRunFontsByFontType(
                    StyleHelper.GetDescendants<Word.RunFonts>(this.stHelper.GetDefaultStyle(StyleHelper.DefaultStyleType.Character)),
                    fontType);
            }

            // Search in docDefault
            if (fontName == null)
            {
                fontName = this.getFontNameFromRunFontsByFontType(
                    this.stHelper.GetDocDefaults<Word.RunFonts>(StyleHelper.DocDefaultsType.Character),
                    fontType);
            }

            // Still can't find, use w:hint
            if ((fontName == null) && (runFonts != null))
            {
                if (runFonts.Hint != null)
                    fontName = runFonts.Hint.InnerText;
            }

            return getBaseFontByFontName(fontName, fontType, (Word.Languages)this.stHelper.GetAppliedElement<Word.Languages>(run));
        }

        /// <summary>
        /// Only handle container (i.e. table and paragraph)
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public iTSText.IElement Dispatcher(object item)
        {
            iTSText.IElement element = null;

            Type itemtype = item.GetType();
            if (itemtype == typeof(Word.Paragraph))
            {
                element = BuildParagraph(item as Word.Paragraph);
            }
            else if (itemtype == typeof(Word.Table))
            {
                element = BuildTable(item as Word.Table);
            }
            return element;
        }

        enum ProcessingLevel
        {
            DirectFormatting,
            Style,
            DefaultStyle,
            DocDefaults,
            End,
        }

        /// <summary>
        /// Set paragraph's spacingBefore & spacingAfter. Before calling this method, make sure pg.Font.Size was set.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pg"></param>
        private void setParagraphSpacing(Word.Paragraph paragraph, iTSText.Paragraph pg)
        {
            // If add chunk > phrase > paragraph, to set leading to paragraph only take
            // effect for the leading between paragraphs. To set leading between lines,
            // have to set leading of phrases as well.

            float spacingAfter = -1f, spacingBefore = -1f, linespacing = -1f;
            ProcessingLevel nextLevel = ProcessingLevel.DirectFormatting;
            Word.Style st = null;
            while (nextLevel != ProcessingLevel.End)
            {
                Word.SpacingBetweenLines space = null;
                switch (nextLevel)
                {
                    case ProcessingLevel.DirectFormatting:
                        if (paragraph.ParagraphProperties != null)
                        {
                            space = this.stHelper.GetAppliedElement<Word.SpacingBetweenLines>(paragraph.ParagraphProperties);
                            nextLevel = ProcessingLevel.Style;
                        }
                        else
                        {
                            nextLevel = ProcessingLevel.DefaultStyle;
                            continue;
                        }
                        break;
                    case ProcessingLevel.Style:
                        if (st == null)
                        { // the first time to enter ProcessingLevel.Style case
                            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphStyleId != null)
                                st = this.stHelper.GetStyleById(paragraph.ParagraphProperties.ParagraphStyleId.Val);
                        }
                        else
                        { // not the first time, go upstream to style.BasedOn
                            if (st.BasedOn != null && st.BasedOn.Val != null)
                                st = this.stHelper.GetStyleById(st.BasedOn.Val);
                            else
                                st = null;
                        }
                        if (st == null)
                        { // TODO: shouldn't go to default style first, should check 
                          // parents (e.g. maybe the paragraph is in table) have spacing
                          // element or not
                            nextLevel = ProcessingLevel.DefaultStyle;
                            continue;
                        }
                        else
                            space = this.stHelper.GetAppliedElement<Word.SpacingBetweenLines>(st);
                        break;
                    case ProcessingLevel.DefaultStyle:
                        space = this.stHelper.GetAppliedElement<Word.SpacingBetweenLines>(this.stHelper.GetDefaultStyle(StyleHelper.DefaultStyleType.Paragraph));
                        nextLevel = ProcessingLevel.DocDefaults;
                        break;
                    case ProcessingLevel.DocDefaults:
                        space = (this.stHelper.GetDocDefaults<Word.SpacingBetweenLines>(StyleHelper.DocDefaultsType.Paragraph));
                        nextLevel = ProcessingLevel.End;
                        break;
                }

                if (space != null)
                {
                    if (linespacing < 0f) // only set when never set before
                    {
                        if (space.LineRule != null && space.Line != null && space.LineRule.HasValue && space.Line.HasValue)
                        {
                            switch (space.LineRule.Value)
                            {
                                case Word.LineSpacingRuleValues.AtLeast: // interpreted as twip
                                    float spacePoint = Tools.ConvertToPoint(space.Line.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                                    float wordDefaultAtLeastLineSpacing = 1.3f; // magic: word default leading
                                    if (spacePoint >= (pg.Font.CalculatedSize * wordDefaultAtLeastLineSpacing))
                                        linespacing = spacePoint;
                                    else
                                        linespacing = wordDefaultAtLeastLineSpacing * pg.Font.CalculatedSize;
                                    break;
                                case Word.LineSpacingRuleValues.Exact: // interpreted as twip
                                    linespacing = Tools.ConvertToPoint(space.Line.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                                    break;
                                case Word.LineSpacingRuleValues.Auto: // interpreted as 240th of a line
                                    linespacing = (Convert.ToSingle(space.Line.Value) / 240) * pg.Font.CalculatedSize;
                                    break;
                            }
                        }
                    }

                    if (spacingAfter < 0f) // only set when never set before
                    {
                        if (space.After != null && space.After.HasValue)
                            spacingAfter = Tools.ConvertToPoint(space.After.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        else if (space.AfterLines != null && space.AfterLines.HasValue)
                            spacingAfter = Tools.ConvertToPoint(space.After.Value, Tools.SizeEnum.TwoHundredFoutiesthOfLine, pg.Font.CalculatedSize);
                    }

                    if (spacingBefore < 0f) // only set when never set before
                    {
                        if (space.Before != null && space.Before.HasValue)
                            spacingBefore = Tools.ConvertToPoint(space.Before.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        else if (space.BeforeLines != null && space.BeforeLines.HasValue)
                            spacingBefore = Tools.ConvertToPoint(space.Before.Value, Tools.SizeEnum.TwoHundredFoutiesthOfLine, pg.Font.CalculatedSize);
                    }
                }
            }
            
            //// adjust spacingBefore based on linespacing
            //if (spacingBefore > 0f)
            //{ // minus linespacing from spacingBefore
            //    spacingBefore -= ((pg.TotalLeading - (pg.Font.CalculatedSize)) > 0f) ? (pg.TotalLeading - pg.Font.CalculatedSize) : 0f;
            //}

            if (linespacing > 0f)
                pg.SetLeading(linespacing, 0f);
            if (spacingAfter > 0f)
                pg.SpacingAfter = spacingAfter;
            if (spacingBefore > 0f)
                pg.SpacingBefore = spacingBefore;
        }

        /// <summary>
        /// Handle autoSpaceDE and autoSpaceDN of paragraph.
        ///  autoSpaceDE: Automatically Adjust Spacing of Latin and East Asian Text
        ///  autoSpaceDN: Automatically Adjust Spacing of Number and East Asian Text
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pg"></param>
        private void setParagraphAutoSpace(Word.Paragraph paragraph, iTSText.Paragraph pg)
        {
            Word.AutoSpaceDE asde = (Word.AutoSpaceDE)this.stHelper.GetAppliedElement<Word.AutoSpaceDE>(paragraph);
            bool autoSpaceDE = (asde == null) ? true : (asde.Val == null) ? true : (asde.Val.Value); // omitted means true, different from other toggle property

            Word.AutoSpaceDN asdn = (Word.AutoSpaceDN)this.stHelper.GetAppliedElement<Word.AutoSpaceDN>(paragraph);
            bool autoSpaceDN = (asdn == null) ? true : (asdn.Val == null) ? true : (asdn.Val.Value); // omitted means true, different from other toggle property

            if (autoSpaceDE || autoSpaceDN)
            {
                if (pg.Chunks.Count >= 2)
                {
                    // ======
                    // Method 1: Flat Chunks to Paragraph (i.e. remove Phrases if any)
                    //List<iTSText.Chunk> chunks = new List<iTSText.Chunk>(pg.Chunks);
                    List<iTSText.Chunk> chunks = pg.Chunks.Cast<iTSText.Chunk>().ToList();
                    pg.Clear();
                    pg.AddRange(chunks);
                    // ------
                    // Method 2: Flat Chunks to Paragraph (i.e. remove Phrases if any) 
                    // and combine the adjancent Chunks that have the same font style
                    // (in order to make word spacing looks more like Office Word, but
                    // seems no help...)
                    //List<iTSText.Chunk> chunks = new List<iTSText.Chunk>();
                    //int index = pg.Chunks.Count - 1;
                    //while (index >= 0)
                    //{
                    //    iTSText.Font startFont = pg.Chunks[index].Font;
                    //    int j = index - 1;
                    //    if (j < 0) { // only one chunk, just add it then quit
                    //        chunks.Insert(0, pg.Chunks[index]);
                    //        break;
                    //    }
                    //    else {
                    //        while (j >= 0) { // search the front chunks until the font is different
                    //            if (pg.Chunks[j].Font.CompareTo(startFont) != 0)
                    //                break;
                    //            else // same font, go on
                    //                j--;
                    //        }
                    //        String combined = String.Empty;
                    //        for (; index >= j + 1; index--)
                    //            combined = pg.Chunks[index].Content + combined;
                    //        chunks.Insert(0, new iTSText.Chunk(combined, startFont));
                    //        index = j;
                    //    }
                    //}
                    //pg.Clear();
                    //pg.AddRange(chunks);
                    // ======

                    for (int i = pg.Chunks.Count - 2, j = i + 1; i >= 0; i--, j = i + 1)
                    {
                        iTSText.Chunk ich = pg.Chunks[i] as iTSText.Chunk;
                        iTSText.Chunk jch = pg.Chunks[j] as iTSText.Chunk;

                        // bypass line break & page break
                        if (ich.Content.Equals(iTSText.Chunk.NEWLINE.Content) ||
                            ich.Content.Equals(iTSText.Chunk.NEXTPAGE.Content) ||
                            jch.Content.Equals(iTSText.Chunk.NEWLINE.Content) ||
                            jch.Content.Equals(iTSText.Chunk.NEXTPAGE.Content))
                            continue;

                        char frontChar = ich.Content.ToCharArray().Last();
                        char rearChar = jch.Content.ToCharArray().First();
                        FontTypeInfo frontCharType = CodePointRecognizer.GetFontType(frontChar);
                        FontTypeInfo rearCharType = CodePointRecognizer.GetFontType(rearChar);

                        // bypass space and line feed
                        List<char> ignoredChars = new List<char>() { ' ', '\u00A0', '\n' };
                        if (ignoredChars.Contains(frontChar) || ignoredChars.Contains(rearChar))
                            continue;

                        if (((frontCharType.FontType == FontTypeEnum.EastAsian) &&
                            ((rearCharType.FontType != FontTypeEnum.EastAsian) && (rearCharType.FontType != FontTypeEnum.ComplexScript)))
                            ||
                            ((rearCharType.FontType == FontTypeEnum.EastAsian) &&
                            ((frontCharType.FontType != FontTypeEnum.EastAsian) && (frontCharType.FontType != FontTypeEnum.ComplexScript))))
                        {
                            if ((autoSpaceDN && (char.IsNumber(frontChar) || char.IsNumber(rearChar)))
                                ||
                                (autoSpaceDE && (char.IsLetter(frontChar) || char.IsLetter(rearChar))))
                            {
                                //iTSText.Chunk space = new iTSText.Chunk('\u00A0');
                                iTSText.Chunk space = new iTSText.Chunk(' ');
                                // Due to we multiply 0.875 to font size, the other font styles 
                                // (e.g. underline) are scaled as well, so do not duplicate 
                                // font style settings
                                //space.Font = new iTSText.Font(pg.Chunks[i].Font); 
                                //space.Font.SetStyle(pg.Chunks[i].Font.Style);
                                space.Font.Size = (float)(ich.Font.CalculatedSize * 0.875);
                                pg.Insert(j, space);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Handle paragraph indentation and numbering/listing.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pg"></param>
        private void setParagraphIndentation(Word.Paragraph paragraph, iTSText.Paragraph pg)
        {
            if (paragraph == null || pg == null)
                return;
            
            Word.Level level = this.stHelper.GetNumbering(paragraph);

            // Generate numbering/listing text
            // https://msdn.microsoft.com/en-us/library/office/ee922775%28v=office.14%29.aspx
            iTSText.Chunk numbering = null;
            if (level != null)
            {
                String text = String.Empty;
                Word.LevelText lvlText = level.Descendants<Word.LevelText>().FirstOrDefault();
                if (lvlText != null & lvlText.Val != null)
                    text = lvlText.Val.Value;

                Word.NumberingFormat numFmt = level.Descendants<Word.NumberingFormat>().FirstOrDefault();
                if (numFmt != null && numFmt.Val != null)
                {
                    if (numFmt.Val.Value != Word.NumberFormatValues.Bullet)
                    {
                        List<int> current = null;
                        if (level.Parent != null && level.Parent.GetType() == typeof(Word.AbstractNum))
                        {
                            Word.AbstractNum an = (Word.AbstractNum)level.Parent;
                            if (an.AbstractNumberId != null && level.LevelIndex != null)
                                current = this.stHelper.GetNumberingCurrent(an.AbstractNumberId.Value, level.LevelIndex.Value);
                        }

                        for (int i = 0; i < current.Count; i++)
                        {
                            String replacePattern = String.Format(@"%{0}", i + 1);
                            String str = null;
                            switch (numFmt.Val.Value)
                            {
                                case Word.NumberFormatValues.TaiwaneseCountingThousand:
                                    str = Tools.IntToTaiwanese(current[i]);
                                    break;
                                case Word.NumberFormatValues.LowerRoman:
                                    str = Tools.IntToRoman(current[i], false);
                                    break;
                                case Word.NumberFormatValues.UpperRoman:
                                    str = Tools.IntToRoman(current[i], true);
                                    break;
                                case Word.NumberFormatValues.DecimalZero:
                                    str = String.Format("0{0}", current[i]);
                                    break;
                                case Word.NumberFormatValues.Decimal:
                                default:
                                    str = current[i].ToString();
                                    break;
                            }
                            text = text.Replace(replacePattern, str);
                        }
                    }
                }

                // Get bullet font's RunFonts and size
                FontTypeInfo fontType = CodePointRecognizer.GetFontType(text[0]);
                iTSPdf.BaseFont baseFont = null;
                Word.RunFonts bulletRunFonts = null;
                Word.FontSizeComplexScript fscs = null;
                Word.FontSize fs = null;
                //  get from numbering rPr
                bulletRunFonts = (level.NumberingSymbolRunProperties != null && level.NumberingSymbolRunProperties.RunFonts != null) ? 
                    level.NumberingSymbolRunProperties.RunFonts : null;
                if (bulletRunFonts != null)
                {
                    baseFont = getBaseFontByFontName(
                        getFontNameFromRunFontsByFontType(bulletRunFonts, fontType), 
                        fontType, 
                        this.stHelper.GetAppliedElement<Word.Languages>(level.NumberingSymbolRunProperties));
                    if (fontType.FontType == FontTypeEnum.ComplexScript)
                        fscs = StyleHelper.GetDescendants<Word.FontSizeComplexScript>(level);
                    else
                        fs = StyleHelper.GetDescendants<Word.FontSize>(level);
                }
                //  get from paragraph rPr (i.e. RunFonts for paragraph glyph)
                if (baseFont == null && paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphMarkRunProperties != null)
                {
                    bulletRunFonts = paragraph.ParagraphProperties.ParagraphMarkRunProperties.Descendants<Word.RunFonts>().FirstOrDefault();
                    if (bulletRunFonts != null)
                    {
                        baseFont = getBaseFontByFontName(
                            getFontNameFromRunFontsByFontType(bulletRunFonts, fontType),
                            fontType,
                            this.stHelper.GetAppliedElement<Word.Languages>(paragraph.ParagraphProperties.ParagraphMarkRunProperties));
                        if (fontType.FontType == FontTypeEnum.ComplexScript)
                            fscs = StyleHelper.GetDescendants<Word.FontSizeComplexScript>(paragraph.ParagraphProperties.ParagraphMarkRunProperties);
                        else
                            fs = StyleHelper.GetDescendants<Word.FontSize>(paragraph.ParagraphProperties.ParagraphMarkRunProperties);
                    }
                }
                //  get from docDefault rPr
                if (baseFont == null)
                {
                    bulletRunFonts = this.stHelper.GetDocDefaults<Word.RunFonts>(StyleHelper.DocDefaultsType.Character);
                    if (bulletRunFonts != null)
                    {
                        baseFont = getBaseFontByFontName(
                            getFontNameFromRunFontsByFontType(bulletRunFonts, fontType),
                            fontType,
                            this.stHelper.GetDocDefaults<Word.Languages>(StyleHelper.DocDefaultsType.Character));
                        if (fontType.FontType == FontTypeEnum.ComplexScript)
                            fscs = this.stHelper.GetDocDefaults<Word.FontSizeComplexScript>(StyleHelper.DocDefaultsType.Character);
                        else
                            fs = this.stHelper.GetDocDefaults<Word.FontSize>(StyleHelper.DocDefaultsType.Character);
                    }
                }

                iTSText.Font font = null;
                if (baseFont != null)
                {
                    float fontSize = 12f;
                    if (fscs != null && fscs.Val != null)
                        fontSize = Tools.ConvertToPoint(fscs.Val.Value, Tools.SizeEnum.HalfPoint, -1f);
                    else if (fs != null && fs.Val != null)
                        fontSize = Tools.ConvertToPoint(fs.Val.Value, Tools.SizeEnum.HalfPoint, -1f);
                    font = FontFactory.CreateFont(baseFont, fontSize,
                        bulletRunFonts.Descendants<Word.Bold>().FirstOrDefault(),
                        bulletRunFonts.Descendants<Word.Italic>().FirstOrDefault(),
                        bulletRunFonts.Descendants<Word.Strike>().FirstOrDefault(),
                        bulletRunFonts.Descendants<Word.Color>().FirstOrDefault());
                }

                numbering = new iTSText.Chunk(text, font);
            }

            Word.Indentation ind = null;
            // Numbering indentation first
            if (level != null &&
                level.PreviousParagraphProperties != null &&
                level.PreviousParagraphProperties.Indentation != null)
                ind = (Word.Indentation)level.PreviousParagraphProperties.Indentation.CloneNode(true);
            // Paragraph indentation next, can override numbering indentation
            Word.Indentation pgind = this.stHelper.GetAppliedElement<Word.Indentation>(paragraph);
            if (pgind != null)
            {
                if (ind == null)
                    ind = (Word.Indentation)pgind.CloneNode(true);
                else
                {
                    foreach (OpenXmlAttribute attr in pgind.GetAttributes())
                        if (attr.Value != null) ind.SetAttribute(attr);
                }
            }
            if (ind != null)
            {
                // Character Unit of hanging and firstLine indentation are
                // based on the font size of the first character of paragraph
                // source: https://social.msdn.microsoft.com/Forums/office/en-US/3cfbd59e-453d-4d7e-9bc8-ecb417dbe4a7/how-many-twips-is-a-character-unit?forum=oxmlsdk

                // hanging, use the first character's font size for HaningChars
                float hanging = -1f;
                if (ind.Hanging != null && ind.Hanging.HasValue)
                    hanging = Tools.ConvertToPoint(ind.Hanging.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                else if (ind.HangingChars != null && ind.HangingChars.HasValue)
                    hanging = Tools.ConvertToPoint(ind.HangingChars.Value, Tools.SizeEnum.HundredthsOfCharacter, ((iTSText.Chunk)(pg.Chunks[0])).Font.CalculatedSize);

                // firstLine (only available when no hanging), use the first character's font size for FirstLineChars
                float firstline = -1f;
                if (hanging == -1f)
                {
                    if (ind.FirstLine != null && ind.FirstLine.HasValue)
                        firstline = Tools.ConvertToPoint(ind.FirstLine.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    else if (ind.FirstLineChars != null && ind.FirstLineChars.HasValue)
                        firstline = Tools.ConvertToPoint(ind.FirstLineChars.Value, Tools.SizeEnum.HundredthsOfCharacter, ((iTSText.Chunk)(pg.Chunks[0])).Font.CalculatedSize);
                }

                // Character Unit of start and end are based on the font size
                // of paragraph style hierachy
                float fs = 12f;
                Word.FontSize fontSize = (Word.FontSize)this.stHelper.GetAppliedElement<Word.FontSize>(paragraph);
                if (fontSize != null) { if (fontSize.Val != null) fs = Tools.ConvertToPoint(fontSize.Val.Value, Tools.SizeEnum.HalfPoint, -1f); }

                float dist = 0f;

                // start
                dist = 0f;
                if (ind.Left != null && ind.Left.HasValue)
                    dist = Tools.ConvertToPoint(ind.Left.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                else if (ind.Start != null && ind.Start.HasValue)
                    dist = Tools.ConvertToPoint(ind.Start.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                else if (ind.LeftChars != null && ind.LeftChars.HasValue)
                    dist = Tools.ConvertToPoint(ind.LeftChars.Value, Tools.SizeEnum.HundredthsOfCharacter, fs);
                else if (ind.StartCharacters != null && ind.StartCharacters.HasValue)
                    dist = Tools.ConvertToPoint(ind.StartCharacters.Value, Tools.SizeEnum.HundredthsOfCharacter, fs);
                if (hanging >= 0f)
                { // first line indentation is based on IndentationLeft to add/reduce
                    pg.IndentationLeft = dist;
                    pg.FirstLineIndent = -hanging;
                }
                else
                { // first line indentation is based on IndentationLeft to add/reduce
                    pg.IndentationLeft = dist;
                    if (firstline >= 0f)
                        pg.FirstLineIndent = firstline;
                }

                // end
                dist = 0f;
                if (ind.Right != null && ind.Right.HasValue)
                    dist = Tools.ConvertToPoint(ind.Right.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                else if (ind.End != null && ind.End.HasValue)
                    dist = Tools.ConvertToPoint(ind.End.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                else if (ind.RightChars != null && ind.RightChars.HasValue)
                    dist = Tools.ConvertToPoint(ind.RightChars.Value, Tools.SizeEnum.HundredthsOfCharacter, fs);
                else if (ind.EndCharacters != null && ind.EndCharacters.HasValue)
                    dist = Tools.ConvertToPoint(ind.EndCharacters.Value, Tools.SizeEnum.HundredthsOfCharacter, fs);
                if (dist != 0f)
                    pg.IndentationRight = dist;
            }

            if (numbering != null)
            {
                float addNumberingSpaceWidth = pg.IndentationLeft - (pg.IndentationLeft + pg.FirstLineIndent + numbering.GetWidthPoint());
                if (addNumberingSpaceWidth > 0f)
                {
                    iTSText.Chunk space = new iTSText.Chunk(" ", numbering.Font);
                    space.SetHorizontalScaling(addNumberingSpaceWidth / space.GetWidthPoint());
                    pg.Insert(0, space);

                    // iTextSharp 5.5.5 & 5.5.6 bug:
                    //   when paragraph in tablecell, using SetHorizontalScaling may 
                    //   results in the first line of paragraph out of the tablecell. It 
                    //   mostly happens when the first line is almost fill the tablecell
                    //   width even without adding numbering (latter description is TBC)
                }
                pg.Insert(0, numbering);
            }
        }

        public iTSText.IElement BuildParagraph(Word.Paragraph paragraph)
        {
            if (paragraph == null)
                return null;

            iTSText.Paragraph pg = new iTSText.Paragraph();

            // TODO: handle tabs

            foreach (OpenXmlElement element in paragraph.Elements())
            {
                if (element.GetType() == typeof(Word.Run))
                {
                    pg.AddAll(this.BuildRun((Word.Run)element));
                }
                else if (element.GetType() == typeof(Word.Hyperlink))
                {
                    iTSText.Anchor anchor = this.BuildHyperlink((Word.Hyperlink)element);
                    if (anchor != null) pg.Add(anchor);
                }
                // TODO: handle tab
                //else if (element.GetType() == typeof(Word.TabChar))
                //    ;
            }

            // Horizontal Justification
            Word.Justification jc = this.stHelper.GetAppliedElement<Word.Justification>(paragraph);
            if (jc != null && jc.Val != null)
            {
                switch (jc.Val.Value)
                {
                    case Word.JustificationValues.Center:
                        pg.Alignment = iTSText.Element.ALIGN_CENTER;
                        break;
                    case Word.JustificationValues.Left:
                        pg.Alignment = iTSText.Element.ALIGN_LEFT;
                        break;
                    case Word.JustificationValues.Right:
                        pg.Alignment = iTSText.Element.ALIGN_RIGHT;
                        break;
                    case Word.JustificationValues.Both: // justify text between both margins equally, but inter-character spacing is not affected.
                    case Word.JustificationValues.Distribute: // justify text between both margins equally, and both inter-word and inter-character spacing are affected. iTextSharp doesnt support this.
                        pg.Alignment = iTSText.Element.ALIGN_JUSTIFIED;
                        break;
                    default:
                        break;
                }
            }

            // w:pPr w:keepLines: all lines of this paragraph are maintained on a single page whenever possible
            pg.KeepTogether = StyleHelper.GetToggleProperty(this.stHelper.GetAppliedElement<Word.KeepLines>(paragraph));

            // wpPr w:widowControl: prevent a single line of this paragraph from being displayed on a separate page from the remaining content at display time by moving the line onto the following page
            // TODO: BELOW IMPLEMENTATION WAS WRONG
            //Word.WidowControl widow = this.getAppliedElement<Word.WidowControl>(paragraph);
            //pg.KeepTogether = this.getToggleProperty(widow);

            // Handle empty paragraph
            if (pg.Chunks.Count == 0)
            {
                if (!StyleHelper.GetToggleProperty(this.stHelper.GetAppliedElement<Word.Vanish>(paragraph)))
                {
                    //iTSText.Chunk newline = new iTSText.Chunk("\n");
                    //iTSText.Chunk newline = new iTSText.Chunk(iTSText.Chunk.NEWLINE);
                    iTSText.Chunk emptyPg = new iTSText.Chunk(" ");
                    Word.FontSize fontSize = this.stHelper.GetAppliedElement<Word.FontSize>(paragraph);
                    if (fontSize != null && fontSize.Val != null)
                        emptyPg.Font.Size = Tools.ConvertToPoint(fontSize.Val.Value, Tools.SizeEnum.HalfPoint, -1f);
                    pg.Add(emptyPg);
                }
                else
                    return null;
            }

            // Handle autoSpaceDE and autoSpaceDN
            //  must be called after all chunks are ready, it's because the autospaces
            //  generated on the point of two adjacent texts with different codepoints
            this.setParagraphAutoSpace(paragraph, pg);

            // Line indentation (including numbering process):
            //  must be called after all chunks are ready and autospace, it's because 
            //  1. the hanging and first line indentation are based on the font size of 
            //     the first character of paragraph
            //  2. to avoid numbering text be processed by autospace
            this.setParagraphIndentation(paragraph, pg);

            // Line spacing
            // *** Trick for iTextSharp to calculate leading ***
            // Leading is calculated by Paragraph.Font.Size, but we added chunks to 
            // Pharses then add to paragraph, that results in Paragraph has no Font and
            // iTextSharp use default size (12pt) for leading calculation
            foreach (iTSText.Chunk chunk in pg.Chunks)
                if (chunk.Font != null && chunk.Font.CalculatedSize > pg.Font.Size)
                    pg.Font.Size = chunk.Font.CalculatedSize;
            // The first parameter is the fixed leading: if you want a leading of 15 no matter which font size is used, you can choose fixed = 15 and multiplied = 0.
            // The second parameter is a factor: for instance if you want the leading to be twice the font size, you can choose fixed = 0 and multiplied = 2. In this case, the leading for a paragraph with font size 12 will be 24, for a font size 10, it will be 20, and son on.
            pg.SetLeading(0f, 1.5f); // magic: Word default leading
            this.setParagraphSpacing(paragraph, pg);
            // TODO: magic: remove first line leading
            //pg.SpacingBefore -= (float)(pg.Font.CalculatedSize * 0.1);

            return pg;
        }

        private iTSText.IElement BuildText(Word.Text text)
        {
            if (text == null)
                return null;

            String str = null;
            if (text.Space != null)
            { // handle preserved spaces
                if (text.Space.Value == SpaceProcessingModeValues.Preserve)
                    //str = text.InnerText.Replace(' ', '\u00A0');
                    str = text.InnerText;
                else
                    //str = text.InnerText.Trim().Replace(' ', '\u00A0');
                    str = text.InnerText.Trim();
            }
            else
                //str = text.InnerText.Trim().Replace(' ', '\u00A0');
                str = text.InnerText.Trim();

            return (str != null && str.Length > 0) ? new iTSText.Chunk(str) : null;
        }

        private iTSText.Anchor BuildHyperlink(Word.Hyperlink hyperlink)
        {
            iTSText.Anchor anchor = null;

            if (hyperlink == null)
                return anchor;

            iTSText.Phrase ph = new iTSText.Phrase();
            foreach (OpenXmlElement element in hyperlink.Elements<Word.Run>())
                ph.AddRange(this.BuildRun(element as Word.Run));

            if (ph.Count > 0)
            {
                anchor = new iTSText.Anchor(ph);
                if (hyperlink.Id != null)
                    anchor.Reference = String.Format("{0}{1}",
                        this.stHelper.GetHyperlinkById(hyperlink.Id.Value),
                        (hyperlink.Anchor != null && hyperlink.Anchor.HasValue) ? "#" + hyperlink.Anchor.Value : "");
            }
            
            return anchor;
        }

        private Vml.Shapetype findShapeTypeById(IEnumerable<Vml.Shapetype> shapeTypes, string id)
        {
            return (shapeTypes != null) ? 
                shapeTypes.FirstOrDefault(c => {return (c.Id != null) ? (id.IndexOf(c.Id.Value, StringComparison.OrdinalIgnoreCase) >= 0) : false; }) : null;
        }

        private iTSText.Image BuildPicture(Word.Picture pict)
        {
            iTSText.Image ret = null;

            if (pict == null)
                return ret;

            float maxWidth = 0f, maxHeight = 0f;
            String xmlStr = "";
            IEnumerable<Vml.Shapetype> shapeTypes = pict.Descendants<Vml.Shapetype>();
            foreach (Vml.Shape shape in pict.Descendants<Vml.Shape>())
            {
                OpenXmlCompositeElement tmp = (Vml.Shape)shape.CloneNode(true);

                Vml.Shapetype shapeType = (shape.Type != null) ? 
                    findShapeTypeById(shapeTypes, shape.Type.Value) : null;
                if (shapeType != null)
                {
                    // copy elements
                    foreach (OpenXmlElement e in shapeType.Elements())
                        tmp.Append(e.CloneNode(true));

                    // copy inexistent attributes
                    foreach (OpenXmlAttribute a in shapeType.GetAttributes())
                        if (tmp.GetAttributes().FirstOrDefault(c => c.LocalName.ToLower() == a.LocalName.ToLower()).LocalName == null)
                            // default value of OpenXmlAttribute is an object but with all the variables as null
                            tmp.SetAttribute(a);
                }

                if (shape.Style != null)
                {
                    float x = 0f, y = 0f, width = 0f, height = 0f;

                    Match match = Regex.Match(shape.Style.Value.ToLower(), @"margin-left:([A-Za-z0-9.]+)");
                    if (match.Groups.Count >= 2)
                        x = Tools.ConvertToPoint(match.Groups[1].Value, Tools.SizeEnum.String, -1f);

                    match = Regex.Match(shape.Style.Value.ToLower(), @"margin-top:([A-Za-z0-9.]+)");
                    if (match.Groups.Count >= 2)
                        y = Tools.ConvertToPoint(match.Groups[1].Value, Tools.SizeEnum.String, -1f);

                    match = Regex.Match(shape.Style.Value.ToLower(), @"width:([A-Za-z0-9.]+)");
                    if (match.Groups.Count >= 2)
                        width = Tools.ConvertToPoint(match.Groups[1].Value, Tools.SizeEnum.String, -1f);

                    match = Regex.Match(shape.Style.Value.ToLower(), @"height:([A-Za-z0-9.]+)");
                    if (match.Groups.Count >= 2)
                        height = Tools.ConvertToPoint(match.Groups[1].Value, Tools.SizeEnum.String, -1f);

                    if ((width + x) > maxWidth)
                        maxWidth = (width + x);
                    if ((height + y) > maxHeight)
                        maxHeight = (height + y);
                }

                OpenXmlAttribute j;

                //// Change shape element to rect
                //j = tmp.GetAttributes().FirstOrDefault(c => c.LocalName.ToLower() == "oned");
                //if (j != null && j.LocalName != null)
                //{
                //    if (j.Value == "t")
                //        tmp.
                //}

                // Fix color expression issue
                //   Sometime Word set color as "#5b9bd5 [3204]", which makes SVG can't
                //   display the correct color
                foreach (OpenXmlAttribute t in tmp.GetAttributes())
                {
                    if (t.LocalName.ToLower().IndexOf("color", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Match match = Regex.Match(t.Value, @"(#[A-Za-z0-9]{6})"); // #00FF00
                        if (match.Groups.Count >= 2)
                            tmp.SetAttribute(new OpenXmlAttribute(t.LocalName, t.NamespaceUri, match.Groups[1].Value));
                    }
                }

                // Fix path issue
                //   Sometime Word omit zeros for position (e.g. "L0,1" => "L,1"), this 
                //   makes SVG render failes to draw the path
                j = tmp.GetAttributes().FirstOrDefault(c => c.LocalName.ToLower() == "path");
                if (j != null && j.LocalName != null)
                {
                    String path = j.Value;
                    for (int index = path.IndexOf(','); index >= 0; index = path.IndexOf(',', index + 1))
                    {
                        bool isNum = false;
                        if (index > 0)
                            isNum = Char.IsNumber(path[index - 1]);
                        if (!isNum)
                        {
                            path = path.Insert(index, "0");
                            index += 1;
                        }

                        isNum = false;
                        if (index < (path.Length - 1))
                            isNum = Char.IsNumber(path[index + 1]);
                        if (!isNum)
                        {
                            path = path.Insert(index + 1, "0");
                            index += 1;
                        }
                    }
                    if (path != j.Value)
                        tmp.SetAttribute(new OpenXmlAttribute(j.LocalName, j.NamespaceUri, path));
                }

                xmlStr += tmp.OuterXml;
            }

            //Image img = Tools.ConvertVmlToImage(xmlStr, maxWidth, maxHeight);
            //if (img != null)
            //{
            //    ret = iTSText.Image.GetInstance(img, System.Drawing.Imaging.ImageFormat.Png);
            //    ret.ScaleAbsolute(maxWidth, maxHeight);
            //}
            return ret;
        }

        private List<iTSText.Chunk> BuildRun(Word.Run run)
        {
            List<iTSText.Chunk> ret = new List<iTSText.Chunk>();

            if (run == null)
                return ret;

            // Handle RUN properties
            // http://officeopenxml.com/WPtextFormatting.php

            // vanish
            if (StyleHelper.GetToggleProperty(this.stHelper.GetAppliedElement<Word.Vanish>(run)))
                return ret;

            // Toggle property
            Word.Bold bold = (Word.Bold)this.stHelper.GetAppliedElement<Word.Bold>(run);
            Word.Italic italic = (Word.Italic)this.stHelper.GetAppliedElement<Word.Italic>(run);
            Word.Strike strike = (Word.Strike)this.stHelper.GetAppliedElement<Word.Strike>(run);
            Word.Caps caps = (Word.Caps)this.stHelper.GetAppliedElement<Word.Caps>(run);

            Word.Underline uline = (Word.Underline)this.stHelper.GetAppliedElement<Word.Underline>(run);

            Word.FontSize size = (Word.FontSize)this.stHelper.GetAppliedElement<Word.FontSize>(run);
            float fontSize = -1f;
            if (size != null && size.Val != null) fontSize = Tools.ConvertToPoint(size.Val.Value, Tools.SizeEnum.HalfPoint, -1f); // OpenXml font size unit is half-point, PDF is point

            Word.FontSizeComplexScript csSize = (Word.FontSizeComplexScript)this.stHelper.GetAppliedElement<Word.FontSizeComplexScript>(run);
            float csFontSize = -1f;
            if (csSize != null && csSize.Val != null) csFontSize = Tools.ConvertToPoint(csSize.Val.Value, Tools.SizeEnum.HalfPoint, -1f); // OpenXml font size unit is half-point, PDF is point

            Word.Color color = (Word.Color)this.stHelper.GetAppliedElement<Word.Color>(run);
            // TODO: theme color

            Word.Kern kern = (Word.Kern)this.stHelper.GetAppliedElement<Word.Kern>(run);
            float kernSize = -1f;
            if (kern != null && kern.Val != null) kernSize = Tools.ConvertToPoint(kern.Val.Value, Tools.SizeEnum.HalfPoint, -1f);

            // Handle OpenXML child elements
            foreach (var element in run.Elements())
            {
                Type elementType = element.GetType();
                if (elementType == typeof(Word.Text))
                {
                    Word.Text t = element as Word.Text;
                    iTSText.Chunk ch = BuildText(t) as iTSText.Chunk;
                    if (ch != null)
                    {
                        if (StyleHelper.GetToggleProperty(caps))
                            ch = new iTSText.Chunk(ch.Content.ToUpper());

                        iTSPdf.BaseFont baseFont = this.getBaseFont(run, ch.Content);
                        if (baseFont != null)
                        {
                            FontTypeInfo fontType = CodePointRecognizer.GetFontType(ch.Content.ToCharArray()[0]);
                            ch.Font = FontFactory.CreateFont(baseFont,
                                ((csFontSize > 0) && (fontType.FontType == FontTypeEnum.ComplexScript)) ? csFontSize : fontSize,
                                bold, italic, strike, color);
                        }

                        // TODO: kerning
                        //if (kernSize > 0f && ch.Font.CalculatedSize > kernSize)
                        //    ch.SetCharacterSpacing(0f);
                        //ch.SetWordSpacing(0f);

                        // special handle for preserved spaces
                        if (t.Space != null && t.Space.Value == SpaceProcessingModeValues.Preserve)
                        {
                            // Word treats Space width as half width of 'n'. If specify
                            // BalanceSingleByteDoubleByteWidth in compatibility settings
                            // (settings.xml), extends Space width as the same width as 
                            // 'n'.
                            if (this.stHelper.BalanceSingleByteDoubleByteWidth)
                            {
                                if (ch.Content.Trim().Length == 0) // all spaces
                                {
                                    float space = baseFont.GetWidthPoint(' ', ch.Font.CalculatedSize);
                                    float n = baseFont.GetWidthPoint('n', ch.Font.CalculatedSize);
                                    ch.SetHorizontalScaling(n / space);
                                }
                            }
                        }

                        if (uline != null && uline.Val != null)
                        {
                            switch (uline.Val.Value)
                            {
                                case Word.UnderlineValues.None:
                                    break;
                                default:
                                    ch.SetUnderline(0.07f * ch.Font.CalculatedSize, -0.2f * ch.Font.CalculatedSize);
                                    //ch.Font.SetStyle(iTSText.Font.UNDERLINE);
                                    //ch.SetUnderline(iTSText.BaseColor.RED, 0.1f, 0f, 0f, 0f, iTSPdf.PdfContentByte.LINE_CAP_ROUND);
                                    //ch.SetUnderline()
                                    // TODO: handle underline: support color and dashed/double/...etc., http://stackoverflow.com/questions/29260730/how-do-you-underline-text-with-dashedline-in-itext-pdf
                                    break;
                            }
                        }

                        ret.Add(ch);
                    }
                }
                else if (elementType == typeof(Word.Break))
                {
                    Word.Break br = element as Word.Break;
                    if (br.Type == null)
                        ret.Add(iTSText.Chunk.NEWLINE);
                    else
                    {
                        switch (br.Type.Value)
                        {
                            case Word.BreakValues.TextWrapping:
                                ret.Add(iTSText.Chunk.NEWLINE);
                                break;
                            case Word.BreakValues.Page:
                                ret.Add(iTSText.Chunk.NEXTPAGE);
                                break;
                            case Word.BreakValues.Column: // TODO: handle br:column
                            default:
                                ret.Add(iTSText.Chunk.NEWLINE);
                                break;
                        }
                    }
                }
                else if (elementType == typeof(Word.SymbolChar))
                {
                    Word.SymbolChar sym = element as Word.SymbolChar;
                    char c = (char)(Int32.Parse(sym.Char.Value, System.Globalization.NumberStyles.HexNumber) - 0xF000);
                    FontTypeInfo fontType = CodePointRecognizer.GetFontType(c);
                    iTSPdf.BaseFont baseFont = this.getBaseFontByFontName(sym.Font.Value, fontType, null);
                    if (baseFont != null)
                    {
                        iTSText.Chunk ch = new iTSText.Chunk(c.ToString());
                        ch.Font = FontFactory.CreateFont(baseFont,
                            ((csFontSize > 0) && (fontType.FontType == FontTypeEnum.ComplexScript)) ? csFontSize : fontSize,
                            bold, italic, strike, color);
                        ret.Add(ch);
                    }
                }
                //else if (elementType == typeof(Word.Picture))
                //{
                //    iTSText.Image image = BuildPicture((Word.Picture)element);
                //    if (image != null)
                //        this.pdfDoc.Add(image);
                //}
                //else if (elementType == typeof(AlternateContent))
                //{
                //    AlternateContent alternateContent = element as AlternateContent;
                //    AlternateContentFallback fallback = alternateContent.Descendants<AlternateContentFallback>().FirstOrDefault();
                //    if (fallback != null)
                //    {
                //        Word.Picture pict = fallback.Descendants<Word.Picture>().FirstOrDefault();
                //        if (pict != null)
                //        {
                //            iTSText.Image image = BuildPicture((Word.Picture)pict);
                //            if (image != null)
                //                this.pdfDoc.Add(image);
                //        }
                //    }
                //}
                // don't need to handle w:lastRenderedPageBreak
                //else if (element.GetType() == typeof(Word.LastRenderedPageBreak))
                //{
                //    ret.Add(iTSText.Chunk.NEXTPAGE);
                //}
            }

            return ret;
        }

        /// <summary>
        /// Convert DocumentFormat.OpenXml.Wordprocessing.Table to iTextSharp.text.pdf.PdfPTable. 
        /// </summary>
        /// <param name="table">DocumentFormat.OpenXml.Wordprocessing.Table.</param>
        /// <returns>iTextSharp.text.pdf.PdfPTable or null.</returns>
        public iTSText.IElement BuildTable(Word.Table table)
        {
            if (table == null)
                return null;

            TableHelper tblHelper = new TableHelper();
            if (this.stHelper != null)
                tblHelper.StHelper = this.stHelper;
            tblHelper.CellChildElementsProc += this.Dispatcher;
            tblHelper.ParseTable(table);

            // ====== Prepare iTextSharp PdfPTable ======

            // Set table width
            iTSPdf.PdfPTable pt = new iTSPdf.PdfPTable(tblHelper.ColumnLength);
            pt.TotalWidth = tblHelper.ColumnWidth.Sum();
            pt.SetWidths(tblHelper.ColumnWidth);
            pt.LockedWidth = true; // use pt.TotalWidth rather than pt.WidthPercentage (iTextSharp default is WidthPercentage)

            // Table justification
            Word.TableJustification jc = this.stHelper.GetAppliedElement<Word.TableJustification>(table);
            pt.HorizontalAlignment = (jc != null && jc.Val != null) ? (
                (jc.Val.Value == Word.TableRowAlignmentValues.Center) ? iTSText.Element.ALIGN_CENTER :
                (jc.Val.Value == Word.TableRowAlignmentValues.Left) ? iTSText.Element.ALIGN_LEFT :
                (jc.Val.Value == Word.TableRowAlignmentValues.Right) ? iTSText.Element.ALIGN_RIGHT : iTSText.Element.ALIGN_LEFT
                ) : iTSText.Element.ALIGN_LEFT;

            foreach (TableHelperCell pcell in tblHelper)
            {
                // Row height
                float minRowHeight = -1f, exactRowHeight = -1f;
                Word.TableRowHeight trHeight = this.stHelper.GetAppliedElement<Word.TableRowHeight>(pcell.row);
                if (trHeight != null && trHeight.Val != null)
                {
                    if (trHeight.Val != null)
                        minRowHeight = Tools.ConvertToPoint(trHeight.Val.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    try
                    {
                        OpenXmlAttribute hrule = trHeight.GetAttribute("hRule", this.stHelper.NamespaceUri);
                        if (hrule.Value != null)
                        {
                            if (hrule.Value == "exact")
                                exactRowHeight = Tools.ConvertToPoint(trHeight.Val.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        }
                    }
                    catch (Exception) { }
                }

                iTSPdf.PdfPCell cell = new iTSPdf.PdfPCell(); // composite mode, not text mode
                if (exactRowHeight > 0f) cell.FixedHeight = exactRowHeight;
                else if (minRowHeight > 0f) cell.MinimumHeight = minRowHeight;
                //cell.UseAscender = false; // remove whitespace on top of each cell even padding&leading set to 0, http://stackoverflow.com/questions/9672046/itextsharp-4-1-6-pdf-table-how-to-remove-whitespace-on-top-of-each-cell-pad
                //cell.UseDescender = true;
                cell.Rowspan = pcell.RowSpan;
                cell.Colspan = pcell.ColSpan;
                if (pcell.Blank)
                {
                    cell.Border = iTSPdf.PdfPCell.NO_BORDER;
                }
                else if (pcell.cell != null)
                {
                    Word.TableCell c = pcell.cell;

                    // Cell margins
                    Word.LeftMargin left = this.stHelper.GetAppliedElement<Word.LeftMargin>(c);
                    if (left != null && left.Width != null)
                        cell.PaddingLeft = Tools.ConvertToPoint(left.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    else
                    {
                        Word.TableCellLeftMargin tblPrLeft = this.stHelper.GetAppliedElement<Word.TableCellLeftMargin>(c);
                        if (tblPrLeft != null && tblPrLeft.Width != null)
                            cell.PaddingLeft = Tools.ConvertToPoint(tblPrLeft.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    }
                    Word.RightMargin right = this.stHelper.GetAppliedElement<Word.RightMargin>(c);
                    if (right != null && right.Width != null)
                        cell.PaddingRight = Tools.ConvertToPoint(right.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    else
                    {
                        Word.TableCellRightMargin tblPrRight = this.stHelper.GetAppliedElement<Word.TableCellRightMargin>(c);
                        if (tblPrRight != null && tblPrRight.Width != null)
                            cell.PaddingRight = Tools.ConvertToPoint(tblPrRight.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    }
                    Word.TopMargin top = this.stHelper.GetAppliedElement<Word.TopMargin>(c);
                    cell.PaddingTop = (top != null && top.Width != null) ?
                            Tools.ConvertToPoint(top.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f) : 0f;
                    Word.BottomMargin bottom = this.stHelper.GetAppliedElement<Word.BottomMargin>(c);
                    cell.PaddingBottom = (bottom != null && bottom.Width != null) ?
                            Tools.ConvertToPoint(bottom.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f) : 0f;

                    // Vertical alignment
                    Word.TableCellVerticalAlignment va = this.stHelper.GetAppliedElement<Word.TableCellVerticalAlignment>(c);
                    if (va != null && va.Val != null)
                    {
                        if (va.Val.Value == Word.TableVerticalAlignmentValues.Top)
                            cell.VerticalAlignment = iTSText.Element.ALIGN_TOP;
                        else if (va.Val.Value == Word.TableVerticalAlignmentValues.Bottom)
                            cell.VerticalAlignment = iTSText.Element.ALIGN_BOTTOM;
                        else if (va.Val.Value == Word.TableVerticalAlignmentValues.Center)
                            cell.VerticalAlignment = iTSText.Element.ALIGN_MIDDLE;
                    }

                    // Shading
                    Word.Shading sh = this.stHelper.GetAppliedElement<Word.Shading>(c);
                    if (sh != null)
                    {
                        if (sh.Fill != null && sh.Fill.HasValue)
                        {
                            if (sh.Fill.Value != "auto")
                                cell.BackgroundColor = new iTSTextColor(Convert.ToInt32(sh.Fill.Value, 16));
                        }
                    }

                    // Border
                    //  top border
                    Word.TopBorder topbr = pcell.Borders.TopBorder;
                    if (topbr == null || topbr.Val == null ||
                        topbr.Val.Value == Word.BorderValues.Nil || topbr.Val.Value == Word.BorderValues.None)
                        cell.Border &= ~iTSText.Rectangle.TOP_BORDER;
                    else
                    {
                        if (topbr.Color != null && topbr.Color.Value != "auto")
                            cell.BorderColorTop = new iTSTextColor(Convert.ToInt32(topbr.Color.Value, 16));
                        if (topbr.Size != null)
                            cell.BorderWidthTop = Tools.ConvertToPoint(topbr.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                    }
                    //  bottom border
                    Word.BottomBorder bottombr = pcell.Borders.BottomBorder;
                    if (bottombr == null || bottombr.Val == null ||
                        bottombr.Val.Value == Word.BorderValues.Nil || bottombr.Val.Value == Word.BorderValues.None)
                        cell.Border &= ~iTSText.Rectangle.BOTTOM_BORDER;
                    else
                    {
                        if (bottombr.Color != null && bottombr.Color.Value != "auto")
                            cell.BorderColorBottom = new iTSTextColor(Convert.ToInt32(bottombr.Color.Value, 16));
                        if (bottombr.Size != null)
                            cell.BorderWidthBottom = Tools.ConvertToPoint(bottombr.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                    }
                    //  left border
                    Word.LeftBorder leftbr = pcell.Borders.LeftBorder;
                    if (leftbr == null || leftbr.Val == null ||
                        leftbr.Val.Value == Word.BorderValues.Nil || leftbr.Val.Value == Word.BorderValues.None)
                        cell.Border &= ~iTSText.Rectangle.LEFT_BORDER;
                    else
                    {
                        if (leftbr.Color != null && leftbr.Color.Value != "auto")
                            cell.BorderColorLeft = new iTSTextColor(Convert.ToInt32(leftbr.Color.Value, 16));
                        if (leftbr.Size != null)
                            cell.BorderWidthLeft = Tools.ConvertToPoint(leftbr.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                    }
                    //  right border
                    Word.RightBorder rightbr = pcell.Borders.RightBorder;
                    if (rightbr == null || rightbr.Val == null ||
                        rightbr.Val.Value == Word.BorderValues.Nil || rightbr.Val.Value == Word.BorderValues.None)
                        cell.Border &= ~iTSText.Rectangle.RIGHT_BORDER;
                    else
                    {
                        if (rightbr.Color != null && rightbr.Color.Value != "auto")
                            cell.BorderColorRight = new iTSTextColor(Convert.ToInt32(rightbr.Color.Value, 16));
                        if (rightbr.Size != null)
                            cell.BorderWidthRight = Tools.ConvertToPoint(rightbr.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                    }
                    //cell.Border = iTSText.Rectangle.BOX; // TODO: for debug
                }
                // add elements to cell
                bool adjustPadding = false; // magic: simulate Word cell padding
                cell.AddElement(new iTSText.Paragraph(0f, "\u00A0")); // magic: to allow Paragraph.SpacingBefore take effect
                foreach (iTSText.IElement element in pcell.elements)
                {
                    if (element.GetType() == typeof(iTSText.Paragraph))
                    {
                        iTSText.Paragraph pg = element as iTSText.Paragraph;
                        pg.SetLeading(pg.TotalLeading * 0.9f, 0f); // magic: paragraph leading in table becomes 0.9
                        
                        // ------
                        // magic: simulate Word cell padding
                        if (!adjustPadding)
                        {
                            cell.PaddingTop -= (pg.TotalLeading * 0.25f);
                            cell.PaddingBottom += (pg.TotalLeading * 0.25f);
                            adjustPadding = true;
                        }
                        // ------
                    }
                    cell.AddElement(element);
                }
                pt.AddCell(cell);
            }
            
            return pt;
        }

        public List<iTSText.IElement> BuildHeader(Word.SectionProperties sectPr)
        {
            List<iTSText.IElement> ret = new List<iTSText.IElement>();

            if (sectPr == null)
                return ret;

            // Get header rId
            String id = null;
            Word.HeaderReference refHdr = sectPr.Descendants<Word.HeaderReference>().FirstOrDefault();
            if (refHdr != null && refHdr.Id != null)
                id = refHdr.Id.Value;
            else
                return ret;

            // Find header by rId
            HeaderPart hdrpt = this.doc.MainDocumentPart.HeaderParts.FirstOrDefault(c => {
                return this.doc.MainDocumentPart.GetIdOfPart(c) == id;
            });
            if (hdrpt != null && hdrpt.Header != null)
            {
                foreach (OpenXmlElement element in hdrpt.Header.Elements())
                {
                    iTSText.IElement t = this.Dispatcher(element);
                    if (t != null)
                        ret.Add(t);
                }
            }
            return ret;
        }

        /// <summary>
        /// Get the height of a set of IElements outupt, must provide output width for reference.
        /// </summary>
        /// <param name="contents"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        private float calculateHeight(List<iTSText.IElement> contents, float width)
        {
            float diff = 0f;

            if (contents == null || (contents.Count() == 0))
                return diff;

            using (MemoryStream ms = new MemoryStream())
            {
                iTSText.Document doc = new iTSText.Document();
                iTSPdf.PdfWriter writer = iTSPdf.PdfWriter.GetInstance(doc, ms);
                doc.Open();
                iTSPdf.ColumnText ct = new iTSPdf.ColumnText(writer.DirectContent);
                iTSText.Rectangle rect = new iTSText.Rectangle(0f, 0f, width, 1000f);
                ct.SetSimpleColumn(rect);
                foreach (iTSText.IElement t in contents)
                    ct.AddElement(t);
                float beforeY = ct.YLine;
                ct.Go(); // do not simulate because no page makes doc.Close() generates exception
                diff = beforeY - ct.YLine;
                doc.Close();
            }
            return diff;
        }

        /// <summary>
        /// For drawing header and footer.
        /// </summary>
        private class pdfPage : iTSPdf.PdfPageEventHelper
        {
            DocxToPdf obj = null;

            public pdfPage(DocxToPdf obj)
            {
                this.obj = obj;
            }

            public override void OnEndPage(iTSPdf.PdfWriter writer, iTSText.Document doc)
            {
                Word.PageMargin margin = StyleHelper.GetElement<Word.PageMargin>(this.obj.stHelper.CurrentSectPr);

                // Draw header
                List<iTSText.IElement> contents = this.obj.BuildHeader(this.obj.stHelper.CurrentSectPr);
                if (contents.Count() > 0)
                {
                    float hdrMargin = (margin != null && margin.Header != null) ?
                        Tools.ConvertToPoint(margin.Header.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f) : 0f;
                    iTSPdf.ColumnText ct = new iTSPdf.ColumnText(writer.DirectContent);
                    // in iTextSharp page coordinate concept, for Y-axis, the top edge of
                    // the page has the maximum value (doc.PageSize.Height) and the 
                    // bottom edge of the page is zero
                    iTSText.Rectangle rect = new iTSText.Rectangle(
                        doc.LeftMargin,
                        doc.PageSize.Height - hdrMargin,
                        doc.PageSize.Width - doc.RightMargin, // not the width, should be the maximum X coordinate
                        doc.BottomMargin); // not the height, should be the smallest Y coordinate
                    ct.SetSimpleColumn(rect);
                    foreach (iTSText.IElement e in contents)
                        ct.AddElement(e);
                    ct.Go();
                }

                // TODO: draw footer
            }
        }
    }
}
