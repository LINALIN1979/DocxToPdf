using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToPdf
{
    class StyleHelper
    {
        #region NumberingCounter
        private class NumberingCounter
        {
            private class LevelCounter
            {
                private readonly int ilvl;
                private int start;
                private int current;

                public LevelCounter(int ilvl, int start)
                {
                    this.ilvl = ilvl;
                    this.Start = start;
                }

                /// <summary>
                /// Get iLevel.
                /// </summary>
                public int iLevel { get { return this.ilvl; } }

                /// <summary>
                /// Reset current numbering to start value.
                /// </summary>
                public void Restart() { this.current = this.start; }

                /// <summary>
                /// Get current numbering and increase by one.
                /// </summary>
                public int Current { get { return ++this.current; } }

                /// <summary>
                /// Get current numbering but without changing it.
                /// </summary>
                public int CurrentStatic { get { return this.current; } }

                /// <summary>
                /// Set start numbering (must be not a negative value).
                /// </summary>
                public int Start { set { if (value > 0) this.start = this.current = value - 1; } }
            }

            /// <summary>
            /// Store abstractNums, the key is abstractNumId.
            /// </summary>
            private Dictionary<int, List<LevelCounter>> abstractNums = new Dictionary<int, List<LevelCounter>>();

            public void SetStart(int abstractNumId, int ilvl, int start)
            {
                if (start < 0) return;

                if (abstractNums.ContainsKey(abstractNumId))
                {
                    LevelCounter lc = abstractNums[abstractNumId].FirstOrDefault(c => c.iLevel == ilvl);
                    if (lc != null)
                        lc.Start = start;
                    else
                        abstractNums[abstractNumId].Add(new LevelCounter(ilvl, start));
                }
                else
                {
                    abstractNums[abstractNumId] = new List<LevelCounter>();
                    abstractNums[abstractNumId].Add(new LevelCounter(ilvl, start));
                }
            }

            /// <summary>
            /// Restart the numbering value of ilvl, and all the levels larger than ilvl will be restarted as well.
            /// </summary>
            /// <param name="abstractNumId"></param>
            /// <param name="ilvl"></param>
            public void Restart(int abstractNumId, int ilvl)
            {
                if (abstractNums.ContainsKey(abstractNumId))
                {
                    while (true)
                    {
                        LevelCounter lc = abstractNums[abstractNumId].FirstOrDefault(c => c.iLevel == ilvl);
                        if (lc != null)
                        {
                            lc.Restart();
                            ilvl++;
                        }
                        else
                            break;
                    }
                }
            }

            /// <summary>
            /// Get a list of numbering value from level-0 to level-ilvl. Call this method will 
            ///   1. increase the numbering value of level-ilvl by one automatically
            ///   2. all the levels larger than level-ilvl will be restarted
            /// </summary>
            /// <param name="abstractNumId"></param>
            /// <param name="ilvl"></param>
            /// <returns></returns>
            public List<int> GetCurrent(int abstractNumId, int ilvl)
            {
                List<int> ret = new List<int>();
                if (abstractNums.ContainsKey(abstractNumId))
                {
                    LevelCounter lc;
                    for (int i = 0; i < ilvl; i++)
                    { // get the numbering value from the levels smaller than ilvl
                        lc = abstractNums[abstractNumId].FirstOrDefault(c => c.iLevel == i);
                        if (lc != null)
                            ret.Add(lc.CurrentStatic);
                    }

                    lc = abstractNums[abstractNumId].FirstOrDefault(c => c.iLevel == ilvl);
                    if (lc != null)
                    { // get the numbering value from ilvl
                        ret.Add(lc.Current);
                        Restart(abstractNumId, ilvl + 1); // all the levels larger than ilvl should restart
                    }
                }
                return ret;
            }
        }
        #endregion

        private WordprocessingDocument doc = null;
        private Word.DocDefaults docDefaults = null;
        private List<Word.SectionProperties> sections = new List<Word.SectionProperties>();
        private int currentSectPr = 0;
        private IEnumerable<Word.Style> styles = null;
        private ThemePart theme = null;
        private Word.Numbering numbering = null;
        private NumberingCounter nc = new NumberingCounter();
        private IEnumerable<HyperlinkRelationship> hyperlinkRelationships = null;

        public String NamespaceUri { get { return (this.docDefaults != null) ? this.docDefaults.NamespaceUri : null; } }

        private bool balanceSingleByteDoubleByteWidth = false;
        /// <summary>
        /// The flag indicates whether to show the width of single byte character as double byte character.
        /// 
        /// https://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.balancesinglebytedoublebytewidth%28v=office.14%29.aspx
        /// </summary>
        public bool BalanceSingleByteDoubleByteWidth { get { return this.balanceSingleByteDoubleByteWidth; } }

        public StyleHelper(WordprocessingDocument doc)
        {
            if (doc != null)
            {
                this.doc = doc;
                this.docDefaults = this.doc.MainDocumentPart.StyleDefinitionsPart.Styles.Descendants<Word.DocDefaults>().FirstOrDefault();
                this.sections = this.doc.MainDocumentPart.Document.Body.Descendants<Word.SectionProperties>().ToList();
                this.styles = this.doc.MainDocumentPart.StyleDefinitionsPart.Styles.Descendants<Word.Style>();
                this.theme = this.doc.MainDocumentPart.ThemePart;
                if (this.doc.MainDocumentPart.NumberingDefinitionsPart != null &&
                    this.doc.MainDocumentPart.NumberingDefinitionsPart.Numbering != null)
                {
                    this.numbering = this.doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;

                    // set all ilevel's start value for all abstractNums
                    foreach (Word.AbstractNum an in this.numbering.Descendants<Word.AbstractNum>())
                    {
                        int abstractNumId = an.AbstractNumberId.Value;
                        foreach (Word.Level lvl in an.Descendants<Word.Level>())
                        {
                            if (lvl.LevelIndex != null && lvl.StartNumberingValue != null)
                                this.nc.SetStart(abstractNumId, lvl.LevelIndex.Value, lvl.StartNumberingValue.Val);
                        }
                    }
                }
                this.hyperlinkRelationships = this.doc.MainDocumentPart.HyperlinkRelationships;
                Word.Compatibility compat = this.doc.MainDocumentPart.DocumentSettingsPart.Settings.Descendants<Word.Compatibility>().FirstOrDefault();
                if (compat != null && compat.BalanceSingleByteDoubleByteWidth != null)
                    this.balanceSingleByteDoubleByteWidth = StyleHelper.GetToggleProperty(compat.BalanceSingleByteDoubleByteWidth);
            }
        }

        /// <summary>
        /// Get printable page width.
        /// </summary>
        public float PrintablePageWidth
        {
            get
            {
                Word.PageSize size = StyleHelper.GetElement<Word.PageSize>(this.CurrentSectPr);
                Word.PageMargin margin = StyleHelper.GetElement<Word.PageMargin>(this.CurrentSectPr);
                return Tools.ConvertToPoint(size.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, 0) -
                    Tools.ConvertToPoint(margin.Left.Value, Tools.SizeEnum.TwentiethsOfPoint, 0) -
                    Tools.ConvertToPoint(margin.Right.Value, Tools.SizeEnum.TwentiethsOfPoint, 0);
            }
        }

        public enum DocDefaultsType
        {
            Paragraph,
            Character,
        }

        /// <summary>
        /// Get object from docDefaults.
        /// </summary>
        /// <typeparam name="T">Target class type.</typeparam>
        /// <returns></returns>
        public T GetDocDefaults<T>(DocDefaultsType type)
        {
            T ret = default(T);

            if (this.docDefaults == null)
                return ret;
            if (type == DocDefaultsType.Character)
                ret = StyleHelper.GetDescendants<T>(this.docDefaults.RunPropertiesDefault);
            else if (type == DocDefaultsType.Paragraph)
                ret = StyleHelper.GetDescendants<T>(this.docDefaults.ParagraphPropertiesDefault);
            return ret;
        }

        /// <summary>
        /// Get current SectionProperties.
        /// 
        /// http://openxmldeveloper.org/discussions/formats/f/13/t/678.aspx
        /// </summary>
        public Word.SectionProperties CurrentSectPr
        {
            get
            {
                return (this.currentSectPr < 0 || this.currentSectPr >= this.sections.Count) ? 
                    null : this.sections[this.currentSectPr];
            }
        }

        /// <summary>
        /// Get current SectionProperties is the end or not.
        /// </summary>
        public bool IsCurrentSectPrTheEnd
        {
            get
            {
                return this.currentSectPr == (this.sections.Count - 1);
            }
        }

        /// <summary>
        /// Get next SectionProperties.
        /// </summary>
        public Word.SectionProperties NextSectPr
        {
            get
            {
                this.currentSectPr++;
                return CurrentSectPr;
            }
        }

        /// <summary>
        /// Get the element index of current SectionProperties in Document.Body.
        /// </summary>
        public int CurrentSectPrIndexInBody
        {
            get
            {
                if (this.currentSectPr < 0 || this.currentSectPr >= this.sections.Count)
                    return -1;
                Word.SectionProperties sectPr = this.sections[this.currentSectPr];
                if (sectPr.Parent.GetType() == typeof(Word.Body))
                { // section.parent = body ==> last section
                    return this.doc.MainDocumentPart.Document.Body.Elements().ToList().IndexOf(sectPr);
                }
                else
                { // otherwise, section is in paragraph.pPr
                    return this.doc.MainDocumentPart.Document.Body.Elements().ToList().IndexOf(sectPr.Parent.Parent);
                }
            }
        }

        /// <summary>
        /// Get the nearest applied element (e.g. RunFonts, Languages) in the style hierachy. It searches upstream from current $obj til reach the top of the hierachy if no found.
        /// </summary>
        /// <typeparam name="T">Target element type (e.g. RunFonts.GetType()).</typeparam>
        /// <param name="obj">The OpenXmlElemet to search from.</param>
        /// <returns>Return found element or null if not found.</returns>
        public T GetAppliedElement<T>(OpenXmlElement obj)
        {
            T ret = default(T);

            if (obj == null)
                return ret;

            Type objType = obj.GetType();

            if (objType == typeof(Word.Run))
            { // Run.RunProperties > Run.RunProperties.rStyle > Paragraph.ParagraphProperties.pStyle (> default style > docDefaults)
                // ( ): done in paragraph level
                Word.RunProperties runpr = StyleHelper.GetElement<Word.RunProperties>(obj);
                if (runpr != null)
                {
                    ret = StyleHelper.GetDescendants<T>(runpr);

                    // If has rStyle, go through rStyle before go on.
                    // Use getAppliedStyleElement() is because it will go over all the basedOn styles.
                    if (ret == null && runpr.RunStyle != null)
                        ret = this.GetAppliedElement<T>(this.GetStyleById(runpr.RunStyle.Val));
                }

                if (ret == null)
                { // parent paragraph's pStyle
                    if (obj.Parent != null && obj.Parent.GetType() == typeof(Word.Paragraph))
                    {
                        Word.Paragraph pg = obj.Parent as Word.Paragraph;
                        if (pg.ParagraphProperties != null && pg.ParagraphProperties.ParagraphStyleId != null)
                            ret = this.GetAppliedElement<T>(this.GetStyleById(pg.ParagraphProperties.ParagraphStyleId.Val));
                    }
                }

                if (ret == null) // default run style
                    ret = this.GetAppliedElement<T>(this.GetDefaultStyle(DefaultStyleType.Character));

                if (ret == null) // docDefaults
                    ret = StyleHelper.GetDescendants<T>(this.docDefaults.RunPropertiesDefault);
            }
            else if (objType == typeof(Word.Paragraph))
            { // Paragraph.ParagraphProperties > Paragraph.ParagraphProperties.pStyle > default style > docDefaults

                Word.ParagraphProperties pgpr = StyleHelper.GetElement<Word.ParagraphProperties>(obj);
                if (pgpr != null)
                {
                    ret = StyleHelper.GetDescendants<T>(pgpr);

                    // If has pStyle, go through pStyle before go on.
                    // Use getAppliedStyleElement() is because it will go over the whole Style hierachy.
                    if (ret == null && pgpr.ParagraphStyleId != null)
                        ret = this.GetAppliedElement<T>(this.GetStyleById(pgpr.ParagraphStyleId.Val));
                }

                if (ret == null)
                {
                    if (obj.Parent != null && obj.Parent.GetType() == typeof(Word.TableCell))
                    {
                        for (int i = 0; i < 3 && obj != null; i++)
                            obj = (obj.Parent != null) ? obj.Parent : null;
                        if (obj != null && obj.GetType() == typeof(Word.Table))
                            ret = this.GetAppliedElement<T>(obj);
                    }
                }

                if (ret == null) // default paragraph style
                    ret = this.GetAppliedElement<T>(this.GetDefaultStyle(DefaultStyleType.Paragraph));

                if (ret == null) // docDefaults
                    ret = StyleHelper.GetDescendants<T>(this.docDefaults);
            }
            else if (objType == typeof(Word.Table))
            { // Table.TableProperties > Table.TableProperties.tblStyle > default style
                Word.TableProperties tblpr = StyleHelper.GetElement<Word.TableProperties>(obj);
                if (tblpr != null)
                {
                    ret = StyleHelper.GetDescendants<T>(tblpr);

                    // If has tblStyle, go through tblStyle before go on.
                    // Use getAppliedStyleElement() is because it will go over the whole Style hierachy.
                    if (ret == null && tblpr.TableStyle != null)
                        ret = this.GetAppliedElement<T>(this.GetStyleById(tblpr.TableStyle.Val));
                }

                if (ret == null) // default table style
                    ret = this.GetAppliedElement<T>(this.GetDefaultStyle(DefaultStyleType.Table));
            }
            else if (objType == typeof(Word.TableRow))
            { // TableRow.TableRowProperties > Table.TableProperties.tblStyle (> default style)
                // ( ): done in Table level
                Word.TableRowProperties rowpr = StyleHelper.GetElement<Word.TableRowProperties>(obj);
                if (rowpr != null)
                    ret = StyleHelper.GetDescendants<T>(rowpr);

                if (ret == null)
                {
                    for (int i = 0; i < 1 && obj != null; i++)
                        obj = (obj.Parent != null) ? obj.Parent : null;
                    if (obj != null && obj.GetType() == typeof(Word.Table))
                        ret = this.GetAppliedElement<T>(obj);
                }
            }
            else if (objType == typeof(Word.TableCell))
            { // TableCell.TableCellProperties > Table.TableProperties.tblStyle (> default style)
                // ( ): done in Table level
                Word.TableCellProperties cellpr = StyleHelper.GetElement<Word.TableCellProperties>(obj);
                if (cellpr != null)
                    ret = StyleHelper.GetDescendants<T>(cellpr);

                if (ret == null)
                {
                    for (int i = 0; i < 2 && obj != null; i++)
                        obj = (obj.Parent != null) ? obj.Parent : null;
                    if (obj != null && obj.GetType() == typeof(Word.Table))
                        ret = this.GetAppliedElement<T>(obj);
                }
            }
            else if (objType == typeof(Word.Style))
            {
                Word.Style st = obj as Word.Style;
                ret = StyleHelper.GetDescendants<T>(st);
                if (ret == null)
                {
                    if (st.BasedOn != null)
                        ret = this.GetAppliedElement<T>(this.GetStyleById(st.BasedOn.Val));
                }
            }
            else // unknown type, just get everything can get
                ret = StyleHelper.GetDescendants<T>(obj);

            return ret;
        }

        /// <summary>
        /// Get Theme font name by script tag (e.g. Hant). Only search from theme>fontScheme>minorFont because majorFont is meant to be used with Headings (Heading 1, etc.) and minorFont with "normal text".
        /// </summary>
        /// <param name="scriptTag"></param>
        /// <returns></returns>
        public String GetThemeFontByScriptTag(String scriptTag)
        {
            foreach (SupplementalFont fontValue in
                this.theme.Theme.ThemeElements.FontScheme.MinorFont.Descendants<SupplementalFont>())
            {
                if (fontValue.Script != null &&
                    fontValue.Script.Value.IndexOf(scriptTag, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (fontValue.Typeface != null && fontValue.Typeface.HasValue)
                    {
                        return fontValue.Typeface.Value;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Get Theme font name by type. The possible values are majorBidi/minorBidi, majorHAnsi/minorHAnsi, majorEastAsia/minorEastAsia.
        /// </summary>
        /// <param name="fontName">Font type, the possible values are majorBidi/minorBidi, majorHAnsi/minorHAnsi, majorEastAsia/minorEastAsia.</param>
        /// <returns></returns>
        public String GetThemeFontByType(String fontType)
        {
            // http://blogs.msdn.com/b/officeinteroperability/archive/2013/04/22/office-open-xml-themes-schemes-and-fonts.aspx
            // major fonts are mainly for styles as headings, whereas minor fonts are generally applied to body and paragraph text

            if (fontType.IndexOf("majorBidi", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MajorFont.ComplexScriptFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MajorFont.ComplexScriptFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MajorFont.ComplexScriptFont.Typeface.Value;
            }
            else if (fontType.IndexOf("minorBidi", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MinorFont.ComplexScriptFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MinorFont.ComplexScriptFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MinorFont.ComplexScriptFont.Typeface.Value;
            }
            else if (fontType.IndexOf("majorHAnsi", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MajorFont.LatinFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MajorFont.LatinFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MajorFont.LatinFont.Typeface.Value;
            }
            else if (fontType.IndexOf("minorHAnsi", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MinorFont.LatinFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MinorFont.LatinFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MinorFont.LatinFont.Typeface.Value;
            }
            else if (fontType.IndexOf("majorEastAsia", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MajorFont.EastAsianFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MajorFont.EastAsianFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MajorFont.EastAsianFont.Typeface.Value;
            }
            else if (fontType.IndexOf("minorEastAsia", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (this.theme.Theme.ThemeElements.FontScheme.MinorFont.EastAsianFont != null &&
                    this.theme.Theme.ThemeElements.FontScheme.MinorFont.EastAsianFont.Typeface != null)
                    return this.theme.Theme.ThemeElements.FontScheme.MinorFont.EastAsianFont.Typeface.Value;
            }
            return null;
        }

        private Dictionary<String, Word.Style> cachedStyles = new Dictionary<string, Word.Style>();
        private Word.Style _getStyleById(String styleId)
        { // Only used by getStyleById()
            return this.styles.FirstOrDefault(c =>
            {
                return (c.StyleId != null && c.StyleId.HasValue) ? c.StyleId.Value == styleId : false;
            });
        }
        /// <summary>
        /// Get Word.Style by style ID. This method combines the LinkedStyle together if any.
        /// </summary>
        /// <param name="styleId">Style ID.</param>
        /// <returns>Return Word.Style object if found otherwise return null.</returns>
        public Word.Style GetStyleById(String styleId)
        {
            if (this.cachedStyles.ContainsKey(styleId))
                return this.cachedStyles[styleId];
            else
            {
                Word.Style broAStyle = this._getStyleById(styleId);
                if (broAStyle == null)
                    return null;

                // Retrieve LinkedStyle only for Paragraph & Character
                if ((broAStyle.Type.Value == Word.StyleValues.Paragraph ||
                     broAStyle.Type.Value == Word.StyleValues.Character) && broAStyle.LinkedStyle != null)
                {
                    Word.Style linkedStyle = this._getStyleById(broAStyle.LinkedStyle.Val);
                    if (linkedStyle != null && linkedStyle.StyleRunProperties != null)
                    {
                        if (broAStyle.Type.Value == Word.StyleValues.Paragraph)
                        {
                            broAStyle.StyleRunProperties = (Word.StyleRunProperties)linkedStyle.StyleRunProperties.CloneNode(true);
                        }
                        else
                        {
                            linkedStyle.StyleRunProperties = (Word.StyleRunProperties)broAStyle.StyleRunProperties.CloneNode(true);
                            broAStyle = linkedStyle;
                        }
                    }
                }
                this.cachedStyles[styleId] = broAStyle;
                return broAStyle;
            }
        }

        public enum DefaultStyleType
        {
            Paragraph,
            Character,
            Table,
            Numbering,
        }

        private Dictionary<DefaultStyleType, string> defaultStyleName = new Dictionary<DefaultStyleType, string>()
        {
            {DefaultStyleType.Paragraph, "paragraph"},
            {DefaultStyleType.Character, "character"},
            {DefaultStyleType.Table, "table"},
            {DefaultStyleType.Numbering, "numbering"},
        };

        /// <summary>
        /// Get default style by type, the possible types are Paragraph, Character, Table, and Numbering. Default styles are the styles have attribute w:default="1", not w:docDefaults.
        /// </summary>
        /// <param name="type">Target style type.</param>
        /// <returns></returns>
        public Word.Style GetDefaultStyle(DefaultStyleType type)
        {
            if (this.styles == null)
                return null;

            String typeStr = defaultStyleName[type];
            return this.styles.FirstOrDefault(c =>
            {
                try
                {
                    //// put "default" search in front has performance degradation because
                    //// it generates lots of KeyNotFoundException output to console, that
                    //// results in the terrible performance degradation
                    //OpenXmlAttribute attrDefault = c.GetAttribute("default", docDefaults.NamespaceUri);
                    //if (attrDefault.Value != null && attrDefault.Value == "1")
                    //{
                    //    OpenXmlAttribute attrType = c.GetAttribute("type", docDefaults.NamespaceUri);
                    //    if (attrType.Value != null)
                    //    {
                    //        if (attrType.Value.IndexOf(typeStr, 0, StringComparison.Ordinal) >= 0)
                    //            return true;
                    //    }
                    //}

                    OpenXmlAttribute attrType = c.GetAttribute("type", docDefaults.NamespaceUri);
                    if (attrType.Value != null)
                    {
                        if (attrType.Value.IndexOf(typeStr, 0, StringComparison.Ordinal) >= 0)
                        {
                            OpenXmlAttribute attrDefault = c.GetAttribute("default", docDefaults.NamespaceUri);
                            if (attrDefault.Value != null)
                            {
                                if (attrDefault.Value == "1")
                                    return true;
                            }
                        }
                    }
                }
                catch (Exception) { }
                return false;
            });
        }

        /// <summary>
        /// Get paragraph's numbering level object by searching abstractNums with the numbering level ID.
        /// </summary>
        /// <param name="pg"></param>
        /// <returns></returns>
        public Word.Level GetNumbering(Word.Paragraph pg)
        {
            Word.Level ret = null;

            if (pg == null || this.numbering == null)
                return ret;

            Word.ParagraphProperties pgpr = pg.Elements<Word.ParagraphProperties>().FirstOrDefault();
            if (pgpr == null)
                return ret;

            // direct numbering
            Word.NumberingProperties numPr = pgpr.Elements<Word.NumberingProperties>().FirstOrDefault();
            if (numPr != null && numPr.NumberingId != null)
            {
                int numId = numPr.NumberingId.Val;
                if (numId > 0) // numId == 0 means this paragraph doesn't have a list item
                {
                    int? ilvl = null;
                    String refStyleName = null;

                    if (numPr.NumberingLevelReference != null)
                    { // ilvl included in NumberingProperties
                        ilvl = numPr.NumberingLevelReference.Val;
                    }
                    else
                    { // doesn't have ilvl in NumberingProperties, search by referenced style name
                        Word.Style st = pgpr.Elements<Word.Style>().FirstOrDefault();
                        if (st != null && st.StyleName != null)
                            refStyleName = st.StyleName.Val;
                    }

                    // find abstractNumId by numId
                    Word.NumberingInstance numInstance = this.numbering.Elements<Word.NumberingInstance>().FirstOrDefault(c => c.NumberID.Value == numId);
                    if (numInstance != null)
                    {
                        // find abstractNum by abstractNumId
                        Word.AbstractNum abstractNum = (Word.AbstractNum)this.numbering.Elements<Word.AbstractNum>().FirstOrDefault(c =>
                            c.AbstractNumberId.Value == numInstance.AbstractNumId.Val);
                        if (abstractNum != null)
                        {
                            if (ilvl != null) // search by ilvl
                                ret = abstractNum.Elements<Word.Level>().FirstOrDefault(c => c.LevelIndex == ilvl);
                            else if (refStyleName != null) // search by matching referenced style name
                                ret = abstractNum.Elements<Word.Level>().FirstOrDefault(c =>
                                {
                                    return (c.ParagraphStyleIdInLevel != null) ? c.ParagraphStyleIdInLevel.Val == refStyleName : false;
                                });
                        }
                    }
                }
            }

            // TODO: linked style
            if (ret == null)
            {
                ;
            }

            return ret;
        }

        /// <summary>
        /// Get a list of numbering value from level-0 to level-ilvl. Call this method will 
        ///   1. increase the numbering value of level-ilvl by one automatically
        ///   2. restart all the levels larger than level-ilvl
        /// </summary>
        /// <param name="abstractNumId"></param>
        /// <param name="ilvl"></param>
        /// <returns></returns>
        public List<int> GetNumberingCurrent(int abstractNumId, int ilvl)
        {
            return this.nc.GetCurrent(abstractNumId, ilvl);
        }

        /// <summary>
        /// Restart the numbering from level-ilvl and all other levels behind/larger it.
        /// </summary>
        /// <param name="abstractNumId"></param>
        /// <param name="ilvl"></param>
        public void RestartNumbering(int abstractNumId, int ilvl)
        {
            this.nc.Restart(abstractNumId, ilvl);
        }

        /// <summary>
        /// Get hyperlink URL string by r:id
        /// </summary>
        /// <param name="rid"></param>
        /// <returns>Return hyperlink URL string or null.</returns>
        public String GetHyperlinkById(String rid)
        {
            HyperlinkRelationship hr = this.hyperlinkRelationships.FirstOrDefault(c => c.Id == rid);
            if (hr != null && hr.IsExternal && hr.Uri != null)
                return hr.Uri.OriginalString;
            else
                return null;
        }

        /// <summary>
        /// Return toggle property is ON or OFF. The logic is
        ///  1. if property is null, it's OFF
        ///  2. if property is not null but no attribute "val", it's ON
        ///  3. if property is not null and has attribute "val", on/off state depends on the value of attribute "val"
        /// </summary>
        /// <param name="property">Toggle property.</param>
        /// <returns>Return true means ON, false means OFF.</returns>
        public static bool GetToggleProperty(Word.OnOffType property)
        {
            bool ret = false;
            if (property != null)
                ret = (property.Val == null) ? true : property.Val.Value;
            return ret;
        }

        /// <summary>
        /// Clean dest's all attributes and duplicate from source.
        /// </summary>
        /// <param name="dest">The destination of attributes copy process.</param>
        /// <param name="source">The source to duplicate all attributes from.</param>
        public static void CopyAttributes(OpenXmlElement dest, OpenXmlElement source)
        {
            if (dest != null)
            {
                dest.ClearAllAttributes();
                if (source != null)
                    dest.SetAttributes(source.GetAttributes());
            }
        }

        /// <summary>
        /// Get specified element (e.g. RunFonts, Languages) from OpenXmlElement's elements.
        /// </summary>
        /// <typeparam name="T">Target element type (e.g. RunFonts).</typeparam>
        /// <param name="obj">The OpenXmlElemet to extract the element from.</param>
        /// <returns>Return found element or null if not found.</returns>
        public static T GetElement<T>(OpenXmlElement obj)
        {
            T ret = default(T);
            if (obj != null)
                ret = (T)Convert.ChangeType(obj.Elements().FirstOrDefault(c => c.GetType() == typeof(T)), typeof(T));
            return ret;
        }

        /// <summary>
        /// Get specified element (e.g. RunFonts, Languages) from ALL THE DESCENDANTS of OpenXmlElement.
        /// </summary>
        /// <typeparam name="T">Target element type (e.g. RunFonts).</typeparam>
        /// <param name="obj">The OpenXmlElemet to extract the element from.</param>
        /// <returns>Return found element or null if not found.</returns>
        public static T GetDescendants<T>(OpenXmlElement obj)
        {
            T ret = default(T);
            if (obj != null)
                ret = (T)Convert.ChangeType(obj.Descendants().FirstOrDefault(c => c.GetType() == typeof(T)), typeof(T));
            return ret;
        }
    }
}
