using DocumentFormat.OpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using iTSText = iTextSharp.text;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToPdf
{
    class TableHelperCell
    {
        private TableHelper owner;

        /// <summary>
        /// Cell ID.
        /// </summary>
        public int cellId = 0;

        /// <summary>
        /// Cell's row number.
        /// </summary>
        public int rowId = -1;

        /// <summary>
        /// Cell's column number.
        /// </summary>
        public int colId = -1;

        /// <summary>
        /// Get whether this cell is the start of the row.
        /// rowStart is set to the main cell of col-spanned cells (i.e. the first cell of col-spanned cells).
        /// gridBefore cells do not set as rowStart.
        /// </summary>
        public bool rowStart = false;

        /// <summary>
        /// Get whether this cell is the end of the row.
        /// rowEnd is set to the main cell of col-spanned cells (i.e. the first cell of col-spanned cells).
        /// gridAfter cells do not set as rowEnd.
        /// </summary>
        public bool rowEnd = false;

        public Word.TableRow row = null; // point to Word.TableRow
        public Word.TableCell cell = null; // point to Word.TableCell
        public List<iTSText.IElement> elements = new List<iTSText.IElement>();

        public Word.TableCellBorders Borders = null;

        public TableHelperCell(TableHelper owner, int cellId, int rowId, int colId)
        {
            this.owner = owner;
            this.cellId = cellId;
            this.rowId = rowId;
            this.colId = colId;
        }

        /// <summary>
        /// Get whether this cell is blank.
        /// </summary>
        public bool Blank
        {
            //get { return !(elements.Count > 0); }
            get { return cell == null; }
        }

        /// <summary>
        /// Get cell's row span.
        /// </summary>
        public int RowSpan
        {
            get { return (Blank) ? 0 : this.owner.GetRowSpan(this.cellId); }
        }

        /// <summary>
        /// Get cell's column span.
        /// </summary>
        public int ColSpan
        {
            get { return (Blank) ? 0 : this.owner.GetColSpan(this.cellId); }
        }
    }

    class TableHelper : IEnumerable
    {
        public delegate iTSText.IElement CellChildElementsHandler(OpenXmlElement element);

        private List<TableHelperCell> cells = new List<TableHelperCell>();
        private StyleHelper styleHelper = null;
        private Word.Table table = null;
        private int rowLen = 0;
        private int colLen = 0;

        // ------
        // Below variables are assigned in ParseTable()
        private Word.TableBorders condFmtbr = null;
        private float[] columnWidth = null; // after adjustTableColumnsWidth()
        // ------

        /// <summary>
        /// Get a float array which indicates all the table column width. This array is only availabe after ParseTable() got called.
        /// </summary>
        public float[] ColumnWidth { get { return this.columnWidth; } }

        /// <summary>
        /// Handle OpenXML child elements in cell (e.g. paragraph or table), called in ParseTable().
        /// </summary>
        public event CellChildElementsHandler CellChildElementsProc;

        public StyleHelper StHelper
        {
            get { return this.styleHelper;  }
            set { this.styleHelper = value; }
        }

        // TODO: get row height (not total row height): http://kuujinbo.info/iTextSharp/rowHeights.aspx

        public void ParseTable(Word.Table t)
        {
            this.table = t;

            if (this.table == null)
                return;

            this.colLen = this.getTableGridCols(table).Count();
            int cellId = 0;
            int rowId = 0;
            foreach (Word.TableRow row in table.Elements<Word.TableRow>())
            {
                bool rowStart = true;

                // w:gridBefore
                Word.GridBefore tmpGridBefore = (Word.GridBefore)StHelper.GetAppliedElement<Word.GridBefore>(row);
                int skipGridsBefore = (tmpGridBefore != null && tmpGridBefore.Val != null) ? tmpGridBefore.Val.Value : 0;

                // w:gridAfter
                Word.GridAfter tmpGridAfter = (Word.GridAfter)StHelper.GetAppliedElement<Word.GridAfter>(row);
                int skipGridsAfter = (tmpGridAfter != null && tmpGridAfter.Val != null) ? tmpGridAfter.Val.Value : 0;

                int colId = 0;

                // gridBefore (with same cellId and set as blank)
                if (skipGridsBefore > 0)
                {
                    for (int i = 0; i < skipGridsBefore; i++)
                    {
                        cells.Add(new TableHelperCell(this, cellId, rowId, colId));
                        colId++; // increase for each cells.Add()
                        Debug.Write(String.Format("{0:X} ", cellId));
                    }
                    cellId++;
                }

                foreach (Word.TableCell col in row.Elements<Word.TableCell>())
                {
                    int _cellId = cellId;
                    int _cellcount = 1;

                    TableHelperCell basecell = new TableHelperCell(this, _cellId, rowId, colId);
                    if (rowStart)
                    {
                        basecell.rowStart = true;
                        rowStart = false;
                    }
                    basecell.row = row;
                    basecell.cell = col;

                    // process rowspan and colspan
                    if (col.TableCellProperties != null)
                    {
                        // colspan
                        if (col.TableCellProperties.GridSpan != null)
                            _cellcount = col.TableCellProperties.GridSpan.Val;

                        // rowspan
                        if (col.TableCellProperties.VerticalMerge != null)
                        {
                            // "continue": get cellId from (rowId-1):(colId)
                            if (col.TableCellProperties.VerticalMerge.Val == null ||
                                col.TableCellProperties.VerticalMerge.Val == "continue")
                                _cellId = cells[(this.ColumnLength * (rowId - 1)) + colId].cellId;
                            // "restart": the begin of rowspan
                            else
                                _cellId = cellId;
                        }
                    }
                    basecell.cellId = _cellId;

                    // Handle OpenXML child elements
                    if (this.CellChildElementsProc != null)
                    {
                        foreach (OpenXmlElement element in col.Elements())
                        {
                            iTSText.IElement em = this.CellChildElementsProc(element);
                            if (em != null) basecell.elements.Add(em);
                        }
                    }

                    cells.Add(basecell);
                    colId++; // increase for each cells.Add()
                    Debug.Write(String.Format("{0:X} ", _cellId));

                    for (int i = 1; i < _cellcount; i++)
                    { // Add spanned cells
                        cells.Add(new TableHelperCell(this, _cellId, rowId, colId));
                        colId++; // increase for each cells.Add()
                        Debug.Write(String.Format("{0:X} ", _cellId));
                    }

                    // The latest cellId was used, then we must increase it for future usage
                    if (cellId == _cellId)
                        cellId++;
                }
                int rowEndIndex = cells.Count - 1;
                int rowEndCellId = cells[rowEndIndex].cellId;
                while (rowEndIndex > 0 && cells[rowEndIndex - 1].cellId == rowEndCellId)
                    rowEndIndex--;
                cells[rowEndIndex].rowEnd = true;

                // gridAfter (with same cellId and set as blank)
                if (skipGridsAfter > 0 && colId < this.ColumnLength)
                {
                    for (int i = 0; i < skipGridsAfter; i++)
                    {
                        cells.Add(new TableHelperCell(this, cellId, rowId, colId));
                        colId++; // increase for each cells.Add()
                        Debug.Write(String.Format("{0:X} ", cellId));
                    }
                    cellId++;
                }

                rowId++;
                Debug.Write("\n");
            }
            this.rowLen = rowId;

            //if (cells.Count <= 0)
            //    return;

            // prepare table conditional formatting (border), which will be used in
            // applyCellBorders() so must be called before applyCellBorders()
            this.rollingUpTableBorders();

            // ====== Resolve cell border conflict ======

            for (int r = 0; r < this.RowLength; r++)
            {
                // The following handles the situation where
                // if table innerV is set, and cnd format for first row specifies nill border, then nil border wins.
                // if table innerH is set, and cnd format for first columns specifies nil border, then table innerH wins.

                //// TODO: if row's cellspacing is not zero then bypass this row
                //Word.TableCellSpacing tcspacing = _CvrtCell.GetTableRow(cells, r).TableRowProperties.Descendants<Word.TableCellSpacing>().FirstOrDefault();
                //if (tcspacing.Type.Value != Word.TableWidthUnitValues.Nil)
                //    continue;

                for (int c = 0; c < this.ColumnLength; c++)
                {
                    TableHelperCell me = this.GetCell(r, c);
                    if (me.Blank) continue;

                    if (me.Borders == null)
                        me.Borders = this.applyCellBorders(me.cell.Descendants<Word.TableCellBorders>().FirstOrDefault(),
                            (me.colId == 0) | me.rowStart,
                            (me.colId + this.GetColSpan(me.cellId) == this.ColumnLength) | me.rowEnd,
                            (me.rowId == 0),
                            (me.rowId + this.GetRowSpan(me.cellId) == this.RowLength)
                            );
                    int colspan = this.GetColSpan(me.cellId);
                    int rowspan = this.GetRowSpan(me.cellId);

                    if ((c + (colspan - 1) + 1) < this.ColumnLength) // not last column
                    {
                        // get the cell at the right side of me
                        List<TableHelperCell> rights = new List<TableHelperCell>();
                        for (int i = 0; i < rowspan; i++)
                        {
                            TableHelperCell tmp = this.GetCell(r + i, c + (colspan - 1) + 1);
                            if (tmp != null && !tmp.Blank) rights.Add(tmp);
                        }

                        if (rights.Count > 0)
                        {
                            foreach (TableHelperCell right in rights)
                            {
                                if (right.Borders == null)
                                    right.Borders = this.applyCellBorders(right.cell.Descendants<Word.TableCellBorders>().FirstOrDefault(),
                                        (right.colId == 0) | right.rowStart,
                                        (right.colId + this.GetColSpan(right.cellId) == this.ColumnLength) | right.rowEnd,
                                        (right.rowId == 0),
                                        (right.rowId + this.GetRowSpan(right.cellId) == this.RowLength)
                                        );

                                bool meWin = compareBorder(me.Borders, right.Borders, compareDirection.Horizontal);
                                if (meWin)
                                    StyleHelper.CopyAttributes(right.Borders.LeftBorder, me.Borders.RightBorder);
                            }
                            me.Borders.RightBorder.ClearAllAttributes();
                        }
                    }

                    // can't bypass row-spanned cells because they still have tcBorders property
                    if ((r + 1) < this.RowLength) // not last row
                    {
                        // get the cell below me
                        List<TableHelperCell> bottoms = new List<TableHelperCell>();
                        for (int i = 0; i < colspan; i++)
                        {
                            TableHelperCell tmp = this.GetCell(r + 1, c + i);
                            if (tmp != null && !tmp.Blank) bottoms.Add(tmp);
                        }

                        foreach (TableHelperCell bottom in bottoms)
                        {
                            if (bottom.Borders == null)
                                bottom.Borders = this.applyCellBorders(bottom.cell.Descendants<Word.TableCellBorders>().FirstOrDefault(),
                                    (bottom.colId == 0) | bottom.rowStart,
                                    (bottom.colId + this.GetColSpan(bottom.cellId) == this.ColumnLength) | bottom.rowEnd,
                                    (bottom.rowId == 0),
                                    (bottom.rowId + this.GetRowSpan(bottom.cellId) == this.RowLength)
                                    );

                            bool meWin = compareBorder(me.Borders, bottom.Borders, compareDirection.Vertical);
                            if (meWin)
                                StyleHelper.CopyAttributes(bottom.Borders.TopBorder, me.Borders.BottomBorder);
                        }
                    }
                }
            }

            if (this.cells.Count > 0)
            { // re-process each cell's border conflict with its bottom cell
                for (int i = 0; i < this.cells[this.cells.Count - 1].cellId; i++)
                {
                    TableHelperCell me = this.GetCellByCellId(i);
                    if (me.Blank) // ignore gridBefore/gridAfter cells
                        continue;

                    if (me.RowSpan > 1)
                    { // merge bottom border from the last cell of row-spanned cells
                        TableHelperCell meRowEnd = this.GetCell(me.rowId + (me.RowSpan - 1), me.colId);
                        if (meRowEnd != null && meRowEnd.Borders != null && meRowEnd.Borders.BottomBorder != null)
                            StyleHelper.CopyAttributes(me.Borders.BottomBorder, meRowEnd.Borders.BottomBorder);
                    }

                    if (me.rowId + me.RowSpan < this.RowLength)
                    { // if me is not at the last row, compare the border with the cell below it
                        TableHelperCell bottom = this.GetCellByCellId(this.GetCell(me.rowId + me.RowSpan, me.colId).cellId);
                        bool meWin = compareBorder(me.Borders, bottom.Borders, compareDirection.Vertical);
                        if (!meWin)
                            me.Borders.BottomBorder.ClearAllAttributes();
                    }
                }
            }

            // ====== Adjust table columns width by their content ======

            this.adjustTableColumnsWidth();
        }

        private enum compareDirection { Horizontal, Vertical, }
        /// <summary>
        /// Compare border and return who is win.
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <param name="dir"></param>
        /// <returns>Return ture means a win, false means b win.</returns>
        private bool compareBorder(Word.TableCellBorders a, Word.TableCellBorders b, compareDirection dir)
        {
            // compare line style
            int weight1 = 0, weight2 = 0;
            if (dir == compareDirection.Horizontal)
            {
                if (a.RightBorder != null && a.RightBorder.Val != null)
                    weight1 = (BorderNumber.ContainsKey(a.RightBorder.Val)) ? BorderNumber[a.RightBorder.Val] : 1;
                else if (a.InsideVerticalBorder != null && a.InsideVerticalBorder.Val != null)
                    weight1 = (BorderNumber.ContainsKey(a.InsideVerticalBorder.Val)) ? BorderNumber[a.InsideVerticalBorder.Val] : 1;

                if (b.LeftBorder != null && b.LeftBorder.Val != null)
                    weight2 = (BorderNumber.ContainsKey(b.LeftBorder.Val)) ? BorderNumber[b.LeftBorder.Val] : 1;
                else if (b.InsideVerticalBorder != null && b.InsideVerticalBorder.Val != null)
                    weight2 = (BorderNumber.ContainsKey(b.InsideVerticalBorder.Val)) ? BorderNumber[b.InsideVerticalBorder.Val] : 1;
            }
            else if (dir == compareDirection.Vertical)
            {
                if (a.BottomBorder != null && a.BottomBorder.Val != null)
                    weight1 = (BorderNumber.ContainsKey(a.BottomBorder.Val)) ? BorderNumber[a.BottomBorder.Val] : 1;
                else if (a.InsideHorizontalBorder != null && a.InsideHorizontalBorder.Val != null)
                    weight1 = (BorderNumber.ContainsKey(a.InsideHorizontalBorder.Val)) ? BorderNumber[a.InsideHorizontalBorder.Val] : 1;
                
                if (b.TopBorder != null && b.TopBorder.Val != null)
                    weight2 = (BorderNumber.ContainsKey(b.TopBorder.Val)) ? BorderNumber[b.TopBorder.Val] : 1;
                else if (b.InsideHorizontalBorder != null && b.InsideHorizontalBorder.Val != null)
                    weight2 = (BorderNumber.ContainsKey(b.InsideHorizontalBorder.Val)) ? BorderNumber[b.InsideHorizontalBorder.Val] : 1;
            }

            if (weight1 > weight2)
                return true;
            else if (weight2 > weight1)
                return false;
            
            // compare width
            float size1 = 0f, size2 = 0f;
            if (dir == compareDirection.Horizontal)
            {
                if (a.RightBorder.Size != null && a.RightBorder.Size.HasValue)
                    size1 = Tools.ConvertToPoint(a.RightBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                else if (a.InsideVerticalBorder.Size != null && a.InsideVerticalBorder.Size.HasValue)
                    size1 = Tools.ConvertToPoint(a.InsideVerticalBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                
                if (b.LeftBorder.Size != null && b.LeftBorder.Size.HasValue)
                    size2 = Tools.ConvertToPoint(b.LeftBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                else if (b.InsideVerticalBorder.Size != null && b.InsideVerticalBorder.Size.HasValue)
                    size2 = Tools.ConvertToPoint(b.InsideVerticalBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
            }
            else if (dir == compareDirection.Vertical)
            {
                if (a.BottomBorder.Size != null && a.BottomBorder.Size.HasValue)
                    size1 = Tools.ConvertToPoint(a.BottomBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                else if (a.InsideHorizontalBorder.Size != null && a.InsideHorizontalBorder.Size.HasValue)
                    size1 = Tools.ConvertToPoint(a.InsideHorizontalBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                
                if (b.TopBorder.Size != null && b.TopBorder.Size.HasValue)
                    size2 = Tools.ConvertToPoint(b.TopBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
                else if (b.InsideHorizontalBorder.Size != null && b.InsideHorizontalBorder.Size.HasValue)
                    size2 = Tools.ConvertToPoint(b.InsideHorizontalBorder.Size.Value, Tools.SizeEnum.LineBorder, -1f);
            }

            if (size1 > size2)
                return true;
            else if (size2 > size1)
                return false;

            // compare brightness
            //   TODO: current brightness implementation is based on Luminance 
            //   but ISO $17.4.66 defines the comparisons should be
            //   1. R+B+2G, 2. B+2G, 3. G
            float brightness1 = 0f, brightness2 = 0f;
            if (dir == compareDirection.Horizontal)
            {
                if (a.RightBorder.Color != null && a.RightBorder.Color.HasValue)
                    brightness1 = Tools.RgbBrightness(a.RightBorder.Color.Value);
                else if (a.InsideVerticalBorder.Color != null && a.InsideVerticalBorder.Color.HasValue)
                    brightness1 = Tools.RgbBrightness(a.InsideVerticalBorder.Color.Value);

                if (b.LeftBorder.Color != null && b.LeftBorder.Color.HasValue)
                    brightness2 = Tools.RgbBrightness(b.LeftBorder.Color.Value);
                else if (b.InsideVerticalBorder.Color != null && b.InsideVerticalBorder.Color.HasValue)
                    brightness2 = Tools.RgbBrightness(b.InsideVerticalBorder.Color.Value);
            }
            else if (dir == compareDirection.Vertical)
            {
                if (a.BottomBorder.Color != null && a.BottomBorder.Color.HasValue)
                    brightness1 = Tools.RgbBrightness(a.BottomBorder.Color.Value);
                else if (a.InsideHorizontalBorder.Color != null && a.InsideHorizontalBorder.Color.HasValue)
                    brightness1 = Tools.RgbBrightness(a.InsideHorizontalBorder.Color.Value);

                if (b.TopBorder.Color != null && b.TopBorder.Color.HasValue)
                    brightness2 = Tools.RgbBrightness(b.TopBorder.Color.Value);
                else if (b.InsideHorizontalBorder.Color != null && b.InsideHorizontalBorder.Color.HasValue)
                    brightness2 = Tools.RgbBrightness(b.InsideHorizontalBorder.Color.Value);
            }

            // smaller brightness wins
            if (brightness1 < brightness2)
                return true;
            else if (brightness2 < brightness1)
                return false;
            
            return false; // special trick, especially for vertical comparison
        }

        private void adjustTableColumnsWidth()
        {
            if (this.table == null) return;

            this.columnWidth = this.getTableGridCols(this.table);

            // Get table total width
            float totalWidth = -1f;
            bool autoWidth = false;
            Word.TableWidth tableWidth = this.styleHelper.GetAppliedElement<Word.TableWidth>(this.table);
            if (tableWidth != null && tableWidth.Type != null)
            {
                switch (tableWidth.Type.Value)
                {
                    case Word.TableWidthUnitValues.Nil:
                    case Word.TableWidthUnitValues.Auto: // fits the contents
                    default:
                        autoWidth = true;
                        break;
                    case Word.TableWidthUnitValues.Dxa:
                        if (tableWidth.Width != null)
                            totalWidth = Tools.ConvertToPoint(tableWidth.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        break;
                    case Word.TableWidthUnitValues.Pct:
                        if (tableWidth.Width != null)
                        {
                            totalWidth = this.styleHelper.PrintablePageWidth * Tools.Percentage(tableWidth.Width.Value);
                            //if (table.Parent.GetType() == typeof(Word.Body))
                            //    totalWidth = (float)((pdfDoc.PageSize.Width - pdfDoc.LeftMargin - pdfDoc.RightMargin) * percentage(tableWidth.Width.Value));
                            //else
                            //    totalWidth = this.getCellWidth(table.Parent as Word.TableCell) * percentage(tableWidth.Width.Value);
                        }
                        break;
                }
            }
            Console.WriteLine("Table total width: " + totalWidth);

            if (!autoWidth)
                scaleTableColumnsWidth(ref this.columnWidth, totalWidth);
            else
                totalWidth = this.columnWidth.Sum();

            for (int i = 0; i < this.RowLength; i++)
            {
                // Get all cells in this row
                List<TableHelperCell> cellsInRow = this.cells.FindAll(c => c.rowId == i);
                if (cellsInRow.Count <= 0)
                    continue;

                // Get if any gridBefore & gridAfter
                int skipGridsBefore = 0, skipGridsAfter = 0;
                float skipGridsBeforeWidth = 0f, skipGridsAfterWidth = 0f;
                TableHelperCell head = cellsInRow.FirstOrDefault(c => c.rowStart);
                if (head != null)
                {
                    if (head.row.TableRowProperties != null)
                    {
                        // w:gridBefore
                        var tmpGridBefore = head.row.TableRowProperties.Elements<Word.GridBefore>().FirstOrDefault();
                        if (tmpGridBefore != null && tmpGridBefore.Val != null)
                            skipGridsBefore = tmpGridBefore.Val.Value;

                        // w:wBefore
                        var tmpGridBeforeWidth = head.row.TableRowProperties.Elements<Word.WidthBeforeTableRow>().FirstOrDefault();
                        if (tmpGridBeforeWidth != null && tmpGridBeforeWidth.Width != null)
                            skipGridsBeforeWidth = Tools.ConvertToPoint(Convert.ToInt32(tmpGridBeforeWidth.Width.Value), Tools.SizeEnum.TwentiethsOfPoint, -1f);

                        // w:gridAfter
                        var tmpGridAfter = head.row.TableRowProperties.Elements<Word.GridAfter>().FirstOrDefault();
                        if (tmpGridAfter != null && tmpGridAfter.Val != null)
                            skipGridsAfter = tmpGridAfter.Val.Value;

                        // w:wAfter
                        var tmpGridAfterWidth = head.row.TableRowProperties.Elements<Word.WidthAfterTableRow>().FirstOrDefault();
                        if (tmpGridAfterWidth != null && tmpGridAfterWidth.Width != null)
                            skipGridsAfterWidth = Tools.ConvertToPoint(Convert.ToInt32(tmpGridAfterWidth.Width.Value), Tools.SizeEnum.TwentiethsOfPoint, -1f);
                    }
                }

                int j = 0;
                int edgeEnd = 0;

                // -------
                // gridBefore
                edgeEnd = skipGridsBefore;
                for (; j < edgeEnd; j++)
                    // deduce specific columns width from required width
                    skipGridsBeforeWidth -= this.columnWidth[j];
                if (skipGridsBeforeWidth > 0f)
                    // if required width is larger than the total width of specific columns,
                    // the remaining required width adds to the last specific column 
                    this.columnWidth[edgeEnd - 1] += skipGridsBeforeWidth;

                // ------
                // cells
                while (j < (cellsInRow.Count - skipGridsAfter))
                {
                    float reqCellWidth = 0f;
                    Word.TableCellWidth cellWidth = this.styleHelper.GetAppliedElement<Word.TableCellWidth>(cellsInRow[j].cell);
                    if (cellWidth != null && cellWidth.Type != null)
                    {
                        switch (cellWidth.Type.Value)
                        {
                            case Word.TableWidthUnitValues.Auto:
                                // TODO: calculate the items width
                                if (cellsInRow[j].elements.Count > 0)
                                {
                                    iTSText.IElement element = cellsInRow[j].elements[0];
                                }
                                break;
                            case Word.TableWidthUnitValues.Nil:
                            default:
                                break;
                            case Word.TableWidthUnitValues.Dxa:
                                if (cellWidth.Width != null)
                                    reqCellWidth = Tools.ConvertToPoint(cellWidth.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                                break;
                            case Word.TableWidthUnitValues.Pct:
                                if (cellWidth.Width != null)
                                    reqCellWidth = Tools.Percentage(cellWidth.Width.Value) * totalWidth;
                                break;
                        }
                    }

                    // check row span
                    int spanCount = 1;
                    if (cellsInRow[j].cell != null)
                    {
                        Word.TableCell tmpCell = cellsInRow[j].cell;
                        if (tmpCell.TableCellProperties != null)
                        {
                            Word.GridSpan span = tmpCell.TableCellProperties.Elements<Word.GridSpan>().FirstOrDefault();
                            spanCount = (span != null && span.Val != null) ? span.Val.Value : 1;
                        }
                    }

                    edgeEnd = j + spanCount;
                    for (; j < edgeEnd; j++)
                        // deduce specific columns width from required width
                        reqCellWidth -= this.columnWidth[j];
                    if (reqCellWidth > 0f)
                        // if required width is larger than the total width of specific columns,
                        // the remaining required width adds to the last specific column 
                        this.columnWidth[edgeEnd - 1] += reqCellWidth;
                }

                // ------
                // gridAfter
                edgeEnd = j + skipGridsAfter;
                for (; j < edgeEnd; j++)
                    // deduce specific columns width from required width
                    skipGridsAfterWidth -= this.columnWidth[j];
                if (skipGridsAfterWidth > 0f)
                    // if required width is larger than the total width of specific columns,
                    // the remaining required width adds to the last specific column 
                    this.columnWidth[edgeEnd - 1] += skipGridsAfterWidth;

                if (!autoWidth) // fixed table width, adjust width to fit in
                    scaleTableColumnsWidth(ref this.columnWidth, totalWidth);
                else // auto table width
                    totalWidth = this.columnWidth.Sum();
            }
        }

        private void scaleTableColumnsWidth(ref float[] columns, float totalWidth)
        {
            float sum = columns.Sum();
            if (sum > totalWidth)
            {
                float ratio = totalWidth / sum;
                for (int j = 0; j < columns.Length; j++)
                    columns[j] *= ratio;
            }
        }

        private static Dictionary<string, int> BorderNumber = new Dictionary<string, int>()
        {
            {"single", 1 },
            {"thick", 2 },
            {"double", 3 },
            {"dotted", 4 },
            {"dashed", 5 },
            {"dotDash", 6 },
            {"dotDotDash", 7 },
            {"triple", 8 },
            {"thinThickSmallGap", 9 },
            {"thickThinSmallGap", 10 },
            {"thinThickThinSmallGap", 11 },
            {"thinThickMediumGap", 12 },
            {"thickThinMediumGap", 13 },
            {"thinThickThinMediumGap", 14 },
            {"thinThickLargeGap", 15 },
            {"thickThinLargeGap", 16 },
            {"thinThickThinLargeGap", 17 },
            {"wave", 18 },
            {"doubleWave", 19 },
            {"dashSmallGap", 20 },
            {"dashDotStroked", 21 },
            {"threeDEmboss", 22 },
            {"threeDEngrave", 23 },
            {"outset", 24 },
            {"inset", 25 },
        };

        /// <summary>
        /// Get table gridCols information and convert to twip to Points.
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private float[] getTableGridCols(Word.Table table)
        {
            float[] grids = null;

            if (table == null)
                return grids;

            // Get grids and their width
            Word.TableGrid grid = (Word.TableGrid)table.Elements<Word.TableGrid>().FirstOrDefault();
            if (grid != null)
            {
                List<Word.GridColumn> gridCols = grid.Elements<Word.GridColumn>().ToList();
                if (gridCols.Count > 0)
                {
                    grids = new float[gridCols.Count];
                    for (int i = 0; i < gridCols.Count; i++)
                    {
                        if (gridCols[i].Width != null)
                            grids[i] = Tools.ConvertToPoint(gridCols[i].Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        else
                            grids[i] = 0f;
                    }
                }
            }
            return grids;
        }

        /// <summary>
        /// Rolling up table border property from TableProperties > TableStyle > Default style.
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private void rollingUpTableBorders()
        {
            if (this.table == null) return;

            this.condFmtbr = new Word.TableBorders();

            List<Word.TableBorders> borders = new List<Word.TableBorders>();
            Word.TableProperties tblPrs = this.table.Elements<Word.TableProperties>().FirstOrDefault();
            if (tblPrs != null)
            {
                // get from table properties (priority 1)
                Word.TableBorders tmp = tblPrs.Descendants<Word.TableBorders>().FirstOrDefault();
                if (tmp != null) borders.Insert(0, tmp);

                // get from styles (priority 2)
                if (tblPrs.TableStyle != null && tblPrs.TableStyle.Val != null)
                {
                    Word.Style st = this.styleHelper.GetStyleById(tblPrs.TableStyle.Val);
                    while (st != null)
                    {
                        tmp = st.Descendants<Word.TableBorders>().FirstOrDefault();
                        if (tmp != null) borders.Insert(0, tmp);

                        st = (st.BasedOn != null && st.BasedOn.Val != null) ? this.styleHelper.GetStyleById(st.BasedOn.Val) : null;
                    }
                }

                // get from default table style (priority 3)
                Word.Style defaultTableStyle = this.styleHelper.GetDefaultStyle(StyleHelper.DefaultStyleType.Table);
                if (defaultTableStyle != null)
                {
                    tmp = defaultTableStyle.Descendants<Word.TableBorders>().FirstOrDefault();
                    if (tmp != null) borders.Insert(0, tmp);
                }
            }

            foreach (Word.TableBorders border in borders)
            {
                if (border.TopBorder != null)
                {
                    if (this.condFmtbr.TopBorder == null)
                        this.condFmtbr.TopBorder = (Word.TopBorder)border.TopBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.TopBorder, border.TopBorder);
                }
                if (border.BottomBorder != null)
                {
                    if (this.condFmtbr.BottomBorder == null)
                        this.condFmtbr.BottomBorder = (Word.BottomBorder)border.BottomBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.BottomBorder, border.BottomBorder);
                }
                if (border.LeftBorder != null)
                {
                    if (this.condFmtbr.LeftBorder == null)
                        this.condFmtbr.LeftBorder = (Word.LeftBorder)border.LeftBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.LeftBorder, border.LeftBorder);
                }
                if (border.RightBorder != null)
                {
                    if (this.condFmtbr.RightBorder == null)
                        this.condFmtbr.RightBorder = (Word.RightBorder)border.RightBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.RightBorder, border.RightBorder);
                }
                if (border.InsideHorizontalBorder != null)
                {
                    if (this.condFmtbr.InsideHorizontalBorder == null)
                        this.condFmtbr.InsideHorizontalBorder = (Word.InsideHorizontalBorder)border.InsideHorizontalBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.InsideHorizontalBorder, border.InsideHorizontalBorder);
                }
                if (border.InsideVerticalBorder != null)
                {
                    if (this.condFmtbr.InsideVerticalBorder == null)
                        this.condFmtbr.InsideVerticalBorder = (Word.InsideVerticalBorder)border.InsideVerticalBorder.CloneNode(true);
                    else
                        StyleHelper.CopyAttributes(this.condFmtbr.InsideVerticalBorder, border.InsideVerticalBorder);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellbr"></param>
        /// <param name="firstColumn">Is current cell at the first column?</param>
        /// <param name="lastColumn">Is current cell at the last column?</param>
        /// <param name="firstRow">Is current cell at the first row?</param>
        /// <param name="lastRow">Is current cell at the first row?</param>
        /// <returns></returns>
        private Word.TableCellBorders applyCellBorders(Word.TableCellBorders cellbr,
            bool firstColumn, bool lastColumn, bool firstRow, bool lastRow)
        {
            Word.TableCellBorders ret = new Word.TableCellBorders()
            {
                TopBorder = new Word.TopBorder(),
                BottomBorder = new Word.BottomBorder(),
                LeftBorder = new Word.LeftBorder(),
                RightBorder = new Word.RightBorder(),
                InsideHorizontalBorder = new Word.InsideHorizontalBorder(),
                InsideVerticalBorder = new Word.InsideVerticalBorder(),
                TopLeftToBottomRightCellBorder = new Word.TopLeftToBottomRightCellBorder(),
                TopRightToBottomLeftCellBorder = new Word.TopRightToBottomLeftCellBorder(),
            };

            // cell border first, if no cell border then conditional formating (table border + style)
            if (cellbr != null && cellbr.TopBorder != null)
                StyleHelper.CopyAttributes(ret.TopBorder, cellbr.TopBorder);
            else
            {
                if (this.condFmtbr != null && this.condFmtbr.TopBorder != null)
                    StyleHelper.CopyAttributes(ret.TopBorder, this.condFmtbr.TopBorder);
                if (!firstRow && (this.condFmtbr != null && this.condFmtbr.InsideHorizontalBorder != null))
                    StyleHelper.CopyAttributes(ret.TopBorder, this.condFmtbr.InsideHorizontalBorder);
            }

            if (cellbr != null && cellbr.BottomBorder != null)
                StyleHelper.CopyAttributes(ret.BottomBorder, cellbr.BottomBorder);
            else
            {
                if (this.condFmtbr != null && this.condFmtbr.BottomBorder != null)
                    StyleHelper.CopyAttributes(ret.BottomBorder, this.condFmtbr.BottomBorder);
                if (!lastRow && (this.condFmtbr != null && this.condFmtbr.InsideHorizontalBorder != null))
                    StyleHelper.CopyAttributes(ret.BottomBorder, this.condFmtbr.InsideHorizontalBorder);
            }

            if (cellbr != null && cellbr.LeftBorder != null)
                StyleHelper.CopyAttributes(ret.LeftBorder, cellbr.LeftBorder);
            else
            {
                if (this.condFmtbr != null && this.condFmtbr.LeftBorder != null)
                    StyleHelper.CopyAttributes(ret.LeftBorder, this.condFmtbr.LeftBorder);
                if (!firstColumn && (this.condFmtbr != null && this.condFmtbr.InsideVerticalBorder != null))
                    StyleHelper.CopyAttributes(ret.LeftBorder, this.condFmtbr.InsideVerticalBorder);
            }

            if (cellbr != null && cellbr.RightBorder != null)
                StyleHelper.CopyAttributes(ret.RightBorder, cellbr.RightBorder);
            else
            {
                if (this.condFmtbr != null && this.condFmtbr.RightBorder != null)
                    StyleHelper.CopyAttributes(ret.RightBorder, this.condFmtbr.RightBorder);
                if (!lastColumn && (this.condFmtbr != null && this.condFmtbr.InsideVerticalBorder != null))
                    StyleHelper.CopyAttributes(ret.RightBorder, this.condFmtbr.InsideVerticalBorder);
            }

            if (cellbr != null && cellbr.TopLeftToBottomRightCellBorder != null)
                StyleHelper.CopyAttributes(ret.TopLeftToBottomRightCellBorder, cellbr.TopLeftToBottomRightCellBorder);
            if (cellbr != null && cellbr.TopRightToBottomLeftCellBorder != null)
                StyleHelper.CopyAttributes(ret.TopRightToBottomLeftCellBorder, cellbr.TopRightToBottomLeftCellBorder);

            return ret;
        }

        /// <summary>
        /// Get table column length.
        /// </summary>
        public int ColumnLength
        {
            get { return this.colLen; }
        }

        /// <summary>
        /// Get table row length.
        /// </summary>
        public int RowLength
        {
            get { return this.rowLen; }
        }

        /// <summary>
        /// Get TableRow by row ID.
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public Word.TableRow GetTableRow(int row)
        {
            int max = (row + 1) * this.ColumnLength;
            for (int index = row * this.ColumnLength; index < this.cells.Count && index < max; index++)
            {
                if (this.cells[index].rowId == row && !this.cells[index].Blank && this.cells[index].row != null)
                    return this.cells[index].row;
            }
            return null;
        }

        /// <summary>
        /// Get column span number of cell.
        /// </summary>
        /// <param name="cellId">Cell ID.</param>
        /// <returns></returns>
        public int GetColSpan(int cellId)
        {
            int colspan = 0;
            for (int i = 0; i < cells.Count; i++)
            {
                if (cells[i].cellId == cellId)
                {
                    colspan++;

                    int endOfRow = (((int)(i / this.ColumnLength)) + 1) * this.ColumnLength;
                    while (i + 1 < endOfRow)
                    {
                        i++;
                        if (cells[i].cellId == cellId)
                            colspan++;
                        else
                            break;
                    }
                    break;
                }
            }
            return colspan;
        }

        /// <summary>
        /// Get row span number of cell.
        /// </summary>
        /// <param name="cellId">Cell ID.</param>
        /// <returns></returns>
        public int GetRowSpan(int cellId)
        {
            int rowspan = 0;
            for (int i = 0; i < cells.Count; i++)
            {
                if (cells[i].cellId == cellId)
                {
                    rowspan++;

                    while (i + this.ColumnLength < cells.Count)
                    {
                        i += this.ColumnLength;
                        if (cells[i].cellId == cellId)
                            rowspan++;
                        else
                            break;
                    }
                    break;
                }
            }
            return rowspan;
        }

        /// <summary>
        /// Get cell object by position (row x column).
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Column number.</param>
        /// <returns></returns>
        public TableHelperCell GetCell(int row, int col)
        {
            int index = (row * this.ColumnLength) + col;
            return (index < this.cells.Count) ? this.cells[index] : null;
        }

        /// <summary>
        /// Get cell object by cellId.
        /// </summary>
        /// <param name="cellId">Cell Id</param>
        /// <returns></returns>
        public TableHelperCell GetCellByCellId(int cellId)
        {
            List<TableHelperCell> cellsInRow = this.cells.FindAll(c => c.cellId == cellId).OrderBy(o => o.rowId).ToList();
            if (cellsInRow.Count > 0)
                return cellsInRow.FindAll(c => c.rowId == cellsInRow[0].rowId).OrderBy(o => o.colId).ToList()[0];
            else
                return null;
        }

        /// <summary>
        /// Return useful cells (i.e. the cell has TableCell and iTSText.IElement elements).
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            if (this.cells.Count > 0)
            {
                int maxId = this.cells[this.cells.Count - 1].cellId + 1;
                int id = 0;
                while (id < maxId)
                {
                    foreach (TableHelperCell cell in this.cells)
                    {
                        if (cell.cellId == id)
                        {
                            id++;
                            yield return cell;
                        }
                    }
                }
            }
        }

        private float getCellWidth(Word.TableCell cell)
        {
            float ret = 0f;

            if (cell == null)
                return ret;

            Word.TableCellWidth width = this.styleHelper.GetAppliedElement<Word.TableCellWidth>(cell);
            if (width != null && width.Type != null)
            {
                switch (width.Type.Value)
                {
                    // TODO: don't know how to calculate auto cell width
                    case Word.TableWidthUnitValues.Nil:
                    case Word.TableWidthUnitValues.Auto: // fits the contents
                    default:
                        Word.Table table = cell.Parent.Parent as Word.Table;
                        Word.TableWidth tableWidth = this.styleHelper.GetAppliedElement<Word.TableWidth>(table);
                        if (tableWidth != null && tableWidth.Type != null)
                        {
                            if (tableWidth.Type.Value == Word.TableWidthUnitValues.Dxa)
                                ret = Tools.ConvertToPoint(tableWidth.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        }
                        if (ret <= 0f)
                            ret = this.getTableGridCols(table).Sum();
                        break;
                    case Word.TableWidthUnitValues.Dxa:
                        if (cell != null)
                            ret = Tools.ConvertToPoint(width.Width.Value, Tools.SizeEnum.TwentiethsOfPoint, -1f);
                        break;
                    case Word.TableWidthUnitValues.Pct:
                        if (width.Width != null)
                        {
                            if (cell.Parent.Parent.Parent.GetType() == typeof(Word.Body)) // use page size as width
                                ret = this.styleHelper.PrintablePageWidth * Tools.Percentage(width.Width.Value);
                            else // still in table, get cell width
                                ret = this.getCellWidth(cell.Parent.Parent.Parent as Word.TableCell) * Tools.Percentage(width.Width.Value);
                        }
                        break;
                }
            }

            return ret;
        }
    }
}
