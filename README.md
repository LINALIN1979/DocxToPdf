# DocxToPdf

C# library for converting *.docx* to *.pdf* without Office.

## Prerequisites ##

- .NET Framework 4
- Install [iTextSharp v5](http://sourceforge.net/projects/itextsharp/)
- Install [Open XML SDK 2.5](https://www.microsoft.com/en-us/download/details.aspx?id=30425)

## DOCX Support Status ##

Below lists current support status of DOCX

<table>
 <tr>
  <td bgcolor="gray"><font color="white"><b>Type</b></font></td>
  <td bgcolor="gray"><font color="white"><b>Support</b></font></td>
  <td bgcolor="gray"><font color="white"><b>Remarks</b></font></td>
 </tr>
 <tr>
  <td colspan="3" bgcolor="black"><font color="white"><b>Run</b></font></td>
 </tr>
 <tr>
  <td>Borders</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Fonts</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Underline only supports single-line<br>2. Theme color is not supported</td>
 </tr>
 <tr>
  <td>Shading</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Line breaks</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Column break is not supported</td>
 </tr>
 <tr>
  <td>Kerning and spacing</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Preserved space</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Symbols</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Hyperlinks</td>
  <td align="center"><font color="blue">V</font></td>
  <td>i.e. external links</td>
 </tr>
 <tr>
  <td>Bookmarks</td>
  <td align="center"><font color="red">X</font></td>
  <td>i.e. internal links</td>
 </tr>
 <tr>
  <td>Picture</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>

 <tr>
  <td colspan="3" bgcolor="black"><font color="white"><b>Paragraph</b></font></td>
 </tr>
 <tr>
  <td>Horizontal alignment</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Borders</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>AutoSapceDN</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>AutoSapceDE</td>
  <td align="center"><font color="blue">V</font></td>
  <td></tr>
 <tr>
  <td>Indentation</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Shading</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Paragraph spacing</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Tabs</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Vertical alignment</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Numbering & listing</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Restart is not supported<br>2. Picture as numbering symbol is not supported<br>3. Numbering only supports Decimal, DecimalZero, LowerRoman, UpperRoman, TaiwaneseCountingThousand<br>4. Linked style is not supported</td>
 </tr>
 <tr>
  <td>Text direction</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>

 <tr>
  <td colspan="3" bgcolor="black"><font color="white"><b>Table</b></font></td>
 </tr>
 <tr>
  <td>Borders</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Border only supports signle-line</td>
 </tr>
 <tr>
  <td>Border conflicts</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Caption</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Table header</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Row height</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Table indentation</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Floating table</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Cell margins</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Cell spacing</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Cell shading</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Vertical alignment</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Fixed table width</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Nested table size may be wrong</td>
 </tr>
 <tr>
  <td>Autofit table width</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Conditional formatting</td>
  <td align="center"><font color="red">X</font></td>
  <td>i.e. formatting for firstColumn, firstRow, lastColumn, lastRow, noHBand, noVBand</td>
 </tr>
 <tr>
  <td>Text direction</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>

 <tr>
  <td colspan="3" bgcolor="black"><font color="white"><b>Sections</b></font></td>
 </tr>
 <tr>
  <td>Type</td>
  <td align="center"><font color="red">X</font></td>
  <td>Five types for section properties: continuous, evenPage, oddPage, nextPage, and nextColumn</td>
 </tr>
 <tr>
  <td>Column</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Borders</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Header</td>
  <td align="center"><font color="blue">V</font></td>
  <td>Known issues<br>1. Page number is not supported</td>
 </tr>
 <tr>
  <td>Footer</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Page margins</td>
  <td align="center"><font color="blue">V</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Page number</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Text direction</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>

 <tr>
  <td colspan="3" bgcolor="black"><font color="white"><b>Others</b></font></td>
 </tr>
 <tr>
  <td>Text frames</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>Auto color</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
 <tr>
  <td>TOC</td>
  <td align="center"><font color="red">X</font></td>
  <td></td>
 </tr>
</table>

## Useful Links ##

- [Office Open XML](http://officeopenxml.com/WPcontentOverview.php) - WordprocessingML documentation
- [Office Dev Center](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.aspx) - Open XML SDK API documentation

