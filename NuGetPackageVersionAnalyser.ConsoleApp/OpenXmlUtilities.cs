using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public static class OpenXmlUtilities
{
    public static Cell CreateCell(string? text) => new Cell
    {
        DataType = CellValues.String,
        CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty)
    };

    public static WorksheetPart? AutoSize(WorksheetPart oldSheetPart)
    {
        if (oldSheetPart is null)
            return null;

        SheetData sheetData = oldSheetPart.Worksheet.GetFirstChild<SheetData>();

        if (sheetData is null)
            return null;

        var maxColWidth = GetMaxCharacterWidth(sheetData);

        Columns columns = new Columns();
        //this is the width of my font - yours may be different
        double maxWidth = 7;
        foreach (var item in maxColWidth)
        {
            //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

            Column col = new Column() { BestFit = true,Min = (UInt32)(item.Key + 1),Max = (UInt32)(item.Key + 1),CustomWidth = true,Width = (DoubleValue)width };
            columns.Append(col);
        }

        oldSheetPart.Worksheet.RemoveAllChildren();
        oldSheetPart.Worksheet.Append(columns);
        oldSheetPart.Worksheet.Append(sheetData);
        return oldSheetPart;
    }


    private static Dictionary<int,int> GetMaxCharacterWidth(SheetData sheetData)
    {
        //iterate over all cells getting a max char value for each column
        Dictionary<int,int> maxColWidth = new Dictionary<int,int>();
        var rows = sheetData.Elements<Row>();
        foreach (var r in rows)
        {
            var cells = r.Elements<Cell>().ToArray();

            //using cell index as my column
            for (int i = 0; i < cells.Length; i++)
            {
                var cell = cells[i];
                var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                var cellTextLength = cellValue.Length;

                if (maxColWidth.ContainsKey(i))
                {
                    var current = maxColWidth[i];
                    if (cellTextLength > current)
                    {
                        maxColWidth[i] = cellTextLength;
                    }
                }
                else
                {
                    maxColWidth.Add(i,cellTextLength);
                }
            }
        }

        return maxColWidth;
    }

    static uint GetColumnIndex(string cellReference)
    {
        uint columnIndex = 0;
        if (cellReference is not null)
        {
            foreach (char ch in cellReference)
            {
                if (char.IsLetter(ch))
                {
                    columnIndex = (uint)(ch - 'A' + 1) + columnIndex * 26;
                }
                else
                {
                    break;
                }
            }
        }
        return columnIndex;
    }
}
