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
    private static Fonts CreateFonts()
    {
        return new Fonts(
            new Font( // Index 0 - Normal Font
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" }
            ),
            new Font( // Index 1 - Bold Font
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" },
                new Bold()
            ),
            new Font( // Index 2 - Italic Font
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" },
                new Italic()
            ),
            new Font( // Index 3 - Bold Italic Font
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" },
                new Bold(),
                new Italic()
            ),
            new Font( // Index 4 - 20 point Font
                new FontSize { Val = 20 },
                new FontName { Val = "Calibri" }
            ),
            new Font( // Index 5 - 20 point Bold Font
                new FontSize { Val = 20 },
                new FontName { Val = "Calibri" },
                new Bold()
            )
        );
    }

    private static Fills CreateFills()
    {
        return new Fills(
            new Fill( // Index 0 - Default Fill (No Fill)
                new PatternFill { PatternType = PatternValues.None }
            ),
            new Fill( // Index 1 - Coral Fill
                new PatternFill(
                    new ForegroundColor { Rgb = new HexBinaryValue { Value = "FF7F50" } },
                    new BackgroundColor { Indexed = 64 }
                )
                {
                    PatternType = PatternValues.Solid
                }
            )
        );
    }

    private static CellFormats CreateCellFormats()
    {
        return new CellFormats(
            new CellFormat { FontId = 0,FillId = 0,ApplyFont = true }, // Index 0 - Normal Font, Default Fill
            new CellFormat { FontId = 1,FillId = 0,ApplyFont = true }, // Index 1 - Bold Font, Default Fill
            new CellFormat { FontId = 2,FillId = 0,ApplyFont = true }, // Index 2 - Italic Font, Default Fill
            new CellFormat { FontId = 3,FillId = 0,ApplyFont = true }, // Index 3 - Bold Italic Font, Default Fill
            new CellFormat { FontId = 4,FillId = 0,ApplyFont = true }, // Index 4 - 20pt Font, Default Fill
            new CellFormat { FontId = 5,FillId = 0,ApplyFont = true }, // Index 5 - 20pt Bold Font, Default Fill

            new CellFormat { FontId = 0,FillId = 1,ApplyFont = true,ApplyFill = true }, // Index 6 - Normal Font, Coral Fill
            new CellFormat { FontId = 1,FillId = 1,ApplyFont = true,ApplyFill = true }, // Index 7 - Bold Font, Coral Fill
            new CellFormat { FontId = 2,FillId = 1,ApplyFont = true,ApplyFill = true }, // Index 8 - Italic Font, Coral Fill
            new CellFormat { FontId = 3,FillId = 1,ApplyFont = true,ApplyFill = true }, // Index 9 - Bold Italic Font, Coral Fill
            new CellFormat { FontId = 4,FillId = 1,ApplyFont = true,ApplyFill = true }, // Index 10 - 20pt Font, Coral Fill
            new CellFormat { FontId = 5,FillId = 1,ApplyFont = true,ApplyFill = true }  // Index 11 - 20pt Bold Font, Coral Fill
        );
    }

    public static Stylesheet CreateStylesheet()
    {
        return new Stylesheet(
            CreateFonts(),
            CreateFills(),
            CreateCellFormats()
        );
    }

    public static Cell CreateCell(string? text) => new Cell
    {
        DataType = CellValues.String,
        CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty)
    };

    public static Cell CreateBoldCell(string? text) => new Cell
    {
        DataType = CellValues.String,
        CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty),
        StyleIndex = UInt32Value.FromUInt32(1)
    };

    public static Cell CreateBoldCoralCell(string? text) => new Cell
    {
        DataType = CellValues.String,
        CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty),
        StyleIndex = UInt32Value.FromUInt32(7)
    };

    public static Cell CreateCell(string? text,int styleIndex)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty),
            StyleIndex = UInt32Value.FromUInt32((uint)styleIndex)
        };
    }

    public static void AddStyles(WorksheetPart worksheetPart)
    {
        var stylesheet = CreateStylesheet();
        var workbookStylesPart = worksheetPart.AddNewPart<WorkbookStylesPart>();
        workbookStylesPart.Stylesheet = stylesheet;
        workbookStylesPart.Stylesheet.Save();
    }

    public static WorksheetPart? AutoSize(WorksheetPart oldSheetPart)
    {
        if (oldSheetPart == null)
            return null;

        var sheetData = oldSheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
            return null;

        var columns = CalculateColumnWidths(sheetData);

        var existingColumns = oldSheetPart.Worksheet.GetFirstChild<Columns>();
        if (existingColumns != null)
            oldSheetPart.Worksheet.RemoveChild(existingColumns);

        oldSheetPart.Worksheet.InsertAt(columns,0);
        return oldSheetPart;
    }

    private static Columns CalculateColumnWidths(SheetData sheetData)
    {
        var maxColWidth = GetMaxCharacterWidth(sheetData);
        Columns columns = new();

        double maxWidth = 7.0; // Adjust this for different fonts
        foreach (var item in maxColWidth)
        {
            double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
            columns.Append(new Column()
            {
                Min = (uint)item.Key + 1,
                Max = (uint)item.Key + 1,
                Width = width,
                CustomWidth = true,
                BestFit = true
            });
        }
        return columns;
    }

    private static Dictionary<int,int> GetMaxCharacterWidth(SheetData sheetData)
    {
        Dictionary<int,int> maxColWidth = new();

        foreach (var row in sheetData.Elements<Row>())
        {
            var cells = row.Elements<Cell>().ToArray();
            for (int i = 0; i < cells.Length; i++)
            {
                int cellTextLength = cells[i].CellValue?.InnerText.Length ?? 0;
                if (maxColWidth.TryGetValue(i,out int currentMax))
                {
                    maxColWidth[i] = Math.Max(currentMax,cellTextLength);
                }
                else
                {
                    maxColWidth[i] = cellTextLength;
                }
            }
        }
        return maxColWidth;
    }
}
