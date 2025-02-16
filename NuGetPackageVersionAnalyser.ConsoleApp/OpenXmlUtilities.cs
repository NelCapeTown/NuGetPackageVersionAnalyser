using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using X = DocumentFormat.OpenXml.Spreadsheet;
using SkiaSharp;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NuGetPackageVersionAnalyser.ConsoleApp;

public static class OpenXmlUtilities
{
    public static void CreatePackage(string pathToFile, string sheetName,X.SheetData xSheetData)
    {
        using (var pkg = SpreadsheetDocument.Create(pathToFile,SpreadsheetDocumentType.Workbook))
        {
            CreateParts(pkg,sheetName,xSheetData);
        }
    }
    private static void CreateParts(SpreadsheetDocument pkg,string sheetName,X.SheetData xSheetData)
    {
        var workbookPart = pkg.AddWorkbookPart();
        GenerateWorkbookPart(workbookPart,sheetName);
        workbookPart.Workbook.Save();

        var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("R0b3c0d8bb108422b");
        GenerateWorkbookStylesPart(workbookStylesPart);
        workbookStylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("R5957bc4f19734625");
        GenerateWorksheetPart(worksheetPart,xSheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void GenerateWorkbookPart(WorkbookPart part,string sheetName)
    {
        var xWorkbook = new X.Workbook();
        xWorkbook.AddNamespaceDeclaration("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        var xSheets = new X.Sheets();
        var xSheet = new X.Sheet
        {
            Name = sheetName,
            SheetId = 1u,
            Id = "R5957bc4f19734625"
        };
        xSheet.AddNamespaceDeclaration("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        xSheets.Append(xSheet);
        xWorkbook.Append(xSheets);

        part.Workbook = xWorkbook;
    }

    private static void GenerateWorksheetPart(WorksheetPart part,X.SheetData xSheetData)
    {
        var xWorksheet = new X.Worksheet();
        xWorksheet.AddNamespaceDeclaration("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        xWorksheet.Append(xSheetData);
        part.Worksheet = xWorksheet;

        X.Columns columns = CalculateColumnWidths(xSheetData);

        // Remove existing columns
        var existingColumns = part.Worksheet.GetFirstChild<X.Columns>();
        if (existingColumns != null)
        {
            part.Worksheet.RemoveChild(existingColumns);
        }

        part.Worksheet.InsertAt(columns,0);
    }

    private static void AddHeaderRow(X.SheetData sheetData)
    {
        var headerRow = new X.Row();
        AddCell(headerRow,"Project Name",true);
        AddCell(headerRow,"Package Name",true);
        AddCell(headerRow,"Is Transitive",true);
        AddCell(headerRow,"Requested Version",true);
        AddCell(headerRow,"Resolved Version",true);
        sheetData.Append(headerRow);
    }

    private static void AddDataRows(X.SheetData sheetData,List<NuGetPackageInfo> packages)
    {
        foreach (var package in packages)
        {
            var dataRow = new X.Row();
            AddCell(dataRow,package.ProjectName);
            AddCell(dataRow,package.PackageName);
            AddCell(dataRow,package.IsTransitive);
            AddCell(dataRow,package.RequestedVersion);
            AddCell(dataRow,package.ResolvedVersion);
            sheetData.Append(dataRow);
        }
    }

    private static void AddCell(X.Row row,string? text,bool isHeader = false)
    {
        var cell = new X.Cell
        {
            DataType = X.CellValues.String,
            CellValue = new X.CellValue(text)
        };
        if (isHeader)
        {
            cell.StyleIndex = 1u; // Assuming style index 1 is for header
        }
        row.Append(cell);
    }


    private static float GetTextWidth(string text,SKFont typeface,float fontSize)
    {
        return typeface.MeasureText(text);
    }

    private static Columns CalculateColumnWidths(SheetData sheetData)
    {
        var maxColWidth = new Dictionary<int,double>();
        var typeface = SKTypeface.FromFamilyName("Calibri");
        float fontSize = 11;
        SKFont font = new SKFont(typeface,fontSize);

        foreach (var row in sheetData.Elements<Row>())
        {
            int colIndex = 0;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellText = cell.CellValue?.Text ?? string.Empty;
                var cellWidth = GetTextWidth(cellText,font,fontSize) / 7; // Convert pixel width to Excel column width

                if (maxColWidth.ContainsKey(colIndex))
                {
                    if (cellWidth > maxColWidth[colIndex])
                    {
                        maxColWidth[colIndex] = cellWidth;
                    }
                }
                else
                {
                    maxColWidth[colIndex] = cellWidth;
                }

                colIndex++;
            }
        }

        Columns columns = new Columns();
        foreach (var item in maxColWidth)
        {
            double width = Math.Truncate((item.Value + 5) / 7 * 256) / 256;
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

    private static void GenerateWorkbookStylesPart(WorkbookStylesPart part)
    {
        var xStylesheet = new X.Stylesheet();
        xStylesheet.AddNamespaceDeclaration("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        var xFonts = new X.Fonts();

        var xFont = new X.Font();
        var xFontSize = new X.FontSize { Val = 11D };
        var xFontName = new X.FontName { Val = "Calibri" };
        xFont.Append(xFontSize);
        xFont.Append(xFontName);
        xFonts.Append(xFont);

        xFont = new X.Font();
        xFontSize = new X.FontSize { Val = 11D };
        xFontName = new X.FontName { Val = "Calibri" };
        var xBold = new X.Bold();
        xFont.Append(xFontSize);
        xFont.Append(xFontName);
        xFont.Append(xBold);
        xFonts.Append(xFont);

        xStylesheet.Append(xFonts);

        var xFills = new X.Fills();
        var xFill = new X.Fill();
        var xPatternFill = new X.PatternFill { PatternType = X.PatternValues.None };
        xFill.Append(xPatternFill);
        xFills.Append(xFill);

        xFill = new X.Fill();
        xPatternFill = new X.PatternFill { PatternType = X.PatternValues.Solid };
        xPatternFill.ForegroundColor = new X.ForegroundColor { Rgb = "FFFF7F50" }; // Coral color
        xPatternFill.BackgroundColor = new X.BackgroundColor { Indexed = 64u };
        xFill.Append(xPatternFill);
        xFills.Append(xFill);

        xFill = new X.Fill();
        xPatternFill = new X.PatternFill { PatternType = X.PatternValues.Solid };
        xPatternFill.ForegroundColor = new X.ForegroundColor { Rgb = "FF008080" }; // Teal color
        xPatternFill.BackgroundColor = new X.BackgroundColor { Indexed = 64u };
        xFill.Append(xPatternFill);
        xFills.Append(xFill);

        xStylesheet.Append(xFills);

        var xBorders = new X.Borders();

        // Default border (no borders)
        var xDefaultBorder = new X.Border();
        xDefaultBorder.Append(new X.LeftBorder());
        xDefaultBorder.Append(new X.RightBorder());
        xDefaultBorder.Append(new X.TopBorder());
        xDefaultBorder.Append(new X.BottomBorder());
        xDefaultBorder.Append(new X.DiagonalBorder());
        xBorders.Append(xDefaultBorder);

        // Border with all edges
        var xBorder = new X.Border();
        xBorder.Append(new X.LeftBorder { Style = X.BorderStyleValues.Thin });
        xBorder.Append(new X.RightBorder { Style = X.BorderStyleValues.Thin });
        xBorder.Append(new X.TopBorder { Style = X.BorderStyleValues.Thin });
        xBorder.Append(new X.BottomBorder { Style = X.BorderStyleValues.Thin });
        xBorder.Append(new X.DiagonalBorder());
        xBorders.Append(xBorder);

        xStylesheet.Append(xBorders);

        var xCellStyleFormats = new X.CellStyleFormats();
        var xCellFormat = new X.CellFormat { FontId = 0U,FillId = 0U,BorderId = 0U };
        xCellStyleFormats.Append(xCellFormat);
        xStylesheet.Append(xCellStyleFormats);

        var xCellFormats = new X.CellFormats();
        xCellFormat = new X.CellFormat { FontId = 0U,FillId = 0U,BorderId = 0U,ApplyFont = true };
        xCellFormats.Append(xCellFormat);

        xCellFormat = new X.CellFormat { FontId = 1U,FillId = 0U,BorderId = 0U,ApplyFont = true };
        xCellFormats.Append(xCellFormat);

        xStylesheet.Append(xCellFormats);

        part.Stylesheet = xStylesheet;
    }


}
