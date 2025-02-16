using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;
using SkiaSharp;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NuGetPackageVersionAnalyser.ConsoleApp;

public static class OpenXmlUtilities
{
    public static SpreadsheetDocument? CreatePackage(string pathToFile)
    {
        try
        {
            var pkg = SpreadsheetDocument.Create(pathToFile,SpreadsheetDocumentType.Workbook);
            var workbookPart = pkg.AddWorkbookPart();
            workbookPart.Workbook = new X.Workbook();

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("R0b3c0d8bb108422b");
            GenerateWorkbookStylesPart(workbookStylesPart);
            workbookStylesPart.Stylesheet.Save();

            return pkg;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.CreatePackage: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return null;
        }
    }

    public static void AddSheetData(SpreadsheetDocument pkg,string sheetName,X.SheetData xSheetData)
    {
        try
        {
            var workbookPart = pkg.WorkbookPart;
            if (workbookPart == null)
            {
                throw new InvalidOperationException("WorkbookPart is null.");
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new X.Worksheet(xSheetData);

            var sheets = workbookPart.Workbook.GetFirstChild<X.Sheets>() ?? workbookPart.Workbook.AppendChild(new X.Sheets());
            var sheet = new X.Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)(sheets.Count() + 1),
                Name = sheetName
            };
            sheets.Append(sheet);

            AutoSize(worksheetPart);
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.AddSheetData: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static void GenerateWorkbookStylesPart(WorkbookStylesPart part)
    {
        try
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
            xPatternFill.BackgroundColor = new X.BackgroundColor { Rgb = "FFFF7F50" }; // Coral color
            xFill.Append(xPatternFill);
            xFills.Append(xFill);

            xFill = new X.Fill();
            xPatternFill = new X.PatternFill { PatternType = X.PatternValues.Solid };
            xPatternFill.ForegroundColor = new X.ForegroundColor { Rgb = "FF008080" }; // Teal color
            xPatternFill.BackgroundColor = new X.BackgroundColor { Rgb = "FF008080" }; // Teal color
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

            xCellFormat = new X.CellFormat { FontId = 1U,FillId = 1U,BorderId = 0U,ApplyFont = true,
                ApplyFill = true };
            xCellFormats.Append(xCellFormat);

            xStylesheet.Append(xCellFormats);

            part.Stylesheet = xStylesheet;

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.GenerateWorkbookStylesPart: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    public static void AddCell(X.Row row,string? text,bool isHeader = false)
    {
        try
        {
            var cell = new X.Cell
            {
                DataType = X.CellValues.String,
                CellValue = new X.CellValue(text ?? string.Empty)
            };
            if (isHeader)
            {
                cell.StyleIndex = 1u; // Assuming style index 1 is for header
            }
            row.Append(cell);

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.AddCell: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static float GetTextWidth(string text,SKFont font)
    {
        try
        {
            return font.MeasureText(text);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.GetTextWidth: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return 0;
        }
    }


    private static X.Columns CalculateColumnWidths(X.SheetData sheetData)
    {
        try
        {
            var maxColWidth = new Dictionary<int,double>();
            var typeface = SKTypeface.FromFamilyName("Calibri");
            var boldTypeface = SKTypeface.FromFamilyName("Calibri",SKFontStyle.Bold);
            float fontSize = 11;
            SKFont font = new SKFont(typeface,fontSize);
            SKFont boldFont = new SKFont(boldTypeface,fontSize);

            foreach (var row in sheetData.Elements<X.Row>())
            {
                int colIndex = 0;
                foreach (var cell in row.Elements<X.Cell>())
                {
                    var cellText = cell.CellValue?.Text ?? string.Empty;
                    var isHeader = cell.StyleIndex != null && cell.StyleIndex == 1u;
                    var cellWidth = (GetTextWidth(cellText,isHeader ? boldFont : font) * 1.4) ;

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

            X.Columns columns = new X.Columns();
            foreach (var item in maxColWidth)
            {
                double width = Math.Truncate((item.Value + 10) / 7 * 256) / 256; // Add a more generous buffer
                Console.WriteLine($"Column {item.Key + 1}: Width = {width}"); // Debug statement
                columns.Append(new X.Column()
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
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.CalculateColumnWidths: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return new X.Columns();
        }
    }


    private static void AutoSize(WorksheetPart worksheetPart)
    {
        try
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<X.SheetData>();
            if (sheetData == null)
            {
                throw new InvalidOperationException("SheetData is null.");
            }

            var columns = CalculateColumnWidths(sheetData);
            if (columns == null)
            {
                Console.WriteLine("No columns calculated."); // Debug statement
                return;
            }

            var existingColumns = worksheetPart.Worksheet.GetFirstChild<X.Columns>();
            if (existingColumns != null)
            {
                worksheetPart.Worksheet.RemoveChild(existingColumns);
            }
            worksheetPart.Worksheet.InsertAt(columns,0);
            worksheetPart.Worksheet.Save(); // Ensure worksheet is saved after adjusting columns

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.AutoSize: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}


