using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Office2013.Theme;
using A = DocumentFormat.OpenXml.Drawing;
using VT = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using SkiaSharp;

namespace NuGetPackageVersionAnalyser.ConsoleApp;

public static class OpenXmlUtilities
{
    public static SpreadsheetDocument? CreatePackage(string pathToFile)
    {
        try
        {
            var pkg = SpreadsheetDocument.Create(pathToFile, SpreadsheetDocumentType.Workbook);
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

    public static void AddSheetData(SpreadsheetDocument pkg, string sheetName, X.SheetData xSheetData)
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
            xStylesheet.AddNamespaceDeclaration("mc","http://schemas.openxmlformats.org/markup-compatibility/2006");
            xStylesheet.AddNamespaceDeclaration("x14ac","http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            xStylesheet.AddNamespaceDeclaration("x16r2","http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            xStylesheet.AddNamespaceDeclaration("xr","http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            var markupCompatibilityAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac x16r2 xr" };
            xStylesheet.MCAttributes = markupCompatibilityAttributes;

            var xFonts = new X.Fonts { Count = 2u,KnownFonts = true };

            var xFont = new X.Font();
            var xFontSize = new X.FontSize { Val = 11D };
            var xFontName = new X.FontName { Val = "Calibri" };
            xFont.Append(xFontSize);
            xFont.Append(xFontName);
            xFonts.Append(xFont);

            xFont = new X.Font();
            var xBold = new X.Bold();
            xFont.Append(xBold);
            xFontSize = new X.FontSize { Val = 11D };
            xFont.Append(xFontSize);
            xFontName = new X.FontName { Val = "Calibri" };
            xFont.Append(xFontName);
            xFonts.Append(xFont);

            xStylesheet.Append(xFonts);

            var xFills = new X.Fills { Count = 3u };

            var xFill = new X.Fill();
            var xPatternFill = new X.PatternFill { PatternType = X.PatternValues.None };
            xFill.Append(xPatternFill);
            xFills.Append(xFill);

            xFill = new X.Fill();
            xPatternFill = new X.PatternFill { PatternType = X.PatternValues.Gray125 };
            xFill.Append(xPatternFill);
            xFills.Append(xFill);

            xFill = new X.Fill();
            xPatternFill = new X.PatternFill { PatternType = X.PatternValues.Solid };
            var xForegroundColor = new X.ForegroundColor { Rgb = "FFFF7F50" };
            xPatternFill.Append(xForegroundColor);
            var xBackgroundColor = new X.BackgroundColor { Rgb = "FFFF7F50" };
            xPatternFill.Append(xBackgroundColor);
            xFill.Append(xPatternFill);
            xFills.Append(xFill);

            xStylesheet.Append(xFills);

            var xBorders = new X.Borders { Count = 1u };

            var xBorder = new X.Border();
            var xLeftBorder = new X.LeftBorder();
            xBorder.Append(xLeftBorder);
            var xRightBorder = new X.RightBorder();
            xBorder.Append(xRightBorder);
            var xTopBorder = new X.TopBorder();
            xBorder.Append(xTopBorder);
            var xBottomBorder = new X.BottomBorder();
            xBorder.Append(xBottomBorder);
            var xDiagonalBorder = new X.DiagonalBorder();
            xBorder.Append(xDiagonalBorder);
            xBorders.Append(xBorder);

            xStylesheet.Append(xBorders);

            var xCellStyleFormats = new X.CellStyleFormats { Count = 1u };
            var xCellFormat = new X.CellFormat { NumberFormatId = 0u,FontId = 0u,FillId = 0u,BorderId = 0u };
            xCellStyleFormats.Append(xCellFormat);
            xStylesheet.Append(xCellStyleFormats);

            var xCellFormats = new X.CellFormats { Count = 2u };
            xCellFormat = new X.CellFormat { NumberFormatId = 0u,FontId = 0u,FillId = 0u,BorderId = 0u,FormatId = 0u };
            xCellFormats.Append(xCellFormat);
            xCellFormat = new X.CellFormat { NumberFormatId = 0u,FontId = 1u,FillId = 2u,BorderId = 0u,FormatId = 0u,ApplyFont = true,ApplyFill = true };
            xCellFormats.Append(xCellFormat);
            xStylesheet.Append(xCellFormats);

            var xCellStyles = new X.CellStyles { Count = 1u };
            var xCellStyle = new X.CellStyle { Name = "Normal",FormatId = 0u,BuiltinId = 0u };
            xCellStyles.Append(xCellStyle);
            xStylesheet.Append(xCellStyles);

            var xDifferentialFormats = new X.DifferentialFormats { Count = 0u };
            xStylesheet.Append(xDifferentialFormats);

            var xTableStyles = new X.TableStyles { Count = 0u,DefaultTableStyle = "TableStyleMedium2",DefaultPivotStyle = "PivotStyleLight16" };
            xStylesheet.Append(xTableStyles);

            var xColors = new X.Colors();
            var xMruColors = new X.MruColors();
            var xColor = new X.Color { Rgb = "FFFF7F50" };
            xMruColors.Append(xColor);
            xColors.Append(xMruColors);
            xStylesheet.Append(xColors);

            var xStylesheetExtensionList = new X.StylesheetExtensionList();
            var xStylesheetExtension = new X.StylesheetExtension { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            xStylesheetExtension.AddNamespaceDeclaration("x14","http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            var x14SlicerStyles = new X14.SlicerStyles { DefaultSlicerStyle = "SlicerStyleLight1" };
            xStylesheetExtension.Append(x14SlicerStyles);
            xStylesheetExtensionList.Append(xStylesheetExtension);

            xStylesheetExtension = new X.StylesheetExtension { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            xStylesheetExtension.AddNamespaceDeclaration("x15","http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            var x15TimelineStyles = new X15.TimelineStyles { DefaultTimelineStyle = "TimeSlicerStyleLight1" };
            xStylesheetExtension.Append(x15TimelineStyles);
            xStylesheetExtensionList.Append(xStylesheetExtension);

            xStylesheet.Append(xStylesheetExtensionList);

            part.Stylesheet = xStylesheet;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.GenerateWorkbookStylesPart: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    public static void AddCell(X.Row row, string? text, bool isHeader = false)
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

    private static float GetTextWidth(string text, SKFont font)
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
            var maxColWidth = new Dictionary<int, double>();
            var typeface = SKTypeface.FromFamilyName("Calibri");
            var boldTypeface = SKTypeface.FromFamilyName("Calibri", SKFontStyle.Bold);
            float fontSize = 11;
            SKFont font = new SKFont(typeface, fontSize);
            SKFont boldFont = new SKFont(boldTypeface, fontSize);

            foreach (var row in sheetData.Elements<X.Row>())
            {
                int colIndex = 0;
                foreach (var cell in row.Elements<X.Cell>())
                {
                    var cellText = cell.CellValue?.Text ?? string.Empty;
                    var isHeader = cell.StyleIndex != null && cell.StyleIndex == 1u;
                    var cellWidth = (GetTextWidth(cellText, isHeader ? boldFont : font) * 1.4);

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
            worksheetPart.Worksheet.InsertAt(columns, 0);
            worksheetPart.Worksheet.Save(); // Ensure worksheet is saved after adjusting columns

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OpenXmlUtilities.AutoSize: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}

