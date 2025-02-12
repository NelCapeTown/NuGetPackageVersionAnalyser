using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace NuGetPackageVersionAnalyser.ConsoleApp;
public static class ReportExcelFlatList
{
    public static void CreateExcelReport(List<NuGetPackageInfo> packages,string fileName)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Create(fileName,SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Packages"
        };
        sheets.Append(sheet);

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        var headerRow = new Row();
        headerRow.Append(
            CreateCell("Project Name"),
            CreateCell("Package Name"),
            CreateCell("IsTransitive"),
            CreateCell("Requested Version"),
            CreateCell("Resolved Version")
        );
        sheetData.Append(headerRow);

        foreach (var package in packages)
        {
            var row = new Row();
            row.Append(
                CreateCell(package.ProjectName),
                CreateCell(package.PackageName),
                CreateCell(package.IsTransitive),
                CreateCell(package.RequestedVersion),
                CreateCell(package.ResolvedVersion)
            );
            sheetData.Append(row);
        }

        AutoFitColumns(worksheetPart);

        workbookPart.Workbook.Save();
    }

    static Cell CreateCell(string? text)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = text != null ? new CellValue(text) : new CellValue(string.Empty)
        };
    }

    static void AutoFitColumns(WorksheetPart worksheetPart)
    {
        var columns = new Columns();
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var maxColumnIndex = sheetData.Elements<Row>().SelectMany(r => r.Elements<Cell>())
            .Max(c => GetColumnIndex(c.CellReference));

        for (uint i = 1; i <= maxColumnIndex; i++)
        {
            var column = new Column
            {
                Min = i,
                Max = i,
                Width = 15, // Default width, can be adjusted
                CustomWidth = true
            };
            columns.Append(column);
        }

        worksheetPart.Worksheet.InsertAt(columns,0);
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
