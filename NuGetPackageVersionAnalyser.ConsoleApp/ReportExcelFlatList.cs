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
            OpenXmlUtilities.CreateCell("Project Name"),
            OpenXmlUtilities.CreateCell("Package Name"),
            OpenXmlUtilities.CreateCell("IsTransitive"),
            OpenXmlUtilities.CreateCell("Requested Version"),
            OpenXmlUtilities.CreateCell("Resolved Version")
        );
        sheetData.Append(headerRow);

        foreach (var package in packages)
        {
            var row = new Row();
            row.Append(
                OpenXmlUtilities.CreateCell(package.ProjectName),
                OpenXmlUtilities.CreateCell(package.PackageName),
                OpenXmlUtilities.CreateCell(package.IsTransitive),
                OpenXmlUtilities.CreateCell(package.RequestedVersion),
                OpenXmlUtilities.CreateCell(package.ResolvedVersion)
            );
            sheetData.Append(row);
        }

        OpenXmlUtilities.AutoSize(worksheetPart);

        workbookPart.Workbook.Save();
    }
}
