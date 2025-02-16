using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public static class ReportExcelFlatList
{
    public static void CreateExcelReport(List<NuGetPackageInfo> packages,string fileName)
    {
        try
        {
            var pkg = OpenXmlUtilities.CreatePackage(fileName);
            var sheetData = CreateSheetData(packages);
            OpenXmlUtilities.AddSheetData(pkg,"Packages",sheetData);
            pkg.Dispose();

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in ReportExcelFlatList.CreateExcelReport: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static X.SheetData? CreateSheetData(List<NuGetPackageInfo> packages)
    {
        try
        {
            var sheetData = new X.SheetData();

            var headerRow = new X.Row();
            OpenXmlUtilities.AddCell(headerRow,"Project Name",true);
            OpenXmlUtilities.AddCell(headerRow,"Package Name",true);
            OpenXmlUtilities.AddCell(headerRow,"Is Transitive",true);
            OpenXmlUtilities.AddCell(headerRow,"Requested Version",true);
            OpenXmlUtilities.AddCell(headerRow,"Resolved Version",true);
            sheetData.Append(headerRow);

            foreach (var package in packages)
            {
                var row = new X.Row();
                OpenXmlUtilities.AddCell(row,package.ProjectName);
                OpenXmlUtilities.AddCell(row,package.PackageName);
                OpenXmlUtilities.AddCell(row,package.IsTransitive);
                OpenXmlUtilities.AddCell(row,package.RequestedVersion);
                OpenXmlUtilities.AddCell(row,package.ResolvedVersion);
                sheetData.Append(row);
            }

            return sheetData;

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in ReportExcelFlatList.CreateSheetData: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return null;
        }
    }
}

