using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public static class ReportExcelPackagesByProject
{
    public static void PreparePackagesByProject(List<NuGetPackageInfo> packages,string filePath)
    {
        // Generate summary table
        var distinctPackages = packages
            .GroupBy(p => new { p.PackageName,p.IsTransitive })
            .Select(g => new NuGetPackageVersion
            {
                PackageName = g.Key.PackageName,
                IsTransitive = g.Key.IsTransitive,
                RequestedVersions = g.Where(p => !string.IsNullOrEmpty(p.RequestedVersion))
                                     .Select(p => p.RequestedVersion)
                                     .Distinct()
                                     .OrderBy(v => v)
                                     .ToList(),
                ResolvedVersions = g.Select(p => new ProjectResolvedVersion
                {
                    ProjectName = p.ProjectName,
                    ResolvedVersion = p.ResolvedVersion
                }).ToList(),
            })
            .OrderBy(p => p.IsTransitive)
            .ThenBy(p => p.PackageName)
            .ToList();

        var projects = packages
            .Select(p => p.ProjectName)
            .Distinct()
            .OrderBy(p => p)
            .ToList();

        CreateExcelReport((List<NuGetPackageVersion>)distinctPackages,projects,filePath);
    }

    private static void CreateExcelReport(List<NuGetPackageVersion> distinctPackages,List<string> projects,string fileName)
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
            Name = "NuGet Summary"
        };
        OpenXmlUtilities.AddStyles(worksheetPart);
        sheets.Append(sheet);

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        // Add headers
        var headerRow = new Row();
        headerRow.Append(OpenXmlUtilities.CreateBoldCoralCell("Package Name"),OpenXmlUtilities.CreateBoldCoralCell("Is Transitive"),OpenXmlUtilities.CreateBoldCoralCell("Requested Version"));
        foreach (var project in projects)
        {
            headerRow.Append(OpenXmlUtilities.CreateBoldCoralCell(project));
        }
        sheetData.Append(headerRow);

        // Add data rows
        foreach (var package in distinctPackages)
        {
            var row = new Row();
            row.Append(OpenXmlUtilities.CreateCell(package.PackageName),OpenXmlUtilities.CreateCell(package.IsTransitive));

            // Handle requested versions
            if (package.IsTransitive == "No" && package.RequestedVersions.Count > 1)
            {
                var concatenatedVersions = string.Join(", ",package.RequestedVersions);
                var highlightedCell = OpenXmlUtilities.CreateCell(concatenatedVersions);
                row.Append(highlightedCell);
            }
            else
            {
                row.Append(OpenXmlUtilities.CreateCell(package.IsTransitive == "No" ? string.Join(", ",package.RequestedVersions) : ""));
            }

            // Add resolved versions for each project
            foreach (var project in projects)
            {
                var resolvedVersion = package.ResolvedVersions.FirstOrDefault(p => p.ProjectName == project)?.ResolvedVersion;
                row.Append(OpenXmlUtilities.CreateCell(resolvedVersion ?? ""));
            }
            sheetData.Append(row);
        }

        OpenXmlUtilities.AutoSize(worksheetPart);
        workbookPart.Workbook.Save();
    }
}
