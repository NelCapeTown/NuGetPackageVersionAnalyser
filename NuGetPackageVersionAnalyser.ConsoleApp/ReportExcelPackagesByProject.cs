using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public static class ReportExcelPackagesByProject
{
    public static void PreparePackagesByProject(List<NuGetPackageInfo> packages,string filePath)
    {
        try
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

            var pkg = OpenXmlUtilities.CreatePackage(filePath);
            if (pkg is not null)
            {
                var sheetData = CreateSheetData(distinctPackages,projects);
                OpenXmlUtilities.AddSheetData(pkg,"NuGet Summary",sheetData);
                pkg.Dispose();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in ReportExcelPackagesByProject.PreparePackagesByProject: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static X.SheetData? CreateSheetData(List<NuGetPackageVersion> distinctPackages,List<string?> projects)
    {
        try
        {
            var sheetData = new X.SheetData();

            // Add headers
            var headerRow = new X.Row();
            OpenXmlUtilities.AddCell(headerRow,"Package Name",true);
            OpenXmlUtilities.AddCell(headerRow,"Is Transitive",true);
            OpenXmlUtilities.AddCell(headerRow,"Requested Version",true);
            foreach (var project in projects)
            {
                OpenXmlUtilities.AddCell(headerRow,project,true);
            }
            sheetData.Append(headerRow);

            // Add data rows
            foreach (var package in distinctPackages)
            {
                var row = new X.Row();
                OpenXmlUtilities.AddCell(row,package.PackageName);
                OpenXmlUtilities.AddCell(row,package.IsTransitive);

                // Handle requested versions
                if (package.IsTransitive == "No" && package.RequestedVersions.Count > 1)
                {
                    var concatenatedVersions = string.Join(", ",package.RequestedVersions);
                    OpenXmlUtilities.AddCell(row,concatenatedVersions);
                }
                else
                {
                    OpenXmlUtilities.AddCell(row,package.IsTransitive == "No" ? string.Join(", ",package.RequestedVersions) : "");
                }

                // Add resolved versions for each project
                foreach (var project in projects)
                {
                    var resolvedVersion = package.ResolvedVersions.FirstOrDefault(p => p.ProjectName == project)?.ResolvedVersion;
                    OpenXmlUtilities.AddCell(row,resolvedVersion ?? "");
                }
                sheetData.Append(row);
            }

            return sheetData;

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in ReportExcelPackagesByProject.CreateSheetData: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return null;
        }
    }
}

