using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the path to the solution folder:");
        string? solutionPath = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(solutionPath) || !Directory.Exists(solutionPath))
        {
            Console.WriteLine("Invalid folder path. Exiting...");
            return;
        }

        try
        {
            // Step 1: Run the dotnet list package command
            Console.WriteLine("Running dotnet list package command...");
            string commandOutput = RunDotnetListPackage(solutionPath);

            // Step 2: Parse the output
            Console.WriteLine("Parsing output...");
            var packages = ParseDotnetOutput(commandOutput);

            // Step 3: Generate Excel report
            Console.WriteLine($"Found {packages.Count} packages in the solution. Do you want findings summarised? (Y/N)");
            string? summarise = Console.ReadLine();
            if (string.IsNullOrEmpty(summarise) || summarise.ToLower() != "y")
            {
                string combinedPath = Path.Combine(solutionPath,"NuGetPackageListReport.xlsx");
                ReportExcelFlatList.CreateExcelReport(packages,combinedPath);
                Console.WriteLine($"Report successfully generated: {combinedPath}");
                return;
            }
            string outputFileName = Path.Combine(solutionPath,"NuGetPackageReport.xlsx");
            Console.WriteLine($"Generating Excel report: {outputFileName}");
            ReportExcelPackagesByProject.PreparePackagesByProject(packages,outputFileName);

            Console.WriteLine($"Report successfully generated: {Path.GetFullPath(outputFileName)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static string RunDotnetListPackage(string solutionPath)
    {
        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = "list package --include-transitive",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = solutionPath
            }
        };

        process.Start();
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            throw new Exception($"dotnet command failed: {error}");
        }

        return output;
    }

    static List<NuGetPackageInfo> ParseDotnetOutput(string output)
    {
        var packages = new List<NuGetPackageInfo>();
        var lines = output.Split(Environment.NewLine);
        string? currentProject = null;
        bool isTransitiveSection = false;

        foreach (var line in lines)
        {
            string trimmedLine = line.Trim();
            Console.WriteLine($"Processing line: {trimmedLine}");

            // Check for project header
            if (trimmedLine.StartsWith("Project"))
            {
                currentProject = Regex.Match(trimmedLine,@"'(.+)'").Groups[1].Value;
                Console.WriteLine($"Found project: {currentProject}");
                isTransitiveSection = false;
            }
            else if (trimmedLine.StartsWith("Transitive"))
            {
                isTransitiveSection = true;
                Console.WriteLine("Switched to Transitive section.");
            }
            else
            {
                Match match;
                if (isTransitiveSection)
                {
                    // Use the simpler regex for transitive packages
                    match = Regex.Match(trimmedLine,@">\s*([\w\.\-]+)\s+([\d+(\.\d+)*\w\-]*)");
                }
                else
                {
                    match = Regex.Match(trimmedLine,@">\s*([\w\.\-]+)\s*(\(A\))?\s+((?:[\[\(]?\d+(\.\d+)*[\w\-]*[\]\), ]?)*)\s+([\w\.\(\)\[\],]+)");
                }

                if (match.Success && currentProject != null)
                {
                    string reqVersion = string.Empty;
                    string resVersion = string.Empty;
                    if (isTransitiveSection)
                    {
                        Console.WriteLine($"Transitive package match: {match.Groups[1].Value}");
                        Console.WriteLine($"Resolved Version: {match.Groups[2].Value}");
                        resVersion = match.Groups[2].Value;
                    }
                    else
                    {
                        Console.WriteLine($"Matched package: {match.Groups[1].Value}");
                        Console.WriteLine($"SDK or other auto-referenced package (A): {match.Groups[2].Value}");
                        Console.WriteLine($"Requested Version: {match.Groups[3].Value}");
                        reqVersion = match.Groups[3].Value;
                        Console.WriteLine($"Resolved Version: {match.Groups[5].Value}");
                        resVersion = match.Groups[5].Value;
                    }


                    packages.Add(new NuGetPackageInfo
                    {
                        ProjectName = currentProject,
                        PackageName = match.Groups[1].Value,
                        RequestedVersion = reqVersion,
                        ResolvedVersion = resVersion,
                        IsTransitive = isTransitiveSection ? "Yes" : "No"
                    });
                }
                else
                {
                    Console.WriteLine($"No match for line: {trimmedLine} or otherwise no current project scrope.");
                }
            }

        }
        return packages;

    }
}

