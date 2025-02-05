using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public class NuGetPackageVersion
{
    public string? PackageName
    {
        get; set;
    }

    public string? IsTransitive
    {
        get; set;
    }

    public List<string?>? RequestedVersions
    {
        get; set;
    }

    public List<ProjectResolvedVersion>? ResolvedVersions
    {
        get; set;
    }
}

public class ProjectResolvedVersion
{
    public string? ProjectName
    {
        get; set;
    }

    public string? ResolvedVersion
    {
        get; set;
    }
}