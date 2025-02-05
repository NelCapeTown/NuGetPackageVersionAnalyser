using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuGetPackageVersionAnalyser.ConsoleApp;
public class NuGetPackageInfo
{
    public string? ProjectName
    {
        get; set;
    }

    public string? PackageName
    {
        get; set;
    }

    public string? IsTransitive
    {
        get; set;
    }

    public string? RequestedVersion
    {
        get; set;
    }

    public string? ResolvedVersion
    {
        get; set;
    }

}
