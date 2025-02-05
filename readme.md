# NuGetPackageVersionAnalyser

This utility applet analyzes the versions of NuGet packages used in a .NET solution and generates a report in Excel format.

## Features

- Issues `dotnet list package --include-transitive`
command to get the required information.
- Lists all NuGet packages used in the solution.
- Identifies transitive dependencies.
- Generates an Excel report with package details.

## Requirements

- .NET 8.0 SDK
- Visual Studio 2022

## Usage

1. Clone the repository.
2. Open the solution in Visual Studio.
3. Build the solution.
4. Run the application and provide the path to your .NET solution folder.

## License

This project is licensed under the MIT License.