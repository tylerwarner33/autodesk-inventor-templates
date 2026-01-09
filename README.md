# Autodesk Inventor Addin Templates

![Inventor Versions](https://img.shields.io/badge/Inventor-2023--2026-blue.svg)
[![.NET Versions](https://img.shields.io/badge/.NET-4.8--8.0-blue.svg)](https://dotnet.microsoft.com/en-us/download/dotnet/6.0)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

This repository is based on the templates created by Curtis Waguespack & provided in his Autodesk University 2023 class: 
[Bridging the Gap Between iLogic Automation and Inventor Add-Ins](https://www.autodesk.com/autodesk-university/class/Bridging-the-Gap-Between-iLogic-Automation-and-Inventor-Add-Ins-2023).

- The templates have been modernized to the `dotnet new` template format.
- The projects have been modernized to the `SDK-style` project format.
- The solution has been modernized to the `slnx` solution format.

## Install

```bash
dotnet new install Autodesk.Inventor.Templates
```

## Uninstall

```bash
dotnet new uninstall Autodesk.Inventor.Templates
```

## Usage

### Visual Studio UI

1. Open `Visual Studio`
2. `File` => `New` => `Project`
3. Search for `Inventor`
4. Select `Inventor Add-In Basic Template`
5. Enter project name and location
6. Click `Create`

### Command Line

```bash
# Create a new add-in with default settings (Inventor 2026)
dotnet new inventoraddin -n MyInventorAddIn

# Create with specific Inventor version
dotnet new inventoraddin -n MyInventorAddIn --InventorVersion 2025

# Create in a specific output directory
dotnet new inventoraddin -n MyInventorAddIn -o ./src/MyInventorAddIn
```

## Setup

### Build

```bash
dotnet pack -o ./nupkg

dotnet new install ./nupkg/Autodesk.Inventor.Templates.*.nupkg

dotnet new uninstall Autodesk.Inventor.Templates
```

### Release

#### GitHub UI

1. Open `https://github.com/tylerwarner33/autodesk-inventor-templates`
2. Select `Actions` tab
3. Select `Build and Release` workflow
4. Click `Run workflow`
5. Enter `Release Version`
6. Click `Run workflow`

#### Command Line

Ex. version 1.0.0

```bash
git tag 1.0.0

git push --tags
```

## Resources

- Microsoft Learn: [Custom Templates For `dotnet new`](https://learn.microsoft.com/en-us/dotnet/core/tools/custom-templates)
- Microsoft GitHub: [Reference For Template.json](https://github.com/dotnet/templating/wiki/Reference-for-template.json)

## License

This sample is licensed under the terms of the [MIT License](http://opensource.org/licenses/MIT).
Please see the [LICENSE](LICENSE) file for more details.
