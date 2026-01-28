# Autodesk Inventor Addin Templates

![Inventor Versions](https://img.shields.io/badge/Inventor-2023--2026-blue.svg)
[![.NET Versions](https://img.shields.io/badge/.NET-4.8--8.0-blue.svg)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/autodesk.inventor.templates?logo=nuget&label=NuGet&color=brightgreen)](https://www.nuget.org/packages/Autodesk.Inventor.Templates)


This is a collection of multi-version Visual Studio templates for creating Autodesk Inventor Add-Ins. Started from the templates created by Curtis Waguespack & provided in his Autodesk University 2023 class: 
[Bridging the Gap Between iLogic Automation and Inventor Add-Ins](https://www.autodesk.com/autodesk-university/class/Bridging-the-Gap-Between-iLogic-Automation-and-Inventor-Add-Ins-2023).

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

- List all installed Inventor templates.
  ``` bash
  dotnet new list inventor
  ```
- Display all available options for the basic template.
  ``` bash
  dotnet new inventoraddin-basic --help
  ```
- Create a new add-in with the basic template and default settings (Inventor 2026).
  ``` bash
  dotnet new inventoraddin-basic --name YourAddInName
  ```
- Create a new add-in with the stacked buttons template and default settings (Inventor 2026).
  ``` bash
  dotnet new inventoraddin-stacked --name YourAddInName
  ```
- Create a new add-in with the basic template and specified Inventor version.
  ``` bash
  dotnet new inventoraddin-basic --name YourAddInName --InventorVersion 2025
  ```
- Create a new add-in with the basic template, specified Inventor version and output directory.
  ``` bash
  dotnet new inventoraddin-basic --name YourAddInName --InventorVersion 2025 --output C:\Projects\YourAddInName
  ```

## Setup

### Build

1. Generate the package locally
    ```bash
    dotnet pack -o ./nupkg -p:PackageVersion=1.0.0
    ```
2. Install the locally generated package.
    ```bash
    dotnet new install ./nupkg/Autodesk.Inventor.Templates.1.0.0.nupkg
    ```
    - View installed packages.
      ```bash
      dotnet new uninstall
      ```
3. Uninstall the locally generated & installed package.
    ```bash
    dotnet new uninstall Autodesk.Inventor.Templates
    ```

### Release

#### GitHub UI

1. Open `https://github.com/tylerwarner33/autodesk-inventor-templates`
2. Select `Actions` tab
3. Select `Release` workflow
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
