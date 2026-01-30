# Stacked Button Template

## Build & Installation

This project uses MSBuild targets to automatically install or uninstall the Autodesk Inventor add-in when you build or clean the project.

### How It Works

The `.vbproj` file contains custom MSBuild targets that handle deployment:

1. **BuildAddinOutputPath** (runs after Build):
   - Copies `PackageContents.xml` to the add-in bundle folder
   - Copies all build output files to the bundle's `Contents` folder
   - Destination: `$(AddinOutputPath)\$(MSBuildProjectName).bundle\`

2. **CleanAddinOutputPath** (runs after Clean):
   - Removes the add-in bundle folder when you clean the project

### Add-in Output Location

The `AddinOutputPath` property in the project file determines where the add-in is installed. By default, it installs to the **Per User, Version Dependent** location:
`%AppData%\Autodesk\Inventor $(InventorVersion)\Addins`

You can change this by uncommenting a different option in the `.vbproj` file:

| Location | Path | Notes |
|----------|------|-------|
| All Users, Version Independent | `$(ProgramData)\Autodesk\Inventor Addins` | Shared across Inventor versions |
| All Users, Version Dependent (2024+) | `$(ProgramFiles)\Autodesk\Inventor $(InventorVersion)\Bin\Addins` | Trusted location, requires Admin |
| All Users, Version Dependent (2023-) | `$(ProgramData)\Autodesk\Inventor $(InventorVersion)\Addins` | For older Inventor versions |
| Per User, Version Independent | `$(AppData)\Autodesk\ApplicationPlugins` | User-specific, shared across versions |
| Per User, Version Dependent | `$(AppData)\Autodesk\Inventor $(InventorVersion)\Addins` | **Default** |

### Usage

1. **Build** the project in Visual Studio with Inventor closed to automatically install the add-in.
2. **Open Inventor** to load the add-in automatically.
3. **Clean** the project in Visual Studio with Inventor closed to uninstall the add-in.

### Documentation

For more information about add-in deployment locations, see the official [Autodesk Documentation](
https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=GUID-52422162-1784-4E8F-B495-CDB7BE9987AB)
