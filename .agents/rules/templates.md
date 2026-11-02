---
name: Templates
description: The parallel C# and Visual Basic template structure, template parameters, and the traps in a C# Inventor add-in
triggers:
  - Editing a file under Source/Templates/CSharp/ or Source/Templates/VisualBasic/
  - Editing a template.json, or adding a template or an Inventor version
  - Editing Source/Autodesk.Inventor.Templates.csproj
globs:
  - "Source/Templates/CSharp/**"
  - "Source/Templates/VisualBasic/**"
  - "Source/Autodesk.Inventor.Templates.csproj"
---

## How the 2 languages connect

Each template has one folder per language. Both versions have the same `shortName` and `groupIdentity`, and each has its own `identity` and `language` tag.
This gives `dotnet new` the `--language` option.

`dotnet new` selects C# when the user does not give a language.
The user can change this default with the `DOTNET_NEW_PREFERRED_LANG` environment variable.

## Keep these the same in both languages

- The buttons: labels, icons, tool tips, panels, and stacks.
- The template parameters and the `InventorVersion` choices in `template.json`.
- The project settings, the `AddinOutputPath` options, and the build and clean targets.
- The files in `Resources/`. Git stores identical files once, so a copy adds no size to the repository.

## Folder and file names

Do not use spaces in a folder name or a file name.
The exception is `My Project` in the Visual Basic templates, because that is the folder that Visual Studio expects.

## Differences that are correct

| Item | C# | Visual Basic |
|---|---|---|
| Add-in class name in `InventorAddIn.X.manifest` | `InventorAddIn.StandardAddInServer` | `InventorAddIn.InventorAddIn.StandardAddInServer`, because Visual Basic adds the root namespace |
| Resources | `Properties/` | `My Project/` |
| Shared values | `Globals.cs` | A module in `StandardAddInServer.vb` |
| Ribbon icons | `AxHost` in `PictureDispConverter.cs` | `OleCreatePictureIndirect` in `PictureConverter.vb` |
| Late binding for iLogic | None: `IiLogicAutomation` | Late-bound calls on `Automation` |
| Inventor interop types | Not embedded | Embedded |
| `stdole` | `Bin\stdole.dll` of Inventor | NuGet package |
| `Button` base class | Not included | `UtilityTools/Button.vb`, which nothing uses |

## To add a template parameter or an Inventor version

1. Change both `template.json` files of the template.
2. Change the 2 files of the other template too, if the change applies to it.
3. For a new Inventor version, add the choice and the `TargetFramework` case to all 4 `template.json` files. Inventor 2027 uses `net10.0-windows`.
4. Do not change `defaultValue` without an instruction from the owner. `README.md` names the default version.

## Traps in a C# Inventor add-in

- **Connect an event only once.** Visual Basic `WithEvents` moves the handler when a field changes. A C# `+=` adds one more handler, and the command then runs 2 times for one click. Use `ConnectButton` in `StandardAddInServer.cs`. A button on 3 tabs, and a ribbon reset, both give the same button definition again.
- **Keep the events objects in fields.** Inventor stops sending events when the garbage collector releases the object.
- **Name conflicts.** The `Inventor` namespace has `Application`, `File`, `Path`, `Environment`, and `Color`. Write `Inventor.Application`, `System.IO.File`, `System.IO.Path`, and `System.Environment`.
- **`System.IO` is not an implicit using here.** `UseWPF` removes it. Write `System.IO.Directory`.
- **`net48` needs `LangVersion`.** Set it to `latest`, because the default is C# 7.3.
- **Do not use `dynamic` for iLogic.** Cast `ApplicationAddIn.Automation` to `IiLogicAutomation` from `Autodesk.iLogic.Interfaces.dll`, as `RunExternaliLogicRule.cs` does.
- **Do not embed the interop types in C#, and do not copy the Inventor assemblies.** The iLogic interfaces use the real `Inventor.Document` type. So the C# project sets `EmbedInteropTypes` to `False`, and sets `Private` to `False` on each Inventor reference.
- **Do not use the `stdole` NuGet package in C#.** The package is version 17.0.0.0, and the interop assembly uses version 7.0.3300.0 from the Inventor `Bin` folder. A build passes, but `ButtonDefinition.LargeIcon` then fails at run time with `Method not found`. Reference `Bin\stdole.dll`.
