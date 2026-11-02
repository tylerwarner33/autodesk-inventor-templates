---
name: Validation
description: How to pack, create, and build each template, and what each check does not prove
triggers:
  - Before you finish a change to a template or to the package
  - A template build fails
  - Running dotnet pack, dotnet new, or a build of a generated project
---

## Pack and create

Use a custom hive, so the test does not change the templates that are installed on the machine.
Replace `<test folder>` with a folder outside the repository.

```
dotnet pack Autodesk.Inventor.Templates.slnx -c Release -p:PackageVersion=9.9.9 -o <test folder>/pkg
dotnet new install <test folder>/pkg/Autodesk.Inventor.Templates.9.9.9.nupkg --debug:custom-hive <test folder>/hive
dotnet new inventoraddin-basic --language C# --InventorVersion 2026 --name Test --output <test folder>/Test --debug:custom-hive <test folder>/hive
```

Do this for each of the 2 templates, in each of the 2 languages.

## Build

A build copies the add-in to the Inventor add-ins folder of the user.
Set `AddinOutputPath` to stop this, unless you want to load the add-in in Inventor.

| Inventor version | Framework | Command |
|---|---|---|
| 2027 | .NET 10 | `dotnet build <project> -p:AddinOutputPath=<test folder>/drop` |
| 2025 and 2026 | .NET 8 | `dotnet build <project> -p:AddinOutputPath=<test folder>/drop` |
| 2024 and earlier | .NET Framework 4.8 | `MSBuild.exe <project> -restore -p:AddinOutputPath=<test folder>/drop` |

- Use the Visual Studio `MSBuild.exe` for `net48`. `dotnet build` fails with `MSB3822` and `MSB3823`, because it cannot embed the bitmap resources for .NET Framework.
- Keep `<test folder>` short. A long path can make the copy step fail.
- The interop assembly comes from the Inventor install folder. A version that is not installed fails with `MSB3245`.

## Check the C# style

Copy `.editorconfig` in to the generated C# project and build with `-p:EnforceCodeStyleInBuild=true`.
The build must show 0 warnings.

## What each check does not prove

- A pack does not compile a template. `EnableDefaultItems` is false in the package project.
- A build does not prove that the add-in loads, that the buttons show, or that a click runs the command. A missing method only fails at run time (ex. the `stdole` trap in `templates.md`).
- Test in Inventor after a change to `StandardAddInServer`, an event, a resource, a project reference, or a `.addin` or `.manifest` file.

## Test in Inventor

Use the Inventor MCP server. It supports Inventor 2025, 2026, and 2027. Test each version that the change can affect.

1. Create a test copy of each template in each language, with a stable name (ex. `TestStackedCs`). A new copy has a new GUID, so create the copies once.
2. Build each copy with no `AddinOutputPath` property. The build installs the add-in in the per-user add-ins folder.
3. In the `.addin` file of the Visual Basic copies, set `LoadOnStartUp` to `0`. Both languages make the same ribbon tabs, so load one language at a time.
4. In the test copies, replace the call to `Welcome.AddinHello` with a line that sets `StatusBarText`. The message box of that command stops all calls.
5. Start Inventor with `inventor_start`, and wait with `inventor_session`.
6. Do these checks with `inventor_eval_csharp`:

| Check | Expected result |
|---|---|
| The ribbon of each environment | The tabs and panels of each template. The Stacked Button template has a popup, a button popup, and a split button on the Drawing tab. |
| `LargeIcon` and `StandardIcon` of each button | Not null |
| `ExternalRuleDirectories` of iLogic | The `ExternalRules` folder of each loaded add-in |
| A new drawing | `StatusBarText` shows the event ran |
| `Execute2(true)` on each symbol button | One more symbol on the sheet for each click |
| `Execute2(true)` on each sheet color button | The sheet has the color of the button |
| `Execute2(true)` on the balloon button | One more balloon tip. This also runs an external rule. |
| `Deactivate`, then `Activate`, then one click (C# only) | The tabs go, the tabs come back, and one click adds one symbol |

7. Do not click a button that shows a message box (ex. Detect Theme, Show Date). The box stops all calls.
8. To test the other language, deactivate the first add-ins and activate the others.

### Dialogs that need the owner

The server does not answer these dialogs. Tell the owner, and wait.

- **Add-in Manager Security Alert.** Inventor blocks each unsigned add-in with a new GUID. The owner unblocks it in the Add-in Manager, once in each version.
- **Log In.** Inventor shows this in a Vault project. Never type credentials.

### Clean up

1. Remove the test folders from `ExternalRuleDirectories`. An add-in adds its folder when it loads, and the setting stays after the session. Keep all other folders.
2. Close only the documents that you opened, with no save.
3. Close Inventor, and then delete the `Test*.bundle` folders from the add-ins folder of each version.
