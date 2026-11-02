# Autodesk Inventor Add-In Templates

A NuGet package of `dotnet new` project templates for Autodesk Inventor add-ins, in C# and in Visual Basic.
See `README.md` for install and usage.

These are the instructions for a coding agent that works on this repository (ex. Claude Code, GitHub Copilot).

## Layout

| Path | Contents |
|---|---|
| `Source/Autodesk.Inventor.Templates.csproj` | The package project. It compiles nothing and packs `Source/Templates/`. |
| `Source/Templates/CSharp/` | The C# templates, and their `UtilityTools/` |
| `Source/Templates/VisualBasic/` | The Visual Basic templates, and their `UtilityTools/` |
| `Source/Templates/ExternalRules/` | The iLogic rules. All 4 templates use them. |
| `Docs/` | The Autodesk University handout and presentation |
| `.github/workflows/` | `Build.yml` packs each push. `Release.yml` publishes to NuGet. |
| `.agents/rules/` | On-demand rules for agents that work on this repository (listed below) |

The folders `UtilityTools/` and `ExternalRules/` are not in a template folder in the repository.
The package project copies them in to each template when it packs.
So a template folder does not build where it is. Create a project from the package to build it.

## Working Principles

- **Keep the 2 languages the same.** A user-visible change to a template (ex. a button, a label, a parameter) goes in the C# and the Visual Basic version. Open `.agents/rules/templates.md`.
- **Keep the README simple.** `README.md` is also the NuGet package page. Do not add sections to it unless the owner asks.
- **A successful pack does not prove that a template works.** Create a project from the package and build it. Open `.agents/rules/validation.md`.
- **Do not change a published identifier without a reason.** A change to `identity` or `shortName` breaks the commands and the Visual Studio entries of existing users.

## Naming Conventions

`.editorconfig` holds the C# rules. Use the same names in Visual Basic.

- Use PascalCase for types, methods, properties, events, and constants.
- Use `_camelCase` for private and internal fields.
- Use camelCase for parameters and local variables.
- Start interface names with `I`.
- Use full names, not abbreviations (ex. `controlDefinitions`, not `controlDefs`).

The exception is iLogic. Open `.agents/rules/ilogic.md` before you edit a rule.

## Code Conventions

- Use tabs with a width of 3 in C#.
- Do not use `var`. Write the type.
- Use file-scoped namespaces. The namespace matches the folder.
- Write `x is not null` and `x is false`, not `x != null` and `!x`.
- Write each XML doc tag on its own line. Indent its content one tab. Write one sentence per line.

## Writing

- Write in ASD-STE100: short sentences, one instruction per sentence, active voice.
- Use `-`, never an em dash or an en dash. Use `ex.` instead of `e.g.`.

## No Private Data in the Repository

- Do not commit a secret (ex. an API key, a token, or a password).
- Do not commit client names, user names, machine names, or absolute user paths.

## On-Demand Rules

The rules are not loaded automatically. Before acting on a task, scan this list and open any rule whose trigger fits the request or the files about to change. If several match, open all of them.

- **`.agents/rules/templates.md`** - Open before you edit anything under `Source/Templates/CSharp/` or `Source/Templates/VisualBasic/`, a `template.json`, or the package project. Covers the parallel structure, template parameters, and the traps in a C# Inventor add-in.
- **`.agents/rules/ilogic.md`** - Open before you edit or add a file under `Source/Templates/ExternalRules/`. Covers the `o` naming prefix.
- **`.agents/rules/validation.md`** - Open before you finish a change to a template or to the package, or when a template build fails. Covers how to pack, create, and build each template, how to test in Inventor, and what each check does not prove.
