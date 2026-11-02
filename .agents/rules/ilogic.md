---
name: iLogic Rules
description: The naming convention for iLogic external rules
triggers:
  - Editing or adding a file under Source/Templates/ExternalRules/
  - Writing iLogic code in a sample or in documentation
globs:
  - "Source/Templates/ExternalRules/**"
---

## Naming

Only iLogic uses this modified Hungarian notation.
All other code uses the conventions in `AGENTS.md`.

Start each variable name with `o`, and then write the full name.
Use the `o` for every type. A string or a boolean gets the same prefix as an object.

```vb
Dim oBalloonTip As BalloonTip
Dim oMessage As String
Dim oIsDrawing As Boolean
```

Do not abbreviate the name after the prefix (ex. `oBtip`, `oMsg`).

## Why

iLogic can change the case of a variable to match a type name.
Then `balloonTip` becomes `BalloonTip`, the variable has the name of the object type, and the rule fails.
The `o` prefix prevents this, because the name is then never the same as a type name.

## The rule file

- Use `ThisApplication` in a rule. `InventorApplication` and other add-in members do not exist in iLogic.
- A file in this folder goes to all 4 templates. The add-in runs a rule by its file name, so change the C# and the Visual Basic command if you rename a rule.
