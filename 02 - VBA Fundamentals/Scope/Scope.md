---
type: concept
domain: VBA
topic: Architecture
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 08 - Debugging, Error Handling & Procedure Scope]]"
related_concepts:
  - "[[Variables]]"
  - "[[Procedures & Functions]]"
tags:
  - vba
  - scope
  - architecture
created: 2026-10-03
updated: 2026-10-03
---

# Scope (Variable & Procedure Visibility)

> [!important]
> **Summary Statement**: Scope determines where in a codebase a variable, constant, or procedure can be seen, referenced, and modified, as well as its lifetime in memory.

---

## 📊 The 3 Scope Levels in VBA

```mermaid
flowchart TD
    subgraph GlobalScope["1. Project / Global Level (Public)"]
        G[Visible to all modules in workbook while workbook is open]
    end

    subgraph ModuleScope["2. Module Level (Private / Dim at top of module)"]
        M[Visible only to procedures within that specific module]
    end

    subgraph ProcedureScope["3. Procedure / Local Level (Dim inside Sub/Function)"]
        P[Created on entry, destroyed when Sub/Function exits]
    end

    GlobalScope --> ModuleScope
    ModuleScope --> ProcedureScope
```

| Scope Level | Declaration Syntax | Lifetime | Visibility |
| :--- | :--- | :--- | :--- |
| **Procedure (Local)** | `Dim x As Long` inside Sub/Function | During procedure execution | Only within the enclosing procedure |
| **Module (Private)** | `Private x As Long` or `Dim x As Long` at top declarations | While workbook remains open | All procedures in that module |
| **Global (Public)** | `Public x As Long` at top of standard module | While workbook remains open | Any procedure in any module |

---

## ⚠️ Best Practice Rule
> [!tip]
> **Default to the narrowest possible scope (Local)**. Only elevate to Module or Public scope when data genuinely needs to be shared across multiple procedures or UserForm events. Global variables increase coupling and make debugging difficult.

---

## 🔗 Connected Notes
- **Procedure Types**: [[Procedures & Functions]]
- **Variables**: [[Variables]]
