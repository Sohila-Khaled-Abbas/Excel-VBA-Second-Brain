---
title: Regular Expression Text Parser in VBA
type: snippet
domain: code_cookbook
status: evergreen
tags:
  - vba/regex
  - vba/text-parsing
  - vba/data-cleaning
  - vba/snippet
related:
  - "[[Variables & Data Types]]"
  - "[[Functions]]"
created: 2026-10-26
updated: 2026-10-26
---

# Regular Expression Text Parser in VBA

> **Purpose**: Use `VBScript.RegExp` via Late Binding to perform pattern matching, extract emails/phone numbers/tax IDs, and clean unstructured strings directly inside Excel macros or User-Defined Functions (UDFs).

---

## Production Function: RegEx Extract Pattern

```vba
Option Explicit

'''
''' Extracts matching regex group or full match from an input string.
''' Can be used as a worksheet formula: =RegExExtract(A2, "\d{4}-\d{2}-\d{2}")
'''
Public Function RegExExtract(ByVal Text As String, ByVal Pattern As String, _
                            Optional ByVal MatchIndex As Long = 0) As String
    Dim reg As Object
    Dim matches As Object
    
    If Len(Trim(Text)) = 0 Or Len(Trim(Pattern)) = 0 Then
        RegExExtract = ""
        Exit Function
    End If
    
    ' Late Binding: No reference needed for VBScript_RegExp_55
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Global = True
        .IgnoreCase = True
        .Pattern = Pattern
        
        If .Test(Text) Then
            Set matches = .Execute(Text)
            If MatchIndex < matches.Count Then
                RegExExtract = matches(MatchIndex).Value
            Else
                RegExExtract = matches(0).Value
            End If
        Else
            RegExExtract = ""
        End If
    End With
    
    Set matches = Nothing
    Set reg = Nothing
End Function
```

---

## Production Function: RegEx Replace (Cleaning Text)

```vba
Public Function RegExClean(ByVal Text As String, ByVal Pattern As String, _
                          Optional ByVal Replacement As String = "") As String
    Dim reg As Object
    
    If Len(Trim(Text)) = 0 Then
        RegExClean = ""
        Exit Function
    End If
    
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Global = True
        .IgnoreCase = True
        .Pattern = Pattern
        RegExClean = .Replace(Text, Replacement)
    End With
    
    Set reg = Nothing
End Function
```

---

## Common Regex Patterns for Excel Data Cleaning

| Data Target | Regex Pattern | Example Formula |
| :--- | :--- | :--- |
| **Email Address** | `[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}` | `=RegExExtract(A2, "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")` |
| **ISO Date (YYYY-MM-DD)**| `\b\d{4}-\d{2}-\d{2}\b` | `=RegExExtract(A2, "\b\d{4}-\d{2}-\d{2}\b")` |
| **Extract Numbers Only** | `[^\d]` (replace with `""`) | `=RegExClean(A2, "[^\d]", "")` |
| **Strip Extra Whitespace**| `\s+` (replace with `" "`) | `=RegExClean(A2, "\s+", " ")` |
| **Invoice Number (INV-12345)**| `\bINV-\d{5}\b` | `=RegExExtract(A2, "\bINV-\d{5}\b")` |
