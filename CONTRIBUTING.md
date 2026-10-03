# Contributing to Excel VBA Second Brain

Thank you for your interest in contributing to this personal knowledge base and automation repository!

## 🎯 Repository Purpose
This vault is a structured **Second Brain** and technical reference for Excel VBA, automation architecture, and data engineering workflows. It transforms learning from course material into evergreen, reusable engineering knowledge.

## 🛡️ Public vs. Private Content Guardrails
We maintain strict compliance with copyright and licensing:
1. **Never commit course assets**: No videos, official slide PDFs, proprietary workbooks, or verbatim instructor transcripts.
2. **Commit only original synthesis**: Your own explanations, mental models, code snippets, diagrams, and debugging insights.
3. **Keep binary workbooks out**: Commit VBA source code as Markdown blocks or exported plain text (`.bas`, `.cls`, `.frm`).

## 📐 Architecture & Contribution Standards
When adding or refining notes:
- **YAML Frontmatter**: Ensure consistent metadata (`type`, `domain`, `status`, `tags`, etc.).
- **Wikilinks**: Connect concepts aggressively using ``[[Variables]]``.
- **Mermaid Diagrams**: Use standard Obsidian-compatible Mermaid syntax.
- **Code Blocks**: Always specify language as `vba` with clean indentation, comments, and explicit variable declarations (`Option Explicit`).
- **Personal Understanding Layer**: Include the 5-layer model:
  1. Technical Definition
  2. Personal Understanding
  3. Simple Analogy
  4. Real Example
  5. Analytics / Production Application

## 🚀 Pull Request Checklist
- [ ] No proprietary `.xlsm` or course files added
- [ ] Wikilinks verified and working
- [ ] Frontmatter valid YAML
- [ ] Code formatted with `Option Explicit` and error handling
