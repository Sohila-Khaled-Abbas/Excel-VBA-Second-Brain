## 📋 Pull Request Summary

> Replace this with a **concise, human-readable description** of what this PR changes and why.

---

## 🔗 Related Issues

Closes # <!-- Add issue number if applicable -->

---

## 📁 Changes Made

> List the notes, scripts, or config files added or modified:

| Type | File Path | Description |
| :--- | :--- | :--- |
| `feat` | `12 - Code Cookbook/Patterns/...md` | Added new snippet for ... |
| `fix` | `01 - Course/Lectures/...md` | Fixed broken wikilink to Section 09 |
| `docs` | `README.md` | Updated architecture diagram |

---

## ✅ Pre-Merge Checklist

### Content Quality
- [ ] No proprietary course files (`.xlsm`, slide PDFs, video files) added
- [ ] All new notes have valid YAML frontmatter (all required fields present)
- [ ] VBA code uses `Option Explicit`, `On Error GoTo`, and COM cleanup
- [ ] No hardcoded absolute Windows paths in notes (use wikilinks or config variables)

### Vault Integrity
- [ ] Ran `python scripts/validate_vault.py` — exits with **code 0** (zero broken links)
- [ ] Ran `python scripts/sync_vault.py` if `config/course_manifest.json` was modified

### Git & Style
- [ ] Commit message follows [Conventional Commits](https://www.conventionalcommits.org/): `feat/fix/docs/chore(scope): description`
- [ ] PR title is descriptive and matches the main change category

---

## 🧪 Validation Output

> Paste the output of `python scripts/validate_vault.py`:

```
Notes analyzed: XXX
Wikilinks checked: X,XXX
Broken links: 0
✅ Vault integrity check passed.
```

---

## 📸 Screenshots (if UI or Dashboard changes)

> Add Obsidian screenshots or GitHub preview screenshots here if the change affects a dashboard or visual output.
