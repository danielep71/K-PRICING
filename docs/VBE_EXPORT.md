> [!IMPORTANT]
> **Migration scope — 2026-09-24.** Imported from [danielep71/KPR at `f26450d`](https://github.com/danielep71/KPR/blob/f26450d1fa7b11261162e901dedba062f21c99a7/docs/VBE_EXPORT.md).
> Source version labels (including v0.0.2 and v0.0.3), roadmap statements and
> historical evidence refer to KPR. They do not set K-PRICING release versions
> or claim destination certification. The frozen implementation is now migrated
> into K-PRICING; exact-source Windows Excel compilation and parity verification
> remain pending in [K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).
> [Migration provenance](MIGRATION_PROVENANCE.md) records hashes, attribution
> and the limited editorial adaptations.

<div align="center">

# 🧩 VBE Export Format

### Canonical source exchange between Git, the Visual Basic Editor and Windows Excel

**Export fidelity · Stable component identity · Reviewable source · Exact-candidate evidence**

<br>

![VBE Export](https://img.shields.io/badge/VBE-Export_Format-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Tracked Source](https://img.shields.io/badge/Source-Tracked_VBA-0969da?style=for-the-badge)
![Encoding](https://img.shields.io/badge/Encoding-ASCII-6f42c1?style=for-the-badge)
![Line Endings](https://img.shields.io/badge/Line_Endings-CRLF-d97706?style=for-the-badge)
![Validation](https://img.shields.io/badge/Validation-Static_Gate-2ea44f?style=for-the-badge)

</div>

---

Every tracked VBA file in this repository must use the Visual Basic Editor's
native export format. Hand-written approximations are not accepted. This format keeps exported source, the repository static gate and Windows
VBE import/export aligned.

> [!IMPORTANT]
> This page defines the source exchange format. It does **not**
> establish that tracked source imports, compiles or runs in Excel. That claim
> belongs exclusively to the exact-source Windows certification recorded in
> source issue [danielep71/KPR#29](https://github.com/danielep71/KPR/issues/29);
> destination parity is tracked in [K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).

## 🧾 Format contract

A VBE export of a standard module begins with its component name:

```text
Attribute VB_Name = "KPR_DATES_DAYS"
```

The imported format contract requires the following rules. Specialist gate
coverage is migrated in [K-PRICING #13](https://github.com/danielep71/K-PRICING/issues/13);
this document alone does not establish that every rule is already enforced:

| Rule | Requirement |
|---|---|
| **Presence** | Every `.bas`, `.cls` and `.frm` file declares `Attribute VB_Name`. |
| **Form** | Use the canonical VBE spelling `Attribute VB_Name = "<ComponentName>"`, with no leading whitespace and no alternative spacing. |
| **Position** | A `.bas` module declares the attribute on line 1. A `.cls` or `.frm` export opens with its own `VERSION`/`BEGIN` header block, so the declaration may appear anywhere in that leading header region. |
| **Identity** | The declared component name matches the file name exactly, including case. |
| **Legality** | The component name is a legal VBA identifier of no more than 31 characters. |
| **Uniqueness** | No two tracked VBA files declare names that are equal under VBA's case-insensitive component-name semantics. VBA components share one flat project namespace across `src/`, `tests/` and `examples/`. |
| **Declarations** | `Option Explicit` is present, as it was before this export format was adopted. |

> [!WARNING]
> Procedure-level attributes such as `Attribute VB_Description` and
> `Attribute VB_ProcData.VB_Invoke_Func` are rejected. Function and argument
> descriptions belong exclusively to the `Application.MacroOptions` manifest.
> These attributes are invisible in the editor but survive export, creating a
> second description mechanism that can silently disagree with the manifest.

### 🔤 Encoding and line endings

Encoding and line endings follow `.gitattributes` and `.editorconfig`:

| Property | Repository policy |
|---|---|
| **Working-tree line endings** | CRLF |
| **End of file** | Final newline required |
| **Tracked VBA character set** | ASCII only |
| **Reason** | The VBE writes text in the Windows system code page rather than UTF-8. |

Typographic quotes, en dashes and accented characters in VBA comments are the
usual causes of a file no longer round-tripping cleanly.

## Procedure and component ownership

[Installation](../INSTALLATION.md) owns import, export, replacement and recovery
procedures. [Repository structure](REPOSITORY_STRUCTURE.md) owns all component
placements, including classes, forms, workbook components and examples under
`examples/modules/`. This format contract does not redefine those paths.

Exact-candidate round-trip evidence belongs to
[release evidence](RELEASE_EVIDENCE.md). A static format check does not prove
that Excel imported, compiled or executed a component.
