> [!IMPORTANT]
> **Migration scope — 2026-09-24.** Imported from [danielep71/KPR at `f26450d`](https://github.com/danielep71/KPR/blob/f26450d1fa7b11261162e901dedba062f21c99a7/docs/VBE_EXPORT.md).
> Source version labels (including v0.0.2 and v0.0.3), roadmap statements and
> historical evidence refer to KPR. They do not set K-PRICING release versions
> or claim destination implementation or certification. The destination still
> contains the neutral starter; migration runtime verification is pending in
> [K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).
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
native export format. Hand-written approximations are not accepted. Following
this procedure keeps exported source, the repository static gate and Windows
VBE import/export aligned.

> [!IMPORTANT]
> This page defines a text format and a contributor procedure. It does **not**
> establish that tracked source imports, compiles or runs in Excel. That claim
> belongs exclusively to the exact-source Windows certification recorded in
> source issue [danielep71/KPR#29](https://github.com/danielep71/KPR/issues/29);
> destination parity is tracked in [K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).

## 🧭 Workflow at a glance

| Stage | Contributor action | Required outcome |
|---:|---|---|
| **① Export** | Export the component from the VBE into its owning repository directory. | File name and `Attribute VB_Name` remain identical. |
| **② Review** | Inspect the complete Git diff before committing. | No accidental whitespace, line-ending or hidden-attribute drift. |
| **③ Remove** | Remove any same-named component already loaded in the target VBA project. | The import cannot create a suffixed duplicate. |
| **④ Import** | Import the tracked source into the clean target project. | The VBE retains the intended component identity. |
| **⑤ Certify** | At the release gate only, round-trip the exact candidate source. | Evidence identifies the exact candidate SHA. |

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

## 📤 Export from Excel

1. In the VBE, select the component in **Project Explorer**.
2. Choose **File → Export File**.
3. Save it into the directory that owns the component:

   | Component role | Destination |
   |---|---|
   | Production facade | `src/modules/` |
   | Production core | `src/core/` |
   | Regression | `tests/modules/` |
   | Examples | `examples/` (no KPR demo builder is imported) |

4. Keep the file name identical to the component name. Do not rename the file
   afterwards and do not hand-edit `Attribute VB_Name` to another value.
5. Review the resulting diff before committing.

> [!NOTE]
> An export that differs only in line endings or whitespace usually means the
> editor or a Git setting has overridden the repository policy.

## 📥 Import into Excel

1. If a component of the same name is already loaded, right-click it, choose
   **Remove**, and decline the export prompt.
2. Choose **File → Import File** and select the exported file.

> [!CAUTION]
> Import does not replace an existing same-named component. Excel retains both
> and renames the newcomer: `KPR_DATES_DAYS` silently becomes
> `KPR_DATES_DAYS1`, leaving two copies of the same procedures in the project.
> **Always remove the existing component before importing.**

## 🔁 Certification round trip

A normalized export/import round trip consists of:

1. importing the exact candidate source;
2. exporting it again without modification; and
3. comparing the normalized result with the candidate source.

This is release evidence, not a routine contributor step. It is performed once
against the recorded candidate SHA during Windows certification and remains out
of scope for ordinary changes and the hosted static gate. The hosted gate
validates text shape only.
