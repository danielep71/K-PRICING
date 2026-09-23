# ✍️ VBA House Style

[![Scope: VBA exports](https://img.shields.io/badge/scope-VBA%20exports-217346)](REPOSITORY_STRUCTURE.md)
[![Change: presentation](https://img.shields.io/badge/change-presentation-0969da)](#review-and-validation)

This document owns source presentation for the starter modules and new VBA
procedures. Component ownership remains in
[Repository Structure](REPOSITORY_STRUCTURE.md); the supported public surface
remains in [PUBLIC_API.txt](PUBLIC_API.txt).

The style follows Daniele Penza's established module and procedure presentation.
The rules below are self-contained and adapt that presentation to this project's
component roles. No external reference repository is required to use them.

## 🧱 Module layout

Keep the following order:

1. `Attribute VB_Name` on the first physical line of a `.bas` export.
2. Full-width module banner: module name, purpose, public surface, dependencies,
   state ownership, error policy, relevant worksheet safety/test seam,
   compatibility, usage where applicable, updated date and author.
3. `MODULE SETTINGS`: `Option Explicit` and the role's existing visibility.
4. Separate `MODULE CONSTANTS` and `MODULE STATE` sections when needed.
5. Centered, uppercase procedure-group headings, followed by the procedures.

Use an apostrophe followed by 78 equals signs for banners and 78 hyphens for
section rules: 79 characters in total. Do not pad centered titles with trailing
spaces. Source decorations use ASCII, even when Markdown documentation uses
glyphs and badges.

The facade remains externally visible. Core, tests and examples keep
`Option Private Module`. Public members inside a project-private module allow
in-project calls; they do not become the supported external API. Copying a
reference module's visibility would change that contract.

## 📝 Procedure layout

Put the procedure signature first, followed immediately by its comment banner
inside the procedure. Use a centered procedure name, then the sections needed
to explain its contract:

| Section | Explain |
| --- | --- |
| `PURPOSE` | Why the procedure exists and the responsibility it owns |
| `INPUTS` | Meaning, valid values, units and ownership where relevant |
| `RETURNS` | Result meaning, output parameters and failure sentinels |
| `ERROR POLICY` | Expected errors, propagation, containment and cleanup |
| `DEPENDENCIES` | Boundaries that matter to callers or maintainers |
| `STATE OWNERSHIP` / `SIDE EFFECTS` | State read, changed, retained or released |
| `USAGE` | Entry-point instructions or a useful calling example |
| `UPDATED` | Date this procedure's code or documentation was last revised |

Document every starter procedure, including assertion, reporting, reset and
cleanup helpers. Keep short helpers' contracts concise; omit empty sections.
Comments should explain intent and invariants rather than narrate assignments.
Each local variable, module-state variable and constant has an inline comment
explaining its role, meaning or lifetime. Align the trailing comment column
within each declaration group. A descriptive name does not replace this short
annotation: distinguish captured errors, host snapshots and retained report
state explicitly.

Under every executable section banner, add an indented explanatory comment
before the code. Explain the purpose of that phase and any ordering constraint
or failure consequence. Add further comments before distinct steps within a
larger phase, particularly error capture, cleanup, validation and propagation.
Section titles alone do not supply this explanation. Put a handler's entry
label before its explanatory comment so the comment accompanies the statements
executed at that label.

Keep module dates at least as recent as the procedures changed in that module;
do not refresh untouched procedure dates merely to make all dates identical.
Preserve accurate authorship when adapting the starter.

Separate local declarations under `DECLARE`. Divide the executable body into
meaningful phases such as `GUARD ENTRY`, `RUN SUITE`, `CLEANUP AND REPORT` and
`HANDLE RUNNER ERROR`. A small helper generally needs only one body section.

## 📐 Indentation and wrapping

- Use spaces in four-column levels; no tabs or trailing whitespace.
- Indent module settings and local declarations four spaces. Align `As` within
  each declaration group, with at least three spaces before the aligned column.
- Indent module constants and state declarations eight spaces. Body statements
  start at eight spaces, below the four-space intent-comment level; nested
  blocks and expression continuations add four spaces.
- Keep procedure declarations, `End Sub` / `End Function`, jump labels and
  conditional-compilation directives at column one.
- Split parameterized signatures onto continuation lines, one parameter per
  line at four spaces. Put a function's return type on a separate continuation
  line. Preserve argument order, defaults, types and passing modes.
- Separate procedures with two empty lines. Keep one empty line between body
  phases and before the terminating `End Sub` or `End Function`.
- Aim for 79-column prose and readable code around 100 columns. Preserve a
  longer literal or declaration when wrapping it would introduce an unrelated
  expression change. This is a presentation target, not a new compiler limit.

For example, the facade signature is formatted as:

```vb
Public Function ProjectRatio( _
    ByVal numerator As Double, _
    ByVal denominator As Double) _
    As Double
```

Its procedure banner follows that signature; it does not precede it. The full
`src/modules/ProjectFacade.bas` export is the working example.

Inside that procedure, annotations accompany the declarations and body:

```vb
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep every error field needed to preserve the core failure contract.
    Dim savedNumber        As Long      'Original error number for later re-raise

'------------------------------------------------------------------------------
' CALL CORE
'------------------------------------------------------------------------------
    'Delegate the arithmetic to the core; this boundary owns only the
    'caller-facing error source.
        On Error GoTo HandleError
```

This is an excerpt: the full export retains all declarations and the handler.

## 💾 Export compatibility

Keep source CP1252-compatible and free of a byte-order mark. The starter uses
ASCII, which is compatible with CP1252. Working-tree VBA exports use CRLF;
Git's index uses normalized LF under the existing `.gitattributes` rules.
Do not enable a new encoding conversion or alter the export header as part of
formatting. Import and export procedures remain in
[Installation](../INSTALLATION.md).

<a id="review-and-validation"></a>

## 🔎 Review and validation

For a formatting-only change, compare the old and new ordered logical code
lines after excluding comments and normalizing whitespace outside strings and
explicit line continuations. Retain statement boundaries, labels, compiler
directives, identifiers, operators and literal contents in that comparison.
Review the comments against the implementation separately.

Names, signatures, visibility, constant expressions, error fields, calculations,
test cases, assertion counts, cleanup and machine-readable output must remain
unchanged. A discovered semantic defect belongs in a separate change.

Run the canonical repository gate, public API, jump and conditional-compilation
checks, plus the initializer self-test for all three profiles. Verify CRLF
working-tree exports and clean whitespace. These checks confirm source
contracts; they do not execute Excel or prove VBE importability by themselves.

For changed VBA exports, retain host compile/regression evidence for the actual
candidate using the [Excel evidence guidance](EXCEL_EVIDENCE.md). A complete
successful suite run establishes that the exercised code compiled; do not
relabel an earlier source revision's test output as a new execution.
