# 🧩 VBA house style

These conventions apply to every VBA component of the workbook. The
[repository structure](REPOSITORY_STRUCTURE.md) says where each kind of
component lives and which may call which; the
[installation guide](../INSTALLATION.md) says how source moves between the
repository and Excel. Until the project adds its own components, the
repository holds only the [VBA test harness](methodology/TEST_HARNESS.md),
its example suite and the illustrative relation tests in `tests/modules/`, so
the examples below name no existing component other than those test modules.

## Naming and visibility

Replace `<REGIME>` with the project's short name, such as `SACCR` or `LCR`,
once, in this table. The
[static checks](#static-checks) read the facade prefix from this table; while
it is `<REGIME>_`, they accept any project prefix that is not a reserved one
(`CORE_`, `TEST_`, `EX_`, `M_`).

| Element | Convention | Example |
| --- | --- | --- |
| Core standard module | `CORE_` plus subject | `CORE_Validation` |
| Public facade module | `<REGIME>_` plus subject | `<REGIME>_Functions` |
| Host-independent class | `C` plus noun | `CCalculationResult` |
| Excel adapter module | `M_` plus subject | `M_Run` |
| Document module | `ThisWorkbook`, or `sh` plus the tab name in PascalCase | `shInputs` |
| UserForm | `frm` plus role | `frmRunOptions` |
| Test module | `TEST_` plus subject | `TEST_Harness` |
| Example module | `EX_` plus subject | `EX_Inputs` |
| Procedure | PascalCase verb phrase | `ValidateInput` |
| Parameter or local variable | Descriptive PascalCase | `CalculationCurrency` |
| Constant, error constant | `UPPER_SNAKE_CASE`, `ERR_` prefix | `ERR_INVALID_INPUT` |
| Enum and member | `<REGIME>` plus noun; short member prefix | `<REGIME>RunStatus`, `rsFailed` |

- A component's name equals its exported `Attribute VB_Name` and its file
  name stem.
- VBA names are case-insensitive, and the VBE re-cases every use of a name to
  its latest spelling ([HB-01](methodology/VALIDATION_PLAN.md#host-behaviour)).
  Review case changes in a diff before accepting them.
- Single-letter names are for loop counters only.
- There are no mutable public variables.

Every component starts with `Option Explicit`. Core and test standard modules
also use `Option Private Module`. Procedures are `Private` unless another
component calls them. `Public` does not mean supported: only the declarations
listed under `[supported]` in [PUBLIC_API.txt](PUBLIC_API.txt) are the
supported API, and a change to one updates the inventory, its contract and its
tests together. Every other `Public` declaration under `src/` is listed under
`[project-public]`, so making a member `Public` is a reviewed change.

- Declare array bounds explicitly; never use `Option Base`.
- Prefer specific types to `Variant`. Use `Variant` only at a boundary that
  must carry missing, error or mixed values, and validate it there.
- Pass `ByVal` by default; document every `ByRef` argument a procedure
  changes.
- Core code has no Excel object model access and no mutable run state.

A native `Declare` needs its own compatibility review: `PtrSafe` alone does not
make a declaration correct. Use `LongPtr` for pointers and handles, and the
widths the API defines for every other argument; never replace `Long` with
`LongPtr` wholesale. Record the export's ABI, its `arguments` and `returns`,
in the [register of native declarations](REPOSITORY_STRUCTURE.md#native-declarations):
`tools/check_vba_declares.py` then compares each `Declare` of the export with
it in every compilation environment, so that a `Long` handle under VBA7 or a
`LongPtr` in a VBA6 branch is reported.

Every `Declare` imports an export listed in the
[register of native declarations](REPOSITORY_STRUCTURE.md#native-declarations),
`docs/native-declarations.json`. Read a native error from `Err.LastDllError`,
never from a declared `GetLastError`, and copy it to a local variable in the
statement directly after the call: `Err.Clear`, any `On Error` statement and
the next native call replace it. Make the call a statement of its own:
`Name args`, `Call Name(args)` or `<variable> = Name(args)`, with the name
alone or qualified by its own module's name; a call without arguments may
leave out the parentheses. `<variable>` is a variable, not an array element
or a member, whose assignment could run code. A call inside a condition, a
single-line `If`, an expression or an argument is not followed by a capture,
since the code that runs after it can replace the error first. Write the
read as `Err.LastDllError`, never as `.LastDllError` in a `With Err` block,
whose `With` statement already stands between the call and the read. A
named label between the call and the read is a jump target, so it counts as
a statement; a line number does not. When the register marks an export's
zero result as ambiguous, clear the error with `SetLastError 0` directly
before the call, and assign the result, with nothing added to it, to a local
variable, so that a zero with no error reads as success:

```vb
SetLastError 0
SlotValue = TlsGetValue(SlotIndex)
DllError = Err.LastDllError
```

`SetLastError` is a `Declare` of its own, from `kernel32`, registered like
any other export.

## Contracts, units and errors

Each component and each procedure has a [header](#layout) with the same nine
fields. `INPUTS` gives each argument's meaning, type, unit, currency or date
basis, valid range and ownership. `DEPENDENCIES` also cites the
[calculation contract](methodology/CALCULATION_CONTRACTS.md) and the source the
procedure implements, and `BEHAVIOR` or `NOTES` says which state it owns.
Comments explain invariants and reasons, not obvious assignments.

- Use `Double` for amounts and proportions, `Long` for counts, `Boolean` for
  flags, `String` for identifiers and currency codes, and `Date` for dates on
  the host side. Where a date crosses a boundary as text, state its format.
- Names or contracts distinguish decimal rates from basis points, and amounts
  in currency units from scaled display amounts.
- Never invent a rounding rule, day count, regulatory parameter or valid
  range: record it in the calculation contract first.
- Never turn missing input into zero silently, and use no undocumented
  default.

Errors raised by the project use named `ERR_` constants, allocated once in one
documented range: `vbObjectError + 2048` through `vbObjectError + 4063`. The
range `vbObjectError + 4064` through `vbObjectError + 4095` is reserved for the
test harness and its example suite (`TEST_Harness`, `TEST_Examples`).
`Err.Source` names the component and procedure. A message describes the
condition without repeating confidential input. Error codes distinguish
invalid input, unsupported scope, numerical failure, and host or cleanup
failure.

Core procedures raise documented errors. Facade entry points translate them as
their contract says: a macro returns or reports a structured failure; a
worksheet function may return a documented `CVErr` value in a `Variant`.
Worksheet functions change no Excel state and show no dialog. A facade never
turns an unexpected failure into an amount that looks valid.

`On Error Resume Next` guards one operation only. Check and clear `Err`
immediately after it, then restore the intended handler. Never use it around
a whole calculation or cleanup sequence.

## Operation and cleanup outcomes

An operation that changes host state reports both outcomes to its caller:

| Field | Meaning |
| --- | --- |
| `OperationSucceeded` | Whether the intended work completed |
| `PrimaryError` | The saved error number, source and description, or none |
| `CleanupSucceeded` | Whether every required restoration succeeded |
| `CleanupErrors` | Every restoration that failed, with its error |
| `RecoveryRequired` | True when host state could not be restored |

A successful operation whose cleanup failed is not a success, and a failed
operation stays failed even if its cleanup succeeded. Callers check both.
These are required fields of the result, not an existing type.

- Capture each setting before changing it, and record which settings were
  captured and which were changed. Register the restoration before attempting
  the change, so that a setter that fails half way is still restored.
- Never restore a value that was not captured. Restore only what the
  operation tried to change, to the captured value; a nested operation keeps
  its caller's state.
- Normal and error paths both enter cleanup. Save `Err.Number`, `Err.Source`
  and `Err.Description` into variables before any other call. The
  [error check](#static-checks), `tools/check_vba_errors.py`, reports a read
  that follows one.
- Restore each setting through its own helper and handler, and attempt every
  restoration even when one fails. Cleanup never replaces the saved primary
  error.
- Return both outcomes, or publish the cleanup failures through a documented
  output before raising the saved primary error again. If only cleanup fails,
  report a cleanup failure, not success.
- Disable the procedure's own handler with `On Error GoTo 0` before you raise
  the saved primary error again, then pass the saved number, source and
  description to `Err.Raise`. After `Resume <label>`, the
  `On Error GoTo <label>` handler is enabled again: it would trap the raise
  and resume to the same cleanup, which raises again, and Excel can hang.
- Never reset `Calculation`, `EnableEvents`, `ScreenUpdating`, `DisplayAlerts`
  or `StatusBar` to an assumed default. Capture only what the operation
  changes. Sheets and other objects an operation creates need the same
  explicit cleanup.

Failure-path tests cover partial capture, a failed change, a primary failure,
each failed restoration, and a primary and cleanup failure together.

## Source representation

- Exported `.bas`, `.cls` and `.frm` files are ASCII only, without a byte
  order mark. Represent other characters at run time, for example with
  `ChrW`, or keep them in a separately specified resource, so the source does
  not depend on a Windows code page.
- A form's binary `.frx` stays beside its `.frm`.
- Git stores VBA source with LF line endings and checks it out with CRLF
  (`.gitattributes`). Keep the `VERSION`, `BEGIN`, `END` and `Attribute` lines
  that the VBE writes.
- Indent with four spaces; no tabs and no trailing white space. Keep
  procedure declarations, `End` statements and labels at their outer level.
  Separate procedures with one blank line. The [layout](#layout) below
  gives the order of a component's parts.
- A form's header, which the VBE writes, is exempt: it keeps its bytes, such
  as the blank at the end of its `Begin` line, because the next export writes
  them again. **Repository integrity** checks the code after it.
- A generated component names its generator and is regenerated, never edited
  by hand.

`python tools/format_vba.py --write` applies the mechanical part of these
rules to the `.bas`, `.cls` and `.frm` files that Git tracks under `src/`,
`tests/modules/` and `examples/modules/`, or to the modules and folders you
name:

- CRLF line endings and one final line ending;
- no trailing white space, and four spaces for each indentation tab;
- procedure declarations, their `End` statements and labels at the outer
  level;
- one blank line after each procedure's `End` statement.

It never changes a string literal, and it keeps the `VERSION`, `BEGIN`, `END`,
`Object` and `Attribute` lines that the VBE writes. A statement keyword and a
colon, such as `Else:` or `Loop:`, is not a label and keeps its indentation.
A line that ends with a space or tab and `_` continues on the next one, in a
comment as in code. No blank line is added to or removed from one logical
line: each line a continuation carries over stays as it is, blank or not, the
blank line after an `End` statement follows the end of what it continues, and
a blank line that ends a continuation at the end of the file stays. It
neither compiles nor checks the other rules of this guide.

**Repository integrity** runs it with `--check`. Re-import a module it
rewrote into the VBE before committing, so that the workbook and the source
stay the same; where Excel evidence is bound to the module, the rewrite also
needs a new run.

A project that imports existing modules can adopt the layout one module at a
time with a baseline, `tools/vba_layout_baseline.json`, which lists the modules
not yet fully in layout. `tools/format_vba.py` (statement layout) and
`tools/check_vba_layout.py` (banners and headers) read the same file. Without
that file, every module must be in layout. To create it, run
`python tools/format_vba.py --check --no-baseline` and
`python tools/check_vba_layout.py --no-baseline`, and list each module that
either one reports, by its path from the repository root with `/` separators,
sorted and without duplicates:

```json
{
  "schema_version": 1,
  "files": [
    "src/modules/PROJ_Imported.bas"
  ]
}
```

Both tools then report each listed module as baselined, and fail on:

- a module out of layout that the baseline does not list;
- a listed module that is in layout for both tools, until you remove its entry;
- a listed path that is not a tracked module under the source directories.

To bring a listed module into layout, run
`python tools/format_vba.py --write src/modules/PROJ_Imported.bas`, add its
banners and headers by hand, re-import it, and remove its entry in the same
commit; delete the file once its list is empty. `--write` says when a module is
in statement layout but still needs its banners and headers. `--baseline FILE`
reads another file, and `--no-baseline` ignores it. A missing `--baseline`
file, a malformed one, or an entry that is absolute, uses `\` or `..`, or
breaks the order is an error with exit status 2, distinct from the exit status
1 of a layout finding.

## Layout

Write every component, new or imported, in this layout. `python
tools/format_vba.py` applies its whitespace rules, and
[`tools/check_vba_layout.py`](#static-checks) checks its banners and headers.
The `.frm` designer block that the VBE writes, and every `.frx`, stay as the
VBE wrote them.

**Top of the file.** The `Attribute` lines, one blank line, `Option Explicit`
(with `Option Private Module` and `Option Compare` where used), one blank
line, then the component header.

**Banners.** A banner is an apostrophe and 78 dashes, 79 characters in all; a
test module may use equals signs instead of dashes in its component header.
A *section banner* has an upper-case title, left-aligned after one space. A
*procedure title banner* has an upper-case verb phrase centred in the 78
columns, not the procedure's identifier. A *group banner* separates families
of procedures: a lone `'`, a banner, an empty comment line, a centred
upper-case title, an empty comment line, a banner and a lone `'`.

**Headers.** A component starts with a lone `'`, a banner, `' MODULE: <name>`
or `' CLASS: <name>` (a UserForm or a document module also uses `MODULE:`),
the nine fields and a closing banner. A procedure's header sits below its
signature, after one blank line, or directly below a signature that spans
several lines: a lone `'`, its title banner and the nine fields. Each field is
its upper-case label after one space, with its text indented three spaces and a
lone `'` between fields. Omit none; write `None` instead.

| Field | Content |
| --- | --- |
| `PURPOSE` | What it does, in one or two sentences. |
| `WHY THIS EXISTS` | The problem it solves, or why it has this shape. |
| `INPUTS` | `None`, or each parameter on its own line and its description beneath it. |
| `RETURNS` | `Nothing` for a `Sub`; else the value and its edge cases. |
| `BEHAVIOR` | What happens, in order; bullets are `-` items. |
| `ERROR POLICY` | What it raises, what it suppresses and what is best effort. |
| `DEPENDENCIES` | Components, object models and members it relies on; the calculation contract and source it implements. |
| `NOTES` | Decisions, caveats and call examples, or `None`. |
| `UPDATED` | The date of the last change to its behavior, as `YYYY-MM-DD`. |

Keep `UPDATED` when you only reformat; set it when behavior changes.

**Declarations.** Below the header, under section banners in this order and
leaving out those that are empty: `PRIVATE TYPES`, `WINDOWS API DECLARATIONS`,
`PRIVATE CONSTANTS`, `PUBLIC CONSTANTS`, `PUBLIC ENUMS`, `PUBLIC TYPES`,
`PUBLIC STATE`, `PRIVATE STATE`; a class has `PRIVATE CONSTANTS` and `PRIVATE
VARIABLES`. A component may add sections of its own with a clear name. Indent
what stands under a banner by four spaces, align the `As` keywords in one
column and the trailing comments in another within a block, and give each
declaration a trailing comment. Separate a section from the first group
banner by two blank lines.

**Procedures.** Every parameter has an explicit `ByVal` or `ByRef` and an
explicit type, as does every function, and every procedure an explicit
visibility. Separate procedures with one blank line. A procedure body is a
sequence of section banners:

- `DECLARE`: `Const PROC_NAME As String = "<Component>.<Procedure>"`, then
  every `Dim`;
- `INITIALIZE`: `On Error GoTo ErrorHandler` and the initial state;
- one section per logical step, titled with a short upper-case verb phrase,
  such as `VALIDATE INPUTS`;
- `EXIT PROCEDURE`: the clean exit, after any cleanup;
- `ERROR HANDLER`: the label at the outer level, then capture, cleanup and
  re-raise as [Operation and cleanup outcomes](#operation-and-cleanup-outcomes)
  describes.

A trivial procedure keeps only the sections it needs. Under a section, put a
one-line intent comment at four spaces, with no space after the apostrophe,
and its statements at eight; nested blocks add four spaces each. The comment
says why or what, in sentence case and without a final period.

**Statements.** Qualify library calls, `VBA.Len`, `VBA.Trim$` and the like,
preferring the `$` string variants, and test for empty with `VBA.Len(x) = 0`
or `VBA.LenB(x) = 0`. Give `InStr` and `StrComp` an explicit compare mode. Name
the counter in `Next`. Write one statement per line, with no `:` separator; a
single-line `If ... Then` is allowed for a guard or an `Exit`. Break a long
line after an operator or comma and indent its continuation four more spaces;
keep an `Err.Raise` number and source on its first line. Keep lines to about
100 characters and never past 120 outside a string literal. Release an object
variable with `Set x = Nothing` on the exit path.

The layout is not retrofitted by [`format_vba.py`](#source-representation): a
module that predates it needs its headers written by hand, and its code must
stay the same, which `tools/vba_code_identity.py` compares without comments.

## Static checks

**Repository integrity** runs the static checks below on every tracked `.bas`,
`.cls` and `.frm` file. Each also runs on its own, for example
`python tools/check_vba_source.py`; with no VBA source, each passes and says
it examined no component. The call check is advisory: **Repository
integrity** reports its findings as a warning, not a failure.

| Check | Enforces |
| --- | --- |
| `tools/check_vba_source.py` | Each component in its home with its role's prefix and a `VB_Name` equal to its file name, names unique ignoring case and at most 31 characters; ASCII only; `Option Explicit`, and `Option Private Module` in core and test modules; no `Option Base`; no tab or trailing white space, except in a form's VBE header; closed string literals; each form beside its `.frx`; LF storage with CRLF checkout; no workbook committed except `src/workbook/Template.xlsx` and the regulatory source copies under `docs/methodology/sources/` |
| `tools/check_vba_layout.py` | The [layout](#layout) of each component under the source directories, with the [baseline](#source-representation) that `tools/format_vba.py` reads: every banner is 79 characters; the component and each procedure carry a header whose nine fields are all present, once each and in order, each with text, and whose `UPDATED` is one `YYYY-MM-DD` date. It reads comments only and does not judge their content |
| `tools/check_vba_conditionals.py` | Well-formed `#If`, `#ElseIf`, `#Else` and `#End If` blocks, evaluated in each compilation environment (VBA6 Win32, VBA7 Win32 and VBA7 Win64); every `Declare` reachable under VBA7 is `PtrSafe`. A module-level `#Const Name = <value>` defines `Name` from its line on, in each environment that reaches the line, so a `#Const` inside an `#If` branch exists only where that branch is active; its value is `True`, `False`, `-1`, `0` or an expression over the compiler constants `VBA6`, `VBA7`, `Win32` and `Win64` and the module's earlier `#Const` names, with `Not`, `And`, `Or`, `=`, `<>` and parentheses. A condition is not evaluated in an environment that cannot enter its branch. A finding, so that the check fails closed: any other symbol, such as a project's conditional-compilation argument or a `#Const` used before or outside its definition; a `#Const` without a name or `=`, with a value that is not Boolean, such as `2`, that redefines a compiler constant, or defined twice in one environment |
| `tools/check_vba_jumps.py` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target, and the line number in `If x Then 100`, is a label of the same procedure, in each environment; no label is defined twice in a procedure; no `Err.Raise` or `Error` statement follows a `Resume <label>` while one of the procedure's own `On Error GoTo <label>` handlers is enabled |
| `tools/check_vba_public_api.py` | Explicit visibility on every declaration; the public declarations of `src/modules/` match the `[supported]` tier of [PUBLIC_API.txt](PUBLIC_API.txt), and every other public declaration under `src/` its `[project-public]` tier, one to one, with normalized signatures; no declaration listed in the wrong tier; no public variable or public `Declare` in a standard module; no public name in two standard modules, which Excel reports as an ambiguous name |
| `tools/check_vba_dependencies.py` | The [dependency direction](REPOSITORY_STRUCTURE.md#dependencies): core and classes have no Excel object model, file, registry or process access, dynamic call or reference to a component above them; core has no module-level or `Static` state; examples use only the facade |
| `tools/check_vba_identifiers.py` | In each environment: every name a procedure assigns (`=`, `Let`, `Set`, `LSet`, `RSet`, `For`, `For Each`) is declared, by the procedure, the module, a form control or a standard module's public name, and `ReDim` declares as in VBA; every statement-level call (`Name args`, `Call Name(args)`) names a procedure or `Declare` of the module or a public one of a standard module, or a documented built-in; no procedure, parameter, variable, constant, Enum, Enum member, Type, `Declare` or `Event` is named like a VBA keyword such as `Empty`, `Date`, `Error` or `Name`; a statement it cannot read is a finding |
| `tools/check_vba_errors.py` | In each environment, on the straight-line code of each procedure: no read of `Err.Number`, `Err.Source` or `Err.Description` after an `On Error` statement or `Err.Clear`, also written `Call Err.Clear`, which clears `Err`, or, in an error handler, after a call or member access (`Helper`, `x = Helper(1)`, `x = Helper` without parentheses, `Debug.Print`, `Application.StatusBar = False`, `New`) or an assignment to a property (`Status = x` with a `Property Let Status`), which may change it. A copy into a variable or into a field of a variable declared `As` a user-defined `Type` of the project (`Saved.Number = Err.Number`), a literal, a constant or an Enum member, also qualified (`RunStatus.rsFailed`, `Module.LIMIT`), a VBA function such as `CLng`, `CStr` or `Len`, also written `VBA.CLng`, and a VBA statement such as `ReDim`, `Erase`, `Open` or `Print #` change nothing. After `On Error Resume Next`, `Err` belongs to the last statement that runs anything, since any can fail, even `Number = 1`: a read checks that statement, and only a declaration, a comment, `With Err`, `End With` or a write of a property of `Err` whose value runs nothing (`Err.Number = 0`) comes between `On Error` or `Err.Clear` and a read reported as stale. A read inside the arguments of a call, as in `CLng(Err.Number)`, comes before the call, and a call or member access that the same statement evaluates before a read counts as one before it; assigning a property of `Err` is not a read, and `.Number` inside `With Err` reads `Err.Number` |
| `tools/check_vba_declares.py` | In each environment: every `Declare` imports an export that `docs/native-declarations.json` lists with its library, matched by its `Alias`, or else its name, ignoring case; no `Declare` of `GetLastError`; `Err.LastDllError` is read only as `<local> = Err.LastDllError`, into a `Dim` or `Static` variable, in the statement directly after a statement that is one `Declare` call (`Name args`, `Call Name(args)` or `<variable> = Name(args)`, the name alone or qualified by its module's name), never after a call inside a condition, a single-line `If`, an expression or an argument, or one assigned to an array element or a member, and never as `.LastDllError` in a `With Err` block or through another qualifier; a call to an export registered as ambiguous on zero assigns its result, and nothing else, to a local variable, directly after `SetLastError 0` and directly before that capture, and no `Declare Sub` is registered as ambiguous on zero; a `Declare` whose export has an ABI in the register has its number of parameters, each one's type and `ByVal` or `ByRef`, and `Sub` or `Function` with its return type, with a registered `pointer` as `LongPtr` under VBA7 and `Long` under VBA6; the register has the documented shape |
| `tools/check_vba_calls.py` (advisory) | In each environment, for every call that resolves to a procedure of the project, a statement call (`Name args`, `Call Name(args)`, `Module.Name args`) or a call with parentheses in an expression (`x = Helper(1)`): no more arguments than the procedure accepts and every required one passed, counting `Optional` parameters, a `ParamArray` and named arguments (`Name:=Value`); no named argument the procedure does not declare; no qualified call to a `Private` procedure of another module |

What they cannot prove:

- they neither compile nor run VBA, and they resolve calls by name only: a
  call made through a variable or an object is missed. The identifier check
  resolves the names of statement-level calls only; only the call check
  resolves qualified calls and calls inside an expression, and checks
  arguments;
- they check the three compilation environments statically, not on a 32-bit
  or 64-bit host;
- a private procedure name may repeat across modules, as VBA allows; only a
  public name in two standard modules is a collision;
- the identifier check reads neither member access (`x.y = 1`, `.y` in a
  `With` block) nor qualified calls (`Module.Name`, `obj.Method`), nor calls
  inside an expression (`x = Helper(1)`), and it does not see the controls of
  a worksheet, which a document module may name. Its built-in callees, such as
  `MsgBox` and `DoEvents`, are listed in `BUILTIN_CALLS`, a form's or document
  module's own methods, such as `Hide` or `Protect`, in `HOST_CALLS`, and the
  host properties assignable without a declaration in `HOST_TARGETS`, such as
  `Range`, `Cells` and `Selection` anywhere, a UserForm's `Caption` or
  `Width` in a form, a workbook's `Saved` or a worksheet's `Visible` in a
  document module. Extend the list when code calls or assigns a built-in it
  does not know. A procedure of another VBA project that this one references
  is not seen either; call it qualified by its module;
- the call check resolves only a name that no parameter, variable,
  constant or control shadows, to a procedure of the calling module, the one
  public procedure of that name in a standard module, or, qualified, a
  procedure of the standard module named. It says nothing about a member
  call (`obj.Method`, `.Add x` in a `With` block), an array element
  (`Values(1)`), a `Property` procedure, a `Declare`, a built-in, a name that
  two standard modules make public, or a call through `Application.Run` or
  `CallByName`. Parentheses after a function without parameters may index its
  result, so their arguments are not counted, and it checks neither argument
  types nor `ByRef` passing;
- the jump check follows the error handler from a `Resume <label>` target
  along the text of the procedure, merging the branches of `If`,
  `Select Case` and loop blocks. It follows no jump, and `Resume` and
  `Resume Next` start no path, because the statement they continue at is
  known only at run time. A `Resume <label>` that runs only under a
  condition, as in a retry, starts no path either. Where branches, later
  passes of a loop or a `GoSub` may leave the handler in different states,
  it reports nothing; after a `GoSub`, that is the case unless the
  subroutine runs straight to its `Return` with no branch, jump, raise or
  `On Error` statement. It checks a raise in a loop body in the first pass
  only, and never inside a subroutine. After a bare `Do ... Loop` around the
  `Resume` target, it follows only an `Exit Do` reached in the first pass;
- the error check follows straight-line code only, and says nothing where
  it cannot prove a read stale:
  - a label and a branch or loop statement (`If`, `Select Case`, `Do`,
    `For`, `While`, `Exit`, `Resume`, `GoTo`) end what it tracks, and it
    does not read the conditions of `ElseIf`, `Case` and `Loop While`, so a
    stale read after a branch, a loop or a jump is missed. A single-line `If`
    is read as its condition and then each branch on its own, from the state
    after the condition;
  - inside a function, `Fetch(True)` calls the function again, unless its
    declared result is an array or a `Variant`, as when it has no type: then
    `Fetch(1)` counts as an element of the result, and a recursive call is
    missed;
  - it cannot tell a procedure that keeps `Err` from one that clears it, so
    in a handler it treats every call, member access and property
    assignment as changing `Err`. A name without parentheses that is not a
    variable, a constant, an Enum member or a component counts as a call: a
    function, property or `Declare` of the project, or a host global such
    as `ActiveCell`; a name that starts with `vb` or `xl` counts as a
    constant. A dotted name counts as a member access unless it reads a
    field of a variable declared `As` a user-defined `Type` of the project
    (`Saved.Number`, `Saved(i).Inner.Field`, `.Number` in `With Saved`), an
    Enum member (`RunStatus.rsFailed`, `Module.RunStatus.rsFailed`), or a
    constant, Enum member or variable of a standard module that its name
    qualifies (`Module.LIMIT`): a variable of a class or object type, a
    class or document module and `Me` stay member accesses, and so does an
    index before a dot that holds a parenthesis, as in
    `Saved(Count(1)).Number`. In a dotted name, an index is an element only
    on a name declared as an array (`Name() As T`, `Name(1 To 3) As T`, a
    `ReDim`): on an object, a `Variant` or an undeclared type it may call a
    default property, as `Owner(1)` does on a `Collection`, so it counts as
    a member access, even on a `Variant` that holds an array. A bare `C(1)`,
    with no dot, is read as an element, even on a `Collection`, whose `Item`
    it calls. A function's own name holds its result: in a `Function` or
    `Property Get` of a user-defined `Type`, `TryRun.Number = Err.Number`
    copies into that result. The part a dotted name reaches must
    hold data, a value type such as `Long` or `String`, an Enum or a
    user-defined `Type`: reading or assigning an object, a class or a
    `Variant` calls its default property, so `Number = rec.Owner` counts as
    a member access, even when the `Variant` holds a plain value. So does
    `Set Book = rec.Owner`, since a variable declared `As New` runs its
    class's `Class_Initialize` when it is first read. An assignment to
    anything but a variable, such a field or value, an array element,
    `Date`, `Time` or `Mid` counts as setting a property, and a statement
    that starts with a name the check does not list as a VBA statement
    counts as a call, `Load` and `Unload` included, since they run the
    form's event procedures. A listed VBA statement written with `Call`, as
    in `Call SaveSetting(...)`, is still that statement. Operators, an array
    element and the listed VBA statements can fail but run no code, so in a
    handler they do not;
  - as in VBA, a procedure, property or `Declare` of the project shadows a
    VBA function, statement or constant of its name, such as a project
    `Format`, while `VBA.Format` or `VBA.Strings.Format` still names the VBA
    function; `Err` and `Erl` always mean VBA's. Within the project, a name
    means what the nearest scope level declares: the procedure's parameters
    and variables, then its module's declarations, then the public ones of
    the other components, then a component, so a module's property `Saved`
    hides another module's public variable `Saved`. A dotted name counts as
    a member access when it starts, unqualified, with a public variable of
    two components, when it goes through a `Type` or an `Enum` that two
    components declare, or when its type is qualified by anything but the
    component that declares it, such as `As VBA.Collection`; a name
    qualified by its component, as `Module.Kept`, is resolved there;
  - within a statement, it follows VBA's evaluation order: a read inside the
    arguments of a call runs before the call, and only what runs before a
    read counts for it, as between statements. In a handler,
    `Report Helper(), Err.Number` is a finding and
    `Report Err.Number, Helper()` is not. An operator written before a read
    may run after it, as `&` in `a & b / Err.Number`; it counts as running
    before;
  - reading or writing any property of `Err` changes nothing, but only a
    read of `Number`, `Source` or `Description` is checked: a stale read of
    `HelpFile`, `HelpContext` or `LastDllError` is not reported. A VBA module
    name alone, as in `Strings.Len`, counts as VBA's only when no parameter,
    variable, component or procedure of the project has that name;
  - it does not follow `Err` read through its default member (`n = Err`,
    `If Err Then`), `Err` passed to a procedure, or `Erl`. The default
    member of another object (`x = Target` for a `Range`) counts as a copy,
    as does a form control assigned without a qualifier, and
    `Set x = Nothing`, although it may run the released object's
    `Class_Terminate`;
  - after `On Error Resume Next`, it counts every statement that runs
    anything, even `Number = 1`, as one that can fail, and does not check
    which of several such statements a read was meant for;
- the native declaration check reads the statements of a procedure in
  order: it follows no jump, so a `GoTo` past the call is not seen, nor an
  error replaced inside the call itself. It compares a `Declare`'s
  signature only with an ABI the register records, and it cannot know
  whether that ABI is the one the library documents. A `Declare` is private
  to its module, so it resolves a call to a `Declare`
  of the calling module only, by its name alone or qualified by the
  module's name. It cannot tell a variable from a `Property` of the same
  class assigned without a qualifier, and it does not report a register
  entry that no `Declare` uses;
- the formatter and these checks enforce the mechanical rules only; review
  covers the rest of this guide.

While the call check is advisory, read its findings in the **Repository
integrity** output and fix each confirmed one, or record it as an issue. My
rule for making it blocking: once it has run on a project's VBA with no
false positive and every finding it reported there is fixed, I remove
`advisory=True` from its entry in `tools/check_repository.py`. If a false
positive appears later, I make it advisory again and record the case as an
issue.

Changed VBA is imported, compiled and run on the supported host before its
behaviour is accepted. Repository checks are not compilation. Record the full
commit, the Excel build and bitness, the references and the results, as the
[validation plan](methodology/VALIDATION_PLAN.md#evidence-record) describes.
