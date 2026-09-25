> [!IMPORTANT]
> **Migration scope — 2026-09-24.** Imported from frozen source commit `f26450d`.
> The frozen source labelled this contract v0.0.2 and reserved calendars for
> v0.0.3. K-PRICING adapts those labels to its own milestone sequence: this
> contract targets v0.0.4 and `KPR_Cal_*` is reserved for v0.0.5 (see the
> [roadmap](ROADMAP.md)). Historical evidence refers to the frozen source and
> does not claim destination certification. The frozen implementation is now migrated
> into K-PRICING; exact-source Windows Excel compilation and parity passed on one
> Windows 64-bit host in [K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).
> [Migration provenance](MIGRATION_PROVENANCE.md) records hashes, attribution
> and the limited editorial adaptations.

<div align="center">

# 📅 KPR Date-Layer Behavioral Contract

### Normative v0.0.4 reference for the Gregorian date primitive layer

**Strict inputs · Deterministic dates · Native Excel errors · Shape-preserving arrays**

<br>

![Milestone](https://img.shields.io/badge/Milestone-v0.0.4-217346?style=for-the-badge)
![Authority](https://img.shields.io/badge/Authority-Normative-6f42c1?style=for-the-badge)
![Surface](https://img.shields.io/badge/Public_Surface-22_Functions-0969da?style=for-the-badge)
![Date Window](https://img.shields.io/badge/Date_Window-1900--03--01_to_9999--12--31-d97706?style=for-the-badge)
![Capacity](https://img.shields.io/badge/Array_Cap-100%2C000-2ea44f?style=for-the-badge)

</div>

---

> [!IMPORTANT]
> **Normative status.** This is the v0.0.4 behavioral contract for
> implementation, registration, regression tests, demo material, and user
> documentation.

This document freezes observable behavior. It and
[the roadmap](ROADMAP.md) are complementary and non-overlapping in authority:

- the roadmap governs scope, milestones, and sequencing;
- the [migration records](MIGRATION_PLAN.md) govern the historical migration
  architecture and evidence;
- this contract governs observable behavior; and
- where another document states a behavioral rule, this contract governs.

> [!WARNING]
> If descriptive material elsewhere conflicts with this document, this document
> governs v0.0.4 behavior. A contract change requires an explicit, reviewed
> behavioral decision rather than an incidental implementation change.

> [!NOTE]
> The migrated calculation implementation implements this contract. Static checks
> and the focused Excel results recorded during its implementation
> support those conclusions. That evidence is deliberately narrower
> than release certification: it does not make the current `main` branch a
> supported release or certify the final candidate's complete import,
> compilation, regression, UI, demo, and source round trip. Those exact-source
> claims require separate certification evidence.

## 🎯 1. Scope and terminology

The v0.0.4 date layer contains pure proleptic-Gregorian date primitives,
strict date and control parsing, date arithmetic, weekday locators, and pillar
conversion. Its supported date window is inclusive:

```text
1900-03-01 .. 9999-12-31
```

The lower bound deliberately excludes Excel serial 60, the fictitious
29-Feb-1900, and every earlier serial/date. All accepted date inputs and all
successful date outputs are normalized to date-only values.

In this document:

- **value argument** means a required argument whose values may determine the
  output shape. The first 21 functions below have one or more value arguments.
- **control argument** means an optional `Opt_` argument. A control never
  determines or expands the output shape.
- **scalar** means a non-array value or a one-cell/one-element wrapper.
- **non-scalar** means a multi-cell Range or multi-element in-memory array.
- **call-level error** means one scalar Excel error returned for the whole call;
  no output array is produced.
- **element-level error** means an Excel error stored only at the affected
  output position while valid neighbours continue to evaluate.

## 🧩 2. Exact supported public surface

Every public declaration returns `Variant` so it can return either its semantic
success type or a native Excel error. Every public argument is `ByVal`. Required
value arguments precede optional controls, and every optional control is last.

```vb
Public Function KPR_Dates_DayOfWeek(ByVal DateIn As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_DaysInMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_DaysInYear(ByVal YearIn As Variant) As Variant
Public Function KPR_Dates_BeginOfMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfMonth(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_BeginOfQuarter(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfQuarter(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_BeginOfYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_EndOfYear(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsMonthEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsQuarterEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsYearEnd(ByVal DateIn As Variant) As Variant
Public Function KPR_Dates_IsLeapYear(ByVal YearIn As Variant) As Variant
Public Function KPR_Dates_AddDays(ByVal DateIn As Variant, ByVal nDays As Variant) As Variant
Public Function KPR_Dates_AddWeeks(ByVal DateIn As Variant, ByVal nWeeks As Variant) As Variant
Public Function KPR_Dates_AddMonths(ByVal DateIn As Variant, ByVal nMonths As Variant, Optional ByVal Opt_KeepEOM As Variant = False) As Variant
Public Function KPR_Dates_AddYears(ByVal DateIn As Variant, ByVal nYears As Variant, Optional ByVal Opt_KeepEOM As Variant = False) As Variant
Public Function KPR_Dates_NthWeekdayOfMonth(ByVal YearIn As Variant, ByVal MonthIn As Variant, ByVal WdIndex As Variant, ByVal n As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_LastWeekdayOfMonth(ByVal YearIn As Variant, ByVal MonthIn As Variant, ByVal WdIndex As Variant, Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
Public Function KPR_Dates_PillarFromDates(ByVal StartDate As Variant, ByVal EndDate As Variant, Optional ByVal Opt_Rounding As Variant = "NEAREST") As Variant
Public Function KPR_Dates_DateFromPillar(ByVal StartDate As Variant, ByVal Pillar As Variant) As Variant
Public Function KPR_Dates_HostDateSystem() As Variant
```

The plural pre-release name `KPR_Dates_DatesFromPillar` is replaced by singular
`KPR_Dates_DateFromPillar`. There is no compatibility alias. There are no
`_Spill` twins and no `KPR_Dates_Spill.bas` module.

| # | Supported name | Semantic scalar result | Vectorization |
| ---: | --- | --- | --- |
| 1 | `KPR_Dates_DayOfWeek` | `Long`, 1 through 7 | Array-capable |
| 2 | `KPR_Dates_DaysInMonth` | `Long`, 28 through 31 | Array-capable |
| 3 | `KPR_Dates_DaysInYear` | `Long`, 365 or 366 | Array-capable |
| 4 | `KPR_Dates_BeginOfMonth` | `Date` | Array-capable |
| 5 | `KPR_Dates_EndOfMonth` | `Date` | Array-capable |
| 6 | `KPR_Dates_BeginOfQuarter` | `Date` | Array-capable |
| 7 | `KPR_Dates_EndOfQuarter` | `Date` | Array-capable |
| 8 | `KPR_Dates_BeginOfYear` | `Date` | Array-capable |
| 9 | `KPR_Dates_EndOfYear` | `Date` | Array-capable |
| 10 | `KPR_Dates_IsMonthEnd` | `Boolean` | Array-capable |
| 11 | `KPR_Dates_IsQuarterEnd` | `Boolean` | Array-capable |
| 12 | `KPR_Dates_IsYearEnd` | `Boolean` | Array-capable |
| 13 | `KPR_Dates_IsLeapYear` | `Boolean` | Array-capable |
| 14 | `KPR_Dates_AddDays` | `Date` | Array-capable |
| 15 | `KPR_Dates_AddWeeks` | `Date` | Array-capable |
| 16 | `KPR_Dates_AddMonths` | `Date` | Array-capable |
| 17 | `KPR_Dates_AddYears` | `Date` | Array-capable |
| 18 | `KPR_Dates_NthWeekdayOfMonth` | `Date` | Array-capable |
| 19 | `KPR_Dates_LastWeekdayOfMonth` | `Date` | Array-capable |
| 20 | `KPR_Dates_PillarFromDates` | `String` | Array-capable |
| 21 | `KPR_Dates_DateFromPillar` | `Date` | Array-capable |
| 22 | `KPR_Dates_HostDateSystem` | `Long`, 1900 or 1904 | Scalar-only; no value argument |

Exactly these 22 names are supported public API in v0.0.4.

## 🧾 3. Date, integer, control, and token inputs

### 📅 3.1 Date-value matrix

Date parsing is independent of Windows and Excel locale. Production code must
not use `CDate`, `IsDate`, or an equivalent permissive conversion to interpret
text.

| Input at one value position | Condition ID | Result |
| --- | --- | --- |
| Incoming native Excel error (`CVErr`) | `INPUT_ERROR_PROPAGATED` | Propagate the same error value verbatim. |
| Native VBA `Date` | — | Accept, remove the time component, then apply the supported-window gate. |
| Native numeric scalar | — | Interpret under the call's documented 1900 serial contract, remove any fractional time component, then apply the supported-window gate. |
| Text exactly `YYYY-MM-DD` | — | Parse ASCII digits and hyphens only; validate year, month, day, month length, and leap day without rollover. |
| Locale-formatted text such as `31/12/2026` or `12/31/2026` | `DATE_TEXT_LOCALE` | `#VALUE!`. |
| Numeric-looking text such as `61`, `45292`, or `45292.5` | `DATE_TEXT_NUMERIC` | `#VALUE!`; never reinterpret it as a serial. |
| Malformed ISO text, empty text, or whitespace-padded text | `DATE_TEXT_FORMAT` | `#VALUE!`. |
| Syntactically exact but impossible ISO date such as `2025-02-29` | `DATE_TEXT_IMPOSSIBLE` | `#VALUE!`. |
| Required blank cell or `Empty` | `INPUT_BLANK_REQUIRED` | `#VALUE!`. |
| `Null` or Boolean | `DATE_TYPE_REJECTED` | `#VALUE!`. |
| Accepted date value outside the supported window, including serial 60 and below | `DATE_WINDOW` | `#NUM!`. |
| Range or supported in-memory array | — | Classify through the shape rules in section 5, then apply this matrix per element. |
| Non-Range object or unsupported array form | `SHAPE_UNSUPPORTED` | Call-level `#VALUE!`. |

A numeric serial carrying a time component, such as a cell holding `NOW()`, is
accepted: the time is removed first and the supported-window gate is applied to
the resulting date-only value. Truncation is never applied to integer
arguments, which reject fractions outright.

ISO text is exactly ten characters. It accepts no alternate separator, omitted
zero, localized digit, time suffix, leading/trailing whitespace, or DateSerial
rollover. Whitespace-padded and time-suffixed text are `DATE_TEXT_FORMAT`;
`DATE_TEXT_LOCALE` is reserved for text shaped like a locale-formatted date.

Component existence is decided before the window. Text whose year is inside
1900 through 9999 but which names a date that does not exist is
`DATE_TEXT_IMPOSSIBLE` and returns `#VALUE!`, even when the named date would
also fall outside the window. `1900-02-29` is therefore `DATE_TEXT_IMPOSSIBLE`
rather than `DATE_WINDOW`: it is the fictitious Excel leap day, and reporting it
as merely out of range would understate the problem. A year outside that span is
`DATE_WINDOW` without any component check, because a year below 1900 cannot be
constructed safely. A syntactically exact ISO date outside the supported window returns
`#NUM!` under `DATE_WINDOW`; an impossible component such as `2025-02-29`
returns `#VALUE!` under `DATE_TEXT_IMPOSSIBLE`.

### 🔢 3.2 Integer-value matrix

`nDays`, `nWeeks`, `nMonths`, `nYears`, `YearIn`, `MonthIn`, `WdIndex`, and `n`
use one strict integer parser at each output position. `YearIn` is the argument
of `DaysInYear`, `IsLeapYear`, `NthWeekdayOfMonth`, and `LastWeekdayOfMonth`.

| Input | Condition ID | Result |
| --- | --- | --- |
| Native numeric value that is mathematically integral and inside the VBA `Long` range | — | Accept as `Long`. |
| Fractional numeric value inside the VBA `Long` range | `INTEGER_FRACTION` | `#VALUE!`; no truncation, banker's rounding, or other rounding is permitted. |
| Numeric value outside the VBA `Long` range | `INTEGER_RANGE` | `#NUM!`. |
| Boolean, `Date`, text (including numeric-looking text), blank, `Empty`, `Null`, or unsupported object | `INTEGER_TYPE_REJECTED` | `#VALUE!`. |
| Incoming native Excel error | `INPUT_ERROR_PROPAGATED` | Propagate verbatim at that output position. |

After parsing, function-specific domains apply:

- `YearIn` is 1900 through 9999; a value outside that range returns `#VALUE!`
  under `DOMAIN_YEAR`. For the weekday locators, a requested result before
  1900-03-01 is still outside the supported window and returns `#NUM!` under
  `RESULT_WINDOW`. `DaysInYear` and `IsLeapYear` construct no date, so no
  window gate applies to them and year 1900 is fully in their domain.
- `MonthIn` must be 1 through 12 (`DOMAIN_MONTH`), `WdIndex` must be 1 through 7
  (`DOMAIN_WEEKDAY`), and occurrence `n` must be 1 through 5
  (`DOMAIN_OCCURRENCE`). A parsed integer outside one of these argument domains
  returns `#VALUE!`.
- Date-shift integers may be negative, zero, or positive. A valid shift that
  cannot produce a supported result returns `#NUM!` under `RESULT_WINDOW`.

Range precedes integrality when both apply. A numeric value outside the `Long`
range returns `INTEGER_RANGE` and `#NUM!` even when it also has a fractional
part. Only after the range test may integrality be checked and the conversion
performed.

### 🎛️ 3.3 Optional-control matrix

An `Opt_` argument may be omitted, a scalar, or a 1x1 Range/array. It never
vectorizes. A multi-cell or multi-element optional argument returns call-level
`#VALUE!` under `CONTROL_NOT_SCALAR`.

Omitting a control and supplying `Empty`, including a reference to a blank
cell, both select the documented default. This is the one place where `Empty` is
not contract-invalid: at a required value position it remains `#VALUE!` under
`INPUT_BLANK_REQUIRED`.

| Control | Accepted values | Default | Invalid value |
| --- | --- | --- | --- |
| `Opt_WeekBaseMonday` | Omitted, `Empty`, or native Boolean `True`/`False` | `True` | Call-level `#VALUE!` |
| `Opt_KeepEOM` | Omitted, `Empty`, or native Boolean `True`/`False` | `False` | Call-level `#VALUE!` |
| `Opt_Rounding` | Omitted, `Empty`, or text `NEAREST`, `FLOOR`, or `CEILING`; outer whitespace is trimmed and the token is compared case-insensitively | `NEAREST` | Call-level `#VALUE!` |

The Boolean controls accept a native Boolean only. Numeric `0` and `1`,
Boolean-looking text such as `"TRUE"`, `Null`, dates, and objects are rejected
with call-level `#VALUE!` under `CONTROL_TYPE_REJECTED`. `Opt_Rounding` accepts
text only and rejects every other type under the same identifier; a text value
that is not one of the three recognized tokens returns call-level `#VALUE!`
under `CONTROL_TOKEN_UNKNOWN`.

An incoming scalar Excel error supplied as an optional control propagates
verbatim as a call-level result under `CONTROL_ERROR_PROPAGATED`.

### 🏷️ 3.4 Pillar-token matrix

The accepted grammar is deliberately wider than the emitted grammar. Parsing
accepts every form a caller may reasonably write; formatting emits one canonical
subset.

#### Accepted grammar

At each `Pillar` value position, `KPR_Dates_DateFromPillar` accepts
case-insensitive ASCII text with this grammar:

```text
(ON | O/N | TN | T/N) | [+|-] [0-9]+[YMWD] {1,4 components, each unit at most once}
```

- Outer whitespace is trimmed before parsing. Internal whitespace is rejected:
  `1 M` and `1M 2W` are malformed.
- The optional sign applies to the complete token.
- `ON` and `O/N` mean `1D`; `TN` and `T/N` mean `2D`. Aliases are matched whole
  and never combined with other components.
- **An alias never carries a sign.** `-ON` and `+T/N` are rejected under
  `PILLAR_ALIAS_SIGNED`. An alias names a fixed point at the short end of the
  curve rather than a quantity, so there is nothing for a sign to negate. A
  caller wanting a backward one-day shift writes `-1D`, which is unambiguous.
- Units are case-insensitive and may appear in any order, so `2W3D` and `3D2W`
  are the same interval.
- **A unit may appear at most once.** `1M2M` is rejected under
  `PILLAR_DUPLICATE_UNIT` rather than accumulated into `3M`. A repeated unit is
  far more likely a typing error than an intended sum.
- Year and month components form one signed month delta; week and day components
  form one signed day delta. The month delta is applied first with clip
  semantics, followed by the exact day delta.
- No decimal quantity, missing quantity, unknown unit, or partial parse is
  accepted.

| Token condition | Condition ID | Result |
| --- | --- | --- |
| Non-text payload | `PILLAR_TYPE_REJECTED` | `#VALUE!`. |
| Malformed token: internal whitespace, missing quantity, decimal quantity, unknown unit, sign with no body, empty text, or partial parse | `PILLAR_TOKEN_MALFORMED` | `#VALUE!`. |
| Repeated unit such as `1M2M` or `3D4D` | `PILLAR_DUPLICATE_UNIT` | `#VALUE!`. |
| Signed alias such as `-ON` or `+T/N` | `PILLAR_ALIAS_SIGNED` | `#VALUE!`. |
| Valid token whose aggregate or resulting date falls outside the numerical or supported date domain | `PILLAR_AGGREGATE_RANGE` | `#NUM!`. |

#### Emitted grammar

`KPR_Dates_PillarFromDates` emits only this canonical subset:

```text
[-] (0D | [1-9][0-9]*D | [1-3]W | [1-9][0-9]*Y | [1-9][0-9]*M | [1-9][0-9]*Y[1-9][0-9]*M)
```

Day tokens are emitted only for an absolute interval shorter than seven days.
Week tokens are emitted only for `1W`, `2W`, and `3W`; see the week-family cap
in section 8.4. An alias, a leading `+`, outer whitespace, a mixed
week-and-month form such as `1M2W`, and a zero-quantity component are never
emitted. Every emitted token is re-parseable by `DateFromPillar`; the reverse
does not hold, because the accepted grammar is wider.

## 🖥️ 4. Caller and workbook date-system contract

Every public date-layer call classifies `Application.Caller` once at its public
boundary, before element conversion or traversal. Date-system authority comes
only from an identifiable worksheet caller:

1. If `Application.Caller` identifies a worksheet `Range`, read
   `Application.Caller.Parent.Parent.Date1904` from that exact workbook.
2. If that workbook uses the 1900 date system, evaluate normally.
3. If that workbook uses the 1904 date system, every value-taking public
   function returns one call-level library-produced `#N/A` under
   `HOST_DATE1904`. It must never return a plausible value shifted by
   1,462 days.
4. If an identifiable worksheet host should be readable but its date system
   cannot be resolved reliably, return one call-level library-produced `#N/A`
   under `HOST_UNRESOLVED`. This condition is reserved for exactly that
   sequence: a worksheet `Range` was identified, and reading its own
   workbook's date system then failed. A failure while obtaining or
   classifying `Application.Caller` itself is not `HOST_UNRESOLVED`; it
   follows the same path as any other non-`Range` caller, meaning no
   worksheet host could be identified and the documented 1900 serial
   contract applies. That raising case belongs in the certification probe matrix.
5. If no worksheet host can be identified, apply the documented 1900 serial
   contract. Never consult `ActiveWorkbook`, `ThisWorkbook`, or another
   unrelated workbook as a fallback.

Direct VBA, the Immediate window, `Application.Run`, and the regression harness
are the certified non-Range uses of the no-worksheet-host path. The caller owns
the interpretation of values it constructs. In particular, passing a Range as
a VBA argument does not turn a direct-VBA invocation into a worksheet-Range
caller.

A non-Range `Application.Caller` is not proof of direct VBA. Macro-attached
shapes may provide a String caller, and data validation, chart series, defined
names, and related Excel evaluation contexts may provide Error or other caller
forms. Those contexts are probed during Windows certification and are outside
the v0.0.4 compatibility claim.

### 🔎 `KPR_Dates_HostDateSystem()`

`KPR_Dates_HostDateSystem()` is the sole deliberately volatile date-layer
function. It calls `Application.Volatile True` at entry and returns:

- `1900` for an identifiable 1900 worksheet caller;
- `1904` for an identifiable 1904 worksheet caller;
- `1900` when no worksheet host can be identified, under the documented 1900
  serial contract; or
- `#N/A` only when an identifiable worksheet host should be readable but its
  date system cannot be resolved reliably.

Changing the caller context or workbook date system and performing an ordinary
recalculation must re-evaluate the diagnostic; a full calculation rebuild is
not required. The other 21 functions remain non-volatile.

## 🧮 5. Scalar, array, shape, and capacity contract

### 📐 5.1 Supported shapes

- A scalar or a 1x1 Range/array is scalar.
- A one-dimensional initialized VBA array is interpreted as a `1xN` row.
- A two-dimensional array or a contiguous single-area Range preserves its
  exact rows, columns, and orientation.
- A scalar value argument expands to the target shape selected by the
  non-scalar value arguments.
- Multiple non-scalar value arguments must have exactly identical row and
  column dimensions.
- There is no row/column outer product, implicit cross-broadcasting, jagged
  traversal, multi-area Range support, or higher-dimensional support.
- A supplied Range is evaluated at its requested dimensions. It is never
  intersected with or shortened to `UsedRange`.

An all-scalar call returns a scalar `Variant`, not a 1x1 array. A non-scalar
call returns a two-dimensional `Variant` array even when the resolved shape is
one row or one column.

### ⚠️ 5.2 Call-level and element-level outcomes

The public boundary applies these stages once per call:

1. classify the caller and enforce the host/date-system rule;
2. classify value/control shapes and reject unsupported wrappers;
3. validate scalar controls and resolve exact-shape broadcasting;
4. enforce the output-element cap; and
5. traverse and evaluate elements.

Stage 2 classifies every value argument from its type and dimensions alone and
reads no content. One shape rejection necessarily follows stage 4: an array
whose elements are themselves arrays cannot be detected without inspecting
elements, and inspection is forbidden before the cap, so a jagged input is
reported as `SHAPE_UNSUPPORTED` during materialization rather than at stage 2.

Consequently:

- a 1904 or unreadable identifiable worksheet host returns call-level `#N/A`
  (`HOST_DATE1904`, `HOST_UNRESOLVED`);
- an unsupported input shape, shape mismatch, or non-scalar `Opt_` argument
  returns call-level `#VALUE!` (`SHAPE_UNSUPPORTED`, `SHAPE_MISMATCH`,
  `CONTROL_NOT_SCALAR`);
- a resolved target of at most 100,000 elements is permitted;
- a resolved target of 100,001 or more elements returns call-level `#NUM!`
  (`CAPACITY_EXCEEDED`), decided from the dimensions alone before any Range is
  read or any element is inspected;
- a blank or otherwise invalid required element returns `#VALUE!` only at that
  position;
- an element result outside the supported numerical/date domain returns
  `#NUM!` only at that position; and
- an incoming native error propagates unchanged at that position without
  suppressing valid neighbours (`INPUT_ERROR_PROPAGATED`).

Output positions are traversed row-major: every column of the first row,
then the second row, and so on. The order is deterministic and observable only
through evaluation sequence, since evaluation has no side effects.

Where more than one value argument at the same output position contains an
incoming error, the first error in signature order is propagated. Call-level
failures necessarily take precedence because no element traversal occurs.

After call-level guards pass, value arguments at each output position are
evaluated in signature order. The first failing argument determines that
element's result, whether the failure is an incoming native error or a
library-classified parse, domain, or range failure.

## 🚦 6. Native-error taxonomy and provenance

Public functions return intentional native Excel errors, never message strings,
`MsgBox` output, or plausible-looking fallback values.

| Error | Library condition |
| --- | --- |
| `#VALUE!` | Uninterpretable or contract-invalid input: malformed/non-ISO text, impossible ISO date, prohibited type, blank required value, fractional integer, invalid Boolean/control/domain token, unsupported shape, or shape mismatch. |
| `#NUM!` | Well-formed value outside the supported numerical/date domain: date/result outside 1900-03-01 through 9999-12-31, integer outside `Long`, impossible valid shift, absent requested fifth weekday, oversized aggregate, or output target above 100,000 elements. |
| Library-produced `#N/A` | A result is unavailable in the identifiable worksheet host configuration, including a 1904 worksheet call or an unreadable identifiable worksheet date system. |
| Any incoming Excel error | Propagate the same native error value verbatim, subject to the call-level guards in section 5. |

A library-produced host-configuration `#N/A` and a propagated incoming `#N/A`
are value-identical Excel errors. Provenance cannot be recovered from the
returned value alone. Test evidence must retain distinct case IDs and
originating-condition metadata. For an identifiable worksheet caller, volatile
`KPR_Dates_HostDateSystem()` supplies context: `1904` identifies host refusal;
`1900` leaves propagation or another documented input condition as the source.

Provenance is carried by the condition identifiers in section 7, not by the
returned error value.

## 🆔 7. Condition identifier registry

Every condition that produces an error or reports host unavailability has a
stable semantic identifier. Successful evaluation has no identifier.

Identifiers exist so that fixtures, the evidence schema, and certification
records can cite an originating condition that the returned Excel value cannot
express. Two conditions that return the same Excel error remain distinguishable
by identifier; `HOST_DATE1904` and `INPUT_ERROR_PROPAGATED` both surface as
`#N/A` when the propagated error is `#N/A`, and only the identifier separates
them.

Registry rules:

- Condition ID, expected Excel error, and call/element level are three separate
  fields. Never merge them, and never derive one from another.
- Identifiers are semantic, never numbered.
- An identifier is never renamed, renumbered, or reused. A retired identifier
  stays retired and its meaning is not reassigned to a different condition.
- A new documented condition requires a new identifier in this registry.
- Fixtures (#38), the evidence schema (#39), and certification records cite
  these identifiers.

| Condition ID | Condition | Excel error | Level |
| --- | --- | --- | --- |
| `DATE_TEXT_FORMAT` | Text is not exactly `YYYY-MM-DD`: alternate separator, omitted zero, time suffix, padding, empty text | `#VALUE!` | Element |
| `DATE_TEXT_LOCALE` | Locale-formatted date text | `#VALUE!` | Element |
| `DATE_TEXT_NUMERIC` | Numeric-looking text offered as a date | `#VALUE!` | Element |
| `DATE_TEXT_IMPOSSIBLE` | Syntactically exact ISO text naming a date that does not exist | `#VALUE!` | Element |
| `DATE_TYPE_REJECTED` | `Null`, Boolean, or another prohibited type at a date position | `#VALUE!` | Element |
| `DATE_WINDOW` | Accepted date or serial outside 1900-03-01 through 9999-12-31, including serial 60 and below | `#NUM!` | Element |
| `INPUT_BLANK_REQUIRED` | Blank cell or `Empty` at a required value position | `#VALUE!` | Element |
| `INPUT_ERROR_PROPAGATED` | Incoming native Excel error at a value position, returned verbatim | Incoming error | Element |
| `INTEGER_FRACTION` | Fractional numeric value inside the VBA `Long` range at an integer position | `#VALUE!` | Element |
| `INTEGER_TYPE_REJECTED` | Boolean, `Date`, text including numeric-looking text, `Null`, or object at an integer position | `#VALUE!` | Element |
| `INTEGER_RANGE` | Numeric value outside the VBA `Long` range, including fractional values | `#NUM!` | Element |
| `DOMAIN_YEAR` | `YearIn` outside 1900 through 9999 | `#VALUE!` | Element |
| `DOMAIN_MONTH` | `MonthIn` outside 1 through 12 | `#VALUE!` | Element |
| `DOMAIN_WEEKDAY` | `WdIndex` outside 1 through 7 | `#VALUE!` | Element |
| `DOMAIN_OCCURRENCE` | Occurrence `n` outside 1 through 5 | `#VALUE!` | Element |
| `OCCURRENCE_ABSENT` | Valid locator request for an occurrence that does not exist in that month | `#NUM!` | Element |
| `RESULT_WINDOW` | Valid operation whose result or intermediate value falls outside the supported window, including arithmetic overflow | `#NUM!` | Element |
| `PILLAR_TYPE_REJECTED` | Non-text payload at a `Pillar` position | `#VALUE!` | Element |
| `PILLAR_TOKEN_MALFORMED` | Token violates the accepted grammar: internal whitespace, missing or decimal quantity, unknown unit, sign with no body, partial parse | `#VALUE!` | Element |
| `PILLAR_DUPLICATE_UNIT` | A unit appears more than once in one token | `#VALUE!` | Element |
| `PILLAR_ALIAS_SIGNED` | A whole-token alias carries a leading sign, such as `-ON` or `+T/N` | `#VALUE!` | Element |
| `PILLAR_AGGREGATE_RANGE` | Grammatically valid token whose aggregate or resulting date leaves the supported domain | `#NUM!` | Element |
| `CONTROL_TYPE_REJECTED` | Optional control of a prohibited type: non-Boolean for a Boolean control, non-text for `Opt_Rounding`, or `Null` | `#VALUE!` | Call |
| `CONTROL_TOKEN_UNKNOWN` | `Opt_Rounding` text that is not `NEAREST`, `FLOOR`, or `CEILING` after trimming and case folding | `#VALUE!` | Call |
| `CONTROL_NOT_SCALAR` | Optional control larger than 1x1 | `#VALUE!` | Call |
| `CONTROL_ERROR_PROPAGATED` | Incoming native Excel error supplied as an optional control, returned verbatim | Incoming error | Call |
| `SHAPE_UNSUPPORTED` | Multi-area Range, jagged, empty, higher-dimensional, or non-Range object input | `#VALUE!` | Call |
| `SHAPE_MISMATCH` | Non-scalar value arguments whose row/column dimensions differ | `#VALUE!` | Call |
| `CAPACITY_EXCEEDED` | Resolved output target of 100,001 or more elements | `#NUM!` | Call |
| `HOST_DATE1904` | Identifiable worksheet caller in a 1904 workbook | `#N/A` | Call |
| `HOST_UNRESOLVED` | Identifiable worksheet host whose date system cannot be resolved reliably | `#N/A` | Call |

### 🛡️ 7.1 Unexpected internal failures

The registry deliberately has no `INTERNAL_UNEXPECTED` condition identifier.
Defensive catch-all handlers are containment only and must be unreachable in
contract-conforming execution. Activating one is a defect against this
contract and a regression/certification failure, never an expected fixture or
evidence outcome. It must not be normalized into a passing `#VALUE!`, `#NUM!`,
or `#N/A` case.

## ⚙️ 8. Function semantics

All functions use the parsing, host, array, and error rules above.

### 📍 8.1 Boundaries and predicates

- `DayOfWeek` returns 1 through 7. With `Opt_WeekBaseMonday=True`, Monday is 1
  and Sunday is 7. With `False`, Sunday is 1 and Saturday is 7.
- `DaysInMonth` returns the Gregorian length of the containing month.
- `DaysInYear` returns 366 for a Gregorian leap year and 365 otherwise. A year
  is leap when divisible by 4 except centuries not divisible by 400. It takes a
  calendar year, not a date.
- `BeginOfMonth` and `EndOfMonth` return the first and last dates of the
  containing month.
- `BeginOfQuarter` and `EndOfQuarter` use calendar quarters Jan-Mar, Apr-Jun,
  Jul-Sep, and Oct-Dec.
- `BeginOfYear` and `EndOfYear` return 1 January and 31 December of the
  containing year.
- `IsMonthEnd`, `IsQuarterEnd`, and `IsYearEnd` test those same boundaries.

A boundary is a result and is gated like one. `BeginOfYear` of any 1900 date
and `BeginOfQuarter` of any date in the first quarter of 1900 name
1900-01-01, which is outside the supported window, and return `#NUM!` under
`RESULT_WINDOW`. The rejection is specific to that result: `EndOfQuarter` and
`EndOfYear` of the same inputs return 1900-03-31 and 1900-12-31, and
`BeginOfQuarter` of a second-quarter 1900 date returns 1900-04-01.
- `IsLeapYear` applies the same Gregorian rule and also takes a calendar year.

`DaysInYear` and `IsLeapYear` are the two year-taking functions. Their argument
is `YearIn`, parsed by the strict integer rules in section 3.2 with the domain
1900 through 9999. A date, a date serial, or ISO date text is not accepted and
returns `#VALUE!` under `INTEGER_TYPE_REJECTED`; a caller holding a date
supplies `YEAR(A1)` or an equivalent year value.

Both functions are pure calendar predicates on a year. They construct no date,
so the supported-window gate does not apply to them and `DATE_WINDOW` and
`RESULT_WINDOW` cannot arise. Year 1900 is therefore in the domain even though
1900-01-01 through 1900-02-28 are outside the supported date window:
`KPR_Dates_IsLeapYear(1900)` returns `False` and `KPR_Dates_DaysInYear(1900)`
returns 365, which are the correct Gregorian answers.

A year argument cannot also accept a date. Every value in 1900 through 9999 is
itself a valid Excel serial, so a dual-meaning argument would have to guess, and
guessing is what the strict-parsing rules in section 3 exist to remove.

`DaysInMonth` keeps a `DateIn` argument because it needs both a month and a
year, and no comparable ambiguity arises. The mixed surface is deliberate: each
function takes the smallest input that determines its answer.

### ➕ 8.2 Date arithmetic

- `AddDays` adds the parsed signed integer as exact calendar days.
- `AddWeeks` adds exactly seven times the parsed signed integer as calendar
  days.
- `AddMonths` shifts by calendar months. In normal clip mode, retain the input
  day when it exists in the target month and otherwise use the target month's
  last day. With `Opt_KeepEOM=True`, an input that is its month's last day maps
  to the target month's last day; a non-EOM input still uses clip mode.
- `AddYears` is the same operation as `AddMonths` with a 12-month multiple, so
  clipping, leap-day handling, and optional EOM preservation cannot diverge.

Every intermediate and final result is range-gated before conversion to a VBA
`Date`. Arithmetic overflow or a result outside the supported window returns
`#NUM!` rather than a runtime error or rollover.

### 🗓️ 8.3 Weekday locators

`NthWeekdayOfMonth` returns occurrence `n` of `WdIndex` in `YearIn`/`MonthIn`
under the selected weekday base. A valid request for an occurrence that does
not exist, such as a fifth weekday absent from that month, returns `#NUM!`
under `OCCURRENCE_ABSENT`.

`LastWeekdayOfMonth` returns the final occurrence of `WdIndex` in the requested
year/month. Its argument domains and weekday-base interpretation are identical.
Any otherwise valid locator whose result is before 1900-03-01 returns `#NUM!`
under `RESULT_WINDOW`.

### 🧱 8.4 Pillar formatting and rounding

`PillarFromDates` returns `0D` for equal dates. An absolute interval shorter
than seven days returns the exact signed day token. Longer intervals compare
whole-week anchors and calendar-month anchors generated from `StartDate` in the
direction of `EndDate`. Month anchors use the same clip semantics as
`DateFromPillar`; month counts of 12 or more format as years plus residual
months.

#### Week-family cap

A week anchor is a candidate only for whole-week counts of 1, 2, and 3. Beyond
`3W` the week family is not considered and the month family owns the label.

The cap restricts the candidate set, not the selection rule, so it applies
identically under `NEAREST`, `FLOOR`, and `CEILING`.

The cap exists because week anchors and month anchors are not comparable on
distance alone. A whole-week anchor always lands within three days of any
target, while a calendar-month anchor can sit roughly fifteen days away, so an
unrestricted distance rule gives the week family almost every long interval: an
interval of ten years and two weeks would emit `524W`, which is arithmetically
correct and useless as a curve bucket. Market convention runs weeks only at the
short end.

The candidate set is the same under every mode: the `1W`, `2W` and `3W`
anchors, the floor month anchor (the largest positive month count whose anchor
does not pass `EndDate` in the interval direction) and the ceiling month anchor
(the floor count plus one). Anchors are generated from the original `StartDate`
in the direction of `EndDate` using signed month counts, so an emitted negative
token re-parses with the same clipping `DateFromPillar` applies. Only positive
month counts are admitted.

Two consequences are deliberate and must not be treated as defects:

- `4W` and every longer week token are unreachable.
- Under `NEAREST` the `3W` to `1M` boundary is derived from distance, not
  pinned to a day count. With an in-window `1M` anchor, 24 days always emits
  `3W` and 27 days always emits `1M`; 25 and 26 days resolve by calendar-day
  distance to the `1M` anchor, with an equal distance choosing the month. From
  a start date whose `1M` anchor is 31 days out, 25 days is `3W` and 26 is
  `1M`; from 31 January, whose `1M` anchor clips to 28 February, 25 days is
  already `1M`. The boundary is left where the distance rule puts it rather
  than special-cased, because a boundary that cannot be derived from the rule
  is the kind that drifts.

An anchor that would fall outside the supported window is not a candidate. It
is excluded, never approximated: there is no nominal distance for an anchor
that cannot be constructed. The restriction is positional, so a label can
depend on where in the calendar the interval sits. An eleven-day interval emits
`2W` under `NEAREST` everywhere except the final fortnight of year 9999, where
the `2W` and `1M` anchors both leave the window and `1W` is the only remaining
candidate: `NEAREST` and `FLOOR` emit `1W`, and `CEILING` has no candidate that
reaches `EndDate`.

`FLOOR` and `NEAREST` always have a candidate: for any interval of seven days
or more, the `1W` anchor lies between the two in-window dates. `CEILING` may
find every in-window candidate short of `EndDate`, as in the example above. It
then returns `#NUM!` under `RESULT_WINDOW`, because the anchor the mode
requires would leave the supported window. It does not fall back to the
largest available anchor, which would silently be `FLOOR`.

#### Rounding modes

Rounding operates on the absolute magnitude and then restores a leading minus
sign for a negative interval:

- `FLOOR` chooses the greatest candidate magnitude whose anchor does not pass
  `EndDate` in the interval direction.
- `CEILING` chooses the smallest candidate magnitude whose anchor reaches or
  passes `EndDate` in the interval direction.
- `NEAREST` chooses the anchor with the smallest absolute calendar-day distance
  from `EndDate`.

Tie-breaking is deterministic and identical under every mode. Equivalent week
and month anchors use the month representation. An equal outcome between a week
candidate and a month candidate takes the month; an equal outcome between two
month candidates takes the larger count, which is the ceiling month anchor.
Exact anchors remain exact under every mode, because an exact anchor wins each
mode's selection outright.

`DateFromPillar` applies the accepted grammar and month-then-day order in
section 3.4, and `PillarFromDates` emits only the narrower canonical grammar
recorded there.
Pillar parsing and formatting are mutually consistent for every emitted token.
Because `PillarFromDates` may round, this is deliberately not a general
invariant:

```text
DateFromPillar(StartDate, PillarFromDates(StartDate, EndDate)) = EndDate
```

Tests and documentation must not imply that equality for rounded intervals.

## 🪟 9. Excel-version compatibility

Scalar calls and multi-cell calls have separate compatibility claims:

- scalar calls may be claimed only for Excel versions exercised successfully
  during exact-source Windows certification;
- multi-cell calls may be tested, supported, and claimed only on dynamic-array
  Excel; and
- v0.0.4 makes no Ctrl+Shift+Enter or other legacy multi-cell compatibility
  claim.

The implementation performs no Excel-version detection and does not create or
simulate Excel's placement-level `#SPILL!` error. Spill placement remains
Excel's responsibility.

## 🔒 10. Supported API boundary and future namespaces

The 22 names in section 2 are the complete supported v0.0.4 calculation API.
VBA may require other procedures to be technically `Public` for cross-module
calls, `Application.Run`, MacroOptions, RibbonX, CommandBars, tests, or demo
generation. Technical visibility does not make those procedures supported API.

Internal `KPR_Core_*` modules use `Option Private Module`; helpers use the
narrowest practical visibility. Registration, UI callbacks, test runners, and
demo builders are explicitly unsupported infrastructure and contain no date
algorithms.

`KPR_Cal_*` is reserved for v0.0.5. v0.0.4 creates no calendar placeholder and
adds no calendar option to a pure `KPR_Dates_*` function.

The following are explicitly outside v0.0.4:

- calendars and calendar composition;
- weekend masks and holiday sets;
- business-day arithmetic;
- following, preceding, modified, nearest, or other roll conventions;
- duplicate `_Spill` functions or `KPR_Dates_Spill.bas`; and
- generated `.xlsm`, `.xlam`, `.xlsx`, or other Office binaries in git.

## ✅ 11. Contract acceptance traceability

The contract has twelve acceptance criteria. Each is listed below in order
with its normative coverage in this document.

| # | Acceptance criterion | Normative coverage |
| ---: | --- | --- |
| 1 | Every one of the 22 supported names has an exact VBA signature, semantic return type, and vectorization classification | Section 2 |
| 2 | The input-type and native-error matrices contain no locale-dependent parsing path | Sections 3.1 through 3.4, 6, and 7 |
| 3 | The `1900-03-01 .. 9999-12-31` gate and caller/date-system policy are unambiguous | Sections 1 and 4 |
| 4 | The documented conditions produce `#VALUE!`, `#NUM!`, host-configuration `#N/A`, and verbatim propagated errors as specified | Sections 3, 5.2, 6, and 7 |
| 5 | Host-generated and propagated `#N/A` are indistinguishable at the Excel-value level, and `HostDateSystem()` supplies caller context | Sections 4, 6, and 7 |
| 6 | `HostDateSystem()` is volatile and its recalculation behavior is specified | Section 4 |
| 7 | Certified direct-VBA uses are distinguished from unsupported non-Range Excel host contexts | Section 4 |
| 8 | Scalar expansion, optional-argument rules, orientation, 1-D arrays, blanks, the 100,000-element cap, and per-element error propagation are specified | Sections 3.3, 5.1, and 5.2 |
| 9 | Scalar and dynamic-array compatibility claims are stated separately with no CSE claim | Section 9 |
| 10 | Pillar rounding is defined with tie-breaking, negative intervals, and non-invariant round trips | Sections 3.4 and 8.4 |
| 11 | The supported-API boundary and reserved `KPR_Cal_*` namespace are recorded | Section 10 |
| 12 | Calendars, weekend masks, holidays, business-day arithmetic, and roll conventions are explicitly out of scope | Sections 1 and 10 |
