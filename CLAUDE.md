# CLAUDE.md — Project Context for Claude Code

This file is read automatically by Claude Code at the start of every session.
Do not delete or rename it.

---

## Project Overview

**Repository:** https://github.com/geki-research/knime-extensions
**Owner:** geki-research
**Vendor:** DataCraft Labs
**Purpose:** KNIME 5.x community extension nodes for advanced data ingestion
and transformation.

---

## Development Environment

| Item | Detail |
|---|---|
| OS | Debian 12 |
| Java | Sources compile at **17**; bundle declares **JavaSE-17** — see below |
| Build system | Maven 3.9+ with Eclipse Tycho 4.0.6 |
| IDE | Eclipse for RCP and RAP Developers 2024-03 |
| Eclipse workspace | `~/knime-dev/workspace` |
| KNIME target platform | `~/knime-dev/knime-sdk-setup` → `KNIME-AP.target` (1897 plugins) |

### Java level — 17 compile, 17 runtime (this branch)

| Setting | Value | Where |
|---|---|---|
| Source/target compile level | **17** | `tycho-compiler-plugin` `<source>/<target>` in `<build><plugins>` — **authoritative on this branch** |
| `maven.compiler.source/target` | **17** | `pom.xml` properties (kept in sync; **not** what governs here) |
| `Bundle-RequiredExecutionEnvironment` | **JavaSE-17** | both `MANIFEST.MF` files |

There is no compile/runtime divergence on this branch. The KNIME 5.5 target
platform runs on Java 17, the bundles declare `JavaSE-17`, and sources compile at
17 to match. Verified empirically: compiled classes are **major version 61
(Java 17)**.

**Which mechanism sets the compile level differs between branches — check before
changing it.** Tycho's `source`/`target` parameters default to
`${maven.compiler.source}` / `${maven.compiler.target}`, and an explicit
`<configuration>` block on `tycho-compiler-plugin` overrides that default.

On **this branch** that explicit block is present in `<build><plugins>`, so it
**governs at 17** and `maven.compiler.*` is inert — editing the properties alone
would be a silent no-op. `main` and `releases/5.8` work the same way.

On `releases/5.12` the rule is **reversed**: PR #2 declares the plugin only under
`<pluginManagement>` with no `<configuration>`, so there `maven.compiler.*` is
what sets the level, and it is 21. Do not carry a compile-level change between
branches without checking which form applies.

### Why `<source>`/`<target>` are sufficient (ecj, not javac)

**`<source>`/`<target>` do more here than their names suggest. They set the
visible API surface, not just the language level and bytecode version.**

**Tycho compiles with the Eclipse Compiler for Java (ecj), not javac.** This
branch uses **ecj 3.36.0.v20231114-0937**, supplied by `tycho-compiler-plugin`
4.0.6 — read it from any build log:

```
[INFO] Compiling 17 source files … using Eclipse Compiler for Java(TM) 3.36.0.v20231114-0937
```

**ecj applies release semantics from `<source>`/`<target>` by itself.** A call to
an API newer than the declared level is a **compile error**, not a runtime
surprise. Measured on this branch:

| Probe | Introduced | Result |
|---|---|---|
| `Math.clamp(long, int, int)` | Java 21 | **BUILD FAILURE** — `The method clamp(long, int, int) is undefined for the type Math` |
| `java.util.HexFormat` | Java 17 | **BUILD SUCCESS** |

The boundary is therefore **exactly the declared level**, not merely "anything
recent". A post-17 API cannot reach a Java 17 KNIME runtime through this build,
because it cannot get past compilation.

**So `<release>` is unnecessary here.** `tycho-compiler-plugin` 4.0.6 does accept
a `<release>` parameter, and it coexists with `<source>`/`<target>` without error
— but it is **redundant**, because ecj already enforces what it would enforce.
Do not add it believing it closes a hole; there is no hole.

**The javac contrast — this is the part that matters.** Under **javac** the
intuition behind that suggestion is correct:

```
$ javac -source 17 -target 17 Probe.java     # JDK 21, calling Math.clamp
warning: [options] system modules path not set in conjunction with -source 17
→ COMPILES.  Emits major-version-61 bytecode. Throws NoSuchMethodError on Java 17.

$ javac --release 17 Probe.java
error: cannot find symbol   Math.clamp
→ REJECTED.
```

Under javac, `-source`/`-target` really do control only the language level and
bytecode version, and only `--release` restricts the API. **That reasoning is
sound but does not apply to this project**, because this project does not compile
with javac. Do not transplant it here.

**Therefore: do not treat `<source>`/`<target>` as cosmetic.** Changing them
widens the API surface. Raising them on this branch would let post-17 APIs into a
bundle that ships to a Java 17 runtime — code that builds green here and throws
`NoSuchMethodError` in the field.

### A local JDK 21 will emit a warning here

Building this branch on a Java 21 JDK prints:

```
[WARNING] Using JavaSE-21 to fulfill requested profile of JavaSE-17. This might
lead to faulty dependency resolution, consider defining a suitable JDK in the
toolchains.xml.
```

Expected, and not a fault: the BREE is `JavaSE-17` but only a 21 JDK is
installed, so Tycho substitutes it. The build passes and the bytecode is still
major version 61.

**Measured: the warning has no effect on resolution.** It comes from Tycho's OSGi
execution-environment resolution — not from the compiler, and unrelated to the
API-surface question above, which ecj settles correctly on its own. Compared with
and without a JDK 17 toolchain, using `config.ini`'s `osgi.bundles` (the
instrument named under "Do not judge resolution by `.source` bundle download
counts"):

```
without toolchain:  5 warnings,  244 bundles provisioned
with toolchain:     0 warnings,  244 bundles provisioned
diff → identical
```

The toolchain changes the message and nothing else. Treat the warning as a
modelling notice, not a fault.

**One honest caveat.** That is measured for the *current* dependency set —
`org.knime.core`, `org.knime.base` and `org.apache.poi`, none of them
JDK-version-conditional. A future dependency whose resolution hinged on a
`java.*` package present in 21 but not 17 could in principle resolve differently.
Re-measure if the dependency set changes materially.

**A toolchain was considered and declined.** A `~/.m2/toolchains.xml` registering
a JDK 17 does silence the warning, but it fixes a message rather than a defect;
it is machine-local, so it would not help a fresh clone, another developer, or
KNIME's Jenkins; and it would leave one machine building "clean" while every
other host still warned, making the warning look like a local anomaly. A JDK 17
is already installed at `/usr/lib/jvm/java-17-openjdk-amd64`, so this remains a
one-file change if it is ever wanted — no installation required.

---

## Repository Structure

```
knime-extensions/
  CLAUDE.md                                     ← this file
  README.md                                     ← project documentation
  pom.xml                                       ← parent POM (Tycho config)
  org.geki.knime.excelformreader/               ← Excel Reader (non-tabular) plugin
  org.geki.knime.excelformreader.tests/         ← test project
  org.geki.knime.excelformreader.feature/       ← feature project
  org.geki.knime.excelformreader.update/        ← update site
```

---

## Test Project

```
org.geki.knime.excelformreader.tests/
  testdata/
    forms/          ← .xlsx test files
    definitions/    ← CSV form definition tables
  src/org/geki/knime/excelformreader/tests/
                    ← 13 JUnit test classes covering the domain, excel
                      and output layers plus the settings class; see
                      "What Is Not Yet Implemented" for the tally and
                      what remains uncovered
```

Test fixture: Legacy_IT_System_Assessment_Test.xlsx
  - 2 data sheets: Test_01, Test_02 (identical layout)
  - 1 excluded sheet: Config
  - Data types covered: string, int, date
  - Data types not covered (unit tests only): double, boolean

Form definition: it_assessment_form_definition.csv
  - 40 field mappings
  - 4 label fields: reference_date, assessor, system_type, system_name
  - 36 data fields
  - Covers all cell address patterns used in the real form
  - Data types covered: string, int, date
  - Data types not covered (unit tests only): double, boolean

---

## Build

```bash
# Full build
cd ~/knime-dev/knime-extensions
mvn -U clean verify

# Build output (update site) is at:
org.geki.knime.excelformreader.update/target/repository/
```

BUILD SUCCESS is the only acceptable outcome before committing.
Always run the build and confirm success before committing any code changes.

### Always pass `-U`

Tycho's default `cache first` update mode never re-fetches a `content.jar` it
already holds, so a local build can silently resolve months-old target-platform
metadata and then fail on artifacts the server has since deleted. `-U` is the
cure. If a cache becomes corrupt, move `~/.m2/repository/.cache/tycho` aside —
that is the recovery step.

Alternatives were tested and **all failed**: waiting for the cache to expire,
`-Dtycho.p2.transport.min-cache-minutes=0`, and
`-Dtycho.p2.transport.update=forced`.

### Do not judge resolution by `.source` bundle download counts

`.source` bundle download counts are **not a valid instrument** for judging what
the build resolves. `mvn clean` empties `target/` but not
`~/.m2/repository/p2/osgi/bundle/`, so a second consecutive run against the same
repository reports zero downloads whatever the configuration — the count tracks
local cache state, not resolution.

Use `org.geki.knime.excelformreader.tests/target/work/configuration/config.ini`
instead: its `osgi.bundles` property is the literal list of bundles provisioned
into the test runtime. (`target/skippedP2Dependencies.txt` is written by Tycho
regardless of dependency-resolution configuration — measured identical with and
without `optionalDependencies=ignore` — so it does not discriminate on its own.)

---

## Branch Strategy for KNIME Versions

### Per-branch facts

Per-branch facts are maintained on `main` in `CLAUDE.md`. Do not duplicate them
here.

### This branch — `releases/5.5`

| Item | Value |
|---|---|
| Active profile | `knime-5.5`, set `activeByDefault` in `pom.xml` |
| p2 repository | `https://update.knime.com/analytics-platform/5.5` |
| Tycho | **4.0.6** — see below |
| Compile level | 17, set by the `tycho-compiler-plugin` block — see "Java level" above |
| BREE | `JavaSE-17`, both manifests |
| `Require-Bundle` upper bounds | `6.0.0` — see below |

```bash
mvn -U clean verify   # uses the knime-5.5 profile — this branch's default
```

No `-P` is needed on this branch; `knime-5.5` is already the default. KNIME
Jenkins passes `-P knime-5.5` explicitly anyway, which selects the same profile.

**Tycho stays at 4.0.6 here — do not bump it.** 4.0.13 resolves optional
dependencies that 4.0.6 ignores, which is what turned `main`'s build red, and it
mishandles the `/analytics-platform/5.12` redirect. This branch builds green on
4.0.6 and KNIME has sent no Tycho-bump PR for it.

**`Require-Bundle` upper bounds stay narrow — `[5.5.0,6.0.0)` in the plugin
manifest, `[5.3.0,6.0.0)` in the tests manifest.** `main` carries `7.0.0` and
this branch deliberately does not, so **a diff against `main` will show it — that
is intentional, not drift.** `main` needs the wide range because it builds
against nightly, which moves toward 6.x; this branch is pinned to the 5.5 update
site and will never meet a 6.x platform. The narrow bound also states what has
actually been tested: this branch has never been built against KNIME 6.x.

The non-default `knime-5.12` profile in this POM must carry the `lts/` URL.
Tycho 4.0.13 caches the `/analytics-platform/5.12` →
`/analytics-platform/lts/5.12` redirect body under the original URL key and then
cannot read it, so the unqualified URL fails at repository load. **This branch
runs 4.0.6, which follows the redirect correctly, so the `lts/` form is a
precaution here rather than load-bearing** — it matters only if anyone builds
this tree with `-P knime-5.12` on a newer Tycho.

KNIME Jenkins activates the correct profile per branch automatically via
`-P knime-X.Y` in its build command.

**New release branch checklist:**
1. `git checkout -b releases/X.Y` from `main`
2. Set `<activeByDefault>true</activeByDefault>` on the `knime-X.Y` profile
   in `pom.xml` (remove it from `knime-nightly`)
3. Commit, push branch
4. Notify KNIME team to add a build job for the new branch

---

## Git Conventions

**Branching strategy:** see "Branch Strategy for KNIME Versions" above — the
`main` + `releases/X.Y` model is the only one in use.
- `main` — always releasable, always passing build; development happens here
- `releases/X.Y` — one long-lived branch per supported KNIME version
- `feature/<name>` — one branch per node or feature, branched from `main`

There is **no `develop` branch**. Earlier revisions of this file described a
GitHub Flow variant with one; that branch never existed. Do not create it.

**Commit message format:**
```
<type>: <short description>

Types: feat | fix | refactor | chore | docs | test
Examples:
  feat: implement CellValueConverter for all supported data types
  fix: handle merged cells in ExcelFormExtractor
  chore: update Tycho version to 4.0.7
```

**Before every commit:**
1. `mvn -U clean verify` must produce BUILD SUCCESS
2. `git status` must show only intentional changes
3. Push to remote immediately after committing

---

## Plugin Project: Excel Reader (non-tabular)

**Plugin ID:** `org.geki.knime.excelformreader`
**Node name:** Excel Reader (non-tabular)
**Category:** `/datacraft-labs/io`
**KNIME API version:** 5.5.x (minimum)

### What this node does
Reads non-tabular, form-structured Excel worksheets (.xlsx) and extracts
field values into a standard KNIME data table. The form structure
(label → cell address mapping) is provided via an input table, making
the node generic and reusable across any form layout.

### Node ports
| Port | Direction | Type | Description |
|---|---|---|---|
| 0 | Input | BufferedDataTable | Form definition table |
| 0 | Output | BufferedDataTable | Extracted data (wide or long) |
| 1 | Output | BufferedDataTable | Label fields (always produced, may be empty) |

### Form definition table schema
| Column | Required | Type | Description |
|---|---|---|---|
| `Name` | ✅ | String | Output column name |
| `Cell Range` | ✅ | String | Cell address (`C4`) or range (`B10:D15`) |
| `Content Type` | ❌ | String | `data` / `label` — defaults to `data` |
| `Data Type` | ❌ | String | `string`/`int`/`double`/`date`/`boolean` — defaults to `string` |

### Reading modes
| Mode | Description |
|---|---|
| `SINGLE_FILE` + single sheet | One file, one named sheet |
| `SINGLE_FILE` + all sheets | One file, all sheets except excluded ones |
| `FOLDER` | All .xlsx files in a folder, all sheets per file |
| `FOLDER` + recursive | Same, including all subfolders |

Each (file, sheet) pair = one form instance = one output row (wide mode).

### Output formats
- **Wide:** one row per (file, sheet), one column per field
- **Long:** one row per (file, sheet, field) — columns: `field_name`, `value`
- Configurable via dialog toggle

### Dialog settings

Every row below is a control in `ExcelFormReaderNodeDialog`. The **Settings key**
column is the `CFG_*` constant in `ExcelFormReaderSettings` that persists it;
`(none)` marks a control that is **not persisted**. **Shown when** records the
parent control whose selection reveals or enables the row — blank means always
visible and enabled.

| Panel | Setting | Type | Default | Settings key | Shown when |
|---|---|---|---|---|---|
| General / Input | Input mode | Radio | Single File | `cfg_inputMode` | |
| General / Output | Output format | Radio | Wide | `cfg_outputFormat` | |
| General / Output | Include source filename | Boolean | true | `cfg_includeSourceFilename` | |
| General / Output | Include sheet name | Boolean | true | `cfg_includeSheetName` | |
| General / Output | Include label fields in port 0 | Boolean | false | `cfg_includeLabelFields` | |
| General / Output | Output label fields in port 1 | Boolean | true | `cfg_outputLabelPort` | |
| General / Output | Include format condition operator columns | Boolean | false | `cfg_includeFormatCondition` | |
| General / Output | Include validation type columns | Boolean | false | `cfg_includeValidationType` | |
| General / Error Handling | On missing cell | Radio | Warn | `cfg_onMissingCell` | |
| General / Error Handling | On unparseable value | Radio | Warn | `cfg_onBadValue` | |
| File / Input Location | Read from | Combo, one fixed item `Local File System` | — | **(none) — not persisted** | |
| File / Input Location | File path | String | — | `cfg_filePath` | |
| File / Select Sheet(s) | Process single/many sheets | Radio | Single | `cfg_fileManySheets` | |
| File / Select Sheet(s) | Include hidden worksheets (single) | Boolean | false | `cfg_fileSingleHiddenSheets` | Process **single** sheet |
| File / Select Sheet(s) | Sheet selection (single) | Radio | First | `cfg_fileSheetSelection` | Process **single** sheet |
| File / Select Sheet(s) | Sheet name (single) | Dropdown, populated from the selected file | — | `cfg_fileSheetName` | Sheet selection = **By name** |
| File / Select Sheet(s) | Sheet position (single) | Integer spinner, 0–999 | 0 | `cfg_fileSheetPosition` | Sheet selection = **By position** |
| File / Select Sheet(s) | Include hidden worksheets (many) | Boolean | false | `cfg_fileHiddenSheets` | Process **many** sheets |
| File / Select Sheet(s) | Sheet filter mode (many) | Radio | All | `cfg_fileSheetFilterMode` | Process **many** sheets |
| File / Select Sheet(s) | Sheet names (many) | String, comma-separated | — | `cfg_fileSheetFilterNames` | Sheet filter mode = **Blacklist** or **Whitelist** |
| Folder / Input Location | Read from | Combo, one fixed item `Local File System` | — | **(none) — not persisted** | |
| Folder / Input Location | Folder path | String | — | `cfg_folderPath` | |
| Folder / Input Location | Include subfolders | Boolean | false | `cfg_recursive` | |
| Folder / Input Location | Include hidden folders | Boolean | false | `cfg_includeHiddenFolders` | Enabled only while **Include subfolders** is checked; unchecking it clears this box |
| Folder / File Filter | Filter by file extension | Radio | Selected | `cfg_filterByExtension` | |
| Folder / File Filter | File extensions | String | xlsx | `cfg_fileExtensions` | **Filter by file extension** selected |
| Folder / File Filter | Include hidden files | Boolean | false | `cfg_includeHiddenFiles` | |
| Folder / Select Sheet(s) | Process single/many sheets | Radio | Single | `cfg_folderManySheets` | |
| Folder / Select Sheet(s) | Include hidden worksheets (single) | Boolean | false | `cfg_folderSingleHiddenSheets` | Process **single** sheet |
| Folder / Select Sheet(s) | Sheet selection (single) | Radio | First | `cfg_folderSheetSelection` | Process **single** sheet |
| Folder / Select Sheet(s) | Sheet name (single) | String (free text — **not** a dropdown, unlike the File tab) | — | `cfg_folderSheetName` | Sheet selection = **By name** |
| Folder / Select Sheet(s) | Sheet position (single) | Integer spinner, 0–999 | 0 | `cfg_folderSheetPosition` | Sheet selection = **By position** |
| Folder / Select Sheet(s) | Include hidden worksheets (many) | Boolean | false | `cfg_folderHiddenSheets` | Process **many** sheets |
| Folder / Select Sheet(s) | Sheet filter mode (many) | Radio | All | `cfg_folderSheetFilterMode` | Process **many** sheets |
| Folder / Select Sheet(s) | Sheet names (many) | String, comma-separated | — | `cfg_folderSheetFilterNames` | Sheet filter mode = **Blacklist** or **Whitelist** |

**Count check: 35 rows = 33 persisted settings + 2 unpersisted "Read from"
combos.** `ExcelFormReaderSettings` holds exactly 33 `SettingsModel` fields and
33 `CFG_*` keys, each saved, loaded and validated. Every key appears in the
Settings key column exactly once. If those numbers stop agreeing, the table has
drifted — same reasoning as the single authoritative test tally.

Not in the table: the two `Browse...` buttons and the first-sheet-name preview
label, which are actions and display only, and carry no state.

**The two "Read from" combos are not settings — and are deliberately so.** Each
is constructed with the single item `"Local File System"`, added to its Input
Location box, and then never saved, loaded or read; no `CFG_*` key backs either
one. They remain unpersisted **by design**, and the count check above stays as it
is: 35 rows = 33 persisted + 2 unpersisted.

They are **intentional placeholders**, confirmed by the project owner. They serve
two purposes: they mirror the "Read from" control in KNIME's native Excel Reader,
so this node looks consistent with the platform's own file-reading nodes; and
they reserve the position in the layout for future KNIME file-system support.

**Do not remove them as dead code.** An unbacked control carrying one fixed item
looks exactly like leftover scaffolding to anyone reading the code cold — this
note is what should stop a future cleanup from deleting it. Earlier revisions of
this file recorded their status as an open question; it is settled.

### Output Ports

**Port 0 — Main output (wide mode columns in order):**
1. `source_file` (String, optional)
2. `sheet_name` (String, optional)
3. Per field in definition order:
   - `<Name>` — value typed per Data Type
   - `<Name> (Format Condition Operator)` — String, optional (toggle)
   - `<Name> (Validation Type)` — String, optional (toggle)

Label fields included when "Include label fields in port 0" is enabled.

**Port 0 — Main output (long mode columns in order):**
1. `source_file` (String, optional)
2. `sheet_name` (String, optional)
3. `field_name` (String)
4. `value` (String)
5. `Format Condition Operator` (String, optional)
6. `Validation Type` (String, optional)

**Port 1 — Label fields output:**
- Fixed wide format regardless of Port 0 format setting
- One row per label field per (file, sheet) pair
- Columns (in order): `Source File` (optional), `Sheet Name` (optional), `Name`, `Cell Range`, `Cell Content`, `Format Condition Operator` (optional), `Validation Type` (optional)
- Controlled by "Output label fields in port 1" toggle (default: true)
- Always produced as an empty table when toggle is disabled

---

## Package Structure & Class Responsibilities

```
org.geki.knime.excelformreader/
  ExcelFormReaderNodeFactory    — Node registration, XML description
  ExcelFormReaderNodeModel      — configure(), execute() — orchestrator only,
                                  no business logic here
  ExcelFormReaderNodeDialog     — Swing dialog panels
  ExcelFormReaderSettings       — All SettingsModel fields, save/load

  domain/
    FieldMapping                — One row from the definition table
                                  (name, cellRange, contentType, dataType)
                                  isData() / isLabel() convenience methods
    FormDefinition              — List<FieldMapping>, static factory
                                  fromDataTable(BufferedDataTable)
                                  getDataFields() / getLabelFields()
    CellAddress                 — Parses "C4" / "B10:D15" into typed fields,
                                  validates format; toString() produces
                                  canonical address string
    ReadingMode                 — Enum: SINGLE_FILE, FOLDER

  excel/
    WorkbookIterator            — Lazy Iterator<Entry> over (Path, Sheet) pairs
                                  Respects sheet filter, hidden sheet flag,
                                  recursive folder walk
                                  Opens/closes workbooks one at a time
    ExcelFormExtractor          — POI-based: resolves FormDefinition against
                                  a Sheet → Map<String, CellExtractionResult>
                                  CellExtractionResult holds value +
                                  formatConditionOperator + validationType
                                  Evaluates formulas transparently
    CellValueConverter          — POI Cell → KNIME DataCell per data_type
                                  Handles string/int/double/date/boolean
    CellMetadataReader          — Stateless; reads conditional formatting
                                  operators and data validation types per cell
                                  Resolves LIST validation options including
                                  inline lists, same-sheet ranges, cross-sheet
                                  ranges, and named ranges
                                  Always requires Workbook parameter

  output/
    OutputSpecFactory           — Creates DataTableSpec at configure() time
                                  createWideSpec() / createLongSpec() /
                                  createLabelSpec()
    WideOutputBuilder           — One DataRow per (file, sheet)
    LongOutputBuilder           — N DataRows per (file, sheet, field)
    LabelOutputBuilder          — One DataRow per label field per (file, sheet)
                                  Produces Port 1 output
```

---

## Key Implementation Rules

1. **No business logic in NodeModel or NodeDialog** — they orchestrate and
   delegate only. All logic lives in domain/, excel/, output/.

2. **WorkbookIterator must be lazy** — open one workbook at a time, close
   it before opening the next. Never load all workbooks into memory.

3. **OutputSpecFactory runs at both configure() and execute() time** — at
   configure() only the column *names* of the definition table are known
   (`FormDefinition.fromSpec()` returns an empty sentinel with no field
   mappings), so the port 0 spec built there has only the provenance
   columns in WIDE mode. The full per-field spec is only built at execute(),
   once `FormDefinition.fromDataTable()` has read the actual rows. See the
   `configure()` TODO comment in NodeModel and the known limitation below.

4. **Formula evaluation is transparent** — CellValueConverter always uses
   a FormulaEvaluator. Never return formula strings.

5. **Cell ranges (B10:D15)** — read left-to-right, top-to-bottom,
   concatenated with the range delimiter (hardcoded as `", "` in
   `ExcelFormExtractor` — see rule 12; not user-configurable).

6. **Missing/unresolvable cells** — honour the error handling settings:
   WARN logs and returns a missing value; FAIL throws a `RuntimeException`
   (the one deliberate unchecked throw in the extraction path — don't add
   others for cases this setting doesn't cover).

7. **Apache POI is provided by KNIME** — do NOT add POI as a Maven
   dependency. It is declared in MANIFEST.MF as `Require-Bundle`.

8. **All SettingsModel types** — use only KNIME SettingsModel* classes
   (SettingsModelString, SettingsModelBoolean, etc.) in Settings class.
   Never use raw strings/booleans for persistent settings.

9. **Provenance columns** — `source_file` (StringCell) and `sheet_name`
   (StringCell) are always the first two columns when enabled.

10. **Sheet exclusion** — comparison is case-insensitive and trimmed.

11. **Content Type filtering** — use `FormDefinition.getDataFields()` and
    `FormDefinition.getLabelFields()` for filtered iteration. Never filter
    inline in builders.

12. **Cell metadata** — `CellMetadataReader` is stateless. Always pass
    `Workbook` to enable cross-sheet list resolution. Range delimiter is
    hardcoded as `", "` in `ExcelFormExtractor`.

13. **Port 1** — always produced (may be empty table). Empty is simpler and
    faster than an optional port for large volumes.

14. **LIST validation resolution order** — (1) inline list; (2) a direct
    range reference, same-sheet or cross-sheet depending on whether the
    formula contains a `'Sheet'!` qualifier; (3) if that fails to parse as
    a range, treat the formula as a named range and recursively resolve its
    `refersToFormula`; (4) if the named range isn't found or resolvable,
    fall back to the raw formula/name string.

---

## Apache POI Notes

POI is bundled inside KNIME — use these classes:
```java
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellRangeAddress;
```

Open workbooks with try-with-resources:
```java
try (Workbook wb = WorkbookFactory.create(file.toFile(), null, true)) {
    // true = read-only mode, much more memory efficient
}
```

---

## KNIME API Notes

```java
// Correct NodeModel constructor for 1 input, 2 outputs:
super(new PortType[]{BufferedDataTable.TYPE},
      new PortType[]{BufferedDataTable.TYPE,
                     BufferedDataTable.TYPE});

// DataCell types to use:
StringCell, IntCell, DoubleCell, BooleanCell, DateAndTimeCell, MissingCell

// Always use DataType.getMissingCell() for missing values — never null.

// BufferedDataContainer pattern in execute():
BufferedDataContainer container = exec.createDataContainer(spec);
container.addRowToTable(row);
container.close();
return new BufferedDataTable[]{container.getTable()};
```

---

## What Is Not Yet Implemented

**Current tally: 150 tests — 149 passing, 1 skipped.** The single skipped test is
`FormDefinitionTest.testFromDataTable_placeholder`, `@Ignore`d because
`fromDataTable` requires a live `BufferedDataTable`. A build reporting any other
figure than `Tests run: 150, Failures: 0, Errors: 0, Skipped: 1` needs
investigating before commit.

Unit tests — covered so far: `CellAddress`, `FieldMapping`, `FormDefinition`
(construction/filtering), `CellValueConverter`, `CellMetadataReader`,
`ExcelFormExtractor`, `WorkbookIterator`, `ReadingMode`, `OutputSpecFactory`,
`WideOutputBuilder`, `LongOutputBuilder`, `LabelOutputBuilder`,
`ExcelFormReaderSettings`.

Still open:
- `FormDefinition.fromDataTable()` — deliberately deferred (see the
  `@Ignore`d `testFromDataTable_placeholder` in `FormDefinitionTest`); needs
  a live `BufferedDataTable`/`ExecutionContext`, not just a POI fixture.
- `ExcelFormReaderNodeModel`, `ExcelFormReaderNodeDialog`,
  `ExcelFormReaderNodeFactory` — need a live KNIME workflow/UI runtime to
  test meaningfully; not covered at the unit level. The unused
  `testdata/forms/*.xlsx` and `testdata/definitions/*.csv` fixtures are
  integration-test candidates for these, not yet wired up.

Known limitations:
- `configure()` returns partial spec in WIDE mode (see TODO comment above
  `configure()` in NodeModel — accepted, cosmetic only)
- Format condition operator reads `CELL_VALUE_IS` rules for operator name;
  other rule types return the condition type name instead

## Known Open Items

Context a future session would otherwise have to rediscover. Current as of
2026-08-03.

### PR #2 and issue #3 — another branch's business

Recorded here only so nobody re-investigates them from this branch. **Neither
affects `releases/5.5`.**

PR #2 ("Fix for 5.12 builds", by `dsaam94` — Ali Marvi, KNIME) was merged into
**`releases/5.12`** on 2026-08-03. It restructured that branch's POM only. Two
questions about it remain open and live in **issue #3**:
https://github.com/geki-research/knime-extensions/issues/3 — what was actually
failing in the 5.12 Jenkins build, and whether `skipArchive=true` is intentional.

None of that configuration exists on this branch: no `<pluginManagement>` block,
no jgit timestamp provider, no `skipArchive`, no aarch64 environment. This branch
still produces a normal update-site ZIP.

### Forward-port from `main` — DONE for this branch

This branch previously ran 65 tests against `main`'s 150. The 8 missing test
classes and the `Export-Package` line four of them require were forward-ported
from `main`'s commit `af65de0`, together with the `feature.xml` copyright fix
(`09b6aa8`), the `knime-5.12` `lts/` URL precaution (`54bf8b6`) and the
`.gitignore` entry (`16a3ee2`). **This branch now runs 150 / 1, matching `main`.**

Product code was already byte-identical to `main` before the port and was not
touched.

**Two of `main`'s commits were deliberately not ported, and should stay that
way:**

- **`dae2711`** ("standardise compile level on Java 17") — a no-op here that
  conflicts. This branch is already at 17 via the explicit
  `tycho-compiler-plugin` block.
- **`d885e90`** ("drop unused `org.knime.core.ui` bundle requirement") — that
  line has never existed on this branch.

### `<optionalDependencies>ignore</optionalDependencies>` — tried and reverted

Do **not** re-add it. It is a no-op for this project: with and without it the
test module provisions an identical OSGi runtime — 154 `.source` bundles, 361
bundles total — and `skippedP2Dependencies.txt` is written identically either
way (281 entries).

It appeared to work only because `.source` download counts were used as the
measure, and those track local p2 cache state rather than resolution (see the
Build section). Recorded here so nobody re-adds it on the strength of
download-count evidence.

### The dialog is legacy Swing

`ExcelFormReaderNodeDialog` extends `NodeDialogPane` — the legacy Swing API, not
the Modern UI / declarative API. KNIME has asked for migration; it is
**deferred**.

Migration is larger than a dialog swap. KNIME 5.12's documented approach
replaces the `NodeFactory` / `NodeModel` / `NodeDialogPane` triad with
`DefaultNodeFactory` plus a `NodeParameters` settings class, and is
**unavailable below 5.12** — so it cannot be done while `releases/5.5` and
`releases/5.8` are supported from the same source.

### Test fixtures

Two `.xlsx` fixtures are tracked in `testdata/forms/`; **neither is referenced by
any unit test** — both are integration-test candidates.

- `Legacy_IT_System_Assessment_Test.xlsx` — the documented one (see Test
  Project above). Sheet order: `Test_01`, `Test_02`, then hidden `Config`.
- `Legacy IT System Assessment single-system format 1 ITRQ - Test01.xlsx` —
  previously undocumented. Same three sheets and the same six named ranges
  (`EOL_DATE_STATUS`, `LU_LAYER`, `LU_MISSING_EOL_DATE_REASON`, `LU_PROVIDER`,
  `LU_REF_DATE`, `LU_SUPPORT_TYPE`), but the **hidden `Config` sheet comes
  first**. That ordering is what makes it useful: it exercises "first sheet"
  resolution and the include-hidden-worksheets flag, which the other fixture
  cannot distinguish. Note the spaces in the filename.

## Node Icon

Node icon extracted from the KNIME native Excel Reader node for visual
consistency in the node repository.
Located at: `icons/excelformreader.png`
