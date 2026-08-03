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
| Java | Sources compile at **17**; bundle declares **JavaSE-21** — see below |
| Build system | Maven 3.9+ with Eclipse Tycho 4.0.13 on `main` (the `releases/*` branches are still on 4.0.6) |
| IDE | Eclipse for RCP and RAP Developers 2024-03 |
| Eclipse workspace | `~/knime-dev/workspace` |
| KNIME target platform | `~/knime-dev/knime-sdk-setup` → `KNIME-AP.target` (1897 plugins) |

### Java level — 17 compile, 21 runtime (deliberate)

| Setting | Value | Where |
|---|---|---|
| Source/target compile level | **17** | `tycho-compiler-plugin` in `pom.xml` (authoritative) |
| `maven.compiler.source/target` | **17** | `pom.xml` properties (kept in sync; not what governs) |
| `Bundle-RequiredExecutionEnvironment` | **JavaSE-21** | plugin `MANIFEST.MF` |

The divergence is intentional. KNIME 5.12+ ships and runs on Java 21, which is
why the bundle declares `JavaSE-21` as its BREE. Sources nevertheless compile at
**17** so the same code stays portable to the `releases/5.5` and `releases/5.8`
branches, whose target platforms are still on Java 17. Java 17 bytecode runs
correctly under a Java 21 runtime, so nothing is lost on newer platforms.
Verified: compiled classes are major version 61 (Java 17).

`releases/5.12` follows a **different convention** — PR #2 compiles that branch
at 21. Do not assume the 17 rule holds there.

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

**This table is the only place in this file that makes claims about branches
other than the one you are on.** Everywhere else describes the current branch.
Anything that changes another branch — a merged PR, a Tycho bump, a profile
change — **must update this table**, because nothing in the workflow will catch
the drift: the other branches are never built while working on `main`. If you
need a per-branch fact, get it here; if you learn one, record it here and nowhere
else.

| Branch | Tycho | `maven.compiler.*` | `tycho-compiler-plugin` config | What governs compilation | BREE | Default profile → p2 URL | Test tally |
|---|---|---|---|---|---|---|---|
| `main` | 4.0.13 | 17 | `<source>17</source><target>17</target>` in `<build><plugins>` | **`tycho-compiler-plugin` → 17** | JavaSE-21 | `knime-nightly` → `https://update.knime.com/analytics-platform/nightly` | 150 / 1 skipped |
| `releases/5.5` | 4.0.6 | 17 | `<source>17</source><target>17</target>` in `<build><plugins>` | **`tycho-compiler-plugin` → 17** | JavaSE-17 | `knime-5.5` → `https://update.knime.com/analytics-platform/5.5` | not measured |
| `releases/5.8` | 4.0.6 | 17 | `<source>17</source><target>17</target>` in `<build><plugins>` | **`tycho-compiler-plugin` → 17** | JavaSE-17 | `knime-5.8` → `https://update.knime.com/analytics-platform/5.8` | not measured |
| `releases/5.12` | 4.0.13 | **21** | **none** — declared in `<pluginManagement>` with `<version>` only, no `<configuration>` | **`maven.compiler.*` → 21** | JavaSE-21 | `knime-5.12` → `https://update.knime.com/analytics-platform/lts/5.12` | 65 / 1 skipped |

**Why `releases/5.12` differs.** Tycho's `source`/`target` parameters default to
`${maven.compiler.source}` / `${maven.compiler.target}`. On `main`, `5.5` and
`5.8` an explicit `<configuration>` block overrides that, so the properties are
inert. On `releases/5.12` the plugin appears only under `<pluginManagement>` with
a version and no configuration, so nothing overrides the defaults and the
properties are what actually set the compile level. **The rule is reversed on
that branch** — check which form applies before changing a compile level.

**Provenance.** Every value above was read from the branch itself via
`git show origin/<branch>:<path>` against `pom.xml` and
`org.geki.knime.excelformreader/META-INF/MANIFEST.MF`. The `main` compile level
is additionally confirmed by measured bytecode (major version 61). The
"what governs" column is derived from POM structure plus Tycho's documented
parameter defaults, not from measured bytecode on the release branches.

**`not measured` means not measured.** No test run has been performed on
`releases/5.5` or `releases/5.8` in this work. Both carry 5 test classes — the
same count as `releases/5.12`, which runs 65 — but **do not infer a tally from
that**; record one only after actually running the build on those branches. Note
that `mvn -U clean verify -P knime-5.5` on `main` reports 150: that is `main`'s
code against the 5.5 target platform, not the `releases/5.5` branch.

**Latent inconsistency, not currently a fault.** On `releases/5.5` and
`releases/5.8` the *non-default* `knime-5.12` profile still points at the
unqualified `https://update.knime.com/analytics-platform/5.12`, not the `lts/`
URL. Harmless today — both branches are on Tycho 4.0.6, which handles the
redirect, and that profile is not their default — but it would break if either
branch moved to 4.0.13. See the `lts/` note below.

### Active profiles

Maven profiles in `pom.xml` control the active p2 repository. Each release
branch sets its profile as `activeByDefault`; `main` defaults to nightly.

```bash
mvn -U clean verify              # uses nightly profile (main branch default)
mvn -U clean verify -P knime-5.5 # explicitly use 5.5 update site
```

The `5.12` entry must be the `lts/` URL. Tycho 4.0.13 caches the
`/analytics-platform/5.12` → `/analytics-platform/lts/5.12` redirect body under
the original URL key and then cannot read it, so the unqualified URL fails at
repository load.

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

### PR #2 — MERGED into `releases/5.12` on 2026-08-03

"Fix for 5.12 builds", by `dsaam94` (Ali Marvi, KNIME). Reviewed in depth,
assessed sound, and merged as **`38f2515`** — a true merge commit with two
parents, so Ali Marvi's commit `1a7ed37` and authorship are preserved intact.
`releases/5.12` is now at `38f2515` and builds green under `-P knime-5.12` at
**65 tests / 1 skipped**.

It was merged while the contributor was out of office rather than leaving the
5.12 branch blocked for several weeks. That was a deliberate call, made on the
strength of the review, not an assumption that the open questions were settled.

Those two questions are **not abandoned** — they now live in **issue #3**:
https://github.com/geki-research/knime-extensions/issues/3

1. **What was actually failing in the 5.12 Jenkins build?** The PR body is empty
   and no build log is linked; the branch built green locally both before and
   after the change, so the fix could never be checked against its symptom.
2. **Is `skipArchive=true` intentional?** Presumed so — the same engineer owns
   the Jenkins job consuming the output — but unconfirmed.

**Practical consequence of `skipArchive`, measured after the merge:** a
successful build of `releases/5.12` produces **no `.zip`** in
`org.geki.knime.excelformreader.update/target/` — only the expanded
`repository/`. `main` still produces
`org.geki.knime.excelformreader.update-1.0.0-SNAPSHOT.zip`. Anyone hand-building
`releases/5.12` and expecting a distributable archive will not get one.

### `releases/5.12` lags `main` by 85 unit tests

65 tests there vs 150 on `main`. A forward-port is worthwhile and **has not been
done**. Unaffected by the PR #2 merge — still outstanding.

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
