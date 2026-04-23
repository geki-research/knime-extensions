# knime-extensions

Community KNIME extensions developed by [geki-research](https://github.com/geki-research).

## Extensions

### Excel Form Reader
A KNIME node that extracts data from non-tabular, form-structured Excel
worksheets (.xlsx) into a standard KNIME data table.

**Key features:**
- Form structure is defined via a configurable input table (field name → cell address mapping) — supports 60+ fields without cumbersome dialog configuration
- Supports single file, all sheets in a workbook, folder, and recursive folder reading modes — consistent with KNIME's native Excel Reader UX
- Each worksheet is treated as one form instance, producing one row per sheet in wide output mode
- Output format configurable: wide (one row per form) or long (one row per field)
- Per-field sheet overrides via the definition table
- Cell ranges supported — read as delimited strings
- Formula evaluation via Apache POI — computed values, not formula strings
- Configurable sheet exclusion list (default: sheets named "Config")
- Optional provenance columns: source file path and sheet name
- Configurable error handling: fail or warn-and-set-missing-value per error type

**Node inputs:**
| Port | Type | Description |
|---|---|---|
| 0 | Table | Form definition table |

**Form definition table schema:**

| Column | Required | Type | Description |
|---|---|---|---|
| `field_name` | ✅ | String | Output column name |
| `value_cell` | ✅ | String | Cell address (`C4`) or range (`B10:D15`) |
| `data_type` | ❌ | String | `string` / `int` / `double` / `date` / `boolean` |
| `sheet_name` | ❌ | String | Per-field sheet override |

## Development Setup

### Prerequisites
- Debian 12 (or compatible Linux)
- JDK 17 (`openjdk-17-jdk`)
- Maven 3.9+
- Eclipse for RCP and RAP Developers 2024-03+
- KNIME SDK Setup: https://github.com/knime/knime-sdk-setup

### Build
```bash
mvn clean verify
```

### Eclipse Import
1. Import all projects from this repository into Eclipse
2. Activate the KNIME target platform via `knime-sdk-setup`
3. Run a KNIME launch configuration to test nodes interactively

## License
Apache License 2.0 — see [LICENSE](LICENSE)
