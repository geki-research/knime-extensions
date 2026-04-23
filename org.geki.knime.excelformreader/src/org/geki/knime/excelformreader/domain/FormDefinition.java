package org.geki.knime.excelformreader.domain;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;

public class FormDefinition {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(FormDefinition.class);

    private final List<FieldMapping> fields;

    public FormDefinition(final List<FieldMapping> fields) {
        if (fields == null || fields.isEmpty()) {
            throw new IllegalArgumentException("fields must not be null or empty");
        }
        this.fields = Collections.unmodifiableList(new ArrayList<>(fields));
    }

    public List<FieldMapping> getFields() { return fields; }
    public int size() { return fields.size(); }

    // Private no-arg constructor for the configure()-time sentinel (empty fields list).
    private FormDefinition() {
        this.fields = Collections.emptyList();
    }

    /**
     * Validates that the required columns are present in the spec and returns a sentinel
     * FormDefinition with no fields. Used at configure() time when no row data is available.
     */
    public static FormDefinition fromSpec(final DataTableSpec spec)
            throws InvalidSettingsException {
        if (findColumnIndex(spec, "field_name") < 0) {
            throw new InvalidSettingsException(
                "Form definition table is missing required column 'field_name'");
        }
        if (findColumnIndex(spec, "value_cell") < 0) {
            throw new InvalidSettingsException(
                "Form definition table is missing required column 'value_cell'");
        }
        return new FormDefinition();
    }

    public static FormDefinition fromDataTable(final BufferedDataTable table)
            throws InvalidSettingsException {
        final DataTableSpec spec = table.getDataTableSpec();

        final int fieldNameIdx = findColumnIndex(spec, "field_name");
        if (fieldNameIdx < 0) {
            throw new InvalidSettingsException(
                "Form definition table is missing required column 'field_name'");
        }
        final int valueCellIdx = findColumnIndex(spec, "value_cell");
        if (valueCellIdx < 0) {
            throw new InvalidSettingsException(
                "Form definition table is missing required column 'value_cell'");
        }
        final int dataTypeIdx = findColumnIndex(spec, "data_type");
        final int sheetNameIdx = findColumnIndex(spec, "sheet_name");

        final List<FieldMapping> mappings = new ArrayList<>();
        for (final DataRow row : table) {
            final DataCell fieldNameCell = row.getCell(fieldNameIdx);
            final DataCell valueCellCell = row.getCell(valueCellIdx);

            if (fieldNameCell.isMissing() || valueCellCell.isMissing()) {
                LOGGER.warn("Skipping row " + row.getKey() + ": field_name or value_cell is missing");
                continue;
            }

            final String fieldName = fieldNameCell.toString();
            final String valueCell = valueCellCell.toString();
            final String dataType = (dataTypeIdx >= 0 && !row.getCell(dataTypeIdx).isMissing())
                ? row.getCell(dataTypeIdx).toString() : null;
            final String sheetName = (sheetNameIdx >= 0 && !row.getCell(sheetNameIdx).isMissing())
                ? row.getCell(sheetNameIdx).toString() : null;

            try {
                mappings.add(new FieldMapping(fieldName, valueCell, dataType, sheetName));
            } catch (final IllegalArgumentException e) {
                LOGGER.warn("Skipping row " + row.getKey() + ": " + e.getMessage());
            }
        }

        if (mappings.isEmpty()) {
            throw new InvalidSettingsException(
                "Form definition table contains no valid field mappings");
        }

        return new FormDefinition(mappings);
    }

    private static int findColumnIndex(final DataTableSpec spec, final String columnName) {
        for (int i = 0; i < spec.getNumColumns(); i++) {
            if (spec.getColumnSpec(i).getName().equalsIgnoreCase(columnName)) {
                return i;
            }
        }
        return -1;
    }
}
