package org.geki.knime.excelformreader.output;

import org.knime.core.data.DataTableSpec;

import org.geki.knime.excelformreader.domain.FormDefinition;

public class OutputSpecFactory {

    /**
     * Creates a wide-format spec: one column per field, optional provenance columns appended.
     */
    public DataTableSpec createWideSpec(final FormDefinition def, final boolean includeProvenance) {
        // TODO: map each FieldMapping.fieldName → DataColumnSpec using declared dataType;
        //       append file_path (StringCell) and sheet_name (StringCell) if includeProvenance
        throw new UnsupportedOperationException("Not yet implemented");
    }

    /**
     * Creates a long-format spec: field_name, value, optional provenance columns.
     */
    public DataTableSpec createLongSpec(final boolean includeProvenance) {
        // TODO: build spec with columns: field_name (String), value (String),
        //       and optionally file_path (String), sheet_name (String)
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
