package org.geki.knime.excelformreader.domain;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.knime.core.node.BufferedDataTable;

public class FormDefinition {

    private final List<FieldMapping> mappings;

    public FormDefinition(final List<FieldMapping> mappings) {
        this.mappings = Collections.unmodifiableList(new ArrayList<>(mappings));
    }

    public List<FieldMapping> getMappings() { return mappings; }

    public static FormDefinition fromDataTable(final BufferedDataTable table) {
        // TODO: read field_name, value_cell, data_type (optional), sheet_name (optional) columns
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
