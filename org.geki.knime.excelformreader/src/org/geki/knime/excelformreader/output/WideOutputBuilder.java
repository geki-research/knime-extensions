package org.geki.knime.excelformreader.output;

import java.util.Map;

import org.geki.knime.excelformreader.domain.FieldMapping;
import org.geki.knime.excelformreader.domain.FormDefinition;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;

public class WideOutputBuilder {

    private final DataTableSpec spec;
    private final boolean includeProvenance;

    public WideOutputBuilder(final DataTableSpec spec, final boolean includeProvenance) {
        this.spec = spec;
        this.includeProvenance = includeProvenance;
    }

    public DataRow buildRow(final String sourceFile,
                             final String sheetName,
                             final Map<String, DataCell> extractedValues,
                             final FormDefinition definition,
                             final long rowIndex) {
        final DataCell[] cells = new DataCell[spec.getNumColumns()];
        int i = 0;

        if (includeProvenance) {
            cells[i++] = new StringCell(sourceFile != null ? sourceFile : "");
            cells[i++] = new StringCell(sheetName != null ? sheetName : "");
        }

        for (final FieldMapping mapping : definition.getFields()) {
            final DataCell value = extractedValues != null
                ? extractedValues.get(mapping.getFieldName())
                : null;
            cells[i++] = (value != null) ? value : DataType.getMissingCell();
        }

        // Fill any unexpected trailing slots defensively
        while (i < cells.length) {
            cells[i++] = DataType.getMissingCell();
        }

        return new DefaultRow(new RowKey("Row" + rowIndex), cells);
    }
}
