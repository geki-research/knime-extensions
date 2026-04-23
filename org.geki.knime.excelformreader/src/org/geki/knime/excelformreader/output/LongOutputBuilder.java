package org.geki.knime.excelformreader.output;

import java.util.List;
import java.util.Map;

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;

public class LongOutputBuilder {

    /**
     * Builds one DataRow per field from extracted values (long/unpivoted format).
     *
     * @param file   source file path string (used for provenance column)
     * @param sheet  source sheet name (used for provenance column)
     * @param values map from field name to DataCell as produced by ExcelFormExtractor
     * @param spec   the DataTableSpec for the output table
     */
    public List<DataRow> buildRows(final String file, final String sheet,
                                   final Map<String, DataCell> values, final DataTableSpec spec) {
        // TODO: emit one row per entry in values: field_name, value.toString(),
        //       optionally file_path and sheet_name
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
