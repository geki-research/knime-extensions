package org.geki.knime.excelformreader.output;

import java.util.Map;

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;

public class WideOutputBuilder {

    /**
     * Builds a single wide-format DataRow from extracted field values.
     *
     * @param file   source file path string (used for provenance column)
     * @param sheet  source sheet name (used for provenance column)
     * @param values map from field name to DataCell as produced by ExcelFormExtractor
     * @param spec   the DataTableSpec for the output table
     */
    public DataRow buildRow(final String file, final String sheet,
                            final Map<String, DataCell> values, final DataTableSpec spec) {
        // TODO: align values map to spec column order, append provenance cells if present
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
