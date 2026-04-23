package org.geki.knime.excelformreader.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.knime.core.data.DataCell;

public class CellValueConverter {

    /**
     * Converts an Apache POI Cell to a KNIME DataCell using the requested type hint.
     *
     * @param cell     the POI cell (may be null → MissingCell)
     * @param dataType one of "string", "int", "double", "date", "boolean", or null for auto
     */
    public DataCell convert(final Cell cell, final String dataType) {
        // TODO: evaluate formulas, map POI cell types to KNIME DataCell subtypes,
        //       apply dataType coercion, return MissingCell for blanks/errors
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
