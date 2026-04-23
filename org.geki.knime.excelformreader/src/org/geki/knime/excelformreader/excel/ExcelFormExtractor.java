package org.geki.knime.excelformreader.excel;

import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.knime.core.data.DataCell;

import org.geki.knime.excelformreader.domain.FormDefinition;

public class ExcelFormExtractor {

    /**
     * Extracts field values from a single sheet according to the form definition.
     *
     * @return map from field name to extracted DataCell (may be MissingCell on error)
     */
    public Map<String, DataCell> extract(final Sheet sheet, final FormDefinition def) {
        // TODO: for each FieldMapping, resolve sheet override, parse CellAddress,
        //       read cell(s), delegate to CellValueConverter, handle errors per policy
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
