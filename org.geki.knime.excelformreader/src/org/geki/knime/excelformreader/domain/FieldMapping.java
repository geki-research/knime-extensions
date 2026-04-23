package org.geki.knime.excelformreader.domain;

public class FieldMapping {

    private final String fieldName;
    private final String valueCell;
    private final String dataType;
    private final String sheetName;

    public FieldMapping(final String fieldName, final String valueCell,
                        final String dataType, final String sheetName) {
        this.fieldName = fieldName;
        this.valueCell = valueCell;
        this.dataType = dataType;
        this.sheetName = sheetName;
    }

    public String getFieldName() { return fieldName; }
    public String getValueCell() { return valueCell; }
    public String getDataType() { return dataType; }
    public String getSheetName() { return sheetName; }
}
