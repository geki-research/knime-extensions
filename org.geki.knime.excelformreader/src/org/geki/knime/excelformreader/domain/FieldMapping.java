package org.geki.knime.excelformreader.domain;

import java.util.Set;

public class FieldMapping {

    private static final Set<String> VALID_DATA_TYPES = Set.of("string", "int", "double", "date", "boolean");

    private final String fieldName;
    private final CellAddress address;
    private final String dataType;
    private final String sheetName;

    public FieldMapping(final String fieldName, final String valueCell,
                        final String dataType, final String sheetName) {
        if (fieldName == null || fieldName.trim().isEmpty()) {
            throw new IllegalArgumentException("fieldName must not be null or blank");
        }
        this.fieldName = fieldName.trim();
        this.address = CellAddress.parse(valueCell);

        final String normalizedType = (dataType == null || dataType.trim().isEmpty())
            ? "string"
            : dataType.trim().toLowerCase();
        if (!VALID_DATA_TYPES.contains(normalizedType)) {
            throw new IllegalArgumentException(
                "Invalid data type: '" + dataType + "'. Must be one of: " + VALID_DATA_TYPES);
        }
        this.dataType = normalizedType;

        final String trimmedSheet = (sheetName == null) ? null : sheetName.trim();
        this.sheetName = (trimmedSheet == null || trimmedSheet.isEmpty()) ? null : trimmedSheet;
    }

    public String getFieldName() { return fieldName; }
    public CellAddress getAddress() { return address; }
    public String getDataType() { return dataType; }
    public String getSheetName() { return sheetName; }
}
