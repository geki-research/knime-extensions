package org.geki.knime.excelformreader.domain;

import java.util.Set;

public class FieldMapping {

    private static final Set<String> VALID_DATA_TYPES = Set.of("string", "int", "double", "date", "boolean");

    private final String fieldName;
    private final CellAddress address;
    private final String dataType;

    public FieldMapping(final String fieldName, final String valueCell,
                        final String dataType) {
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
    }

    public String getFieldName() { return fieldName; }
    public CellAddress getAddress() { return address; }
    public String getDataType() { return dataType; }
}
