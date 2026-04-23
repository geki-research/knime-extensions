package org.geki.knime.excelformreader.domain;

public enum ReadingMode {
    SINGLE_FILE("Single File"),
    FOLDER("Folder");

    private final String label;

    ReadingMode(final String label) {
        this.label = label;
    }

    @Override
    public String toString() {
        return label;
    }

    public static ReadingMode fromString(final String label) {
        if (label == null) {
            throw new IllegalArgumentException("ReadingMode label must not be null");
        }
        for (final ReadingMode mode : values()) {
            if (mode.label.equalsIgnoreCase(label.trim())) {
                return mode;
            }
        }
        throw new IllegalArgumentException("Unknown ReadingMode: '" + label + "'");
    }
}
