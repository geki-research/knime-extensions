package org.geki.knime.excelformreader.domain;

public class CellAddress {

    private final String sheetName;
    private final int column;
    private final int row;
    private final boolean isRange;
    private final int rangeEndCol;
    private final int rangeEndRow;

    private CellAddress(final String sheetName, final int column, final int row,
                        final boolean isRange, final int rangeEndCol, final int rangeEndRow) {
        this.sheetName = sheetName;
        this.column = column;
        this.row = row;
        this.isRange = isRange;
        this.rangeEndCol = rangeEndCol;
        this.rangeEndRow = rangeEndRow;
    }

    public String getSheetName() { return sheetName; }
    public int getColumn() { return column; }
    public int getRow() { return row; }
    public boolean isRange() { return isRange; }
    public int getRangeEndCol() { return rangeEndCol; }
    public int getRangeEndRow() { return rangeEndRow; }

    /**
     * Parses a cell address string such as "C4", "B10:D15", or "Sheet1!C4".
     */
    public static CellAddress parse(final String address) {
        // TODO: implement parsing for A1 notation, ranges, and optional sheet prefix
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
