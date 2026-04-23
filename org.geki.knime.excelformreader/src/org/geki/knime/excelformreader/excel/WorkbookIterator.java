package org.geki.knime.excelformreader.excel;

import java.nio.file.Path;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;

import org.geki.knime.excelformreader.domain.ReadingMode;

public class WorkbookIterator implements Iterator<WorkbookIterator.Entry> {

    public record Entry(Path filePath, String sheetName, Sheet sheet) {}

    public WorkbookIterator(final Path root, final ReadingMode mode,
                            final Set<String> excludedSheets, final boolean recursive) {
        // TODO: initialize file traversal based on mode and recursive flag
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    public boolean hasNext() {
        // TODO: advance through files and sheets, skipping excluded sheet names
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    public Entry next() {
        // TODO: return next (filePath, sheetName, sheet) entry
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
