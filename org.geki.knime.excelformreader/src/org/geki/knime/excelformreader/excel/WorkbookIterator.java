package org.geki.knime.excelformreader.excel;

import java.io.Closeable;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.geki.knime.excelformreader.domain.ReadingMode;
import org.knime.core.node.NodeLogger;

public class WorkbookIterator implements Iterator<WorkbookIterator.Entry>, Closeable {

    private static final NodeLogger LOGGER = NodeLogger.getLogger(WorkbookIterator.class);

    public static final class Entry {
        public final Path filePath;
        public final String sheetName;
        public final Sheet sheet;
        public final Workbook workbook;

        Entry(final Path filePath, final String sheetName,
              final Sheet sheet, final Workbook workbook) {
            this.filePath = filePath;
            this.sheetName = sheetName;
            this.sheet = sheet;
            this.workbook = workbook;
        }
    }

    private final List<Path> files;
    private final Set<String> excludedSheets;

    private int fileIndex = 0;
    private Workbook currentWorkbook = null;
    private Path currentFilePath = null;
    private int sheetIndex = 0;

    // Pre-fetch state
    private Entry cachedNext = null;
    private boolean fetched = false;

    public WorkbookIterator(final Path rootPath,
                            final ReadingMode mode,
                            final Set<String> excludedSheets,
                            final boolean recursive) throws IOException {
        this.excludedSheets = (excludedSheets != null) ? excludedSheets : Collections.emptySet();

        if (mode == ReadingMode.SINGLE_FILE) {
            this.files = new ArrayList<>(Collections.singletonList(rootPath));
        } else {
            if (recursive) {
                try (Stream<Path> stream = Files.walk(rootPath)) {
                    this.files = stream
                        .filter(Files::isRegularFile)
                        .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".xlsx"))
                        .sorted()
                        .collect(Collectors.toList());
                }
            } else {
                try (Stream<Path> stream = Files.list(rootPath)) {
                    this.files = stream
                        .filter(Files::isRegularFile)
                        .filter(p -> p.getFileName().toString().toLowerCase().endsWith(".xlsx"))
                        .sorted()
                        .collect(Collectors.toList());
                }
            }
        }

        ensureFetched();
    }

    @Override
    public boolean hasNext() {
        ensureFetched();
        return cachedNext != null;
    }

    @Override
    public Entry next() {
        ensureFetched();
        if (cachedNext == null) {
            throw new NoSuchElementException("No more workbook entries");
        }
        final Entry result = cachedNext;
        cachedNext = null;
        fetched = false;
        return result;
    }

    @Override
    public void close() throws IOException {
        closeCurrentWorkbook();
    }

    private void ensureFetched() {
        if (!fetched) {
            cachedNext = findNext();
            fetched = true;
        }
    }

    // Advances internal state and returns the next valid (non-excluded) entry,
    // or null if all files and sheets are exhausted.
    // Workbooks are closed here only after all their sheets have been yielded —
    // so any entry's workbook is still open when the caller receives it.
    private Entry findNext() {
        while (true) {
            if (currentWorkbook != null) {
                while (sheetIndex < currentWorkbook.getNumberOfSheets()) {
                    final Sheet sheet = currentWorkbook.getSheetAt(sheetIndex++);
                    if (!isExcluded(sheet.getSheetName())) {
                        return new Entry(currentFilePath, sheet.getSheetName(), sheet, currentWorkbook);
                    }
                }
                // All sheets of this workbook exhausted — safe to close now because
                // the last entry from this workbook was already returned by next().
                closeCurrentWorkbook();
            }

            if (fileIndex >= files.size()) {
                return null;
            }

            final Path file = files.get(fileIndex++);
            try {
                currentWorkbook = WorkbookFactory.create(file.toFile(), null, true);
                currentFilePath = file;
                sheetIndex = 0;
            } catch (final IOException e) {
                LOGGER.warn("Cannot open workbook '" + file + "': " + e.getMessage());
                // Continue to the next file
            }
        }
    }

    private void closeCurrentWorkbook() {
        if (currentWorkbook != null) {
            try {
                currentWorkbook.close();
            } catch (final IOException e) {
                LOGGER.warn("Failed to close workbook '" + currentFilePath + "': " + e.getMessage());
            } finally {
                currentWorkbook = null;
                currentFilePath = null;
            }
        }
    }

    private boolean isExcluded(final String sheetName) {
        if (excludedSheets.isEmpty()) {
            return false;
        }
        return excludedSheets.contains(sheetName.trim().toLowerCase());
    }
}
