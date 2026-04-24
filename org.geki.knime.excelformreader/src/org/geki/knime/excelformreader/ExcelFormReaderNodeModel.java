package org.geki.knime.excelformreader;

import java.io.File;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

import org.geki.knime.excelformreader.ExcelFormReaderSettings.InputMode;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.OutputFormat;
import org.geki.knime.excelformreader.domain.FormDefinition;
import org.geki.knime.excelformreader.domain.ReadingMode;
import org.geki.knime.excelformreader.excel.ExcelFormExtractor;
import org.geki.knime.excelformreader.excel.WorkbookIterator;
import org.geki.knime.excelformreader.output.LongOutputBuilder;
import org.geki.knime.excelformreader.output.OutputSpecFactory;
import org.geki.knime.excelformreader.output.WideOutputBuilder;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.port.PortType;

public class ExcelFormReaderNodeModel extends NodeModel {

    private static final NodeLogger LOGGER =
        NodeLogger.getLogger(ExcelFormReaderNodeModel.class);

    private final ExcelFormReaderSettings m_settings = new ExcelFormReaderSettings();

    public ExcelFormReaderNodeModel() {
        super(new PortType[]{BufferedDataTable.TYPE},
              new PortType[]{BufferedDataTable.TYPE});
    }

    @Override
    protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
            throws InvalidSettingsException {
        if (inSpecs[0] == null) {
            throw new InvalidSettingsException("Form definition table is not connected");
        }

        final InputMode mode = m_settings.getInputMode();
        final String path = mode == InputMode.SINGLE_FILE
            ? m_settings.getFilePath()
            : m_settings.getFolderPath();
        if (path == null || path.trim().isEmpty()) {
            throw new InvalidSettingsException(
                "No " + (mode == InputMode.SINGLE_FILE ? "file" : "folder") + " path configured");
        }

        final FormDefinition definition = FormDefinition.fromSpec(inSpecs[0]);

        final DataTableSpec spec;
        if (m_settings.getOutputFormat() == OutputFormat.WIDE) {
            spec = OutputSpecFactory.createWideSpec(definition,
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName());
        } else {
            spec = OutputSpecFactory.createLongSpec(
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName());
        }

        return new DataTableSpec[]{spec};
    }

    @Override
    protected BufferedDataTable[] execute(final BufferedDataTable[] inData,
                                           final ExecutionContext exec) throws Exception {
        final FormDefinition definition = FormDefinition.fromDataTable(inData[0]);

        final boolean wide = m_settings.getOutputFormat() == OutputFormat.WIDE;
        final DataTableSpec spec = wide
            ? OutputSpecFactory.createWideSpec(definition,
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName())
            : OutputSpecFactory.createLongSpec(
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName());

        final BufferedDataContainer container = exec.createDataContainer(spec);

        final WideOutputBuilder wideBuilder = wide
            ? new WideOutputBuilder(spec,
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName())
            : null;
        final LongOutputBuilder longBuilder = !wide
            ? new LongOutputBuilder(spec,
                m_settings.isIncludeSourceFilename(),
                m_settings.isIncludeSheetName())
            : null;

        final ExcelFormExtractor extractor = new ExcelFormExtractor(m_settings);

        final InputMode inputMode = m_settings.getInputMode();
        final String path = inputMode == InputMode.SINGLE_FILE
            ? m_settings.getFilePath()
            : m_settings.getFolderPath();
        final ReadingMode readingMode = inputMode == InputMode.FOLDER
            ? ReadingMode.FOLDER
            : ReadingMode.SINGLE_FILE;

        long rowIndex = 0;
        try (WorkbookIterator iterator = new WorkbookIterator(
                Paths.get(path),
                readingMode,
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.getFileSheetSelection()
                    : ExcelFormReaderSettings.SheetSelection.FIRST,
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.getFileSheetName() : "",
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.getFileSheetPosition() : 0,
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.getFileSheetFilterMode()
                    : m_settings.getFolderSheetFilterMode(),
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.getFileSheetFilterNames()
                    : m_settings.getFolderSheetFilterNames(),
                inputMode == InputMode.SINGLE_FILE
                    ? m_settings.isFileIncludeHiddenSheets()
                    : m_settings.isFolderIncludeHiddenSheets(),
                m_settings.isRecursive(),
                m_settings.isIncludeHiddenFiles(),
                m_settings.isIncludeHiddenFolders(),
                m_settings.isFilterByExtension(),
                m_settings.getFileExtensions())) {

            while (iterator.hasNext()) {
                final WorkbookIterator.Entry entry = iterator.next();

                exec.checkCanceled();
                exec.setMessage("Reading sheet '" + entry.sheetName
                    + "' from " + entry.filePath.getFileName());

                final Map<String, DataCell> values =
                    extractor.extract(entry.sheet, definition, entry.workbook);

                final String filePath = entry.filePath.toString();
                final String sheetName = entry.sheetName;

                if (wide) {
                    final DataRow row = wideBuilder.buildRow(
                        filePath, sheetName, values, definition, rowIndex++);
                    container.addRowToTable(row);
                } else {
                    final List<DataRow> rows = longBuilder.buildRows(
                        filePath, sheetName, values, definition, rowIndex);
                    rows.forEach(container::addRowToTable);
                    rowIndex += rows.size();
                }
            }
        }

        container.close();
        return new BufferedDataTable[]{container.getTable()};
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings) {
        m_settings.saveSettings(settings);
    }

    @Override
    protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        m_settings.loadSettings(settings);
    }

    @Override
    protected void validateSettings(final NodeSettingsRO settings)
            throws InvalidSettingsException {
        m_settings.validateSettings(settings);
    }

    @Override
    protected void reset() {}

    @Override
    protected void loadInternals(final File nodeInternDir, final ExecutionMonitor exec)
            throws CanceledExecutionException {}

    @Override
    protected void saveInternals(final File nodeInternDir, final ExecutionMonitor exec)
            throws CanceledExecutionException {}
}
