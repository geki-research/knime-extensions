package org.geki.knime.excelformreader;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import org.geki.knime.excelformreader.domain.ReadingMode;

public class ExcelFormReaderSettings {

    private final SettingsModelString m_filePath =
            new SettingsModelString("file_path", "");

    private final SettingsModelString m_readingMode =
            new SettingsModelString("reading_mode", ReadingMode.SINGLE_FILE.name());

    private final SettingsModelBoolean m_recursive =
            new SettingsModelBoolean("recursive", false);

    private final SettingsModelString m_excludedSheets =
            new SettingsModelString("excluded_sheets", "Config");

    private final SettingsModelString m_outputFormat =
            new SettingsModelString("output_format", "wide");

    private final SettingsModelBoolean m_includeProvenance =
            new SettingsModelBoolean("include_provenance", true);

    private final SettingsModelString m_errorHandling =
            new SettingsModelString("error_handling", "fail");

    public void saveSettings(final NodeSettingsWO settings) {
        m_filePath.saveSettingsTo(settings);
        m_readingMode.saveSettingsTo(settings);
        m_recursive.saveSettingsTo(settings);
        m_excludedSheets.saveSettingsTo(settings);
        m_outputFormat.saveSettingsTo(settings);
        m_includeProvenance.saveSettingsTo(settings);
        m_errorHandling.saveSettingsTo(settings);
    }

    public void loadSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_filePath.loadSettingsFrom(settings);
        m_readingMode.loadSettingsFrom(settings);
        m_recursive.loadSettingsFrom(settings);
        m_excludedSheets.loadSettingsFrom(settings);
        m_outputFormat.loadSettingsFrom(settings);
        m_includeProvenance.loadSettingsFrom(settings);
        m_errorHandling.loadSettingsFrom(settings);
    }

    public String getFilePath() { return m_filePath.getStringValue(); }
    public ReadingMode getReadingMode() { return ReadingMode.valueOf(m_readingMode.getStringValue()); }
    public boolean isRecursive() { return m_recursive.getBooleanValue(); }
    public String getExcludedSheets() { return m_excludedSheets.getStringValue(); }
    public String getOutputFormat() { return m_outputFormat.getStringValue(); }
    public boolean isIncludeProvenance() { return m_includeProvenance.getBooleanValue(); }
    public String getErrorHandling() { return m_errorHandling.getStringValue(); }
}
