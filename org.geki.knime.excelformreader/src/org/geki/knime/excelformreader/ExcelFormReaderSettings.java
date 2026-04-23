package org.geki.knime.excelformreader;

import java.util.Arrays;
import java.util.Set;
import java.util.stream.Collectors;

import org.geki.knime.excelformreader.domain.ReadingMode;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

public class ExcelFormReaderSettings {

    private static final String CFG_READING_MODE    = "cfg_readingMode";
    private static final String CFG_PATH            = "cfg_path";
    private static final String CFG_RECURSIVE       = "cfg_recursive";
    private static final String CFG_DEFAULT_SHEET   = "cfg_defaultSheet";
    private static final String CFG_EXCLUDED_SHEETS = "cfg_excludedSheets";
    private static final String CFG_OUTPUT_FORMAT   = "cfg_outputFormat";
    private static final String CFG_ADD_PROVENANCE  = "cfg_addProvenance";
    private static final String CFG_RANGE_DELIMITER = "cfg_rangeDelimiter";
    private static final String CFG_ON_MISSING_CELL = "cfg_onMissingCell";
    private static final String CFG_ON_BAD_VALUE    = "cfg_onBadValue";

    private final SettingsModelString  m_readingMode    = new SettingsModelString(CFG_READING_MODE, "SINGLE_FILE");
    private final SettingsModelString  m_path           = new SettingsModelString(CFG_PATH, "");
    private final SettingsModelBoolean m_recursive      = new SettingsModelBoolean(CFG_RECURSIVE, false);
    private final SettingsModelString  m_defaultSheet   = new SettingsModelString(CFG_DEFAULT_SHEET, "");
    private final SettingsModelString  m_excludedSheets = new SettingsModelString(CFG_EXCLUDED_SHEETS, "Config");
    private final SettingsModelString  m_outputFormat   = new SettingsModelString(CFG_OUTPUT_FORMAT, "WIDE");
    private final SettingsModelBoolean m_addProvenance  = new SettingsModelBoolean(CFG_ADD_PROVENANCE, true);
    private final SettingsModelString  m_rangeDelimiter = new SettingsModelString(CFG_RANGE_DELIMITER, ", ");
    private final SettingsModelString  m_onMissingCell  = new SettingsModelString(CFG_ON_MISSING_CELL, "WARN");
    private final SettingsModelString  m_onBadValue     = new SettingsModelString(CFG_ON_BAD_VALUE, "WARN");

    // Typed getters

    public ReadingMode getReadingMode() {
        return ReadingMode.valueOf(m_readingMode.getStringValue());
    }

    public String getPath() { return m_path.getStringValue(); }

    public boolean isRecursive() { return m_recursive.getBooleanValue(); }

    public String getDefaultSheet() { return m_defaultSheet.getStringValue(); }

    public Set<String> getExcludedSheets() {
        return Arrays.stream(m_excludedSheets.getStringValue().split(","))
            .map(s -> s.trim().toLowerCase())
            .filter(s -> !s.isEmpty())
            .collect(Collectors.toSet());
    }

    public String getOutputFormat() { return m_outputFormat.getStringValue(); }

    public boolean isAddProvenance() { return m_addProvenance.getBooleanValue(); }

    public String getRangeDelimiter() { return m_rangeDelimiter.getStringValue(); }

    public String getOnMissingCell() { return m_onMissingCell.getStringValue(); }

    public String getOnBadValue() { return m_onBadValue.getStringValue(); }

    // Raw SettingsModel getters for NodeDialog binding

    public SettingsModelString getReadingModeModel() { return m_readingMode; }
    public SettingsModelString getPathModel() { return m_path; }
    public SettingsModelBoolean getRecursiveModel() { return m_recursive; }
    public SettingsModelString getDefaultSheetModel() { return m_defaultSheet; }
    public SettingsModelString getExcludedSheetsModel() { return m_excludedSheets; }
    public SettingsModelString getOutputFormatModel() { return m_outputFormat; }
    public SettingsModelBoolean getAddProvenanceModel() { return m_addProvenance; }
    public SettingsModelString getRangeDelimiterModel() { return m_rangeDelimiter; }
    public SettingsModelString getOnMissingCellModel() { return m_onMissingCell; }
    public SettingsModelString getOnBadValueModel() { return m_onBadValue; }

    public void saveSettings(final NodeSettingsWO settings) {
        m_readingMode.saveSettingsTo(settings);
        m_path.saveSettingsTo(settings);
        m_recursive.saveSettingsTo(settings);
        m_defaultSheet.saveSettingsTo(settings);
        m_excludedSheets.saveSettingsTo(settings);
        m_outputFormat.saveSettingsTo(settings);
        m_addProvenance.saveSettingsTo(settings);
        m_rangeDelimiter.saveSettingsTo(settings);
        m_onMissingCell.saveSettingsTo(settings);
        m_onBadValue.saveSettingsTo(settings);
    }

    public void loadSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_readingMode.loadSettingsFrom(settings);
        m_path.loadSettingsFrom(settings);
        m_recursive.loadSettingsFrom(settings);
        m_defaultSheet.loadSettingsFrom(settings);
        m_excludedSheets.loadSettingsFrom(settings);
        m_outputFormat.loadSettingsFrom(settings);
        m_addProvenance.loadSettingsFrom(settings);
        m_rangeDelimiter.loadSettingsFrom(settings);
        m_onMissingCell.loadSettingsFrom(settings);
        m_onBadValue.loadSettingsFrom(settings);
    }

    public void validateSettings(final NodeSettingsRO settings) throws InvalidSettingsException {
        m_readingMode.validateSettings(settings);
        m_path.validateSettings(settings);
        m_recursive.validateSettings(settings);
        m_defaultSheet.validateSettings(settings);
        m_excludedSheets.validateSettings(settings);
        m_outputFormat.validateSettings(settings);
        m_addProvenance.validateSettings(settings);
        m_rangeDelimiter.validateSettings(settings);
        m_onMissingCell.validateSettings(settings);
        m_onBadValue.validateSettings(settings);
    }
}
