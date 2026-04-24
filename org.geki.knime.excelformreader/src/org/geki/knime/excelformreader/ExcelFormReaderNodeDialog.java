package org.geki.knime.excelformreader;

import org.geki.knime.excelformreader.ExcelFormReaderSettings.ErrorHandling;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.InputMode;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.OutputFormat;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;

import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTabbedPane;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.lang.reflect.Field;

public class ExcelFormReaderNodeDialog extends NodeDialogPane {

    private final ExcelFormReaderSettings m_settings = new ExcelFormReaderSettings();

    // Input group
    private final JRadioButton m_singleFileRadio = new JRadioButton("Single File");
    private final JRadioButton m_folderRadio     = new JRadioButton("Folder");

    // Output group
    private final JRadioButton m_wideRadio             = new JRadioButton("Wide");
    private final JRadioButton m_longRadio             = new JRadioButton("Long");
    private final JCheckBox    m_includeSourceFilename = new JCheckBox("Include source filename");
    private final JCheckBox    m_includeSheetName      = new JCheckBox("Include sheet name");

    // Error handling group
    private final JRadioButton m_missingCellWarn = new JRadioButton("Warn and insert missing value");
    private final JRadioButton m_missingCellFail = new JRadioButton("Fail");
    private final JRadioButton m_badValueWarn    = new JRadioButton("Warn and insert missing value");
    private final JRadioButton m_badValueFail    = new JRadioButton("Fail");

    public ExcelFormReaderNodeDialog() {
        addTab("General", buildGeneralPanel());
        addTab("File",    new JPanel());
        addTab("Folder",  new JPanel());
    }

    private JPanel buildGeneralPanel() {
        // ── Input group ───────────────────────────────────────────────────────
        ButtonGroup inputGroup = new ButtonGroup();
        inputGroup.add(m_singleFileRadio);
        inputGroup.add(m_folderRadio);
        m_singleFileRadio.setSelected(true);

        JPanel inputBox = createGroupBox("Input");
        inputBox.add(m_singleFileRadio);
        inputBox.add(m_folderRadio);

        m_singleFileRadio.addItemListener(e -> updateTabStates());

        // ── Output group ──────────────────────────────────────────────────────
        ButtonGroup outputGroup = new ButtonGroup();
        outputGroup.add(m_wideRadio);
        outputGroup.add(m_longRadio);
        m_wideRadio.setSelected(true);

        JPanel formatRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        formatRow.add(new JLabel("Output format:"));
        formatRow.add(m_wideRadio);
        formatRow.add(m_longRadio);

        JPanel outputBox = createGroupBox("Output");
        outputBox.add(formatRow);
        outputBox.add(m_includeSourceFilename);
        outputBox.add(m_includeSheetName);

        // ── Error handling group ──────────────────────────────────────────────
        ButtonGroup missingGroup = new ButtonGroup();
        missingGroup.add(m_missingCellWarn);
        missingGroup.add(m_missingCellFail);
        m_missingCellWarn.setSelected(true);

        ButtonGroup badValueGroup = new ButtonGroup();
        badValueGroup.add(m_badValueWarn);
        badValueGroup.add(m_badValueFail);
        m_badValueWarn.setSelected(true);

        JPanel missingRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        missingRow.add(new JLabel("On missing cell:"));
        missingRow.add(m_missingCellWarn);
        missingRow.add(m_missingCellFail);

        JPanel badValueRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        badValueRow.add(new JLabel("On unparseable value:"));
        badValueRow.add(m_badValueWarn);
        badValueRow.add(m_badValueFail);

        JPanel errorBox = createGroupBox("Error Handling");
        errorBox.add(missingRow);
        errorBox.add(badValueRow);

        // ── Assemble ──────────────────────────────────────────────────────────
        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.add(inputBox);
        content.add(outputBox);
        content.add(errorBox);

        JPanel root = new JPanel(new BorderLayout());
        root.add(content, BorderLayout.NORTH);
        return root;
    }

    private void updateTabStates() {
        boolean singleFile = m_singleFileRadio.isSelected();
        JTabbedPane pane = getTabbedPane();
        if (pane != null) {
            pane.setEnabledAt(1, singleFile);
            pane.setEnabledAt(2, !singleFile);
        }
    }

    // NodeDialogPane.m_pane is private with no public accessor
    private JTabbedPane getTabbedPane() {
        try {
            Field f = NodeDialogPane.class.getDeclaredField("m_pane");
            f.setAccessible(true);
            return (JTabbedPane) f.get(this);
        } catch (Exception e) {
            return null;
        }
    }

    private JPanel createGroupBox(final String title) {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createTitledBorder(title),
            BorderFactory.createEmptyBorder(5, 5, 5, 5)));
        return panel;
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings)
            throws InvalidSettingsException {
        m_settings.getInputModeModel().setStringValue(
            m_singleFileRadio.isSelected() ? "SINGLE_FILE" : "FOLDER");
        m_settings.getOutputFormatModel().setStringValue(
            m_wideRadio.isSelected() ? "WIDE" : "LONG");
        m_settings.getIncludeSourceFilenameModel().setBooleanValue(
            m_includeSourceFilename.isSelected());
        m_settings.getIncludeSheetNameModel().setBooleanValue(
            m_includeSheetName.isSelected());
        m_settings.getOnMissingCellModel().setStringValue(
            m_missingCellWarn.isSelected() ? "WARN" : "FAIL");
        m_settings.getOnBadValueModel().setStringValue(
            m_badValueWarn.isSelected() ? "WARN" : "FAIL");
        m_settings.saveSettings(settings);
    }

    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings,
            final DataTableSpec[] specs) throws NotConfigurableException {
        try {
            m_settings.loadSettings(settings);
        } catch (InvalidSettingsException e) {
            // use defaults on first open
        }
        boolean singleFile = m_settings.getInputMode() == InputMode.SINGLE_FILE;
        m_singleFileRadio.setSelected(singleFile);
        m_folderRadio.setSelected(!singleFile);

        m_wideRadio.setSelected(m_settings.getOutputFormat() == OutputFormat.WIDE);
        m_longRadio.setSelected(m_settings.getOutputFormat() == OutputFormat.LONG);

        m_includeSourceFilename.setSelected(m_settings.isIncludeSourceFilename());
        m_includeSheetName.setSelected(m_settings.isIncludeSheetName());

        m_missingCellWarn.setSelected(m_settings.getOnMissingCell() == ErrorHandling.WARN);
        m_missingCellFail.setSelected(m_settings.getOnMissingCell() == ErrorHandling.FAIL);

        m_badValueWarn.setSelected(m_settings.getOnBadValue() == ErrorHandling.WARN);
        m_badValueFail.setSelected(m_settings.getOnBadValue() == ErrorHandling.FAIL);

        updateTabStates();
    }
}
