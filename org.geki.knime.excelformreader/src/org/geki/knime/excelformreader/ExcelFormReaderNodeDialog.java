package org.geki.knime.excelformreader;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.ErrorHandling;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.InputMode;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.OutputFormat;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.SheetFilterMode;
import org.geki.knime.excelformreader.ExcelFormReaderSettings.SheetSelection;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;

import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSeparator;
import javax.swing.JSpinner;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.io.File;
import java.lang.reflect.Field;

public class ExcelFormReaderNodeDialog extends NodeDialogPane {

    private static final NodeLogger LOGGER =
        NodeLogger.getLogger(ExcelFormReaderNodeDialog.class);

    private final ExcelFormReaderSettings m_settings = new ExcelFormReaderSettings();

    // ── General tab ───────────────────────────────────────────────────────────

    private final JRadioButton m_singleFileRadio = new JRadioButton("Single File");
    private final JRadioButton m_folderRadio     = new JRadioButton("Folder");

    private final JRadioButton m_wideRadio             = new JRadioButton("Wide");
    private final JRadioButton m_longRadio             = new JRadioButton("Long");
    private final JCheckBox    m_includeSourceFilename = new JCheckBox("Include source filename");
    private final JCheckBox    m_includeSheetName      = new JCheckBox("Include sheet name");

    private final JRadioButton m_missingCellWarn = new JRadioButton("Warn and insert missing value");
    private final JRadioButton m_missingCellFail = new JRadioButton("Fail");
    private final JRadioButton m_badValueWarn    = new JRadioButton("Warn and insert missing value");
    private final JRadioButton m_badValueFail    = new JRadioButton("Fail");

    // ── File tab ──────────────────────────────────────────────────────────────

    private final JComboBox<String> m_fileReadFrom   = new JComboBox<>();
    private final JTextField        m_filePath       = new JTextField(30);
    private final JButton           m_fileBrowse     = new JButton("Browse...");

    private final JRadioButton      m_fileFirstRadio       = new JRadioButton("First");
    private final JRadioButton      m_fileByNameRadio      = new JRadioButton("By name");
    private final JRadioButton      m_fileByPositionRadio  = new JRadioButton("By position");
    private final JLabel            m_fileFirstSheetName   = new JLabel("(no file selected)");
    private final JComboBox<String> m_fileSheetNameCombo   = new JComboBox<>();
    private final JSpinner          m_fileSheetPosition    =
        new JSpinner(new SpinnerNumberModel(0, 0, 999, 1));
    private JPanel                  m_fileByPositionPanel;

    private final JCheckBox         m_fileIncludeHiddenSheets  = new JCheckBox("Include hidden worksheets");
    private final JRadioButton      m_fileSheetFilterAll       = new JRadioButton("All");
    private final JRadioButton      m_fileSheetFilterBlacklist = new JRadioButton("Blacklist");
    private final JRadioButton      m_fileSheetFilterWhitelist = new JRadioButton("Whitelist");
    private final JTextField        m_fileSheetFilterNames     = new JTextField(30);
    private JPanel                  m_fileSheetFilterNamesPanel;

    // ─────────────────────────────────────────────────────────────────────────

    public ExcelFormReaderNodeDialog() {
        addTab("General", buildGeneralPanel());
        addTab("File",    buildFilePanel());
        addTab("Folder",  new JPanel());
    }

    // ── General tab ───────────────────────────────────────────────────────────

    private JPanel buildGeneralPanel() {
        ButtonGroup inputGroup = new ButtonGroup();
        inputGroup.add(m_singleFileRadio);
        inputGroup.add(m_folderRadio);
        m_singleFileRadio.setSelected(true);

        JPanel inputBox = createGroupBox("Input");
        inputBox.add(m_singleFileRadio);
        inputBox.add(m_folderRadio);

        m_singleFileRadio.addItemListener(e -> updateTabStates());

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

        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.add(inputBox);
        content.add(outputBox);
        content.add(errorBox);

        JPanel root = new JPanel(new BorderLayout());
        root.add(content, BorderLayout.NORTH);
        return root;
    }

    // ── File tab ──────────────────────────────────────────────────────────────

    private JPanel buildFilePanel() {
        // ── Input Location group ──────────────────────────────────────────────
        m_fileReadFrom.addItem("Local File System");

        JPanel readFromRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        readFromRow.add(new JLabel("Read from:"));
        readFromRow.add(m_fileReadFrom);

        m_fileBrowse.addActionListener(e -> {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx)", "xlsx"));
            chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                m_filePath.setText(chooser.getSelectedFile().getAbsolutePath());
                refreshSheetNames();
            }
        });

        JPanel filePathRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        filePathRow.add(new JLabel("File:"));
        filePathRow.add(m_filePath);
        filePathRow.add(m_fileBrowse);

        JPanel locationBox = createGroupBox("Input Location");
        locationBox.add(readFromRow);
        locationBox.add(filePathRow);

        // ── Select Sheet group — Part A: sheet selection ──────────────────────
        ButtonGroup sheetSelGroup = new ButtonGroup();
        sheetSelGroup.add(m_fileFirstRadio);
        sheetSelGroup.add(m_fileByNameRadio);
        sheetSelGroup.add(m_fileByPositionRadio);
        m_fileFirstRadio.setSelected(true);

        JPanel firstNamePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        firstNamePanel.setBorder(BorderFactory.createEmptyBorder(0, 20, 0, 0));
        firstNamePanel.add(m_fileFirstSheetName);

        m_fileSheetNameCombo.addItem("(select a file first)");
        m_fileSheetNameCombo.setEnabled(false);
        JPanel byNamePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        byNamePanel.setBorder(BorderFactory.createEmptyBorder(0, 20, 0, 0));
        byNamePanel.add(m_fileSheetNameCombo);

        m_fileByPositionPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        m_fileByPositionPanel.setBorder(BorderFactory.createEmptyBorder(0, 20, 0, 0));
        m_fileByPositionPanel.add(m_fileSheetPosition);
        m_fileByPositionPanel.add(new JLabel("Position starts with 0."));

        m_fileFirstRadio.addItemListener(e -> updateSheetControls());
        m_fileByNameRadio.addItemListener(e -> updateSheetControls());
        m_fileByPositionRadio.addItemListener(e -> updateSheetControls());

        // ── Select Sheet group — Part B: sheet filter ─────────────────────────
        JPanel separatorPanel = new JPanel(new BorderLayout());
        separatorPanel.setBorder(BorderFactory.createEmptyBorder(4, 0, 4, 0));
        separatorPanel.add(new JSeparator(SwingConstants.HORIZONTAL), BorderLayout.CENTER);

        JPanel hiddenSheetsRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        hiddenSheetsRow.add(m_fileIncludeHiddenSheets);

        ButtonGroup filterModeGroup = new ButtonGroup();
        filterModeGroup.add(m_fileSheetFilterAll);
        filterModeGroup.add(m_fileSheetFilterBlacklist);
        filterModeGroup.add(m_fileSheetFilterWhitelist);
        m_fileSheetFilterAll.setSelected(true);

        JPanel filterModeRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        filterModeRow.add(new JLabel("Sheet filter:"));
        filterModeRow.add(m_fileSheetFilterAll);
        filterModeRow.add(m_fileSheetFilterBlacklist);
        filterModeRow.add(m_fileSheetFilterWhitelist);

        m_fileSheetFilterNamesPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        m_fileSheetFilterNamesPanel.add(new JLabel("Sheet names (comma-separated):"));
        m_fileSheetFilterNamesPanel.add(m_fileSheetFilterNames);
        m_fileSheetFilterNamesPanel.setVisible(false);

        m_fileSheetFilterAll.addItemListener(e -> updateSheetFilterControls());
        m_fileSheetFilterBlacklist.addItemListener(e -> updateSheetFilterControls());
        m_fileSheetFilterWhitelist.addItemListener(e -> updateSheetFilterControls());

        // ── Assemble Select Sheet group ───────────────────────────────────────
        JPanel sheetBox = createGroupBox("Select Sheet");
        sheetBox.add(m_fileFirstRadio);
        sheetBox.add(firstNamePanel);
        sheetBox.add(m_fileByNameRadio);
        sheetBox.add(byNamePanel);
        sheetBox.add(m_fileByPositionRadio);
        sheetBox.add(m_fileByPositionPanel);
        sheetBox.add(separatorPanel);
        sheetBox.add(hiddenSheetsRow);
        sheetBox.add(filterModeRow);
        sheetBox.add(m_fileSheetFilterNamesPanel);

        updateSheetControls();

        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.add(locationBox);
        content.add(sheetBox);

        JPanel root = new JPanel(new BorderLayout());
        root.add(content, BorderLayout.NORTH);
        return root;
    }

    private void updateSheetControls() {
        m_fileFirstSheetName.setVisible(m_fileFirstRadio.isSelected());
        m_fileSheetNameCombo.setVisible(m_fileByNameRadio.isSelected());
        m_fileByPositionPanel.setVisible(m_fileByPositionRadio.isSelected());
    }

    private void updateSheetFilterControls() {
        boolean showNames = m_fileSheetFilterBlacklist.isSelected()
            || m_fileSheetFilterWhitelist.isSelected();
        m_fileSheetFilterNamesPanel.setVisible(showNames);
    }

    private void refreshSheetNames() {
        String path = m_filePath.getText().trim();
        m_fileSheetNameCombo.removeAllItems();
        m_fileFirstSheetName.setText("(no file selected)");

        if (path.isEmpty()) return;

        File file = new File(path);
        if (!file.exists() || !file.isFile()) return;

        try (Workbook wb = WorkbookFactory.create(file, null, true)) {
            String firstName = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String name = wb.getSheetName(i);
                m_fileSheetNameCombo.addItem(name);
                if (firstName == null) firstName = name;
            }
            if (firstName != null) {
                m_fileFirstSheetName.setText(firstName);
                m_fileSheetNameCombo.setEnabled(true);
            }
        } catch (Exception e) {
            m_fileFirstSheetName.setText("(error reading file)");
            LOGGER.warn("Could not read sheet names from: " + path, e);
        }
    }

    // ── Tabbed pane helpers ───────────────────────────────────────────────────

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

    // ── Persistence ───────────────────────────────────────────────────────────

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings)
            throws InvalidSettingsException {
        // General tab
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

        // File tab
        m_settings.getFilePathModel().setStringValue(m_filePath.getText().trim());

        String sheetSel = m_fileFirstRadio.isSelected() ? "FIRST"
            : m_fileByNameRadio.isSelected() ? "BY_NAME"
            : "BY_POSITION";
        m_settings.getFileSheetSelectionModel().setStringValue(sheetSel);

        m_settings.getFileSheetNameModel().setStringValue(
            m_fileSheetNameCombo.getSelectedItem() != null
            ? (String) m_fileSheetNameCombo.getSelectedItem() : "");
        m_settings.getFileSheetPositionModel().setIntValue(
            (Integer) m_fileSheetPosition.getValue());
        m_settings.getFileHiddenSheetsModel().setBooleanValue(
            m_fileIncludeHiddenSheets.isSelected());

        String filterMode = m_fileSheetFilterAll.isSelected() ? "ALL"
            : m_fileSheetFilterBlacklist.isSelected() ? "BLACKLIST"
            : "WHITELIST";
        m_settings.getFileSheetFilterModeModel().setStringValue(filterMode);
        m_settings.getFileSheetFilterNamesModel().setStringValue(
            m_fileSheetFilterNames.getText().trim());

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

        // General tab
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

        // File tab
        m_filePath.setText(m_settings.getFilePath());
        refreshSheetNames();

        SheetSelection sel = m_settings.getFileSheetSelection();
        m_fileFirstRadio.setSelected(sel == SheetSelection.FIRST);
        m_fileByNameRadio.setSelected(sel == SheetSelection.BY_NAME);
        m_fileByPositionRadio.setSelected(sel == SheetSelection.BY_POSITION);

        String savedName = m_settings.getFileSheetName();
        if (!savedName.isEmpty()) {
            m_fileSheetNameCombo.setSelectedItem(savedName);
        }

        m_fileSheetPosition.setValue(m_settings.getFileSheetPosition());
        m_fileIncludeHiddenSheets.setSelected(m_settings.isFileIncludeHiddenSheets());

        SheetFilterMode fm = m_settings.getFileSheetFilterMode();
        m_fileSheetFilterAll.setSelected(fm == SheetFilterMode.ALL);
        m_fileSheetFilterBlacklist.setSelected(fm == SheetFilterMode.BLACKLIST);
        m_fileSheetFilterWhitelist.setSelected(fm == SheetFilterMode.WHITELIST);

        m_fileSheetFilterNames.setText(String.join(", ", m_settings.getFileSheetFilterNames()));

        updateSheetControls();
        updateSheetFilterControls();
    }
}
