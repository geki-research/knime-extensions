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

    private final JComboBox<String> m_fileReadFrom = new JComboBox<>();
    private final JTextField        m_filePath     = new JTextField(30);
    private final JButton           m_fileBrowse   = new JButton("Browse...");

    // Sheet mode (top-level toggle)
    private final JRadioButton m_fileSingleSheetRadio = new JRadioButton("Process single sheet");
    private final JRadioButton m_fileManySheetRadio   = new JRadioButton("Process many sheets");
    private JPanel             m_fileSingleSheetPanel;
    private JPanel             m_fileManySheetPanel;

    // Section A — single sheet controls
    private final JRadioButton      m_fileFirstRadio      = new JRadioButton("First");
    private final JRadioButton      m_fileByNameRadio     = new JRadioButton("By name");
    private final JRadioButton      m_fileByPositionRadio = new JRadioButton("By position");
    private final JLabel            m_fileFirstSheetName  = new JLabel("(no file selected)");
    private final JComboBox<String> m_fileSheetNameCombo  = new JComboBox<>();
    private final JSpinner          m_fileSheetPosition   =
        new JSpinner(new SpinnerNumberModel(0, 0, 999, 1));
    private JPanel                  m_fileByPositionPanel;
    private final JCheckBox         m_fileSingleHiddenSheets = new JCheckBox("Include hidden worksheets");

    // Section B — many sheets controls
    private final JCheckBox    m_fileManyHiddenSheets     = new JCheckBox("Include hidden worksheets");
    private final JRadioButton m_fileSheetFilterAll       = new JRadioButton("All");
    private final JRadioButton m_fileSheetFilterBlacklist = new JRadioButton("Blacklist");
    private final JRadioButton m_fileSheetFilterWhitelist = new JRadioButton("Whitelist");
    private final JTextField   m_fileSheetFilterNames     = new JTextField(30);
    private JPanel             m_fileSheetFilterNamesPanel;

    // ── Folder tab ────────────────────────────────────────────────────────────

    private final JComboBox<String> m_folderReadFrom       = new JComboBox<>();
    private final JTextField        m_folderPath           = new JTextField(30);
    private final JButton           m_folderBrowse         = new JButton("Browse...");
    private final JCheckBox         m_recursive            = new JCheckBox("Include subfolders");
    private final JCheckBox         m_includeHiddenFolders = new JCheckBox("Include hidden folders");

    private final JCheckBox    m_filterByExtension = new JCheckBox("Filter by file extension");
    private final JTextField   m_fileExtensions    = new JTextField(20);
    private JPanel             m_extensionsPanel;
    private final JCheckBox    m_includeHiddenFiles = new JCheckBox("Include hidden files");

    private final JCheckBox    m_folderIncludeHiddenSheets  = new JCheckBox("Include hidden worksheets");
    private final JRadioButton m_folderSheetFilterAll       = new JRadioButton("All");
    private final JRadioButton m_folderSheetFilterBlacklist = new JRadioButton("Blacklist");
    private final JRadioButton m_folderSheetFilterWhitelist = new JRadioButton("Whitelist");
    private final JTextField   m_folderSheetFilterNames     = new JTextField(30);
    private JPanel             m_folderSheetFilterNamesPanel;

    // ─────────────────────────────────────────────────────────────────────────

    public ExcelFormReaderNodeDialog() {
        addTab("General", buildGeneralPanel());
        addTab("File",    buildFilePanel());
        addTab("Folder",  buildFolderPanel());
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

        // ── Select Sheet group — top-level mode toggle ────────────────────────
        ButtonGroup sheetsModeGroup = new ButtonGroup();
        sheetsModeGroup.add(m_fileSingleSheetRadio);
        sheetsModeGroup.add(m_fileManySheetRadio);
        m_fileSingleSheetRadio.setSelected(true);
        m_fileSingleSheetRadio.addItemListener(e -> updateSheetModeControls());

        // ── Section A: single sheet panel ─────────────────────────────────────
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

        JPanel singleSepPanel = new JPanel(new BorderLayout());
        singleSepPanel.setBorder(BorderFactory.createEmptyBorder(4, 0, 4, 0));
        singleSepPanel.add(new JSeparator(SwingConstants.HORIZONTAL), BorderLayout.CENTER);

        m_fileSingleSheetPanel = new JPanel();
        m_fileSingleSheetPanel.setLayout(new BoxLayout(m_fileSingleSheetPanel, BoxLayout.Y_AXIS));
        m_fileSingleSheetPanel.add(m_fileFirstRadio);
        m_fileSingleSheetPanel.add(firstNamePanel);
        m_fileSingleSheetPanel.add(m_fileByNameRadio);
        m_fileSingleSheetPanel.add(byNamePanel);
        m_fileSingleSheetPanel.add(m_fileByPositionRadio);
        m_fileSingleSheetPanel.add(m_fileByPositionPanel);
        m_fileSingleSheetPanel.add(singleSepPanel);
        m_fileSingleSheetPanel.add(m_fileSingleHiddenSheets);

        // ── Section B: many sheets panel ──────────────────────────────────────
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

        m_fileManySheetPanel = new JPanel();
        m_fileManySheetPanel.setLayout(new BoxLayout(m_fileManySheetPanel, BoxLayout.Y_AXIS));
        m_fileManySheetPanel.setVisible(false);
        m_fileManySheetPanel.add(m_fileManyHiddenSheets);
        m_fileManySheetPanel.add(filterModeRow);
        m_fileManySheetPanel.add(m_fileSheetFilterNamesPanel);

        // ── Assemble Select Sheet group ───────────────────────────────────────
        JPanel sheetBox = createGroupBox("Select Sheet");
        sheetBox.add(m_fileSingleSheetRadio);
        sheetBox.add(m_fileManySheetRadio);
        sheetBox.add(m_fileSingleSheetPanel);
        sheetBox.add(m_fileManySheetPanel);

        updateSheetControls();

        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.add(locationBox);
        content.add(sheetBox);

        JPanel root = new JPanel(new BorderLayout());
        root.add(content, BorderLayout.NORTH);
        return root;
    }

    private void updateSheetModeControls() {
        boolean single = m_fileSingleSheetRadio.isSelected();
        m_fileSingleSheetPanel.setVisible(single);
        m_fileManySheetPanel.setVisible(!single);
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

    // ── Folder tab ────────────────────────────────────────────────────────────

    private JPanel buildFolderPanel() {
        m_folderReadFrom.addItem("Local File System");

        JPanel readFromRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        readFromRow.add(new JLabel("Read from:"));
        readFromRow.add(m_folderReadFrom);

        m_folderBrowse.addActionListener(e -> {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                m_folderPath.setText(chooser.getSelectedFile().getAbsolutePath());
            }
        });

        JPanel folderPathRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        folderPathRow.add(new JLabel("Folder:"));
        folderPathRow.add(m_folderPath);
        folderPathRow.add(m_folderBrowse);

        JPanel locationBox = createGroupBox("Input Location");
        locationBox.add(readFromRow);
        locationBox.add(folderPathRow);
        locationBox.add(m_recursive);
        locationBox.add(m_includeHiddenFolders);

        m_filterByExtension.setSelected(true);
        m_fileExtensions.setText("xlsx");

        m_extensionsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        m_extensionsPanel.add(new JLabel("File extensions (comma-separated):"));
        m_extensionsPanel.add(m_fileExtensions);

        m_filterByExtension.addItemListener(
            e -> m_extensionsPanel.setVisible(m_filterByExtension.isSelected()));

        JPanel fileFilterBox = createGroupBox("File Filter");
        fileFilterBox.add(m_filterByExtension);
        fileFilterBox.add(m_extensionsPanel);
        fileFilterBox.add(m_includeHiddenFiles);

        ButtonGroup folderFilterModeGroup = new ButtonGroup();
        folderFilterModeGroup.add(m_folderSheetFilterAll);
        folderFilterModeGroup.add(m_folderSheetFilterBlacklist);
        folderFilterModeGroup.add(m_folderSheetFilterWhitelist);
        m_folderSheetFilterAll.setSelected(true);

        JPanel folderFilterModeRow = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        folderFilterModeRow.add(new JLabel("Sheet filter:"));
        folderFilterModeRow.add(m_folderSheetFilterAll);
        folderFilterModeRow.add(m_folderSheetFilterBlacklist);
        folderFilterModeRow.add(m_folderSheetFilterWhitelist);

        m_folderSheetFilterNamesPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        m_folderSheetFilterNamesPanel.add(new JLabel("Sheet names (comma-separated):"));
        m_folderSheetFilterNamesPanel.add(m_folderSheetFilterNames);
        m_folderSheetFilterNamesPanel.setVisible(false);

        m_folderSheetFilterAll.addItemListener(e -> updateFolderSheetFilterControls());
        m_folderSheetFilterBlacklist.addItemListener(e -> updateFolderSheetFilterControls());
        m_folderSheetFilterWhitelist.addItemListener(e -> updateFolderSheetFilterControls());

        JPanel sheetBox = createGroupBox("Select Sheet");
        sheetBox.add(m_folderIncludeHiddenSheets);
        sheetBox.add(folderFilterModeRow);
        sheetBox.add(m_folderSheetFilterNamesPanel);

        JPanel content = new JPanel();
        content.setLayout(new BoxLayout(content, BoxLayout.Y_AXIS));
        content.add(locationBox);
        content.add(fileFilterBox);
        content.add(sheetBox);

        JPanel root = new JPanel(new BorderLayout());
        root.add(content, BorderLayout.NORTH);
        return root;
    }

    private void updateFolderSheetFilterControls() {
        boolean showNames = m_folderSheetFilterBlacklist.isSelected()
            || m_folderSheetFilterWhitelist.isSelected();
        m_folderSheetFilterNamesPanel.setVisible(showNames);
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
        m_settings.getFileManySheeetsModel().setBooleanValue(
            m_fileManySheetRadio.isSelected());
        m_settings.getFileSingleHiddenSheetsModel().setBooleanValue(
            m_fileSingleHiddenSheets.isSelected());
        m_settings.getFileHiddenSheetsModel().setBooleanValue(
            m_fileManyHiddenSheets.isSelected());

        String filterMode = m_fileSheetFilterAll.isSelected() ? "ALL"
            : m_fileSheetFilterBlacklist.isSelected() ? "BLACKLIST"
            : "WHITELIST";
        m_settings.getFileSheetFilterModeModel().setStringValue(filterMode);
        m_settings.getFileSheetFilterNamesModel().setStringValue(
            m_fileSheetFilterNames.getText().trim());

        // Folder tab
        m_settings.getFolderPathModel().setStringValue(m_folderPath.getText().trim());
        m_settings.getRecursiveModel().setBooleanValue(m_recursive.isSelected());
        m_settings.getIncludeHiddenFoldersModel().setBooleanValue(m_includeHiddenFolders.isSelected());
        m_settings.getFilterByExtensionModel().setBooleanValue(m_filterByExtension.isSelected());
        m_settings.getFileExtensionsModel().setStringValue(m_fileExtensions.getText().trim());
        m_settings.getIncludeHiddenFilesModel().setBooleanValue(m_includeHiddenFiles.isSelected());
        m_settings.getFolderHiddenSheetsModel().setBooleanValue(m_folderIncludeHiddenSheets.isSelected());

        String folderFilterMode = m_folderSheetFilterAll.isSelected() ? "ALL"
            : m_folderSheetFilterBlacklist.isSelected() ? "BLACKLIST"
            : "WHITELIST";
        m_settings.getFolderSheetFilterModeModel().setStringValue(folderFilterMode);
        m_settings.getFolderSheetFilterNamesModel().setStringValue(
            m_folderSheetFilterNames.getText().trim());

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

        boolean manySheets = m_settings.getFileManySheets();
        m_fileSingleSheetRadio.setSelected(!manySheets);
        m_fileManySheetRadio.setSelected(manySheets);
        m_fileSingleHiddenSheets.setSelected(m_settings.isFileSingleIncludeHiddenSheets());
        m_fileManyHiddenSheets.setSelected(m_settings.isFileIncludeHiddenSheets());

        SheetFilterMode fm = m_settings.getFileSheetFilterMode();
        m_fileSheetFilterAll.setSelected(fm == SheetFilterMode.ALL);
        m_fileSheetFilterBlacklist.setSelected(fm == SheetFilterMode.BLACKLIST);
        m_fileSheetFilterWhitelist.setSelected(fm == SheetFilterMode.WHITELIST);

        m_fileSheetFilterNames.setText(String.join(", ", m_settings.getFileSheetFilterNames()));

        updateSheetControls();
        updateSheetFilterControls();
        updateSheetModeControls();

        // Folder tab
        m_folderPath.setText(m_settings.getFolderPath());
        m_recursive.setSelected(m_settings.isRecursive());
        m_includeHiddenFolders.setSelected(m_settings.isIncludeHiddenFolders());
        m_filterByExtension.setSelected(m_settings.isFilterByExtension());
        m_fileExtensions.setText(String.join(", ", m_settings.getFileExtensions()));
        m_extensionsPanel.setVisible(m_settings.isFilterByExtension());
        m_includeHiddenFiles.setSelected(m_settings.isIncludeHiddenFiles());
        m_folderIncludeHiddenSheets.setSelected(m_settings.isFolderIncludeHiddenSheets());

        SheetFilterMode ffm = m_settings.getFolderSheetFilterMode();
        m_folderSheetFilterAll.setSelected(ffm == SheetFilterMode.ALL);
        m_folderSheetFilterBlacklist.setSelected(ffm == SheetFilterMode.BLACKLIST);
        m_folderSheetFilterWhitelist.setSelected(ffm == SheetFilterMode.WHITELIST);

        m_folderSheetFilterNames.setText(String.join(", ", m_settings.getFolderSheetFilterNames()));

        updateFolderSheetFilterControls();
    }
}
