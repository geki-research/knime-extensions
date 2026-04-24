package org.geki.knime.excelformreader;

import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentFileChooser;
import org.knime.core.node.defaultnodesettings.DialogComponentString;

public class ExcelFormReaderNodeDialog extends DefaultNodeSettingsPane {

    private final ExcelFormReaderSettings m_settings = new ExcelFormReaderSettings();

    public ExcelFormReaderNodeDialog() {

        // ── Panel 1: File / Folder ─────────────────────────────────────────
        createNewGroup("File / Folder");

        addDialogComponent(new DialogComponentButtonGroup(
            m_settings.getInputModeModel(),  // TODO: full dialog redesign pending
            "Reading mode",
            false,
            new String[]{"SINGLE_FILE", "FOLDER"},
            new String[]{"Single File", "Folder"}
        ));

        addDialogComponent(new DialogComponentFileChooser(
            m_settings.getFilePathModel(),   // TODO: full dialog redesign pending
            "excel_form_reader_path",
            ".xlsx"
        ));

        addDialogComponent(new DialogComponentBoolean(
            m_settings.getRecursiveModel(),
            "Include subfolders"
        ));

        closeCurrentGroup();

        // ── Panel 2: Sheet ─────────────────────────────────────────────────
        createNewGroup("Sheet");

        addDialogComponent(new DialogComponentString(
            m_settings.getFileSheetNameModel(), // TODO: full dialog redesign pending
            "Default sheet name (leave blank for first sheet)",
            true, 30
        ));

        // TODO: sheet filter UI pending full dialog redesign

        closeCurrentGroup();

        // ── Panel 3: Output ────────────────────────────────────────────────
        createNewGroup("Output");

        addDialogComponent(new DialogComponentButtonGroup(
            m_settings.getOutputFormatModel(),
            "Output format",
            false,
            new String[]{"WIDE", "LONG"},
            new String[]{"Wide (one row per form)", "Long (one row per field)"}
        ));

        addDialogComponent(new DialogComponentBoolean(
            m_settings.getIncludeSourceFilenameModel(), // TODO: full dialog redesign pending
            "Add source file and sheet name columns"
        ));

        // TODO: range delimiter UI pending full dialog redesign

        closeCurrentGroup();

        // ── Panel 4: Error Handling ────────────────────────────────────────
        createNewGroup("Error Handling");

        addDialogComponent(new DialogComponentButtonGroup(
            m_settings.getOnMissingCellModel(),
            "On missing cell",
            false,
            new String[]{"WARN", "FAIL"},
            new String[]{"Warn and set missing value", "Fail on missing cell"}
        ));

        addDialogComponent(new DialogComponentButtonGroup(
            m_settings.getOnBadValueModel(),
            "On unparseable value",
            false,
            new String[]{"WARN", "FAIL"},
            new String[]{"Warn and set missing value", "Fail on unparseable value"}
        ));

        closeCurrentGroup();
    }
}
