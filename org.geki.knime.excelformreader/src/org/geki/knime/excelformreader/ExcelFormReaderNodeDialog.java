package org.geki.knime.excelformreader;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.data.DataTableSpec;

import javax.swing.JPanel;

public class ExcelFormReaderNodeDialog extends NodeDialogPane {

    private final ExcelFormReaderSettings m_settings = new ExcelFormReaderSettings();

    public ExcelFormReaderNodeDialog() {
        addTab("General", new JPanel());
        addTab("File",    new JPanel());
        addTab("Folder",  new JPanel());
    }

    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings)
            throws InvalidSettingsException {
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
    }
}
