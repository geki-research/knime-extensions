package org.geki.knime.excelformreader;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeFactory;
import org.knime.core.node.NodeView;

public class ExcelFormReaderNodeFactory extends NodeFactory<ExcelFormReaderNodeModel> {

    private static final String NODE_DESCRIPTION_XML =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        + "<knimeNode icon=\"./icons/excelformreader.png\" type=\"Source\"\n"
        + "    xmlns=\"http://knime.org/node/v2.8\"\n"
        + "    xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n"
        + "    xsi:schemaLocation=\"http://knime.org/node/v2.8 http://knime.org/node/v2.8.xsd\">\n"
        + "  <name>Excel Form Reader</name>\n"
        + "  <shortDescription>\n"
        + "    Reads non-tabular form-structured Excel worksheets into a KNIME data table.\n"
        + "  </shortDescription>\n"
        + "  <fullDescription>\n"
        + "    <intro>\n"
        + "      Reads one or more .xlsx files containing form-structured (non-tabular)\n"
        + "      worksheets and extracts field values into a standard KNIME data table.\n"
        + "      The form structure is defined by a connected input table that maps field\n"
        + "      names to Excel cell addresses (e.g. C4, B10:D15). Each worksheet produces\n"
        + "      one output row in wide mode, or one row per field in long mode. Supports\n"
        + "      single file, all sheets, folder, and recursive folder reading modes.\n"
        + "    </intro>\n"
        + "  </fullDescription>\n"
        + "  <ports>\n"
        + "    <inPort index=\"0\" name=\"Form Definition\">\n"
        + "      A table defining the form structure. Required columns: field_name (String),\n"
        + "      value_cell (String). Optional columns: data_type (String), sheet_name (String).\n"
        + "    </inPort>\n"
        + "    <outPort index=\"0\" name=\"Extracted Data\">\n"
        + "      The extracted form data as a flat table. In wide mode: one row per\n"
        + "      (file, sheet). In long mode: one row per (file, sheet, field).\n"
        + "    </outPort>\n"
        + "  </ports>\n"
        + "</knimeNode>\n";

    @Override
    protected InputStream getPropertiesInputStream() {
        return new ByteArrayInputStream(
            NODE_DESCRIPTION_XML.getBytes(StandardCharsets.UTF_8));
    }

    @Override
    public ExcelFormReaderNodeModel createNodeModel() {
        return new ExcelFormReaderNodeModel();
    }

    @Override
    protected int getNrNodeViews() {
        return 0;
    }

    @Override
    public NodeView<ExcelFormReaderNodeModel> createNodeView(
            final int viewIndex, final ExcelFormReaderNodeModel nodeModel) {
        throw new IllegalStateException("No views");
    }

    @Override
    protected boolean hasDialog() {
        return true;
    }

    @Override
    protected NodeDialogPane createNodeDialogPane() {
        return new ExcelFormReaderNodeDialog();
    }
}
