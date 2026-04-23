package org.geki.knime.excelformreader;

import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeFactory;
import org.knime.core.node.NodeView;

public class ExcelFormReaderNodeFactory extends NodeFactory<ExcelFormReaderNodeModel> {

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
        throw new IndexOutOfBoundsException("No views: " + viewIndex);
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
