package bigreport.performers;

import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import org.apache.poi.ss.usermodel.Cell;

import static bigreport.util.ValueResolver.*;

public class VelocityTemplateConstuctor {

    private final IterationContext iterationContext;

    public VelocityTemplateConstuctor(IterationContext iterationContext) {
        this.iterationContext = iterationContext;
    }

    private void createCell() {
        String value = resolve(iterationContext.getCurrentCell());
        boolean isCellCondition = isCellConditionDirective(value);
        if (checkIsMergedCell()){
            createMergedCell(value, isCellCondition);
            return;
        }
        if (checkIsMergedChildCell()) {
            createMergedChildCell();
            return;
        }
        createCell(isCellCondition, value, value);
    }

    private void createMergedCell(String value, boolean cellCondition) {
        createMergedCell(iterationContext.getCurrentCell(), value, cellCondition);
    }

    private void createMergedChildCell() {
        createMergedChildCell(iterationContext.getOriginalMergedHeadRegion());
    }

    private boolean checkIsMergedCell(){
       return iterationContext.getCellIterator().isMergedCell();
    }

    private boolean checkIsMergedChildCell() {
        return iterationContext.getOriginalMergedHeadRegion() != null;
    }

    private void createMergedChildCell(MockCell mergedRegionHeader) {
        String mergedRegionHeaderValue = iterationContext.getValueAt(mergedRegionHeader);
        createCell(isCellConditionDirective(mergedRegionHeaderValue), mergedRegionHeaderValue, "");
    }

    private void startRowIfNecessary(Cell cell) {
        if (iterationContext.getCellIterator().isNewRow()) {
            addRow(cell, iterationContext.getTemplateBuilder(), iterationContext.getIteratedOutlineLevel());
        }
    }

    private void addRow(Cell cell, VelocityTemplateBuilder templateBuilder, int outlineLevel) {
        templateBuilder.addRow(getRowStyleForCell(cell), outlineLevel);
    }

    public void appendCell(Cell cell) {
        startRowIfNecessary(cell);
        createCell();
        finishRowIfNecessary();
    }

    private void finishRowIfNecessary() {
        if (iterationContext.getCellIterator().isEndOfRow()) {
            iterationContext.getTemplateBuilder().addEndRow();
        }
    }

    private void createCell(boolean isCondition, String value, String displayText) {
        Cell cell = iterationContext.getCurrentCell();
        if (isCondition) {
            iterationContext.getTemplateBuilder().addCell(getPureValueShowIf(displayText),
                    cell.getColumnIndex(),
                    cell.getCellStyle().getIndex(),
                    getCondition(value));
            return;
        }
        iterationContext.getTemplateBuilder().addCell(displayText,
                cell.getColumnIndex(),
                cell.getCellStyle().getIndex()
        );
    }

    private void performDirective(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder, String value) {
        templateBuilder.addDirective(convertToDirective(value));
        cellIterator.skipRow();
    }

    private void createMergedCell(Cell cell, String value, boolean cellConditionDirective) {
        if (cellConditionDirective) {
            iterationContext.getTemplateBuilder().addCell(getPureValueShowIf(value),
                    cell.getColumnIndex(),
                    cell.getCellStyle().getIndex(),
                    iterationContext.getCellIterator().getOffset(cell),
                    getCondition(value)
            );
            return;
        }
        iterationContext.getTemplateBuilder().addCell(value,
                cell.getColumnIndex(),
                cell.getCellStyle().getIndex(),
                iterationContext.getCellIterator().getOffset(cell));
    }
}
