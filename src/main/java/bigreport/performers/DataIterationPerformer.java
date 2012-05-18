package bigreport.performers;

import bigreport.*;
import bigreport.xls.MockCell;
import bigreport.util.ValueResolver;
import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.CellIterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Performs cell iteration on data cells (after "start" directive)
 */
public class DataIterationPerformer implements IterationPerformer {

    private final Map<String, IterationPerformer> graph = new HashMap<String, IterationPerformer>();

    public DataIterationPerformer() {
        EndIterationPerformer endIterationPerformer = new EndIterationPerformer();
        graph.put(Markers.END, endIterationPerformer);
        graph.put(Markers.END_SYNONYM, endIterationPerformer);
        graph.put(Markers.OUT_LINE_START, new OutLineStart());
        graph.put(Markers.OUT_LINE_END, new OutLineEnd());
    }

    public void iterate(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder) throws IOException {
        templateBuilder.start();
        Cell currentCell = cellIterator.getCurrentCell();
        cellIterator.setStartedAt(new MockCell(currentCell.getRowIndex(), currentCell.getColumnIndex()));
        cellIterator.skipRow();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell != null) {
                performCellValue(cellIterator, templateBuilder, cell);
            }
        }
    }

    private void performCellValue(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder, Cell cell) throws IOException {
        String value = ValueResolver.resolve(cell);
        boolean isCellConditionDirective = isCellConditionDirective(value);
        if (!isCellConditionDirective) {
            if (isDirective(value)) {
                templateBuilder.addDirective(convertToDirective(value));
                cellIterator.skipRow();
                return;
            }
            if (graph.containsKey(value)) {
                graph.get(value).iterate(cellIterator, templateBuilder);
                return;
            }
        }

        if (cellIterator.isNewRow()) {
            addRow(cell, templateBuilder, cellIterator.getOutlineLevel());
        }
        if (cellIterator.isMergedCell()) {
            createMergedCell(cellIterator, templateBuilder, cell, value, isCellConditionDirective);
        } else {
            MockCell mergedRegionHeader = cellIterator.getMergedRegionHeader(new MockCell(cell.getRowIndex(),
                                                                            cell.getColumnIndex()));
            if (mergedRegionHeader == null) {
                createCell(templateBuilder, cell, isCellConditionDirective, value, value);
            } else {
                String mergedRegionHeaderValue = cellIterator.getValueAt(mergedRegionHeader.getRow(),
                                                                          mergedRegionHeader.getCol());
                createCell(templateBuilder, cell, isCellConditionDirective(mergedRegionHeaderValue),
                           mergedRegionHeaderValue, "");
            }
        }

        if (cellIterator.isEndOfRow()) {
            templateBuilder.addEndRow();
        }
    }

    private void createMergedCell(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder, Cell cell,
                                  String value, boolean cellConditionDirective) {
        if (cellConditionDirective) {
            templateBuilder.addCell(getPureValueShowIf(value),
                    cell.getColumnIndex(),
                    ((XSSFCell) cell).getCellStyle().getIndex(),
                    cellIterator.getOffset(cell),
                    getCondition(value)
            );
        } else {
            templateBuilder.addCell(value,
                    cell.getColumnIndex(),
                    ((XSSFCell) cell).getCellStyle().getIndex(),
                    cellIterator.getOffset(cell));
        }
    }

    private void createCell(VelocityTemplateBuilder templateBuilder, Cell cell,
                            boolean isCondition, String value, String displayText) {
        if (isCondition) {
            templateBuilder.addCell(getPureValueShowIf(displayText),
                    cell.getColumnIndex(),
                    ((XSSFCell) cell).getCellStyle().getIndex(),
                    getCondition(value));
        } else {
            templateBuilder.addCell(displayText,
                    cell.getColumnIndex(),
                    ((XSSFCell) cell).getCellStyle().getIndex()
            );
        }
    }

    public String getCondition(String value) {
        String cellData = value.replace(Markers.START_CELL_CONDITION_DIRECTIVE, "").replace(Markers.END_CELL_CONDITION_DIRECTIVE, "");
        return cellData.substring(0, cellData.indexOf(Markers.END_CELL_CONDITION));
    }

    public String getPureValueShowIf(String value) {
        if (value.isEmpty()){
            return value;
        }
        int startIndex = value.indexOf(Markers.END_CELL_CONDITION) + Markers.END_CELL_CONDITION.length();
        int endIndex = value.indexOf(Markers.END_CELL_CONDITION_DIRECTIVE);
        return value.substring(startIndex, endIndex);
    }

    private String convertToDirective(String value) {
        return value.substring(0, value.lastIndexOf(Markers.END_DIRECTIVE)).substring(Markers.START_DIRECTIVE.length());
    }

    private boolean isCellConditionDirective(String value) {
        return value.startsWith(Markers.START_CELL_CONDITION_DIRECTIVE)
                && value.contains(Markers.END_CELL_CONDITION)
                && value.endsWith(Markers.END_CELL_CONDITION_DIRECTIVE);
    }

    private boolean isDirective(String value) {
        return value.startsWith(Markers.START_DIRECTIVE) && value.endsWith(Markers.END_DIRECTIVE);
    }

    private void addRow(Cell cell, VelocityTemplateBuilder templateBuilder, int outlineLevel) {
        long styleIndex = ((XSSFRow) cell.getRow()).getCTRow().getS();
        templateBuilder.addRow(styleIndex, outlineLevel);
    }
}
