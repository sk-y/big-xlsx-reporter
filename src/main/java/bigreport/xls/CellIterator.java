package bigreport.xls;

import bigreport.performers.IterationContext;
import bigreport.util.StreamUtil;
import bigreport.util.ValueResolver;
import bigreport.xls.merge.MergeCells;
import bigreport.xls.merge.MergeOffset;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

public class CellIterator implements Iterator<Cell> {
    private Iterator<Row> rowIterator;
    private Iterator<Cell> cellIterator;
    private Sheet sheet;
    private boolean isNewRow = true;
    private Cell currentCell = null;
    private Row currentRow = null;
    private int mergedCount;
    private int outlineLevel;
    private IterationContext iterationContext;

    public CellIterator(XSSFSheet sheet) {
        this.sheet = sheet;
        this.rowIterator = sheet.rowIterator();
        this.mergedCount=0;
        this.outlineLevel=0;
    }

    public boolean hasNext() {
        if (isFinished()) {
            return false;
        }
        if (rowIterator.hasNext()) {
            return true;
        }
        return (cellIterator != null && cellIterator.hasNext());
    }

    public Cell next() {
        if (cellIterator == null || !cellIterator.hasNext()) {
            currentRow = rowIterator.next();
            cellIterator = currentRow.cellIterator();
            isNewRow = true;
        } else {
            isNewRow = false;
        }
        currentCell = cellIterator.next();
        return currentCell;
    }

    public void skipRow(){
        cellIterator=null;
    }

    public void remove() {
        throw new UnsupportedOperationException("Cannot remove items from original file");
    }

    public boolean isNewRow() {
        return isNewRow;
    }

    public boolean isEndOfRow() {
        return cellIterator != null && !cellIterator.hasNext();
    }

    public Cell getCurrentCell() {
        return currentCell;
    }

    public boolean isFinished() {
        return iterationContext.getFinishedAt() != null;
    }

    public boolean isMergedCell() {
        return  iterationContext.getOriginalMergedCells().containsKey(currentCell);
    }

    public MockCell getMergedRegionHeader(MockCell cell){
        return iterationContext.getOriginalMergedCells().getMergedRegionHeader(cell);
    }

    public void addMergedCell(Cell cell) throws IOException {
        mergedCount++;
        MockCell mockCell = new MockCell(cell);
        MergeOffset offset = iterationContext.getOriginalMergedCells().get(mockCell);
        iterationContext.addMergedCell(mockCell, offset);
    }

    public String getValueAt(int row, int col){
        return ValueResolver.resolve(sheet.getRow(row).getCell(col));
    }

    public MergeOffset getOffset(Cell cell) {
        MockCell cellReference = new MockCell(cell);
        return iterationContext.getOriginalMergedCells().get(cellReference);
    }

    public int getMergedCount() {
        return mergedCount;
    }

    public void increaseOutlineLevel() {
        outlineLevel++;
    }

    public void decreaseOutlineLevel() {
        outlineLevel--;
    }

    public int getOutlineLevel() {
        return outlineLevel;
    }

    public void setContext(IterationContext iterationContext) {
        this.iterationContext=iterationContext;
    }
}
