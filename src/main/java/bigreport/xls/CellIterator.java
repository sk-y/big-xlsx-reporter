package bigreport.xls;

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
    private MockCell startedAt;
    private MockCell finishedAt;
    private Iterator<Row> rowIterator;
    private Iterator<Cell> cellIterator;
    private Sheet sheet;
    private boolean isNewRow = true;
    private Cell currentCell = null;
    private Row currentRow = null;
    private final MergeCells originalCells;
    private OutputStream mergedCellsOutputStream;
    private int mergedCount;
    private int outlineLevel;

    public CellIterator(XSSFSheet sheet, OutputStream mergedCellsOutputStream) {
        originalCells= new MergeCells();
        originalCells.detect(sheet);
        this.mergedCellsOutputStream = mergedCellsOutputStream;
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
        rowIterator.remove();
        sheet.removeRow(currentRow);
        cellIterator = null;
    }

    public boolean isNewRow() {
        return isNewRow;
    }

    public boolean isEndOfRow() {
        return cellIterator != null && !cellIterator.hasNext();
    }

    public void setFinishedAt(MockCell finishedAt) {
        this.finishedAt = finishedAt;
    }

    public MockCell getStartedAt() {
        return startedAt;
    }

    public void setStartedAt(MockCell startedAt) {
        this.startedAt = startedAt;
    }

    public Cell getCurrentCell() {
        return currentCell;
    }

    public boolean isFinished() {
        return finishedAt != null;
    }

    public boolean isMergedCell() {
        return originalCells.containsKey(new MockCell(currentCell.getRowIndex(), currentCell.getColumnIndex()));
    }

    public MockCell getMergedRegionHeader(MockCell cell){
        return originalCells.getMergedRegionHeader(cell);
    }

    public void addMergedCell(Cell cell) throws IOException {
        mergedCount++;
        MockCell mockCell = new MockCell(cell.getRowIndex(), cell.getColumnIndex());
        MergeOffset offset = originalCells.get(mockCell);
        StreamUtil.writeMergedCell(mergedCellsOutputStream, mockCell, offset);
    }

    public String getValueAt(int row, int col){
        return ValueResolver.resolve(sheet.getRow(row).getCell(col));
    }

    public MergeOffset getOffset(Cell cell) {
        MockCell cellReference = new MockCell(cell.getRowIndex(), cell.getColumnIndex());
        return originalCells.get(cellReference);
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
}
