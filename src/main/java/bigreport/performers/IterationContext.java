package bigreport.performers;

import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Iterator;

public class IterationContext implements Iterator<Cell> {
    private CellIterator cellIterator;
    private VelocityTemplateBuilder templateBuilder;

    public IterationContext(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder) {
        this.cellIterator = cellIterator;
        this.templateBuilder = templateBuilder;
    }

    public CellIterator getCellIterator() {
        return cellIterator;
    }

    public VelocityTemplateBuilder getTemplateBuilder() {
        return templateBuilder;
    }

    public int getIteratedOutlineLevel(){
        return cellIterator.getOutlineLevel();
    }

    public Cell getCurrentCell(){
       return cellIterator.getCurrentCell();
    }

    public MockCell getOriginalMergedHeadRegion(){
        return getOriginalMergedHeadRegion(cellIterator.getCurrentCell());
    }

    public MockCell getOriginalMergedHeadRegion(Cell cell){
        return  cellIterator.getMergedRegionHeader(new MockCell(cell.getRowIndex(), cell.getColumnIndex()));
    }

    public String getValueAt(int row, int col){
        return cellIterator.getValueAt(row, col);
    }

    public String getValueAt(MockCell mockCell){
        return getValueAt(mockCell.getRow(), mockCell.getCol());
    }

    @Override
    public boolean hasNext() {
        return cellIterator.hasNext();
    }

    @Override
    public Cell next() {
        return cellIterator.next();
    }

    @Override
    public void remove() {
        throw new UnsupportedOperationException("Cannot remove item from original file");
    }
}
