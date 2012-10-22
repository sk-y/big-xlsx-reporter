package bigreport.performers;

import bigreport.util.StreamUtil;
import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeCells;
import bigreport.xls.merge.MergeOffset;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

public class IterationContext implements Iterator<Cell> {
    private CellIterator cellIterator;
    private VelocityTemplateBuilder templateBuilder;
    private final MergeCells originalMergedCells;
    private OutputStream mergedOutputStream;
    private MockCell startedAt;
    private MockCell finishedAt;

    private IterationContext(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder,
                             MergeCells mergeCells, OutputStream mergedOutputStream) {
        this.cellIterator = cellIterator;
        this.templateBuilder = templateBuilder;
        this.originalMergedCells=mergeCells;
        this.mergedOutputStream=mergedOutputStream;
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

    public int getMergedCount(){
        return cellIterator.getMergedCount();
    }

    public MockCell getStartedAt(){
        return startedAt;
    }

    public void setStartedAt(MockCell startedAt) {
        this.startedAt = startedAt;
    }

    public MergeCells getOriginalMergedCells() {
        return originalMergedCells;
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

    public MockCell getFinishedAt() {
        return finishedAt;
    }

    public void setFinishedAt(MockCell finishedAt) {
        this.finishedAt = finishedAt;
    }

    public void setFinishedAt(int row, int col){
        this.finishedAt=new MockCell(row, col);
    }

    public static IterationContext createInstance(XSSFSheet sheet, OutputStream mergedOutputStream){
        VelocityTemplateBuilder templateBuilder = new VelocityTemplateBuilder().reset();
        CellIterator cellIterator = new CellIterator(sheet);
        MergeCells mergeCells=MergeCells.detect(sheet);
        IterationContext iterationContext=new IterationContext(cellIterator, templateBuilder, mergeCells, mergedOutputStream);
        cellIterator.setContext(iterationContext);
        return iterationContext;
    }

    public void addMergedCell(MockCell mockCell, MergeOffset offset) throws IOException {
        StreamUtil.writeMergedCell(mergedOutputStream, mockCell, offset);
    }
}
