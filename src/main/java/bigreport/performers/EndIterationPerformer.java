package bigreport.performers;

import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import bigreport.velocity.VelocityTemplateBuilder;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Performs cell iteration after "end" directive
 */
public class EndIterationPerformer implements IterationPerformer {
    public void iterate(CellIterator cellIterator, VelocityTemplateBuilder velocityTemplateBuilder) {
        Cell currentCell = cellIterator.getCurrentCell();
        cellIterator.setFinishedAt(new MockCell(currentCell.getRowIndex(), currentCell.getColumnIndex()));
        cellIterator.remove();
    }
}
