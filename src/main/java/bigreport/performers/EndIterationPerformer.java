package bigreport.performers;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Performs cell iteration after "end" directive
 */
public class EndIterationPerformer implements IterationPerformer {
    public void iterate(IterationContext iterationContext) {
        Cell currentCell = iterationContext.getCurrentCell();
        iterationContext.getCellIterator().setFinishedAt(currentCell.getRowIndex(), currentCell.getColumnIndex());
    }
}
