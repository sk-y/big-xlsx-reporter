package bigreport.performers;

import org.apache.poi.ss.usermodel.Cell;

import java.io.IOException;

/**
 * Performs cell iteration after "end" directive
 */
public class EndIterationPerformer implements IterationPerformer {
    public void iterate(IterationContext iterationContext) {
        Cell currentCell = iterationContext.getCurrentCell();
        iterationContext.getCellIterator().setFinishedAt(currentCell.getRowIndex(), currentCell.getColumnIndex());
    }

    @Override
    public boolean shouldSwitchToAnotherPerformer(Object value) {
        return false;
    }

    @Override
    public void startAnotherPerformer(IterationContext context, Object value){
        throw new UnsupportedOperationException("No iteration performer should be started after end tag");
    }
}
