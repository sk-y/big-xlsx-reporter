package bigreport.performers;

import bigreport.*;
import bigreport.xls.MockCell;
import bigreport.util.ValueResolver;
import bigreport.xls.CellIterator;
import org.apache.poi.ss.usermodel.Cell;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static bigreport.util.ValueResolver.*;

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

    public void iterate(IterationContext iterationContext) throws IOException {
        iterationContext.getTemplateBuilder().reset();
        CellIterator cellIterator = iterationContext.getCellIterator();
        Cell currentCell = iterationContext.getCurrentCell();
        iterationContext.setStartedAt(new MockCell(currentCell.getRowIndex(), currentCell.getColumnIndex()));
        cellIterator.skipRow();
        VelocityTemplateConstuctor constuctor = new VelocityTemplateConstuctor(iterationContext);
        while (iterationContext.hasNext()) {
            performCellInContext(iterationContext, constuctor);
        }
    }

    public void performCellInContext(IterationContext context, VelocityTemplateConstuctor constuctor) throws IOException {
        Cell cell = context.next();
        if (cell == null) {
            return;
        }
        String value = resolve(cell);
        if (shouldSwitchToAnotherPerformer(value)) {
            startAnotherPerformer(context, value);
            return;
        }
        if (isDirective(value)) {
            performDirective(context, value);
            return;
        }
        constuctor.appendCell(cell);
    }

    public void startAnotherPerformer(IterationContext context, Object value) throws IOException {
        graph.get(value).iterate(context);
    }

    public boolean shouldSwitchToAnotherPerformer(Object value) {
        return graph.containsKey(value);
    }

    private void performDirective(IterationContext context, String value) {
        context.getTemplateBuilder().addDirective(convertToDirective(value));
        context.getCellIterator().skipRow();
    }
}
