package bigreport.performers;

import bigreport.*;
import bigreport.xls.MockCell;
import bigreport.util.ValueResolver;
import bigreport.xls.CellIterator;
import org.apache.poi.ss.usermodel.Cell;

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

    public void iterate(IterationContext iterationContext) throws IOException {
        iterationContext.getTemplateBuilder().start();
        CellIterator cellIterator=iterationContext.getCellIterator();
        Cell currentCell = iterationContext.getCurrentCell();
        cellIterator.setStartedAt(new MockCell(currentCell.getRowIndex(), currentCell.getColumnIndex()));
        cellIterator.skipRow();
        VelocityTemplateConstuctor constuctor=new VelocityTemplateConstuctor(iterationContext);
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell != null) {
                String value = ValueResolver.resolve(cell);
                if (graph.containsKey(value)) {
                    graph.get(value).iterate(iterationContext);
                }  else {
                    constuctor.appendCell(cell);
                }
            }
        }
    }
}
