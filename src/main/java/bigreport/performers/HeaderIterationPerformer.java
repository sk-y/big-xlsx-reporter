package bigreport.performers;

import bigreport.xls.CellIterator;
import bigreport.Markers;
import bigreport.util.ValueResolver;
import bigreport.velocity.VelocityTemplateBuilder;
import org.apache.poi.ss.usermodel.Cell;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Performs cell iteration on header cells.
 */
public class HeaderIterationPerformer implements IterationPerformer {

    private final Map<String, IterationPerformer> graph = new HashMap<String, IterationPerformer>();

    public HeaderIterationPerformer() {
        DataIterationPerformer dataPerformer = new DataIterationPerformer();
        graph.put(Markers.DATA, dataPerformer);
        graph.put(Markers.DATA_SYNONYM, dataPerformer);
        graph.put(Markers.OUT_LINE_START, new OutLineStart());
        graph.put(Markers.OUT_LINE_END, new OutLineEnd());
    }

    @Override
    public void iterate(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder) throws IOException {
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            boolean isMerged=cellIterator.isMergedCell();
            if (cell != null) {
                String value = ValueResolver.resolve(cell);
                if (graph.containsKey(value)) {
                    graph.get(value).iterate(cellIterator, templateBuilder);
                }
            }
            if (isMerged) {
                cellIterator.addMergedCell(cell);
            }
        }
    }
}
