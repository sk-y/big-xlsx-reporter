package bigreport.performers;

import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.CellIterator;

import java.io.IOException;

public class OutLineStart implements IterationPerformer {
    public void iterate(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder) throws IOException {
        cellIterator.increaseOutlineLevel();
        cellIterator.skipRow();
    }
}
