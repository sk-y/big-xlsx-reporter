package bigreport.performers;

import bigreport.xls.CellIterator;
import bigreport.velocity.VelocityTemplateBuilder;

import java.io.IOException;

public interface IterationPerformer {
    void iterate(CellIterator cellIterator, VelocityTemplateBuilder templateBuilder) throws IOException;
}
