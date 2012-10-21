package bigreport.performers;

import java.io.IOException;

public class OutLineStart implements IterationPerformer {
    public void iterate(IterationContext context) throws IOException {
        context.getCellIterator().increaseOutlineLevel();
        context.getCellIterator().skipRow();
    }
}
