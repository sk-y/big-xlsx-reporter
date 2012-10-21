package bigreport.performers;

import java.io.IOException;

public class OutLineEnd implements IterationPerformer {
    public void iterate(IterationContext context) throws IOException {
        context.getCellIterator().decreaseOutlineLevel();
        context.getCellIterator().skipRow();
    }
}
