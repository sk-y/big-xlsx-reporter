package bigreport.performers;

import java.io.IOException;

public class OutLineStart implements IterationPerformer {
    public void iterate(IterationContext context) throws IOException {
        context.getCellIterator().increaseOutlineLevel();
        context.getCellIterator().skipRow();
    }

    @Override
    public boolean shouldSwitchToAnotherPerformer(Object value) {
        return false;
    }

    @Override
    public void startAnotherPerformer(IterationContext context, Object value) throws IOException {
        throw new UnsupportedOperationException("No iteration performer should be started in start of outline tag");
    }
}
