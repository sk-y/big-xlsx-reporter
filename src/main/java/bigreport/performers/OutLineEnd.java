package bigreport.performers;

import java.io.IOException;

public class OutLineEnd implements IterationPerformer {
    public void iterate(IterationContext context) throws IOException {
        context.getCellIterator().decreaseOutlineLevel();
        context.getCellIterator().skipRow();
    }

    @Override
    public boolean shouldSwitchToAnotherPerformer(Object value) {
        return false;
    }

    @Override
    public void startAnotherPerformer(IterationContext context, Object value) throws IOException {
        throw new UnsupportedOperationException("No iteration performer should be started in end of outline tag");
    }
}
