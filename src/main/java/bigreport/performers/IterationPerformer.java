package bigreport.performers;


import java.io.IOException;

import static bigreport.util.ValueResolver.convertToDirective;

public interface IterationPerformer {
    void iterate(IterationContext iterationContext) throws IOException;

    public boolean shouldSwitchToAnotherPerformer(Object value);

    public void startAnotherPerformer(IterationContext context, Object value) throws IOException;
}
