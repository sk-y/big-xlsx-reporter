package bigreport.performers;


import java.io.IOException;

public interface IterationPerformer {
    void iterate(IterationContext iterationContext) throws IOException;
}
