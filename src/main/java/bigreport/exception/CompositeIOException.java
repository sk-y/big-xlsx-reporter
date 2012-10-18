package bigreport.exception;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class CompositeIOException extends IOException {
    private List<Exception> suppressedExceptionList=new ArrayList<Exception>();

    public CompositeIOException() {
    }

    public CompositeIOException(String message) {
        super(message);
    }

    public CompositeIOException(String message, Throwable cause) {
        super(message, cause);
    }

    public CompositeIOException(Throwable cause) {
        super(cause);
    }

    public void addSuppressedException(Exception e){
        suppressedExceptionList.add(e);
    }

    public List<Exception> getSuppressedExceptionList() {
        return suppressedExceptionList;
    }
}
