package bigreport;

import bigreport.exception.CompositeIOException;
import org.apache.poi.POIXMLException;

import java.io.IOException;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public abstract class IncorrectTemplateTester {

    public void doTest(String fileName) throws IOException {
        CompositeIOException mandatoryException = null;
        try {
            testReportCreationForFile(fileName);
        } catch (CompositeIOException e) {
            if (!(e.getCause() instanceof POIXMLException)) {
                throw e;
            }
            mandatoryException = e;
        }
        assertNotNull(mandatoryException);
        assertTrue(mandatoryException.getSuppressedExceptionList().isEmpty());
    }

    abstract void testReportCreationForFile(String fileName) throws IOException;
}
