package bigreport.xls;

import bigreport.TemplateDescription;
import bigreport.exception.CompositeIOException;
import bigreport.performers.HeaderIterationPerformer;
import bigreport.performers.IterationContext;
import bigreport.util.StreamUtil;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static bigreport.util.StreamUtil.close;
import static bigreport.util.StreamUtil.closeStreamAndDeleteFile;

public class WorkBookParser {
    private Workbook workbook;

    public WorkBookParser(Workbook workbook) {
        this.workbook = workbook;
    }

    public Map<String, TemplateDescription> parseWorkBook() throws IOException {
        int sheetCount = workbook.getNumberOfSheets();
        Map<String, TemplateDescription> result = new HashMap<String, TemplateDescription>();
        for (int i = 0; i < sheetCount; i++) {
            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(i);
            TemplateDescription templateDesc = parseSheet(sheet);
            String sheetRef = sheet.getPackagePart().getPartName().getName();
            if (sheetRef.startsWith("/")) {
                sheetRef = sheetRef.substring(1);
            }
            result.put(sheetRef, templateDesc);
        }
        return result;
    }

    private TemplateDescription parseSheet(XSSFSheet sheet) throws IOException {
        FileOutputStream tempMergeOutputStream = null;
        File tempMergeData = null;
        CompositeIOException compositeIOException = null;
        try {
            tempMergeData = File.createTempFile("merge" + sheet.getSheetName(), null);
            tempMergeOutputStream = new FileOutputStream(tempMergeData);
            IterationContext context = IterationContext.createInstance(sheet, tempMergeOutputStream);
            new HeaderIterationPerformer().iterate(context);
            return new TemplateDescription(context, tempMergeData);
        } catch (IOException e) {
            compositeIOException = new CompositeIOException(e);
            closeStreamAndDeleteFile(tempMergeOutputStream, tempMergeData, compositeIOException);
            tempMergeOutputStream=null;
            throw compositeIOException;
        } finally {
            close(tempMergeOutputStream, compositeIOException);
        }
    }
}
