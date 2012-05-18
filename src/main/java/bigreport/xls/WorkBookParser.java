package bigreport.xls;

import bigreport.TemplateDescription;
import bigreport.performers.HeaderIterationPerformer;
import bigreport.velocity.VelocityTemplateBuilder;
import bigreport.xls.merge.MergeInfo;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class WorkBookParser {
    private Workbook workbook;
    private VelocityTemplateBuilder templateBuilder = new VelocityTemplateBuilder();

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
        File tempMergeData=null;
        try {
            tempMergeData = File.createTempFile("merge" + sheet.getSheetName(), null);
            tempMergeOutputStream = new FileOutputStream(tempMergeData);
            CellIterator cellIterator = new CellIterator(sheet, tempMergeOutputStream);
            templateBuilder.start();
            new HeaderIterationPerformer().iterate(cellIterator, templateBuilder);
            MergeInfo mergeInfo = new MergeInfo(tempMergeData, cellIterator.getMergedCount());
            return new TemplateDescription(templateBuilder.getTemplate(), cellIterator, mergeInfo);
        } catch (IOException e) {
            if (tempMergeOutputStream != null) {
                tempMergeOutputStream.close();
                tempMergeOutputStream=null;
            }
            if (tempMergeData!=null && tempMergeData.exists()){
                FileUtils.forceDelete(tempMergeData);
            }
            throw e;
        } finally {
            if (tempMergeOutputStream != null) {
                tempMergeOutputStream.close();
            }
        }
    }
}
