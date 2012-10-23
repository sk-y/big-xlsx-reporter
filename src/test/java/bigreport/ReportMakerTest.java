package bigreport;

import bigreport.model.CalcItem;
import bigreport.util.StreamUtil;
import bigreport.util.ValueResolver;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.velocity.texen.util.FileUtil;
import org.junit.Before;
import org.junit.Test;

import java.io.*;
import java.util.*;

import static org.junit.Assert.*;
import static org.junit.Assert.assertEquals;

/**
 * Unit test for simple ReportMaker.
 */
public class ReportMakerTest {
    private final static String OUTPUT_DIRECTORY_NAME = "out";
    private Random random;
    private static File outputDir;
    private final static String[] HEADER_ARRAY = {"Name", "Price", "Count", "Sum", "", "Date"};

    @Before
    public void init() throws IOException {
        random = new Random(System.currentTimeMillis());
        outputDir = createOutDir();
    }

    @Test
    public void testReportCreationWithCorrectTemplateFile() throws IOException {
        testReportCreation("example.xlsx");
    }

    @Test
    public void testReportCreationWithIncorrectFileFormat() throws IOException {
        testeReportCreateWithIncorrectFile("wrong.txt");
    }

    @Test
    public void testReportCreationWithIncorrectArchiveFormat() throws IOException {
        testeReportCreateWithIncorrectFile("wrong.zip");
    }

    private void testeReportCreateWithIncorrectFile(String fileName) throws IOException {
        new IncorrectTemplateTester() {
            void testReportCreationForFile(String fileName) throws IOException {
                testReportCreation(fileName);
            }
        }.doTest(fileName);
    }

    private void testReportCreation(String fileName) throws IOException {
        File reportFile=null;
        try {
            String templatePath = Thread.currentThread().getContextClassLoader().getResource(fileName).getFile();
            Map<String, Object> bean = createDataBean(random);
            Date startDate = fixStartTime();
            String reportFileName = createReportFilePathAndName(outputDir);
            reportFile = new File(reportFileName);
            new ReportMaker(bean).createReport(templatePath, reportFileName);
            fixFinishedTime(startDate);
            checkData(reportFile, bean);
        } finally {
         //   StreamUtil.forceDelete(reportFile, null);
        }
    }

    private Date fixStartTime() {
        Date startDate = new Date();
        System.out.println("start at " + startDate);
        return startDate;
    }

    private void fixFinishedTime(Date startDate) {
        Date finishedDate = new Date();
        System.out.println("finished at " + finishedDate.getTime());
        System.out.println("Lasts " + (finishedDate.getTime() - startDate.getTime()) + "mils");
    }

    private String createReportFilePathAndName(File outDir) {
        return new File(outDir, "template" + System.currentTimeMillis() + ".xlsx").getAbsolutePath();
    }

    private File createOutDir() throws IOException {
        File outDir = new File(OUTPUT_DIRECTORY_NAME);
        if (outDir.exists()) {
            System.out.println("Output directory exists:" + outDir.getAbsolutePath());
            return outDir;
        }
        if (!outDir.mkdir()) {
            throw new IOException("Error while creating output directory");
        }
        System.out.println("Output directory created:" + outDir.getAbsolutePath());
        return outDir;
    }

    @Test
    public void testReportCreationWithStreamsWithCorrectTemplate() throws IOException {
        testReportCreationWithStreams("example.xlsx");
    }

    @Test
    public void testReportCreationWithStreamsAndIncorrectFileFormat() throws IOException {
        testeReportCreateWithStreamsWithIncorrectFile("wrong.txt");
    }

    @Test
    public void testReportCreationWithSteamsAndIncorrectArchiveFormat() throws IOException {
        testeReportCreateWithStreamsWithIncorrectFile("wrong.zip");
    }

    private void testeReportCreateWithStreamsWithIncorrectFile(String fileName) throws IOException {
        new IncorrectTemplateTester() {
            void testReportCreationForFile(String fileName) throws IOException {
                testReportCreationWithStreams(fileName);
            }
        }.doTest(fileName);
    }

    private void testReportCreationWithStreams(String resourceName) throws IOException {
        String templatePath = Thread.currentThread().getContextClassLoader().getResource(resourceName).getFile();
        FileOutputStream fos = null;
        File reportFile = null;
        try {
            Map<String, Object> bean = createDataBean(random);
            Date startDate = fixStartTime();
            FileInputStream fis = new FileInputStream(templatePath);
            reportFile = new File(createReportFilePathAndName(outputDir));
            fos = new FileOutputStream(reportFile);
            new ReportMaker(bean).createReport(fis, fos);
            fixFinishedTime(startDate);
            checkData(reportFile, bean);
        } finally {
            closeOutputStream(fos);
            StreamUtil.forceDelete(reportFile, null);
        }
    }

    private void closeOutputStream(FileOutputStream fos) throws IOException {
        if (fos != null) {
            fos.flush();
            fos.close();
        }
    }

    private Map<String, Object> createDataBean(Random random) {
        List<CalcItem> itemList = new ArrayList<CalcItem>();
        while (itemList.size() < 1000) {
            itemList.add(new CalcItem("<\"&Unit1" + itemList.size(), random.nextFloat() * 20, random.nextInt(20), new Date()));
        }
        itemList.add(new CalcItem(null, 0, 0, null));
        Map<String, Object> bean = new HashMap<String, Object>();
        bean.put("rows", itemList);
        return bean;
    }


    @Test
    public void convertString() {
        String res = ValueResolver.convertStringForXml("<d>f&g\"'");
        assertEquals(res, "&lt;d&gt;f&amp;g&quot;&apos;");
    }

    private void checkData(File resultFile, Map<String, Object> dataBean) throws IOException {
        FileInputStream resultStream = null;
        try {
            resultStream = new FileInputStream(resultFile);
            Workbook workbook = new XSSFWorkbook(resultStream);
            Sheet sheet = workbook.getSheet("Sheet1");
            Row row = sheet.getRow(0);
            assertHeaders(row, HEADER_ARRAY);
            assertData(sheet, dataBean);
        } finally {
            StreamUtil.close(resultStream, null);
        }
    }

    private void assertData(Sheet sheet, Map<String, Object> dataBean) {
        List itemList = (List) dataBean.get("rows");
        for (int i = 0; i < itemList.size(); i++) {
            CalcItem item = (CalcItem) itemList.get(i);
            assertRow(item, sheet.getRow(i + 1));
        }
    }

    private void assertRow(CalcItem item, Row row) {
        if (item.getPrice() < 10) {
            assertPriceLessThenTen(item, row);
        } else {
            assertPriceGreaterEqualsThenTen(item, row);
        }
    }

    private void assertPriceGreaterEqualsThenTen(CalcItem item, Row row) {
        assertGeneralCells(item, row);
        assertEquals("\"" + item.getName(), row.getCell(0).getStringCellValue());
        if (item.getPrice() < 15) {
            assertEquals(item.getSum(), row.getCell(3).getNumericCellValue(), 0);
        } else {
            assertEquals(item.getSum(), row.getCell(4).getNumericCellValue(), 0);
        }
    }


    private void assertPriceLessThenTen(CalcItem item, Row row) {
        assertEquals(nvl(item.getName()), row.getCell(0).getStringCellValue());
        assertGeneralCells(item, row);
        assertEquals(item.getSum(), row.getCell(3).getNumericCellValue(), 0);
        assertEquals(Cell.CELL_TYPE_BLANK, row.getCell(4).getCellType());
    }

    private void assertGeneralCells(CalcItem item, Row row) {
        assertEquals(item.getPrice(), row.getCell(1).getNumericCellValue(), 0);
        assertEquals(item.getCount(), row.getCell(2).getNumericCellValue(), 0);
        assertEquals(item.getDate(), row.getCell(5).getDateCellValue());
    }

    private void assertHeaders(Row row, String[] headerArray) {
        for (int i = 0; i < headerArray.length; i++) {
            assertEquals(row.getCell(i).getStringCellValue(), headerArray[i]);
        }
    }

    private Object nvl(Object o){
        if (o==null){
            return "";
        }
        return o;
    }
}