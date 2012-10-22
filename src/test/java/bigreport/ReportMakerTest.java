package bigreport;

import bigreport.model.CalcItem;
import bigreport.util.ValueResolver;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * Unit test for simple ReportMaker.
 */
public class ReportMakerTest {
    private final static String OUTPUT_DIRECTORY_NAME ="out";
    private Random random;
    private static File outputDir;

    @Before
    public void init() throws IOException {
        random = new Random(System.currentTimeMillis());
        outputDir=createOutDir();
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
        String templatePath = Thread.currentThread().getContextClassLoader().getResource(fileName).getFile();
        Map<String, Object> bean = createDataBean(random);
        Date startDate = fixStartTime();
        new ReportMaker(bean).createReport(templatePath, createReportFilePathAndName(outputDir));
        fixFinishedTime(startDate);
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
        if (outDir.exists()){
            System.out.println("Output directory exists:"+outDir.getAbsolutePath());
            return outDir;
        }
        if (!outDir.mkdir()){
            throw new IOException("Error while creating output directory");
        }
        System.out.println("Output directory created:"+outDir.getAbsolutePath());
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
        try {
            Map<String, Object> bean = createDataBean(random);
            Date startDate = fixStartTime();
            FileInputStream fis = new FileInputStream(templatePath);
            File reportFile = new File(createReportFilePathAndName(outputDir));
            fos = new FileOutputStream(reportFile);
            new ReportMaker(bean).createReport(fis, fos);
            fixFinishedTime(startDate);
        } finally {
            closeOutputStream(fos);
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
        Map<String, Object> bean = new HashMap<String, Object>();
        bean.put("rows", itemList);
        return bean;
    }


    @Test
    public void convertString() {
        String res = ValueResolver.convertStringForXml("<d>f&g\"'");
        assertEquals(res, "&lt;d&gt;f&amp;g&quot;&apos;");
    }


}