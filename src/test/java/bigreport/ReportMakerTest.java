package bigreport;

import bigreport.model.CalcItem;
import bigreport.util.ValueResolver;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static org.junit.Assert.assertEquals;

/**
 * Unit test for simple ReportMaker.
 */
public class ReportMakerTest {

    @Test
    public void testReportCreation() {
        String templatePath = Thread.currentThread().getContextClassLoader().getResource("example.xlsx").getFile();
        try {
            Random random = new Random();
            Map<String, Object> bean = createDataBean(random);
            File outDir= new File("out");
            outDir.mkdir();
            Date startDate = new Date();
            System.out.println("start at " + startDate);
            new ReportMaker(bean).createReport(templatePath, "out/template"+System.currentTimeMillis()+".xlsx");
            Date finishedDate = new Date();
            System.out.println("finished at " + new Date().getTime());
            System.out.println("Lasts " + (finishedDate.getTime() - startDate.getTime()) + "mils");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testReportCreationWithStreams() {
        String templatePath = Thread.currentThread().getContextClassLoader().getResource("example.xlsx").getFile();
        FileOutputStream fos = null;
        try {
            Random random = new Random();
            Map<String, Object> bean = createDataBean(random);
            Date startDate = new Date();
            System.out.println("start at " + startDate);
            FileInputStream fis = new FileInputStream(templatePath);
            File outDir= new File("out");
            outDir.mkdir();
            System.out.println(outDir.getAbsolutePath());
            File file = new File("out/template"+System.currentTimeMillis()+".xlsx");
            fos = new FileOutputStream(file);
            new ReportMaker(bean).createReport(fis, fos);
            Date finishedDate = new Date();
            System.out.println("finished at " + new Date().getTime());
            System.out.println("Lasts " + (finishedDate.getTime() - startDate.getTime()) + "mils");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (fos!=null){
                try{
                    fos.flush();
                    fos.close();
                } catch (IOException exception){
                    throw new RuntimeException(exception);
                }
            }
        }
    }

    private Map<String, Object> createDataBean(Random random) {
        List<CalcItem> itemList = new ArrayList<CalcItem>();
        while (itemList.size() < 1000) {
            itemList.add(new CalcItem("<&Unit1" + itemList.size(), random.nextFloat() * 20, random.nextInt(20)));
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