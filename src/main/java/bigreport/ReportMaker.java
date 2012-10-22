package bigreport;

import bigreport.exception.CompositeIOException;
import bigreport.exception.FileDeleteException;
import bigreport.velocity.VelocityResolver;
import bigreport.velocity.VelocityResult;
import bigreport.xls.WorkBookParser;
import bigreport.xls.merge.MergeInfo;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.velocity.runtime.directive.Directive;

import java.io.*;
import java.util.Date;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static bigreport.util.StreamUtil.closeZipFile;
import static bigreport.util.StreamUtil.finishZipOutputStream;

public class ReportMaker {

    private VelocityResolver resolver;

    private static final String DEFAULT_CHAR_SET = "UTF-8";

    private String ecnoding;

    public ReportMaker(Map beans) {
        this(new VelocityResolver(beans));
    }

    public ReportMaker(Map beans, List<Class<? extends Directive>> directives) {
        this(new VelocityResolver(beans, directives));
    }

    public ReportMaker(VelocityResolver resolver) {
        this(resolver, DEFAULT_CHAR_SET);
    }

    public ReportMaker(VelocityResolver resolver, String charset) {
        this.resolver = resolver;
        if (charset != null) {
            this.ecnoding = charset;
        } else {
            this.ecnoding = DEFAULT_CHAR_SET;
        }
    }

    public void createReport(InputStream templateInputStream, OutputStream resultOutputStream) throws IOException {
        File templateCopy = File.createTempFile("rm-report-" + new Date().getTime(), null);
        OutputStream templateCopyOutputStream = null;
        CompositeIOException compositeIOException=null;
        try {
            templateCopyOutputStream = new FileOutputStream(templateCopy);
            IOUtils.copy(templateInputStream, templateCopyOutputStream);
            doCreateReport(resultOutputStream, templateCopy);
        } catch (Exception e){
            compositeIOException=new CompositeIOException(e);
            throw compositeIOException;
        }  finally {
            IOUtils.closeQuietly(templateCopyOutputStream);
            deleteTemplateCopy(templateCopy, compositeIOException);
        }
    }

    private void deleteTemplateCopy(File templateCopy, CompositeIOException compositeException) throws IOException {
        if (templateCopy != null && templateCopy.exists()) {
            try {
                FileUtils.forceDelete(templateCopy);
            } catch (IOException e){
                FileDeleteException exception=new FileDeleteException(e, templateCopy);
                if (compositeException==null){
                    throw exception;
                }
                compositeException.addSuppressedException(exception);
                throw compositeException;
            }
        }
    }

    public void createReport(String templatePath, String destFilePath) throws IOException {
        File destFile = new File(destFilePath);
        FileOutputStream destOutputStream = null;
        File templateCopy = null;
        CompositeIOException compositeIOException=null;
        try {
            destOutputStream = new FileOutputStream(destFile);
            File templateFile = new File(templatePath);
            templateCopy = File.createTempFile(destFile.getName(), null);
            FileUtils.copyFile(templateFile, templateCopy);
            doCreateReport(destOutputStream, templateCopy);
        } catch (Exception e){
            compositeIOException=new CompositeIOException(e);
            throw compositeIOException;
        } finally {
            IOUtils.closeQuietly(destOutputStream);
            deleteTemplateCopy(templateCopy, compositeIOException);
        }
    }


    public ReportMaker addDirective(Class<? extends Directive> directiveClass) {
        resolver.addDirective(directiveClass);
        return this;
    }

    private void doCreateReport(OutputStream destOutputStream, File templateCopy) throws IOException {
        ZipOutputStream zipDestFileOutputStream = null;
        ZipFile zippedTemplateCopy = null;
        try {
            zipDestFileOutputStream = new ZipOutputStream(destOutputStream);
            Map<String, TemplateDescription> sheetTemplates = getTemplatePerSheet(templateCopy);
            zippedTemplateCopy = new ZipFile(templateCopy);
            if (sheetTemplates.isEmpty()) {
                return;
            }
            writeWorkBook(zipDestFileOutputStream, zippedTemplateCopy, sheetTemplates);
        } finally {
            closeZipFile(zippedTemplateCopy);
            finishZipOutputStream(zipDestFileOutputStream);
        }

    }

    private void writeWorkBook(ZipOutputStream zipDestFileOutputStream, ZipFile zippedTemplateCopy,
                               Map<String, TemplateDescription> sheetTemplates) throws IOException {
        Enumeration<? extends ZipEntry> zipEntries = zippedTemplateCopy.entries();
        while (zipEntries.hasMoreElements()) {
            ZipEntry zipEntry = zipEntries.nextElement();
            zipDestFileOutputStream.putNextEntry(new ZipEntry(zipEntry.getName()));
            InputStream tempZipInputStream = null;
            try {
                tempZipInputStream = zippedTemplateCopy.getInputStream(zipEntry);
                TemplateDescription description = sheetTemplates.get(zipEntry.getName());
                if (description != null && !description.isEmpty()) {
                    VelocityResult result = resolveVelocityTemplate(zipEntry.getName(), sheetTemplates.get(zipEntry.getName()));
                    String xlsxTemplateString = IOUtils.toString(tempZipInputStream, ecnoding);
                    writeSheet(zipDestFileOutputStream, result, xlsxTemplateString, sheetTemplates.get(zipEntry.getName()));
                    result.deleteAllTempFiles();
                } else {
                    IOUtils.copy(tempZipInputStream, zipDestFileOutputStream);
                    if (description != null && description.getMergeInfo() != null) {
                        description.getMergeInfo().reset();
                    }
                }
                zipDestFileOutputStream.flush();
                zipDestFileOutputStream.closeEntry();
            } finally {
                if (tempZipInputStream != null) {
                    tempZipInputStream.close();
                }
            }
        }
    }

    private VelocityResult resolveVelocityTemplate(String name, TemplateDescription templateDescription) throws IOException {
        Writer fw = null;
        File tempDataFile = null;
        MergeInfo mergeInfo = templateDescription.getMergeInfo();
        try {
            tempDataFile = File.createTempFile(name.substring(name.lastIndexOf('/')).replace(' ', '_'), null);
            FileOutputStream outputStream = new FileOutputStream(tempDataFile);
            fw = new OutputStreamWriter(outputStream, DEFAULT_CHAR_SET);
            resolver.resolve(templateDescription, fw, mergeInfo.getMergedTempFile());
        } finally {
            if (fw != null) {
                fw.close();
            }
        }
        return new VelocityResult(tempDataFile, mergeInfo);
    }

    private void writeSheet(OutputStream output, VelocityResult velocityResult,
                            String xlsxTemplateString, TemplateDescription templateDescription) throws IOException {
        writeBeforeData(templateDescription.getCellFrom().getRow() + 1, xlsxTemplateString, output);
        writeDataFromTempFile(velocityResult.getDataXmlFile(), output);
        writeBetweenDataAndMergedSection(velocityResult, xlsxTemplateString, output);
        writeMergedCellsSection(velocityResult, output);
        writeRestData(velocityResult, xlsxTemplateString, output);
    }

    private Map<String, TemplateDescription> getTemplatePerSheet(File templateCopy) throws IOException {
        XSSFWorkbook wb = null;
        FileInputStream templateCopyInputStream = null;
        try {
            templateCopyInputStream = new FileInputStream(templateCopy.getCanonicalPath());
            wb = new XSSFWorkbook(templateCopyInputStream);
            WorkBookParser parser = new WorkBookParser(wb);
            return parser.parseWorkBook();
        } finally {
            if (wb != null) {
                closePackage(wb.getPackage());
            }
            IOUtils.closeQuietly(templateCopyInputStream);
        }
    }

    private void closePackage(OPCPackage aPackage) {
        try {
            aPackage.close();
        } catch (IOException e) {
            //do nothing
        }
    }


    private void writeRestData(VelocityResult velocityResult, String xlsxTemplateString, OutputStream output) throws IOException {
        int startIndex;
        if (velocityResult.getMergeInfo().getCount() == 0) {
            int startMergedCells = xlsxTemplateString.indexOf(Markers.START_MERGED_CELLS);
            int endMergedCells = xlsxTemplateString.indexOf(Markers.END_MERGED_CELLS) + Markers.END_MERGED_CELLS.length();
            if (startMergedCells >= 0 && endMergedCells >= 0) {
                xlsxTemplateString = new StringBuffer(xlsxTemplateString).replace(startMergedCells, endMergedCells, "").toString();
            }
            startIndex = xlsxTemplateString.indexOf(Markers.END_DATA);
        } else {
            startIndex = xlsxTemplateString.indexOf(Markers.END_MERGED_CELLS);
        }
        if (startIndex >= 0) {
            IOUtils.write(xlsxTemplateString.substring(startIndex), output, DEFAULT_CHAR_SET);
        }
    }

    private void writeMergedCellsSection(VelocityResult velocityResult,
                                         OutputStream output) throws IOException {
        FileInputStream is = null;
        if (velocityResult.getMergeInfo().getCount() == 0) {
            return;
        }
        try {
            String mergeHeader = Markers.START_MERGED_CELLS + "count=\"" + velocityResult.getMergeInfo().getCount() + "\">";
            IOUtils.write(mergeHeader, output, DEFAULT_CHAR_SET);
            is = new FileInputStream(velocityResult.getMergeInfo().getMergedTempFile());
            IOUtils.copy(is, output);
        } finally {
            if (is != null) {
                is.close();
            }
        }
    }

    private void writeBetweenDataAndMergedSection(VelocityResult velocityResult, String xlsxTemplateString, OutputStream output) throws IOException {
        if (velocityResult.getMergeInfo().getCount() > 0) {
            int endOfDataIndex = xlsxTemplateString.indexOf(Markers.END_DATA);
            int startMergedCellsGroup = xlsxTemplateString.indexOf(Markers.START_MERGED_CELLS);
            String beforeMergedCells = "";
            if (startMergedCellsGroup >= 0) {
                beforeMergedCells = xlsxTemplateString.substring(endOfDataIndex, startMergedCellsGroup);
            }
            IOUtils.write(beforeMergedCells, output, DEFAULT_CHAR_SET);
        }
    }

    private void writeDataFromTempFile(File dataXmlFile, OutputStream output) throws IOException {
        InputStream is = null;
        try {
            is = new FileInputStream(dataXmlFile);
            IOUtils.copy(is, output);
        } finally {
            if (is != null) {
                is.close();
            }
        }
    }

    private void writeBeforeData(int rowStart, String xlsxTemplateString, OutputStream output) throws IOException {
        String beforeTemplate = xlsxTemplateString.substring(0, xlsxTemplateString.indexOf("<row r=\"" + rowStart + "\""));
        IOUtils.write(beforeTemplate, output, DEFAULT_CHAR_SET);
    }

}
