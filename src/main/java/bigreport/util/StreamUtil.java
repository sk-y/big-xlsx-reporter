package bigreport.util;

import bigreport.exception.CompositeIOException;
import bigreport.exception.FileDeleteException;
import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeOffset;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

import java.io.*;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class StreamUtil {
    private StreamUtil() {
    }

    public static void writeMergedCell(OutputStream mergedCellsOutputStream, int row, int yOffset, int col, int xOffset) throws IOException {
        String mergedCellAsString = ValueResolver.getMergedCellAsString(row, yOffset, col, xOffset);
        IOUtils.write(mergedCellAsString, mergedCellsOutputStream);

    }

    public static void writeMergedCell(OutputStream mergedCellsOutputStream, MockCell cell, MergeOffset offset) throws IOException {
        writeMergedCell(mergedCellsOutputStream, cell.getRow(), offset.getYOffset(), cell.getCol(), offset.getXOffset());
    }

    public static void close(Closeable tempMergeOutputStream, CompositeIOException ce) throws IOException {
        if (tempMergeOutputStream == null) {
            return;
        }
        try {
            tempMergeOutputStream.close();
        } catch (IOException e) {
            if (ce != null) {
                ce.addSuppressedException(e);
                throw ce;
            }
            throw e;
        }
    }

    public static void forceDelete(File fileToDelete, CompositeIOException ce) throws IOException {
        if (fileToDelete != null && fileToDelete.exists()) {
            try {
                FileUtils.forceDelete(fileToDelete);
            } catch (IOException e){
                FileDeleteException exception=new FileDeleteException(e, fileToDelete);
                if (ce==null){
                    throw exception;
                }
                ce.addSuppressedException(exception);
                throw ce;
            }
        }
    }

    public static void closeStreamAndDeleteFile(Closeable stream, File fileToDelete, CompositeIOException ce) throws IOException {
        if (ce == null) {
            ce = new CompositeIOException();
        }
        try {
            close(stream, ce);
        } finally {
            forceDelete(fileToDelete, ce);
        }
    }

    public static void finishZipOutputStream(ZipOutputStream zipDestFileOutputStream) {
        try {
            if (zipDestFileOutputStream != null) {
                zipDestFileOutputStream.finish();
            }
        } catch (Exception e) {
            //ignore
        }
    }

    public static void closeZipFile(ZipFile zippedTemplateCopy) {
        try {
            if (zippedTemplateCopy != null) {
                zippedTemplateCopy.close();
            }
        } catch (Exception e) {
            //ingnore
        }
    }

    public static void writeFromFileToOutputStream(File file, OutputStream output) throws IOException {
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            IOUtils.copy(is, output);
        } finally {
            if (is != null) {
                is.close();
            }
        }
    }

}
