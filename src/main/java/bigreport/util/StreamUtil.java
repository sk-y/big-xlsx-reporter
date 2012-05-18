package bigreport.util;

import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeOffset;
import org.apache.commons.io.IOUtils;

import java.io.IOException;
import java.io.OutputStream;

public class StreamUtil {
    private StreamUtil() {
    }

    public static void writeMergedCell(OutputStream mergedCellsOutputStream, int row, int yOffset, int col, int xOffset) throws IOException {
        String mergedCellAsString=ValueResolver.getMergedCellAsString(row,yOffset,col,xOffset);
        IOUtils.write(mergedCellAsString, mergedCellsOutputStream);

    }

    public static void writeMergedCell(OutputStream mergedCellsOutputStream, MockCell cell, MergeOffset offset) throws IOException {
        writeMergedCell(mergedCellsOutputStream, cell.getRow(), offset.getYOffset(), cell.getCol(), offset.getXOffset());
    }
}
