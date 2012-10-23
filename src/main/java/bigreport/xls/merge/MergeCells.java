package bigreport.xls.merge;

import bigreport.xls.MockCell;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMergeCell;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MergeCells extends HashMap<MockCell, MergeOffset> {

    private MergeCells findMergedCells(XSSFSheet sheet) {
        List<CTMergeCell> mergeCellList = getListOfMergeCell(sheet);
        for (CTMergeCell mergeCell : mergeCellList) {
            CellRangeAddress rangeAddress = CellRangeAddress.valueOf(mergeCell.getRef());
            MockCell cell = new MockCell(rangeAddress.getFirstRow(), rangeAddress.getFirstColumn());
            MergeOffset mergeOffset = new MergeOffset(rangeAddress.getLastColumn() - rangeAddress.getFirstColumn(),
                    rangeAddress.getLastRow() - rangeAddress.getFirstRow());
            put(cell, mergeOffset);
        }
        return this;
    }

    private List<CTMergeCell> getListOfMergeCell(XSSFSheet sheet) {
        List<CTMergeCell> mergeCells = new ArrayList<CTMergeCell>();
        if (sheet.getCTWorksheet().getMergeCells() == null) {
            return mergeCells;
        }
        long count = sheet.getCTWorksheet().getMergeCells().getCount();
        for (int i = 0; i < count; i++) {
            mergeCells.add(sheet.getCTWorksheet().getMergeCells().getMergeCellArray(i));
        }
        return mergeCells;
    }

    public MockCell getMergedRegionHeader(MockCell cell) {
        for (Map.Entry<MockCell, MergeOffset> entry : entrySet()) {
            Region region = new Region(entry.getKey().getRow(), (short) entry.getKey().getCol(),
                    entry.getKey().getRow() + entry.getValue().getYOffset(),
                    (short) (entry.getKey().getCol() + entry.getValue().getXOffset()));
            if (region.contains(cell.getRow(), (short) cell.getCol())) {
                return entry.getKey();
            }
        }
        return null;
    }

    public static MergeCells detect(XSSFSheet sheet){
        return new MergeCells().findMergedCells(sheet);
    }
}
