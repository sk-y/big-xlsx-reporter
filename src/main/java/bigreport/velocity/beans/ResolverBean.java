package bigreport.velocity.beans;

import bigreport.util.StreamUtil;
import bigreport.util.ValueResolver;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;

public class ResolverBean {
    private int rowOffset = 0;
    private int rowNum = 0;
    private OutputStream mergedCellsOutputStream;
    private int mergedRegionCount = 0;
    private int skippedColumns = 0;

    public ResolverBean(int rowOffset, OutputStream mergedCellsOutputStream) {
        this.rowOffset = (short) rowOffset;
        this.mergedCellsOutputStream = mergedCellsOutputStream;
        mergedRegionCount = 0;
    }

    public ResolverBean(int rowOffset) {
        this.rowOffset = (short) rowOffset;
        mergedRegionCount = 0;
    }

    public String skipColumn() {
        skippedColumns++;
        return "";
    }

    private int getCurrentRow() {
        return rowNum + rowOffset;
    }

    public String addRow(int styleIndex) {
        rowNum++;
        skippedColumns = 0;
        return ValueResolver.getRowAsString(getCurrentRow() + 1, styleIndex);
    }

    public String addRow(int styleIndex, int outlineLevel) {
        rowNum++;
        skippedColumns = 0;
        return ValueResolver.getRowAsString(getCurrentRow() + 1, styleIndex, outlineLevel);
    }

    public String endRow() {
        return ValueResolver.getEndOfRowAsString();
    }

    public String addCell(String value, int columnIndex, int styleIndex, int xOffset, int yOffset) throws IOException {
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, styleIndex);
    }

    public String addCell(String value, int columnIndex, int styleIndex) {
        if (ValueResolver.isNumeric(value)) {
            return ValueResolver.getNumberCellAsString(value, getCurrentRow(), columnIndex - skippedColumns, styleIndex);
        }
        return ValueResolver.getCellAsString(value, getCurrentRow(), columnIndex - skippedColumns, styleIndex);
    }

    public String addCell(String value, int columnIndex) {
        return addCell(value, columnIndex, -1);
    }

    public String addCell(String value, int columnIndex, int xOffset, int yOffset) throws IOException {
        if (ValueResolver.isNumeric(value)) {

        }
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, -1, xOffset, yOffset);
    }

    public String addCell(Number value, int columnIndex, int styleIndex) {
        return ValueResolver.getCellAsString(value, getCurrentRow(), columnIndex - skippedColumns, styleIndex);
    }

    public String addCell(Number value, int columnIndex, int styleIndex, int xOffset, int yOffset) throws IOException {
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, styleIndex);
    }

    private void addMergedCells(int col, int xOffset, int yOffset) throws IOException {
        if (mergedCellsOutputStream != null) {
            mergedRegionCount++;
            StreamUtil.writeMergedCell(mergedCellsOutputStream, getCurrentRow(), yOffset, col - skippedColumns, xOffset);
        }
    }
/*
    public String addCell(double value, int columnIndex) {
        return addCell(value, columnIndex, -1);
    }

    public String addCell(double value, int columnIndex, int styleIndex, int xOffset, int yOffset) throws IOException {
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, styleIndex);
    }

    public String addCell(int value, int columnIndex) {
        return addCell(value, columnIndex, -1);
    }

    public String addCell(int value, int columnIndex, int styleIndex, int xOffset, int yOffset) throws IOException {
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, styleIndex);
    }*/


    public String addCell(Calendar value, int columnIndex, int styleIndex) {
        return addCell(DateUtil.getExcelDate(value, false), columnIndex, styleIndex);
    }

    public String addCell(Calendar value, int columnIndex, int styleIndex, int xOffset, int yOffset) throws IOException {
        addMergedCells(columnIndex, xOffset, yOffset);
        return addCell(value, columnIndex, styleIndex);
    }

    public int getMergedRegionCount() {
        return mergedRegionCount;
    }

    public void setMergedCellsOutputStream(OutputStream mergedCellsOutputStream) {
        this.mergedCellsOutputStream = mergedCellsOutputStream;
    }
}



