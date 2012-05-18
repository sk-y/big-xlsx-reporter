package bigreport.util;

import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeOffset;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.math.BigDecimal;

public class ValueResolver {
    private ValueResolver() {
    }

    public static String resolve(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
        }
        return "";
    }

    public static String convertStringForXml(String value) {
        return value.replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll("\"", "&quot;")
                .replaceAll("'", "&apos;");
    }

    public static String getCellAddress(int row, int col) {
        return new CellReference(row, col).formatAsString().replaceAll("\\$", "");
    }

    public static String getNumberCellAsString(String value, int rowIndex, int columnIndex, int styleIndex) {
        String strValue = ValueResolver.convertStringForXml(value);
        StringBuilder sb = new StringBuilder("");
        String ref = ValueResolver.getCellAddress(rowIndex, columnIndex);
        sb.append("<c r=\"").append(ref).append("\" t=\"n\"");
        if (styleIndex != -1) {
            sb.append(" s=\"").append(styleIndex).append("\"");
        }
        sb.append(">").append("<v>").append(strValue).append("</v>")
                .append("</c>");
        return sb.toString();
    }

    public static String getCellAsString(Number value, int rowIndex, int columnIndex, int styleIndex) {
        String val = (value == null) ? "" : value.toString();
        return getNumberCellAsString(val, rowIndex, columnIndex, styleIndex);
    }

    public static String getCellAsString(String value, int rowIndex, int columnIndex, int styleIndex) {
        String strValue = ValueResolver.convertStringForXml(value);
        StringBuilder sb = new StringBuilder("");
        String ref = ValueResolver.getCellAddress(rowIndex, columnIndex);
        sb.append("<c r=\"").append(ref).append("\"");
        if (!value.isEmpty()) {
            sb.append(" t=\"inlineStr\"");
        }
        if (styleIndex != -1) {
            sb.append(" s=\"").append(styleIndex).append("\"");
        }
        sb.append(">");
        if (!value.isEmpty()) {
            sb.append("<is><t>").append(strValue).append("</t></is>");
        }
        sb.append("</c>");
        return sb.toString();
    }

    public static String getRowAsString(int rowIndex, int styleIndex) {
        StringBuilder sb = getRowBuilder(rowIndex, styleIndex);
        sb.append(">\n");
        return sb.toString();
    }

    public static String getEndOfRowAsString() {
        return "</row>\n";
    }

    public static String getMergedCellAsString(int row, int yOffset, int col, int xOffset) throws IOException {
        StringBuilder sb = new StringBuilder();
        String range = new CellRangeAddress(row, row + yOffset, col, col + xOffset).formatAsString();
        sb.append("<mergeCell ref=\"").append(range).append("\"/>");
        return sb.toString();
    }

    public static String getMergedCellAsString(MockCell cell, MergeOffset offset) throws IOException {
        return getMergedCellAsString(cell.getRow(), offset.getYOffset(), cell.getCol(), offset.getXOffset());
    }

    public static String getRowAsString(int rowIndex, int styleIndex, int outlineLevel) {
        StringBuilder sb = getRowBuilder(rowIndex, styleIndex);
        sb.append(" outlineLevel=\"").append(outlineLevel).append("\"").append(">\n");
        return sb.toString();
    }

    private static StringBuilder getRowBuilder(int rowIndex, int styleIndex) {
        StringBuilder sb = new StringBuilder("");
        sb.append("<row r=\"").append(rowIndex).append("\"");
        if (styleIndex != -1) {
            sb.append(" s=\"").append(styleIndex).append("\"");
        }
        return sb;
    }

    public static boolean isNumeric(String value) {
        try {
            new BigDecimal(value);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public static boolean isString(String value) {
        return !isNumeric(value);
    }

    public static boolean isVariable(String value) {
        if (value == null) {
            return false;
        }
        return value.startsWith("$");
    }
}
