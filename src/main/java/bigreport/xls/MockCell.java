package bigreport.xls;

import org.apache.poi.ss.usermodel.Cell;

public class MockCell {
    private int row;
    private int col;

    public MockCell(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public MockCell(Cell cell){
        this.row=cell.getRowIndex();
        this.col=cell.getColumnIndex();
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof MockCell)) return false;
        MockCell mockCell = (MockCell) o;
        if (col != mockCell.col) return false;
        return (row == mockCell.row);
    }

    @Override
    public int hashCode() {
        int result = row;
        result = 31 * result + col;
        return result;
    }
}
