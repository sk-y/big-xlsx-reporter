package bigreport.xls.merge;

public class MergeOffset{

    private int xOffset;
    private int yOffset;

    public MergeOffset(int xOffset, int yOffset) {
        this.xOffset = xOffset;
        this.yOffset = yOffset;
    }

    public int getXOffset() {
        return xOffset;
    }

    public int getYOffset() {
        return yOffset;
    }
}