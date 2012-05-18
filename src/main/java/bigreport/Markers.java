package bigreport;

public class Markers {
    private Markers() {
    }

    public static final String END_OPERATOR="#end";
    public static final String DATA = "<rm:start>";
    public static final String DATA_SYNONYM = "<rm:data>";
    public static final String END = "<rm:end>";
    public static final String END_SYNONYM = "</rm:data>";
    public static final String END_DATA="</sheetData>";
    public static final String START_DIRECTIVE="<rm:#";
    public static final String CONDITION_OPERATOR="#if";
    public static final String ELSE_OPERATOR="#{else}";
    public static final String START_CELL_CONDITION_DIRECTIVE="<rm:showIf>";
    public static final String END_CELL_CONDITION_DIRECTIVE="</rm:showIf>";
    public static final String END_CELL_CONDITION="<rm:then>";
    public static final String END_DIRECTIVE=">";
    public static final String START_MERGED_CELLS="<mergeCells ";
    public static final String END_MERGED_CELLS="</mergeCells>";
    public static final String OUT_LINE_START="<rm:outline>";
    public static final String OUT_LINE_END="</rm:outline>";
}
