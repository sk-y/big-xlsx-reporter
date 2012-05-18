package bigreport;

import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeInfo;

public class TemplateDescription {
    private String template = "";
    private MockCell cellFrom;
    private MergeInfo mergeInfo;

    public TemplateDescription(String template, CellIterator cellIterator, MergeInfo mergeInfo) {
        this.template = template;
        this.cellFrom = cellIterator.getStartedAt();
        this.mergeInfo = mergeInfo;
    }

    public String getTemplate() {
        return template;
    }

    public MockCell getCellFrom() {
        return cellFrom;
    }

    public boolean isEmpty() {
        return template == null || template.isEmpty();
    }

    public MergeInfo getMergeInfo() {
        return mergeInfo;
    }

}
