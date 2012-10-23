package bigreport;

import bigreport.performers.IterationContext;
import bigreport.xls.CellIterator;
import bigreport.xls.MockCell;
import bigreport.xls.merge.MergeInfo;

import java.io.File;

public class TemplateDescription {
    private String template = "";
    private MockCell cellFrom;
    private MergeInfo mergeInfo;

    public TemplateDescription(IterationContext context, File tempMergeData) {
        this.template = context.getTemplateBuilder().getTemplate();
        this.cellFrom = context.getStartedAt();
        this.mergeInfo = new MergeInfo(tempMergeData, context.getMergedCount());
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
