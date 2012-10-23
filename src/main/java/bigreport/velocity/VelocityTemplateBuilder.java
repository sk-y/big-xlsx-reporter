package bigreport.velocity;

import bigreport.Markers;
import bigreport.xls.merge.MergeOffset;

public class VelocityTemplateBuilder {
    private StringBuilder templateBuilder;

    public VelocityTemplateBuilder reset() {
        templateBuilder = new StringBuilder();
        return this;
    }

    public String getTemplate() {
        return templateBuilder.toString();
    }

    public VelocityTemplateBuilder addCell(String value, int column, short styleIndex) {
        value = performQuotes(value);
        templateBuilder.append("$resolver.addCell(")
                .append("$resolver.nvl(")
                .append(value)
                .append("),")
                .append(column)
                .append(",")
                .append(styleIndex)
                .append(")");
        return this;
    }

    private String performQuotes(String value) {
        if (value.contains("\"")){
            value = "\""+value.replaceAll("\"", "\\$\\{dblqt\\}")+"\"";
        }
        return value;
    }

    public VelocityTemplateBuilder addCell(String value, int column, short styleIndex, MergeOffset mergeOffset) {
        value = performQuotes(value);
        templateBuilder.append("$resolver.addCell(")
                .append("$resolver.nvl(")
                .append(value)
                .append("),")
                .append(column)
                .append(",")
                .append(styleIndex)
                .append(",")
                .append(mergeOffset.getXOffset())
                .append(",")
                .append(mergeOffset.getYOffset())
                .append(")");
        return this;
    }

    public VelocityTemplateBuilder addRow(long styleIndex, int outlineLevel) {
        templateBuilder.append("$resolver.addRow(")
                .append(styleIndex);
        if (outlineLevel > 0) {
            templateBuilder.append(",").append(outlineLevel);
        }
        templateBuilder.append(")");
        return this;
    }

    public VelocityTemplateBuilder addEndRow() {
        templateBuilder.append("$resolver.endRow()");
        return this;
    }

    //N.B. All directives must be placed into cell with 1 row height
    public VelocityTemplateBuilder addDirective(String value) {
        templateBuilder.append("#").append(value).append(" ");
        return this;
    }

    public VelocityTemplateBuilder addCell(String value, int columnIndex, short index, MergeOffset offset, String condition) {
        templateBuilder.append(Markers.CONDITION_OPERATOR)
                .append("(")
                .append(condition)
                .append(")");
        addCell(value, columnIndex, index,offset);
        templateBuilder.append(Markers.ELSE_OPERATOR)
                .append("$resolver.skipColumn()")
                .append(Markers.END_OPERATOR);
        return this;
    }

    public VelocityTemplateBuilder addCell(String value, int columnIndex, short index, String condition) {
        templateBuilder.append(Markers.CONDITION_OPERATOR)
                .append("(")
                .append(condition)
                .append(")");
        addCell(value, columnIndex, index);
        templateBuilder.append(Markers.ELSE_OPERATOR)
                .append("$resolver.skipColumn()")
                .append(Markers.END_OPERATOR);
        return this;
    }
}
