package bigreport.velocity;

import bigreport.TemplateDescription;
import bigreport.velocity.beans.ResolverBean;
import bigreport.xls.merge.MergeInfo;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.directive.Directive;
import org.apache.velocity.runtime.resource.loader.StringResourceLoader;
import org.apache.velocity.runtime.resource.util.StringResourceRepository;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class VelocityResolver {
    private Map beans;

    private List<Class<? extends Directive>> directives;

    public VelocityResolver(Map beans) {
        this(beans, (List<Class<? extends Directive>>) null);
    }

    public VelocityResolver(Map beans, List<Class<? extends Directive>> directives) {
        this.beans = beans;
        this.directives = directives;
    }

    public void resolve(TemplateDescription templateDescription, Writer writer, File mergedCellsTempFile) throws IOException {
        VelocityEngine ve = new VelocityEngine();
        setupVelocityProperties(ve);
        StringResourceRepository repo = StringResourceLoader.getRepository();
        String templateName = "testTemplate";
        repo.putStringResource(templateName, templateDescription.getTemplate());
        Template template = ve.getTemplate(templateName,"UTF-8");
        VelocityContext context = new VelocityContext(beans);
        FileOutputStream fileOutputStream = null;
        ResolverBean resolverBean = new ResolverBean(templateDescription.getCellFrom().getRow() - 1);
        try {
            fileOutputStream = new FileOutputStream(mergedCellsTempFile, true);
            resolverBean.setMergedCellsOutputStream(fileOutputStream);
            context.put("resolver", resolverBean);
            context.put("dblqt","\"");
            template.merge(context, writer);
            writer.flush();
            fileOutputStream.flush();
        } catch (IOException e) {
            throw e;
        } finally {
            try {
                closeStream(fileOutputStream);
            } finally {
                closeWriter(writer);
            }
        }
        MergeInfo mergeInfo = templateDescription.getMergeInfo();
        mergeInfo.setCount(mergeInfo.getCount() + resolverBean.getMergedRegionCount());
    }

    public String getDirectivesString() {
        if (directives == null) {
            return "";
        }
        StringBuilder resultBuilder = new StringBuilder("");
        for (Class<? extends Directive> directiveClass : directives) {
            if (resultBuilder.length() != 0) {
                resultBuilder.append(", ");
            }
            resultBuilder.append(directiveClass.getName());
        }
        return resultBuilder.toString();
    }

    public void addDirective(Class<? extends Directive> directiveClass) {
        if (directives == null) {
            directives = new ArrayList<Class<? extends Directive>>();
        }
        if (!directives.contains(directiveClass)) {
            directives.add(directiveClass);
        }
    }

    private void closeStream(OutputStream outputStream) throws IOException {
        if (outputStream != null) {
            outputStream.close();
        }
    }

    private void closeWriter(Writer writer) throws IOException {
        if (writer != null) {
            writer.close();
        }
    }

    private void setupVelocityProperties(VelocityEngine ve) {
        ve.addProperty("resource.loader", "string");
        ve.addProperty("string.resource.loader.description", "Velocity StringResource loader");
        ve.addProperty("string.resource.loader.class", "org.apache.velocity.runtime.resource.loader.StringResourceLoader");
        ve.addProperty("string.resource.loader.repository.class", "org.apache.velocity.runtime.resource.util.StringResourceRepositoryImpl");
        if (directives != null && !directives.isEmpty()) {
            ve.setProperty("userdirective", getDirectivesString());
        }
        ve.init();
    }
}
