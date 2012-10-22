package bigreport.velocity;

import bigreport.TemplateDescription;
import bigreport.exception.CompositeIOException;
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
    private final static String TEMPLATE_NAME = "testTemplate";
    private final static String DEFAULT_ENCODING = "UTF-8";
    private final static String RESLOVER_NEME = "resolver";
    private final static String QUOTE_OBJECT_NAME = "dblqt";

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
        Template template = createTemplate(templateDescription, ve);
        VelocityContext context = new VelocityContext(beans);
        ResolverBean resolverBean = new ResolverBean(templateDescription.getCellFrom().getRow() - 1);
        CompositeIOException compositeIOException = null;
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(mergedCellsTempFile, true);
            resolverBean.setMergedCellsOutputStream(fileOutputStream);
            initContext(context, resolverBean);
            template.merge(context, writer);
            writer.flush();
            fileOutputStream.flush();
        } catch (IOException e) {
            compositeIOException = new CompositeIOException(e);
            throw compositeIOException;
        } finally {
            try {
                close(fileOutputStream, compositeIOException);
            } finally {
                close(writer, compositeIOException);
            }
        }
        MergeInfo mergeInfo = templateDescription.getMergeInfo();
        mergeInfo.setCount(mergeInfo.getCount() + resolverBean.getMergedRegionCount());
    }

    private void initContext(VelocityContext context, ResolverBean resolverBean) {
        context.put(RESLOVER_NEME, resolverBean);
        context.put(QUOTE_OBJECT_NAME, "\"");
    }

    private Template createTemplate(TemplateDescription templateDescription, VelocityEngine ve) {
        StringResourceRepository repo = StringResourceLoader.getRepository();
        repo.putStringResource(TEMPLATE_NAME, templateDescription.getTemplate());
        return ve.getTemplate(TEMPLATE_NAME, DEFAULT_ENCODING);
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

    private void close(Closeable outputStream, CompositeIOException compositeException) throws IOException {
        if (outputStream == null) {
            return;
        }
        try {
            outputStream.close();
        } catch (IOException e) {
            if (compositeException != null) {
                compositeException.addSuppressedException(e);
                return;
            }
            throw e;
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
