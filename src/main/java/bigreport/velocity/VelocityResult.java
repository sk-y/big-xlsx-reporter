package bigreport.velocity;

import bigreport.xls.merge.MergeInfo;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;

public class VelocityResult {
    private File dataXmlFile;
    private MergeInfo mergeInfo;

    public VelocityResult(File dataXmlFile, MergeInfo mergeInfo) {
        this.dataXmlFile = dataXmlFile;
        this.mergeInfo = mergeInfo;
    }

    public File getDataXmlFile() {
        return dataXmlFile;
    }

    public MergeInfo getMergeInfo() {
        return mergeInfo;
    }

    public void deleteAllTempFiles() throws IOException {
        FileUtils.forceDelete(dataXmlFile);
        if (mergeInfo != null) {
            mergeInfo.reset();
        }
    }
}
