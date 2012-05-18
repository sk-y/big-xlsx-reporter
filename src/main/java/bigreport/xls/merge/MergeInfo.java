package bigreport.xls.merge;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;

public class MergeInfo {
    private File mergedTempFile;
    private int count=0;

    public MergeInfo(File mergedTempFile, int count) {
        this.mergedTempFile = mergedTempFile;
        this.count = count;
    }

    public File getMergedTempFile() {
        return mergedTempFile;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public void reset() throws IOException {
        if (mergedTempFile!=null){
            FileUtils.forceDelete(mergedTempFile);
            count=0;
        }
    }
}
