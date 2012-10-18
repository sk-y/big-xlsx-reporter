package bigreport.exception;

import java.io.File;
import java.io.IOException;

public class FileDeleteException extends IOException {
    private File file;

    public FileDeleteException(File file) {
        super();
        this.file=file;
    }

    public FileDeleteException(String message, File file) {
        super(message);
        this.file=file;
    }

    public FileDeleteException(String message, Throwable cause, File file) {
        super(message, cause);
        this.file=file;
    }

    public FileDeleteException(Throwable cause, File file) {
        super(cause);
        this.file=file;
    }

    public File getFile() {
        return file;
    }
}
