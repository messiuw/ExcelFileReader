import org.apache.commons.io.FilenameUtils;

import java.io.File;

public class ExcelFileReaderFactory {

    private static final XLSXFileReader xlsxFileReader = new XLSXFileReader();
    private static final XLSFileReader xlsFileReader = new XLSFileReader();

    /**
     * This method is used to return the only instance of this factory method. Based on the Fileextension of @param{p_File}
     * the correct excel file reader instance will be returned.
     * @param p_File the Excelfile to be read
     * @return the only instance of @method{XLSXFileReaderIF} or @method{XLSFileReaderIF}. If the filetype does not match
     * <it>null</it> is returned.
     */
    public static FileReaderIF getExcelFileReader(File p_File) {
        if (p_File == null) {
            return null;
        }

        String fileExtension = FilenameUtils.getExtension(p_File.toString());

        if (fileExtension.equalsIgnoreCase("xlsx")) {
            return xlsxFileReader;
        } else if (fileExtension.equalsIgnoreCase("xls")) {
            return xlsFileReader;
        } else {
            return null;
        }
    }

    /**
     *
     * @param p_File the Excelfile to be read
     * @return a <ul>new</ul> instace of @method{XLSXFileReaderIF} or @method{XLSFileReaderIF}. If the filetype does not match
     * <it>null</it> is returned.
     */
    public static FileReaderIF getNewInstanceOfExcelFileReader(File p_File) {
        if (p_File == null) {
            return null;
        }

        String fileExtension = FilenameUtils.getExtension(p_File.toString());

        if (fileExtension.equalsIgnoreCase("xlsx")) {
            return new XLSXFileReader(p_File);
        } else if (fileExtension.equalsIgnoreCase("xls")) {
            return new XLSFileReader(p_File);
        } else {
            return null;
        }
    }
}
