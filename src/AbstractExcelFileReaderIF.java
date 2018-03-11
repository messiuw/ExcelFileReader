import java.io.File;
import java.util.List;
import java.util.Map;

abstract class AbstractExcelFileReader implements FileReaderIF {

    @Override
    public File loadFile(String p_path) {
        File file = null;
        if (p_path != null) {
            file = new File(p_path);
        }
        return file;
    }

    protected abstract Map<String, List<String>> readExcelsheet(final File p_file, final int sheet);

    protected abstract Map<String, Map<String,List<String>>> readExcelFile(final File p_file);

    protected abstract List<String> returnHeaderTitles(final int sheet) throws Exception;

    protected abstract Map<String, List<String>> readExcelsheetOnlyInArea(final File p_file, final int sheet,
                                                                          final int rowStart, final int rowEnd,
                                                                          final int colStart, final int colEnd) throws Exception;

    protected abstract int numberOfSheetsInFile(final File p_file);

    protected abstract int numberOfRowsInSheet(final int sheet);

}
