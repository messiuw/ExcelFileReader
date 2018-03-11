import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.Map;

public class XLSXFileReader extends AbstractExcelFileReader implements FileReaderIF {

    XSSFWorkbook workbook;

    XLSXFileReader(final File p_file) {
        initializeWorkbookFromFile(p_file);
    }

    XLSXFileReader() {}

    @Override
    protected Map<String, String> readExcelFile(final File p_file) {
        readFile(p_file);
        return null;
    }

    @Override
    protected int numberOfSheetsInFile(File p_file) {
        return 0;
    }

    @Override
    protected List<String> returnHeaderTitles(File p_file, String p_sheet) {
        return null;
    }

    @Override
    protected int numberOfRowsInSheet(File p_file, String p_sheet) {
        return 0;
    }

    @Override
    protected Map<String, String> readExcelsheet(File p_file, String p_sheet) {
        return null;
    }

    protected List<String> returnHeaderTitles(File p_file, int p_sheet) {
        return null;
    }

    protected int numberOfRowsInSheet(File p_file, int p_sheet) {
        return 0;
    }

    protected Map<String, String> readExcelsheet(File p_file, int p_sheet) {
        return null;
    }

    private void readFile(final File p_file) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(p_file));

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void initializeWorkbookFromFile(final File p_file) {
        try {
            this.workbook = new XSSFWorkbook(new FileInputStream(p_file));
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

}
