//todo: make usage of file unnecessary, use workbook instead

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class XLSFileReader extends AbstractExcelFileReader implements FileReaderIF {

    private HSSFWorkbook workbook;
    private int headerRows;

    XLSFileReader(final File p_file) {
        initializeWorkbookFromFile(p_file);
    }

    XLSFileReader() {}

    public void readFileToWorkbook(final File p_file) {
        initializeWorkbookFromFile(p_file);
    }

    /**
     * Use this method to read a single excel sheet
     * @param p_file The Excelfile
     * @param sheetNum The sheetnumber which shall be read
     * @return A Map<String, List<String>> object. The key is the row number, with the value being a list of
     * values
     */
    @Override
    protected Map<String, List<String>> readExcelsheet(final File p_file, final int sheetNum) {
        return readSingleSheet(p_file, sheetNum);
    }

    @Override
    protected Map<String, Map<String,List<String>>> readExcelFile(final File p_file) {
        Map<String, Map<String,List<String>>> returnData = new HashMap<>();
        int numberOfSheets = determineNumberOfSheets(p_file);
        if (numberOfSheets > 0) {
            String sheetName;
            for (int sheetNum=0; sheetNum<numberOfSheets; sheetNum++) {
                sheetName = determineSheetName(p_file,sheetNum);
                returnData.put(sheetName, readSingleSheet(p_file,sheetNum));
            }
        }
        return returnData;
    }

    @Override
    protected List<String> returnHeaderTitles(final int sheetNum) throws Exception {
        if (workbook == null) {
            throw new Exception("Workbook has not been initialized. Use readFileToWorkbook() first!");
        }

        if (headerRows != 0) {
            this.headerRows = 0;
        }

        HSSFSheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(0);
        Iterator<Cell> cellIterator = row.cellIterator();

        String cellStringValue;
        List<String> returnData = new ArrayList<>();
        while (cellIterator.hasNext()) {
            Cell currentCell = cellIterator.next();
            if (!currentCell.getStringCellValue().isEmpty()) {
                cellStringValue = currentCell.getStringCellValue();
                returnData.add(cellStringValue);
                headerRows++;
            }
        }
        return returnData;
    }

    @Override
    protected Map<String, List<String>> readExcelsheetOnlyInArea(final File p_file, int sheetNum, int rowStart, int rowEnd,
                                                                 int colStart, int colEnd) throws Exception {
        if (this.workbook.getBytes().length == 0 ) {
            initializeWorkbookFromFile(p_file);
        }
        if (sheetNum == 0) {
            throw new Exception("No sheetnumber given to read from!");
        }

        HSSFSheet sheet = this.workbook.getSheetAt(sheetNum);
        Map<String, List<String>> returnData = new HashMap<>();
        for (int row=rowStart; row<rowEnd; row++) {
            Row rowInSheet = sheet.getRow(row);
            Iterator<Cell> cellIterator = rowInSheet.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = cell.getStringCellValue();
                if(!cellValue.isEmpty()) {
                    String rowNum = String.valueOf(rowInSheet.getRowNum());
                    if (returnData.containsKey(rowNum)) {
                        returnData.get(rowNum).add(cellValue);
                    } else {
                        List<String> list = new ArrayList<>();
                        list.add(cellValue);
                        returnData.put(rowNum,list);
                    }
                }
            }
        }
        return returnData;
    }

    @Override
    protected int numberOfSheetsInFile(final File p_file) {
        return determineNumberOfSheets(p_file);
    }

    @Override
    protected int numberOfRowsInSheet(final int sheetNum) {
        HSSFSheet sheet = workbook.getSheetAt(sheetNum);
        return sheet.getPhysicalNumberOfRows();
    }

    //=========================================================================================//

    private Object checkValue(final String p_string) {
        try {
            return Integer.parseInt(p_string);
        } catch (Exception ex) {
            return p_string;
        }
    }

    private Map<String, List<String>> readSingleSheet(final File p_file, final int p_sheetNum) {
        Map<String, List<String>> returnData = new HashMap<>();
        try {
            HSSFSheet sheet = workbook.getSheetAt(p_sheetNum);

            for (Row row : sheet) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (!cell.getStringCellValue().isEmpty()) {
                        String rowNum = String.valueOf(row.getRowNum());
                        if (returnData.containsKey(rowNum)) {
                            returnData.get(rowNum).add(cell.getStringCellValue());
                        } else {
                            List<String> list = new ArrayList<>();
                            list.add(cell.getStringCellValue());
                            returnData.put(rowNum,list);
                        }
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return returnData;
    }

    private int determineNumberOfSheets(final File p_file) {
        int maxSheets = 0;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(p_file));
            maxSheets = workbook.getNumberOfSheets();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return maxSheets;
    }

    private String determineSheetName(final File p_file, final int p_sheetNum) {
        String sheetName = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(p_file));
            Sheet sheet = workbook.getSheetAt(p_sheetNum);
            sheetName = sheet.getSheetName();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return sheetName;
    }

    private void initializeWorkbookFromFile(final File p_file) {
        try {
            this.workbook = new HSSFWorkbook(new FileInputStream(p_file));
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

}
