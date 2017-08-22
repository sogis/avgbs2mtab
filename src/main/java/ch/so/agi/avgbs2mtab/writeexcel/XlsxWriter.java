package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.XLSXTemplate;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxWriter {

    public void writeXlsx(String filePath) {

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        XSSFWorkbook workbook = xlsxTemplate.createExcelTemplate(filePath);
        workbook = excelData.fillValuesIntoParcelTable(filePath, workbook);
        excelData.fillValuesIntoDPRTable(filePath, workbook);
    }
}
