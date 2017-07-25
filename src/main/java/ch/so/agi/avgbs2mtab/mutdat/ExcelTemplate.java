package ch.so.agi.avgbs2mtab.mutdat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public interface ExcelTemplate {

    public Workbook createWorkbook(String filePath);

    public Workbook createParcelTable(Workbook excelTemplate, int newParcels, int oldParcels);

    public Workbook createDPRTable(Workbook excelTemplate, int parcels, int dpr);
}
