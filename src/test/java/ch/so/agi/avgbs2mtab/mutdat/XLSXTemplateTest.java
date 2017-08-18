package ch.so.agi.avgbs2mtab.mutdat;

import ch.so.agi.avgbs2mtab.writeexcel.ExcelData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class XLSXTemplateTest {

    @Test
    public void WorkbookCreatedOnWritablePath() throws Exception {

        String filePath = "/home/barpastu/Documents/test.xlsx";
        XLSXTemplate xlsxTemplate = new XLSXTemplate();

        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void WorkbookWithColoredCellCreated() throws Exception {
        String filePath = "/home/barpastu/Documents/test.xlsx";
        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();
        List<Integer> oldparcels = new ArrayList<>();
        oldparcels.add(597);
        oldparcels.add(40011);
        oldparcels.add(40012);



        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            newWorkbook = xlsxTemplate.createParcelTable(newWorkbook,filePath, 3,3);
            newWorkbook = excelData.writeOldParcelsInTemplate(oldparcels, filePath, newWorkbook);
            newWorkbook = excelData.writeNewParcelsInTemplate(oldparcels,filePath, newWorkbook);
            newWorkbook = excelData.writeInflowAndOutflows(597, 40011,636,filePath,newWorkbook);
            newWorkbook = excelData.writeNewArea(40011, 638, filePath,newWorkbook);
            newWorkbook = excelData.writeRoundingDifference(597, -1, filePath, newWorkbook);
            newWorkbook = excelData.writeOldArea(597,oldparcels,-1,filePath, newWorkbook);

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }
}
