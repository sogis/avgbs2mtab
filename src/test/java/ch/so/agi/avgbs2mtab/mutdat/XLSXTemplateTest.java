package ch.so.agi.avgbs2mtab.mutdat;

import ch.so.agi.avgbs2mtab.writeexcel.ExcelData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class XLSXTemplateTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();

    @Test
    public void WorkbookCreatedOnWritablePath() throws Exception {

        File excelFile = folder.newFile("test.xlsx");

        String filePath = excelFile.getAbsolutePath();
        XLSXTemplate xlsxTemplate = new XLSXTemplate();

        try {
            xlsxTemplate.createWorkbook(filePath);
        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }


    @Test
    public void OldParcelsCorrectlyWrittenToExcel() throws Exception {

        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();


        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();

        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            newWorkbook = xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size());
            newWorkbook = excelData.writeOldParcelsInTemplate(oldParcels, filePath, newWorkbook);
            Assert.assertTrue(checkOldParcels(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void NewParcelsCorrectlyWrittenToExcel() throws Exception {

        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);
            Assert.assertTrue(checkNewParcels(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void OutflowsCorrectlyWrittenToExcel() throws Exception{
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);
            newWorkbook = insertInflowAndOutflows(filePath, newWorkbook, excelData);
            Assert.assertTrue(checkInflowsOutflows(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void InsertRoundingDifferencesCorrectly() throws Exception{
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);

            newWorkbook = insertRoundingDifferences(filePath, newWorkbook, excelData);

            Assert.assertTrue(checkRoundingDifferences(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void CalculateOldAreasCorrectly() throws Exception {
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);



            newWorkbook = excelData.writeOldArea(695,areasOldParcels695(),-1,filePath, newWorkbook);
            newWorkbook = excelData.writeOldArea(696,areasOldParcels696(),0,filePath, newWorkbook);
            newWorkbook = excelData.writeOldArea(697,areasOldParcels697(),-1,filePath, newWorkbook);


            Assert.assertTrue(checkOldAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void CalculateNewAreasCorrectly() throws Exception {
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();


        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);

            newWorkbook = excelData.writeAreaSum(oldAreas(),newAreas(),-1, filePath, newWorkbook);


            Assert.assertTrue(checkSumOfAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void CalculateSumOfAreasCorrectly() throws Exception {
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();


        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,excelData);

            newWorkbook = insertNewAreas(filePath, newWorkbook, excelData);


            Assert.assertTrue(checkNewAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    private List<Integer> generateOldParcels() {
        List<Integer> oldparcels = new ArrayList<>();
        oldparcels.add(695);
        oldparcels.add(696);
        oldparcels.add(697);
        oldparcels.add(701);
        oldparcels.add(870);
        oldparcels.add(874);

        return oldparcels;
    }

    private List <Integer> generateNewParcels() {
        List<Integer> newParcels = new ArrayList<>();
        newParcels.add(695);
        newParcels.add(696);
        newParcels.add(697);
        newParcels.add(701);
        newParcels.add(870);
        newParcels.add(874);
        newParcels.add(4004);

        return newParcels;
    }

    private boolean checkOldParcels(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allOldParcelsAreCorrect = true;

        Row row = xlsxSheet.getRow(2);
        for (Cell cell : row){
            if (cell.getColumnIndex()== 1){
                if(cell.getNumericCellValue()!=695){
                    allOldParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 2) {
                if (cell.getNumericCellValue() != 696) {
                    allOldParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 3) {
                if (cell.getNumericCellValue() != 697) {
                    allOldParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 4) {
                if (cell.getNumericCellValue() != 701) {
                    allOldParcelsAreCorrect = false;
                }
            }  else if (cell.getColumnIndex()== 5) {
                if (cell.getNumericCellValue() != 870) {
                    allOldParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 6) {
                if (cell.getNumericCellValue() != 874) {
                    allOldParcelsAreCorrect = false;
                }
            }
        }
        return allOldParcelsAreCorrect;
    }

    private XSSFWorkbook insertParcels(String filePath, XLSXTemplate xlsxTemplate, ExcelData excelData) {

        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();

        XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
        newWorkbook = xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size());
        newWorkbook = excelData.writeOldParcelsInTemplate(oldParcels, filePath, newWorkbook);
        newWorkbook = excelData.writeNewParcelsInTemplate(newParcels, filePath, newWorkbook);

        return newWorkbook;

    }

    private boolean checkNewParcels(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allNewParcelsAreCorrect = true;

        for (int i = 1; i <= 7; i++){
            Row row = xlsxSheet.getRow(2 + 2*i);
            Cell cell = row.getCell(0);

            if (i == 1) {
                if(cell.getNumericCellValue()!=695){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 2){
                if(cell.getNumericCellValue()!=696){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 3){
                if(cell.getNumericCellValue()!=697){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 4){
                if(cell.getNumericCellValue()!=701){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 5){
                if(cell.getNumericCellValue()!=870){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 6){
                if(cell.getNumericCellValue()!=874){
                    allNewParcelsAreCorrect = false;
                }
            } else if (i == 7){
                if(cell.getNumericCellValue()!=4004){
                    allNewParcelsAreCorrect = false;
                }
            }
        }

        return allNewParcelsAreCorrect;
    }

    private XSSFWorkbook insertInflowAndOutflows(String filePath, XSSFWorkbook newWorkbook, ExcelData excelData) {
        newWorkbook = excelData.writeInflowAndOutflows(695, 695,416,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(696, 696,507,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(697, 697,687,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(696, 701,1,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(697, 701,1,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(701, 701,1112,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(870, 870,611,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(874, 874,1939,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(695, 4004,242,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(696, 4004,100,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(697, 4004,129,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(701, 4004,1,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(870, 4004,39,filePath,newWorkbook);
        newWorkbook = excelData.writeInflowAndOutflows(874, 4004,81,filePath,newWorkbook);

        return newWorkbook;
    }

    private boolean checkInflowsOutflows(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allInflowsAndOutflowsAreCorrect = true;

        for (int i = 1; i <= 7; i++) {
            Row row = xlsxSheet.getRow(2 + 2 * i);
            if (i == 1) {
                Cell cell = row.getCell(1);
                if (cell.getNumericCellValue() != 416) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 2) {
                Cell cell = row.getCell(2);
                if (cell.getNumericCellValue() != 507) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 3) {
                Cell cell = row.getCell(3);
                if (cell.getNumericCellValue() != 687) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 4) {
                if (row.getCell(2).getNumericCellValue() != 1 ||
                        row.getCell(3).getNumericCellValue() != 1 ||
                        row.getCell(4).getNumericCellValue() != 1112) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 5) {
                Cell cell = row.getCell(5);
                if (cell.getNumericCellValue() != 611) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 6) {
                Cell cell = row.getCell(6);
                if (cell.getNumericCellValue() != 1939) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            } else if (i == 7) {
                if (row.getCell(1).getNumericCellValue() != 242 ||
                        row.getCell(2).getNumericCellValue() != 100 ||
                        row.getCell(3).getNumericCellValue() != 129 ||
                        row.getCell(4).getNumericCellValue() != 1 ||
                        row.getCell(5).getNumericCellValue() != 39 ||
                        row.getCell(6).getNumericCellValue() != 81) {
                    allInflowsAndOutflowsAreCorrect = false;
                }
            }
        }

        return allInflowsAndOutflowsAreCorrect;
    }

    private XSSFWorkbook insertRoundingDifferences(String filePath, XSSFWorkbook newWorkbook, ExcelData excelData){
        newWorkbook = excelData.writeRoundingDifference(695, -1, filePath, newWorkbook);
        newWorkbook = excelData.writeRoundingDifference(697, -1, filePath, newWorkbook);
        newWorkbook = excelData.writeRoundingDifference(701, 1, filePath, newWorkbook);
        return newWorkbook;
    }


    private boolean checkRoundingDifferences(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allRoundingDifferencesAreCorrect = true;

        Row row = xlsxSheet.getRow(18);

        if(row.getCell(1).getNumericCellValue() != -1 ||
                row.getCell(3).getNumericCellValue() != -1 ||
                row.getCell(4).getNumericCellValue() != 1){
            allRoundingDifferencesAreCorrect = false;
        }
        return allRoundingDifferencesAreCorrect;
    }

    private List<Integer> areasOldParcels695(){
        List<Integer> areas = new ArrayList<>();
        areas.add(416);
        areas.add(242);
        return areas;
    }

    private List<Integer> areasOldParcels696(){
        List<Integer> areas = new ArrayList<>();
        areas.add(507);
        areas.add(1);
        areas.add(100);
        return areas;
    }
    private List<Integer> areasOldParcels697(){
        List<Integer> areas = new ArrayList<>();
        areas.add(687);
        areas.add(1);
        areas.add(129);
        return areas;
    }

    private boolean checkOldAreas(XSSFWorkbook workbook) {
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allOldAreasAreCorrect = true;

        Row row = xlsxSheet.getRow(19);

        if (row.getCell(1).getNumericCellValue() != 657 ||
                row.getCell(2).getNumericCellValue() != 608 ||
                row.getCell(3).getNumericCellValue() != 816) {
            allOldAreasAreCorrect = false;
        }
        return allOldAreasAreCorrect;
    }

    private XSSFWorkbook insertNewAreas(String filePath, XSSFWorkbook newWorkbook, ExcelData excelData) {


        newWorkbook = excelData.writeNewArea(695, 416, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(696, 507, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(697, 687, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(701, 1114, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(870, 611, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(874, 1939, filePath, newWorkbook);
        newWorkbook = excelData.writeNewArea(4004, 592, filePath, newWorkbook);

        return newWorkbook;
    }

    private boolean checkNewAreas(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allNewAreasAreCorrect = true;

        for (int i = 1; i <= 7; i++){
            Row row = xlsxSheet.getRow(2+2*i);
            if (i == 1) {
                if (row.getCell(7).getNumericCellValue() != 416) {
                    allNewAreasAreCorrect = false;
                }
            } else if (i == 2){
                if (row.getCell(7).getNumericCellValue() != 507) {
                    allNewAreasAreCorrect = false;
                }
            } else if (i == 3){
                if (row.getCell(7).getNumericCellValue() != 687) {
                    allNewAreasAreCorrect = false;
                }
            } else if (i == 4){
                if (row.getCell(7).getNumericCellValue() != 1114) {
                    allNewAreasAreCorrect = false;
                }
            }  else if (i == 5){
                if (row.getCell(7).getNumericCellValue() != 611) {
                    allNewAreasAreCorrect = false;
                }
            }  else if (i == 6){
                if (row.getCell(7).getNumericCellValue() != 1939) {
                    allNewAreasAreCorrect = false;
                }
            }  else if (i == 7){
                if (row.getCell(7).getNumericCellValue() != 592) {
                    allNewAreasAreCorrect = false;
                }
            }
        }

        return allNewAreasAreCorrect;

    }

    private List<Integer> oldAreas() {
        List<Integer> areas = new ArrayList<>();
        areas.add(657);
        areas.add(608);
        areas.add(816);
        areas.add(1114);
        areas.add(650);
        areas.add(2020);

        return areas;
    }

    private List<Integer> newAreas() {
        List<Integer> areas = new ArrayList<>();
        areas.add(416);
        areas.add(507);
        areas.add(687);
        areas.add(1114);
        areas.add(611);
        areas.add(1939);
        areas.add(592);

        return areas;
    }

    private boolean checkSumOfAreas(XSSFWorkbook workbook) {
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean sumOfAreasIsCorrect = true;

        Row row = xlsxSheet.getRow(19);
        Cell cell = row.getCell(7);
        if (cell.getNumericCellValue() != 5865) {
            sumOfAreasIsCorrect = false;
        }

        return sumOfAreasIsCorrect;
    }
}
