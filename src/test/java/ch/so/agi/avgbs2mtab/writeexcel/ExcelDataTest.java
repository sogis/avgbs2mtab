package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.XLSXTemplate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/*
todo: Bitte lass dich nicht von den notierten todos Fehlleiten - Die Tests in dieser Klasse zeigen mir
dass du über hervorragende logisch-analytische Denkfähigkeit verfügst und diese beim Programmieren anwenden kannst
SEHR GUTE ARBEIT!!!!!!!!!
 */
public class ExcelDataTest {

    @Rule
    public TemporaryFolder parcelTableWriter = new TemporaryFolder();

    //todo sprechender wenn's geht. Welcher Fall ist abgedeckt?
    @Test
    public void oldParcelsCorrectlyWrittenToExcel() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();


        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();

        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            XSSFSheet sheet = newWorkbook.getSheet("Mutationstabelle");
            xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size(), 0);
            parcelTableWriter.writeOldParcelsInTemplate(oldParcels, sheet);
            Assert.assertTrue(checkOldParcels(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    //todo sprechender wenn's geht. Welcher Fall ist abgedeckt?
    @Test
    public void newParcelsCorrectlyWrittenToExcel() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);
            Assert.assertTrue(checkNewParcels(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void outflowsCorrectlyWrittenToExcel() throws Exception{
        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();


        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            insertInflowAndOutflows(parcelTableWriter, xlsxSheet);
            Assert.assertTrue(checkInflowsOutflows(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void insertRoundingDifferencesCorrectly() throws Exception{
        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);

            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");

            insertRoundingDifferences(parcelTableWriter, xlsxSheet);

            Assert.assertTrue(checkRoundingDifferences(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void calculateOldAreasCorrectly() throws Exception {
        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();

        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);

            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");

            parcelTableWriter.writeOldArea(695,657, 7, xlsxSheet);
            parcelTableWriter.writeOldArea(696,608, 7, xlsxSheet);
            parcelTableWriter.writeOldArea(697,816, 7, xlsxSheet);


            Assert.assertTrue(checkOldAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void calculateNewAreasCorrectly() throws Exception {
        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();


        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            parcelTableWriter.writeAreaSum(oldAreas(),newAreas(),-1, xlsxSheet);


            Assert.assertTrue(checkSumOfAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    @Test
    public void calculateSumOfAreasCorrectly() throws Exception {
        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ParcelTableWriter parcelTableWriter = new ParcelTableWriter();


        try {
            XSSFWorkbook newWorkbook = insertParcels(filePath,xlsxTemplate,parcelTableWriter);

            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");

            insertNewAreas(parcelTableWriter, xlsxSheet);


            Assert.assertTrue(checkNewAreas(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }

    //todo sprechender wenn's geht: Sowohl die alten und neuen oder nur die alten oder nur die neuen?
    @Test
    public void writeParcelsCorrectlyInDPRTable() throws Exception{

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        DPRTableWriter dprTableWriter = new DPRTableWriter();

        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();
        List<Integer> parcels = generateParcels();
        List<Integer> dpr = generateDPR();

        try {


            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size(),
                    parcels.size());
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            xlsxTemplate.createDPRTable(newWorkbook, filePath, parcels.size(), dpr.size(),
                    newParcels.size(), oldParcels.size());
            dprTableWriter.writeParcelsAffectedByDPRsInTemplate(parcels, newParcels.size(), xlsxSheet);

            Assert.assertTrue(checkParcels(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }


    @Test
    public void writeAllDPRsCorrectlyInDPRTable() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();


        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        DPRTableWriter dprTableWriter = new DPRTableWriter();

        try {

            XSSFWorkbook newWorkbook = insertDPR(filePath,xlsxTemplate, dprTableWriter);

            Assert.assertTrue(checkDPRs(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }


    }

    @Test
    public void writeAllFlowsCorrectlyInDPRTable() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();


        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        DPRTableWriter dprTableWriter = new DPRTableWriter();

        try {

            XSSFWorkbook newWorkbook = insertDPR(filePath,xlsxTemplate, dprTableWriter);
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            insertDPRFlows(dprTableWriter, xlsxSheet);

            Assert.assertTrue(checkFlows(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    //todo: Brauchts noch einen Test "writeOldAreasCorrectlyInDPRTable()"
    @Test
    public void writeNewAreasCorrectlyInDPRTable() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();


        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        DPRTableWriter dprTableWriter = new DPRTableWriter();

        try {

            XSSFWorkbook newWorkbook = insertDPR(filePath,xlsxTemplate, dprTableWriter);
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            insertDPRFlows(dprTableWriter, xlsxSheet);
            insertNewDPRAreas(dprTableWriter, xlsxSheet);


            Assert.assertTrue(checkDPRNewArea(newWorkbook));

        } catch (Exception e){
            throw new RuntimeException(e);
        }


    }

    @Test
    public void writeRoundingDifferencesCorrectlyInDPRTable() throws Exception {

        File excelFile = parcelTableWriter.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();


        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        DPRTableWriter dprTableWriter = new DPRTableWriter();

        try {

            XSSFWorkbook newWorkbook = insertDPR(filePath,xlsxTemplate, dprTableWriter);
            XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
            insertDPRFlows(dprTableWriter, xlsxSheet);
            insertNewDPRAreas(dprTableWriter, xlsxSheet);
            insertDPRRoundingDifferences(newWorkbook, dprTableWriter);


            Assert.assertTrue(checkDPRRoundingDifferences(newWorkbook));

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

    private List<Integer> generateParcels() {
        List<Integer> parcels = new ArrayList<>();
        parcels.add(2174);
        parcels.add(2175);
        parcels.add(2176);

        return parcels;
    }

    private List<Integer> generateDPR() {
        List<Integer> dpr = new ArrayList<>();
        dpr.add(40053);
        dpr.add(15828);

        return dpr;
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

    private XSSFWorkbook insertParcels(String filePath, XLSXTemplate xlsxTemplate, ParcelTableWriter parcelTableWriter) throws Exception {

        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();
        List<Integer> parcels = generateParcels();

        XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
        XSSFSheet sheet = newWorkbook.getSheet("Mutationstabelle");
        xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size(),
                parcels.size());
        parcelTableWriter.writeOldParcelsInTemplate(oldParcels, sheet);
        parcelTableWriter.writeNewParcelsInTemplate(newParcels, sheet);

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

    private void insertInflowAndOutflows(ParcelTableWriter parcelTableWriter,
                                                 XSSFSheet xlsxSheet) {
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(695, 695,416, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(696, 696,507, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(697, 697,687, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(696, 701,1, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(697, 701,1, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(701, 701,1112, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(870, 870,611, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(874, 874,1939, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(695, 4004,242, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(696, 4004,100, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(697, 4004,129, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(701, 4004,1, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(870, 4004,39, xlsxSheet);
        parcelTableWriter.writeInflowAndOutflowOfOneParcelPair(874, 4004,81, xlsxSheet);

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

    private void insertRoundingDifferences(ParcelTableWriter parcelTableWriter,
                                                   XSSFSheet xlsxSheet){
        parcelTableWriter.writeRoundingDifference(695, -1, 7, xlsxSheet);
        parcelTableWriter.writeRoundingDifference(697, -1, 7, xlsxSheet);
        parcelTableWriter.writeRoundingDifference(701, 1, 7, xlsxSheet);
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



    private boolean checkOldAreas(XSSFWorkbook workbook) {
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allOldAreasAreCorrect = true;

        Row row = xlsxSheet.getRow(19);
        /*System.out.println(row.getCell(1));
        System.out.println(row.getCell(2));
        System.out.println(row.getCell(3));*/

        if (row.getCell(1).getNumericCellValue() != 657 ||
                row.getCell(2).getNumericCellValue() != 608 ||
                row.getCell(3).getNumericCellValue() != 816) {
            allOldAreasAreCorrect = false;
        }
        return allOldAreasAreCorrect;
    }

    private void insertNewAreas(ParcelTableWriter parcelTableWriter,
                                        XSSFSheet xlsxSheet) {


        parcelTableWriter.writeNewArea(695, 416, xlsxSheet);
        parcelTableWriter.writeNewArea(696, 507, xlsxSheet);
        parcelTableWriter.writeNewArea(697, 687, xlsxSheet);
        parcelTableWriter.writeNewArea(701, 1114, xlsxSheet);
        parcelTableWriter.writeNewArea(870, 611, xlsxSheet);
        parcelTableWriter.writeNewArea(874, 1939, xlsxSheet);
        parcelTableWriter.writeNewArea(4004, 592, xlsxSheet);

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

    private HashMap<Integer, Integer> oldAreas() {
        HashMap<Integer, Integer> areas = new HashMap<>();
        areas.put(695, 657);
        areas.put(696, 608);
        areas.put(697, 816);
        areas.put(701, 1114);
        areas.put(870, 650);
        areas.put(874, 2020);

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

    private boolean checkParcels(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allParcelsAreCorrect = true;

        Row row = xlsxSheet.getRow(24);
        for (Cell cell : row){
            if (cell.getColumnIndex()== 1){
                if(cell.getNumericCellValue()!=2174){
                    allParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 2) {
                if (cell.getNumericCellValue() != 2175) {
                    allParcelsAreCorrect = false;
                }
            } else if (cell.getColumnIndex()== 3) {
                if (cell.getNumericCellValue() != 2176) {
                    allParcelsAreCorrect = false;
                }
            }
        }
        return allParcelsAreCorrect;
    }


    private XSSFWorkbook insertDPR(String filePath, XLSXTemplate xlsxTemplate, DPRTableWriter dprTableWriter) {

        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();
        List<Integer> parcels = generateParcels();
        List<Integer> dpr = generateDPR();

        XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
        xlsxTemplate.createParcelTable(newWorkbook,filePath, newParcels.size(),oldParcels.size(),
                parcels.size());
        XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
        xlsxTemplate.createDPRTable(newWorkbook, filePath, parcels.size(), dpr.size(),
                newParcels.size(), oldParcels.size());
        dprTableWriter.writeParcelsAffectedByDPRsInTemplate(parcels, newParcels.size(), xlsxSheet);
        dprTableWriter.writeDPRsInTemplate(dpr, newParcels.size(), xlsxSheet);

        return newWorkbook;

    }

    private boolean checkDPRs(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allDPRsAreCorrect = true;

        if (!xlsxSheet.getRow(26).getCell(0).getStringCellValue().equals("(40053)") ||
                !xlsxSheet.getRow(28).getCell(0).getStringCellValue().equals("(15828)")){
            allDPRsAreCorrect = false;
        }
        return allDPRsAreCorrect;
    }

    private void insertDPRFlows(DPRTableWriter dprTableWriter,
                                        XSSFSheet xlsxSheet) {
        Integer numberNewParcels = generateNewParcels().size();
        dprTableWriter.writeDPRInflowAndOutflows(2174, 40053,1175, numberNewParcels,
                xlsxSheet);
        dprTableWriter.writeDPRInflowAndOutflows(2175, 40053,2481, numberNewParcels,
                xlsxSheet);
        dprTableWriter.writeDPRInflowAndOutflows(2176, 40053,5, numberNewParcels, xlsxSheet);

    }

    private boolean checkFlows(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allFlowsAreCorrect = true;
        Row row = xlsxSheet.getRow(26);

        if(row.getCell(1).getNumericCellValue() != 1175 ||
                row.getCell(2).getNumericCellValue() != 2481 ||
                row.getCell(3).getNumericCellValue() != 5){
            allFlowsAreCorrect = false;
        }

        return allFlowsAreCorrect;
    }

    private void insertNewDPRAreas(DPRTableWriter dprTableWriter,
                                           XSSFSheet xlsxSheet) {

        Integer numberNewParcels = generateNewParcels().size();
        dprTableWriter.writeNewDPRArea(40053,3660, numberNewParcels, xlsxSheet);
        dprTableWriter.writeNewDPRArea(15828,0, numberNewParcels, xlsxSheet);

    }

    private boolean checkDPRNewArea(XSSFWorkbook workbook) {
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allNewAreasAreCorrect = true;

        if (xlsxSheet.getRow(26).getCell(5).getNumericCellValue()!=3660 ||
                !xlsxSheet.getRow(28).getCell(5).getStringCellValue().equals("gelöscht")) {
            allNewAreasAreCorrect = false;
        }

        return allNewAreasAreCorrect;
    }

    private void insertDPRRoundingDifferences(XSSFWorkbook newWorkbook, DPRTableWriter dprTableWriter) {

        Integer numberNewParcels = generateNewParcels().size();
        XSSFSheet xlsxSheet = newWorkbook.getSheet("Mutationstabelle");
        dprTableWriter.writeDPRRoundingDifference(40053, -1, numberNewParcels, xlsxSheet);
    }

    private boolean checkDPRRoundingDifferences(XSSFWorkbook workbook){
        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allRoundingDifferencesAreCorrect = true;

        if (xlsxSheet.getRow(26).getCell(4).getNumericCellValue()!=-1) {
            allRoundingDifferencesAreCorrect = false;
        }

        return allRoundingDifferencesAreCorrect;
    }
}

