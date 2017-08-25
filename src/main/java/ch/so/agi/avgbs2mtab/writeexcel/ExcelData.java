package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.*;
import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.security.x509.AVA;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

public class ExcelData implements WriteExcel {



    public XSSFWorkbook fillValuesIntoParcelTable (String filePath, XSSFWorkbook workbook,
                                                   DataExtractionParcel dataExtractionParcel){


        try {
            OutputStream excelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<Integer> orderedListOfOldParcelNumbers = dataExtractionParcel.getOldParcelNumbers();
            List<Integer> orderedListOfNewParcelNumbers = dataExtractionParcel.getNewParcelNumbers();

            workbook = writeParcelsIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    xlsxSheet, workbook);

            workbook = writeAllInflowsAndOutFlowsIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, workbook, dataExtractionParcel, xlsxSheet);

            workbook = writeAllRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, workbook, dataExtractionParcel, xlsxSheet);

            workbook = writeSumOfRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, workbook, dataExtractionParcel, xlsxSheet);

            workbook = writeAllNewAreasIntoParcelTable(orderedListOfNewParcelNumbers, workbook,
                    dataExtractionParcel, xlsxSheet);

            HashMap<Integer, Integer> oldAreaHashMap = getAllOldAreas(orderedListOfNewParcelNumbers,
                    orderedListOfOldParcelNumbers, dataExtractionParcel);

            workbook = writeAllOldAreasIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    oldAreaHashMap, workbook, xlsxSheet);


            workbook = writeAreaSumIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    oldAreaHashMap, workbook, dataExtractionParcel, xlsxSheet);

            workbook.write(excelFile);
            excelFile.close();
            return workbook;
        } catch (IOException e){
            throw new RuntimeException(e);
        }

    }

    private XSSFWorkbook writeParcelsIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                     List<Integer> orderedListOfNewParcelNumbers,
                                                     XSSFSheet xlsxSheet,
                                                     XSSFWorkbook workbook){

        workbook = writeOldParcelsInTemplate(orderedListOfOldParcelNumbers, workbook, xlsxSheet);
        workbook = writeNewParcelsInTemplate(orderedListOfNewParcelNumbers, workbook, xlsxSheet);

        return workbook;
    }

    @Override
    public XSSFWorkbook writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet) {

        Row rowWithOldParcelNumbers =xlsxSheet.getRow(2);

        Integer column = 1;

        for (Integer parcelNumber : orderedListOfOldParcelNumbers){
            Cell cell =rowWithOldParcelNumbers.getCell(column);
            cell.setCellValue(parcelNumber);
            column++;
        }

        return workbook;
    }

    @Override
    public XSSFWorkbook writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet){

        int rowIndex = 1;

        for (Integer parcelNumber : orderedListOfNewParcelNumbers){
            writeValueIntoCell(2+2*rowIndex, 0, xlsxSheet, parcelNumber);
            rowIndex++;
        }

        return workbook;
    }

    private void writeValueIntoCell(Integer rowIndex,
                                     Integer columnIndex,
                                     XSSFSheet xlsxSheet,
                                     Integer value){

        Row row = xlsxSheet.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        cell.setCellValue(value);

    }


    private XSSFWorkbook writeAllInflowsAndOutFlowsIntoParcelTable (List<Integer> orderedListOfOldParcelNumbers,
                                                                    List<Integer> orderedListOfNewParcelNumbers,
                                                                    XSSFWorkbook workbook,
                                                                    DataExtractionParcel dataExtractionParcel,
                                                                    XSSFSheet xlsxSheet) {


        for (int oldParcel : orderedListOfOldParcelNumbers) {
            for (int newParcel : orderedListOfNewParcelNumbers) {

                Integer area = getAreaOfFlowBetweenOldAndNewParcel(oldParcel, newParcel, dataExtractionParcel);

                if (area != null) {
                    workbook = writeInflowAndOutflowOfOneParcelPair(oldParcel, newParcel, area, workbook, xlsxSheet);
                }
            }
        }

        return workbook;
    }

    private Integer getAreaOfFlowBetweenOldAndNewParcel(int oldParcel,
                                                        int newParcel,
                                                        DataExtractionParcel dataExtractionParcel){

        Integer area;

        if (oldParcel != newParcel) {
            area = dataExtractionParcel.getAddedArea(newParcel, oldParcel);
        } else {
            area = dataExtractionParcel.getRestAreaOfParcel(oldParcel);
        }

        return area;

    }

    @Override
    public XSSFWorkbook writeInflowAndOutflowOfOneParcelPair(int oldParcelNumber,
                                                             int newParcelNumber,
                                                             int area,
                                                             XSSFWorkbook workbook,
                                                             XSSFSheet xlsxSheet) {

        Integer indexOldParcelNumber;
        Integer indexNewParcelNumber;

        indexOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        indexNewParcelNumber = getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        if (indexNewParcelNumber==null || indexOldParcelNumber==null){
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Either the old parcel "
                    + oldParcelNumber + " or the new parcel " + newParcelNumber + " could not be found in the excel.");
        } else {
            writeValueIntoCell(indexNewParcelNumber, indexOldParcelNumber, xlsxSheet, area);
        }

        return workbook;
    }

    private Integer getColumnIndexOfParcelInTable(int ParcelNumber,
                                                  int rowNumber,
                                                  XSSFSheet xlsxSheet){
        Integer indexOldParcelNumber = null;

        Row row = xlsxSheet.getRow(rowNumber);

        for (Cell cell : row) {
            if (cell.getCellTypeEnum() == CellType.NUMERIC && cell.getNumericCellValue() == ParcelNumber) {
                indexOldParcelNumber = cell.getColumnIndex();
                break;
            }
        }
        if (indexOldParcelNumber != null) {
            return indexOldParcelNumber;
        } else {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
                    ParcelNumber);
        }
    }

    private Integer getRowIndexOfNewParcelInTable (int newParcelNumber,
                                                   XSSFSheet xlsxSheet){

        Integer indexNewParcelNumber = null;

        Iterator<Row> rowIterator = xlsxSheet.iterator();
        rowIterator.next();

        while(rowIterator.hasNext()){
            Row row1 = rowIterator.next();
            Cell cell1 = row1.getCell(0);

            if (cell1.getCellTypeEnum() == CellType.NUMERIC && cell1.getNumericCellValue() == newParcelNumber){
                indexNewParcelNumber = cell1.getRowIndex();
                break;
            }
        }

        if (indexNewParcelNumber == null) {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
            newParcelNumber);
        } else {
            return indexNewParcelNumber;
        }
    }



    private XSSFWorkbook writeAllRoundingDifferenceIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                                   List<Integer> orderedListOfNewParcelNumbers,
                                                                   XSSFWorkbook workbook,
                                                                   DataExtractionParcel dataExtractionParcel,
                                                                   XSSFSheet xlsxSheet) {

        for (int oldParcel : orderedListOfOldParcelNumbers) {

            Integer roundingDifference =  dataExtractionParcel.getRoundingDifference(oldParcel);

            if (roundingDifference != null && roundingDifference != 0) {

                workbook = writeRoundingDifference(oldParcel, -roundingDifference, orderedListOfNewParcelNumbers.size(),
                        workbook, xlsxSheet);
            }
        }

        return workbook;
    }

    @Override
    public XSSFWorkbook writeRoundingDifference(int oldParcelNumber,int roundingDifference, int numberOfNewParcels,
                                                XSSFWorkbook workbook, XSSFSheet xlsxSheet) {

        Integer columnOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        Integer rowOldParcelNumber = 5 + 2 * numberOfNewParcels - 1;

        if (columnOldParcelNumber==null){
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "The old parcel "
                    + oldParcelNumber + " could not be found in the excel.");
        } else {
            writeValueIntoCell(rowOldParcelNumber, columnOldParcelNumber, xlsxSheet, roundingDifference);
        }

        return workbook;
    }


    private XSSFWorkbook writeSumOfRoundingDifferenceIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                                     List<Integer> orderedListOfNewParcelNumbers,
                                                                     XSSFWorkbook workbook,
                                                                     DataExtractionParcel dataExtractionParcel,
                                                                     XSSFSheet xlsxSheet) {

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {
            workbook = writeSumOfRoundingDifference(orderedListOfNewParcelNumbers.size(),
                    orderedListOfOldParcelNumbers.size(),
                    calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel), workbook,
                    xlsxSheet);
        }

        return workbook;
    }

    private List<Integer> getAllNewAreas (List<Integer> orderedListOfNewParcelNumbers,
                                          DataExtractionParcel dataExtractionParcel) {

        List<Integer> newAreaList = new ArrayList<>();

        for (int newParcel : orderedListOfNewParcelNumbers) {
            newAreaList.add(dataExtractionParcel.getNewArea(newParcel));
        }

        return newAreaList;
    }

    private Integer calculateRoundingDifference(List<Integer> orderedListOfOldParcelNumbers,
                                                DataExtractionParcel dataExtractionParcel) {

        Integer roundingDifference = 0;

        for (int oldParcel : orderedListOfOldParcelNumbers) {
            Integer newRoundingDifference = dataExtractionParcel.getRoundingDifference(oldParcel);
            if (newRoundingDifference != null) {
                roundingDifference += newRoundingDifference;
            }
        }

        return - roundingDifference;

    }

    @Override
    public XSSFWorkbook writeSumOfRoundingDifference (int NumberOfNewParcels,
                                                      int NumberOfOldParcels,
                                                      int roundingDifferenceSum,
                                                      XSSFWorkbook workbook,
                                                      XSSFSheet xlsxSheet){

        int rowNumber = 5 + 2 * NumberOfNewParcels - 1;
        int columnNumber = NumberOfOldParcels + 1;

        if (roundingDifferenceSum != 0) {
            writeValueIntoCell(rowNumber, columnNumber, xlsxSheet, roundingDifferenceSum);
        }

        return workbook;
    }

    private XSSFWorkbook writeAllNewAreasIntoParcelTable(List<Integer> orderedListOfNewParcelNumbers,
                                                         XSSFWorkbook workbook,
                                                         DataExtractionParcel dataExtractionParcel,
                                                         XSSFSheet xlsxSheet) {

        for (int newParcel : orderedListOfNewParcelNumbers){

            int newArea = dataExtractionParcel.getNewArea(newParcel);

            workbook = writeNewArea(newParcel, newArea, workbook, xlsxSheet);
        }

        return workbook;
    }

    @Override
    public XSSFWorkbook writeNewArea(int newParcelNumber,
                                     int area,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet) {

        Integer rowNewParcelNumber = getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        try {
            Row row = xlsxSheet.getRow(rowNewParcelNumber);
            Integer columnNewParcelNumber = row.getLastCellNum()-1;
            writeValueIntoCell(rowNewParcelNumber, columnNewParcelNumber, xlsxSheet, area);
        } catch (Exception e){
            throw new Avgbs2MtabException("Could not find last row");
        }

        return workbook;
    }

    private HashMap<Integer, Integer> getAllOldAreas(List<Integer> orderedListOfNewParcelNumbers,
                                                     List<Integer> orderedListOfOldParcelNumbers,
                                                     DataExtractionParcel dataExtractionParcel) {

        HashMap<Integer, Integer> oldAreaHashMap = new HashMap<>();
        Integer oldArea;
        Integer area;
        Integer roundingDifference;

        for (int oldParcel : orderedListOfOldParcelNumbers) {
            oldArea = 0;
            for (int newParcel : orderedListOfNewParcelNumbers) {

                area = getFlowBetweenTwoParcels(oldParcel, newParcel, dataExtractionParcel);

                if (area != null) {
                    oldArea += area;
                }
            }

            if (oldArea != null) {
                roundingDifference = dataExtractionParcel.getRoundingDifference(oldParcel);
                if (roundingDifference != null){
                    oldArea = oldArea - roundingDifference;
                }
                oldAreaHashMap.put(oldParcel, oldArea);
            } else {
                throw new Avgbs2MtabException("Area of old parcel must not be null");
            }

        }

        return oldAreaHashMap;

    }

    private Integer getFlowBetweenTwoParcels(int oldParcel,
                                             int newParcel,
                                             DataExtractionParcel dataExtractionParcel){
        Integer area;

        if (oldParcel != newParcel) {

            area = dataExtractionParcel.getAddedArea(newParcel, oldParcel);
        } else {
            area = dataExtractionParcel.getRestAreaOfParcel(oldParcel);
        }

        return area;
    }


    private XSSFWorkbook writeAllOldAreasIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                         List<Integer> orderedListOfNewParcelNumbers,
                                                         HashMap<Integer, Integer> oldAreaHashMap,
                                                         XSSFWorkbook workbook,
                                                         XSSFSheet xlsxSheet) {

        for (int oldParcel : orderedListOfOldParcelNumbers) {
            Integer oldArea = oldAreaHashMap.get(oldParcel);

            workbook = writeOldArea(oldParcel, oldArea, orderedListOfNewParcelNumbers.size(),
                    workbook, xlsxSheet);
        }

        return workbook;
    }


    @Override
    public XSSFWorkbook writeOldArea(int oldParcelNumber,
                                     int oldArea,
                                     int numberOfnewParcels,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet){

        Integer columnOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        Integer rowOldParcelArea = 6 + 2 * numberOfnewParcels -1;

        if (columnOldParcelNumber != null){
            writeValueIntoCell(rowOldParcelArea, columnOldParcelNumber, xlsxSheet, oldArea);
        } else {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "The old parcel "
                    + oldParcelNumber + " could not be found in the excel.");
        }

        return workbook;
    }


    private XSSFWorkbook writeAreaSumIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                     List<Integer> orderedListOfNewParcelNumbers,
                                                     HashMap<Integer, Integer> oldAreaHashMap,
                                                     XSSFWorkbook workbook,
                                                     DataExtractionParcel dataExtractionParcel,
                                                     XSSFSheet xlsxSheet) {

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {

            workbook = writeAreaSum(oldAreaHashMap, getAllNewAreas(orderedListOfNewParcelNumbers, dataExtractionParcel),
                    calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel),
                    workbook, xlsxSheet);
        }

        return workbook;
    }

    @Override
    public XSSFWorkbook writeAreaSum(HashMap<Integer, Integer> oldAreas,
                                     List<Integer> newAreas,
                                     int roundingDifference,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet){

        Integer sumOldAreas = 0;
        Integer sumNewAreas = 0;

        Integer numberOfOldParcels = oldAreas.size();
        Integer numberOfNewParcels = newAreas.size();


        for (Map.Entry<Integer, Integer> entry : oldAreas.entrySet()){
            sumOldAreas += entry.getValue();
        }

        for (int area : newAreas){
            sumNewAreas += area;
        }
        sumNewAreas = sumNewAreas + roundingDifference;

        if (sumOldAreas.equals( sumNewAreas)){
            writeValueIntoCell(5 + 2 * numberOfNewParcels, 1 + numberOfOldParcels, xlsxSheet,
                    sumOldAreas);
        } else {
            throw new Avgbs2MtabException("The sum of the old areas does not equal with the sum of the new areas.");
        }

        return workbook;
    }



    public XSSFWorkbook fillValuesIntoDPRTable (String filePath, XSSFWorkbook workbook,
                                                DataExtractionDPR dataExtractionDPR,
                                                MetadataOfParcelMutation metadataOfParcelMutation) {

        try {
            OutputStream ExcelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<Integer> orderedListOfParcelNumbers = dataExtractionDPR.getParcelsAffectedByDPRs();
            List<Integer> orderedListOfDPRs = dataExtractionDPR.getNewDPRs();


            Integer numberOfNewParcelsInParcelTable = metadataOfParcelMutation.getNumberOfNewParcels();

            workbook = writeParcelsAndDPRsIntoTable(orderedListOfParcelNumbers, orderedListOfDPRs,
                    numberOfNewParcelsInParcelTable, workbook, xlsxSheet);


            workbook = writeAllFlowsIntoDPRTable(orderedListOfParcelNumbers, orderedListOfDPRs,
                    numberOfNewParcelsInParcelTable, workbook, dataExtractionDPR, xlsxSheet);


            workbook = writeAllRoundingDifferencesIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    workbook, dataExtractionDPR, xlsxSheet);


            workbook = writeAllNewAreasIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable, workbook,
                    dataExtractionDPR, xlsxSheet);


            workbook.write(ExcelFile);
            ExcelFile.close();
            return workbook;
        } catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    private XSSFWorkbook writeParcelsAndDPRsIntoTable(List<Integer> orderedListOfParcelNumbers,
                                                      List<Integer> orderedListOfDPRs,
                                                      int numberOfNewParcelsInParcelTable,
                                                      XSSFWorkbook workbook,
                                                      XSSFSheet xlsxSheet) {

        workbook = writeParcelsAffectedByDPRsInTemplate(orderedListOfParcelNumbers, numberOfNewParcelsInParcelTable,
                workbook, xlsxSheet);
        workbook = writeDPRsInTemplate(orderedListOfDPRs, numberOfNewParcelsInParcelTable, workbook, xlsxSheet);

        return workbook;
    }

    @Override
    public XSSFWorkbook writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                             int newParcelNumber,
                                                             XSSFWorkbook workbook,
                                                             XSSFSheet xlsxSheet) {

        Integer column = 1;

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);


        for (Integer parcelNumber : orderedListOfParcelNumbersAffectedByDPRs){

            writeValueIntoCell(indexOfParcelRow, column, xlsxSheet, parcelNumber);

            column++;
        }

        return workbook;
    }

    private int calculateIndexOfParcelRow (int newParcelNumber,
                                           int constant) {
        int indexOfParcelRow;

        if (newParcelNumber == 0){
            indexOfParcelRow = 2 + constant;
        } else {
            indexOfParcelRow = newParcelNumber*2 + constant;
        }

        return indexOfParcelRow;

    }

    @Override
    public XSSFWorkbook writeDPRsInTemplate(List<Integer> orderedListOfDPRs,
                                            int newParcelNumber,
                                            XSSFWorkbook workbook,
                                            XSSFSheet xlsxSheet) {

        Integer rowIndex = calculateIndexOfParcelRow(newParcelNumber, 12);

        for (Integer dpr : orderedListOfDPRs){

            writeValueIntoCell(rowIndex, 0, xlsxSheet, dpr.toString());

            rowIndex++;
            rowIndex++;
        }

        return workbook;
    }

    private void writeValueIntoCell(Integer rowIndex,
                               Integer columnIndex,
                               XSSFSheet xlsxSheet,
                               String value){

        XSSFRow row = xlsxSheet.getRow(rowIndex);
        XSSFCell cell =row.getCell(columnIndex);
        cell.setCellValue("(" + value + ")");

    }

    private XSSFWorkbook writeAllFlowsIntoDPRTable(List<Integer> orderedListOfParcelNumbers,
                                                   List<Integer> orderedListOfDPRs,
                                                   int numberOfNewParcelsInParcelTable,
                                                   XSSFWorkbook workbook,
                                                   DataExtractionDPR dataExtractionDPR,
                                                   XSSFSheet xlsxSheet){

        for (int parcel : orderedListOfParcelNumbers) {
            for (int dpr : orderedListOfDPRs) {

                Integer area = dataExtractionDPR.getAddedAreaDPR(parcel, dpr);

                workbook = writeDPRInflowAndOutflows(parcel, dpr, area, numberOfNewParcelsInParcelTable,
                        workbook, xlsxSheet);
            }
        }

        return workbook;
    }

    @Override
    public XSSFWorkbook writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                                  int dpr,
                                                  int area,
                                                  int newParcelNumber,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        int indexParcel = getColumnIndexOfParcelInTable(parcelNumberAffectedByDPR, indexOfParcelRow, xlsxSheet);

        int indexDPR = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        writeValueIntoCell(indexDPR, indexParcel, xlsxSheet, area);

        return workbook;
    }

    private Integer getRowIndexOfDPRInTable(int indexOfParcelRow,
                                            int dpr,
                                            XSSFSheet xlsxSheet){

        int lastRow = xlsxSheet.getLastRowNum();
        Integer indexDPR = null;

        for (int i = indexOfParcelRow + 2; i <= lastRow; i++){
            Row row1 = xlsxSheet.getRow(i);
            Cell cell1 = row1.getCell(0);

            if (cell1.getCellTypeEnum() == CellType.STRING){

                int dprNumber = getDPRNumberFromCell(cell1);

                if (dprNumber == dpr){
                    indexDPR = cell1.getRowIndex();
                    break;
                }
            }
        }

        if (indexDPR != null) {
            return indexDPR;
        } else {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not finde DPR " +
                    dpr + " in DPR-Table.");
        }
    }

    private Integer getDPRNumberFromCell(Cell cell1){
        String dprString = cell1.getStringCellValue();
        int dprStringLength = dprString.length();

        return Integer.parseInt(dprString.substring(1, (dprStringLength-1)));
    }

    private XSSFWorkbook writeAllRoundingDifferencesIntoDPRTable(List<Integer> orderedListOfDPRs,
                                                                 int numberOfNewParcelsInParcelTable,
                                                                 XSSFWorkbook workbook,
                                                                 DataExtractionDPR dataExtractionDPR,
                                                                 XSSFSheet xlsxSheet){

        for (int dpr : orderedListOfDPRs) {

            Integer roundingDifference = dataExtractionDPR.getRoundingDifferenceDPR(dpr);

            if (roundingDifference != null && roundingDifference != 0) {
                workbook = writeDPRRoundingDifference(dpr, -roundingDifference, numberOfNewParcelsInParcelTable,
                        workbook, xlsxSheet);
            }
        }

        return workbook;

    }

    @Override
    public XSSFWorkbook writeDPRRoundingDifference(int dpr,
                                                   int roundingDifference,
                                                   int newParcelNumber,
                                                   XSSFWorkbook workbook,
                                                   XSSFSheet xlsxSheet) {

        Integer rowDPRNumber;
        Integer columnRoundingDifference;

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        rowDPRNumber = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        columnRoundingDifference = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-2;

        writeValueIntoCell(rowDPRNumber, columnRoundingDifference, xlsxSheet, roundingDifference);

        return workbook;
    }

    private XSSFWorkbook writeAllNewAreasIntoDPRTable(List<Integer> orderedListOfDPRs,
                                                      int numberOfNewParcelsInParcelTable,
                                                      XSSFWorkbook workbook,
                                                      DataExtractionDPR dataExtractionDPR,
                                                      XSSFSheet xlsxSheet){

        for (int dpr : orderedListOfDPRs) {
            Integer newArea = dataExtractionDPR.getNewAreaDPR(dpr);
            workbook = writeNewDPRArea(dpr, newArea, numberOfNewParcelsInParcelTable, workbook, xlsxSheet);
        }

        return workbook;
    }



    @Override
    public XSSFWorkbook writeNewDPRArea(int dpr,
                                        int area,
                                        int newParcelNumber,
                                        XSSFWorkbook workbook,
                                        XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        Integer rowDPRNumber = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        Integer columnNewArea = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-1;

        writeNewDPRAreaValueIntoCell(rowDPRNumber, columnNewArea, area, xlsxSheet);

        return workbook;
    }

    private void writeNewDPRAreaValueIntoCell(int rowDPRNumber,
                                              int columnNewArea,
                                              int area,
                                              XSSFSheet xlsxSheet){

        Row rowFlows = xlsxSheet.getRow(rowDPRNumber);
        Cell cellFlows = rowFlows.getCell(columnNewArea);

        if (area > 0) {
            cellFlows.setCellValue(area);
        } else {
            cellFlows.setCellValue("gelöscht");
        }
    }
}
