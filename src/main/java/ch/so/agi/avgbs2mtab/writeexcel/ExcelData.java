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

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * The Class ExcelData gets the data from the container and writes it into the prepared exceltemplate
 */
public class ExcelData implements WriteExcel {

    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());

    /**
     * gets the parcel data from the container and writes it into the parcel table from the prepared exceltemplate
     * @param filePath                  Path, where the excel-template should be writen to
     * @param workbook                  Excel-workbook
     * @param dataExtractionParcel      Methods to get data from container
     */
    public void fillValuesIntoParcelTable (String filePath,
                                           XSSFWorkbook workbook,
                                           DataExtractionParcel dataExtractionParcel){


        try {
            OutputStream excelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<Integer> orderedListOfOldParcelNumbers = dataExtractionParcel.getOldParcelNumbers();
            List<Integer> orderedListOfNewParcelNumbers = dataExtractionParcel.getNewParcelNumbers();

            writeParcelsIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers, xlsxSheet);

            writeAllInflowsAndOutFlowsIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    dataExtractionParcel, xlsxSheet);

            writeAllRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    dataExtractionParcel, xlsxSheet);

            writeSumOfRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    dataExtractionParcel, xlsxSheet);

            writeAllNewAreasIntoParcelTable(orderedListOfNewParcelNumbers, dataExtractionParcel, xlsxSheet);

            HashMap<Integer, Integer> oldAreaHashMap = getAllOldAreas(orderedListOfNewParcelNumbers,
                    orderedListOfOldParcelNumbers, dataExtractionParcel);

            writeAllOldAreasIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    oldAreaHashMap, xlsxSheet);


            writeAreaSumIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    oldAreaHashMap, dataExtractionParcel, xlsxSheet);

            workbook.write(excelFile);
            excelFile.close();

        } catch (IOException e){
            LOGGER.log(Level.SEVERE, "Could not write values into parcel table " + e.getMessage());
            throw new RuntimeException(e);
        }

    }

    /**
     * Writes numbers of parcels into parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    private void writeParcelsIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                             List<Integer> orderedListOfNewParcelNumbers,
                                             XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write number of new and old parcels into parcel table");

        writeOldParcelsInTemplate(orderedListOfOldParcelNumbers, xlsxSheet);
        writeNewParcelsInTemplate(orderedListOfNewParcelNumbers, xlsxSheet);
    }

    /**
     * writes the numbers of old parcels into specific row of parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    @Override
    public void writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                          XSSFSheet xlsxSheet) {

        Row rowWithOldParcelNumbers =xlsxSheet.getRow(2);

        Integer column = 1;

        for (Integer parcelNumber : orderedListOfOldParcelNumbers){
            Cell cell =rowWithOldParcelNumbers.getCell(column);
            cell.setCellValue(parcelNumber);
            column++;
        }
    }

    /**
     * writes the numbers of new parcels into specific column of parcel table
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    @Override
    public void writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                          XSSFSheet xlsxSheet){

        int rowIndex = 1;

        for (Integer parcelNumber : orderedListOfNewParcelNumbers){
            writeValueIntoCell(2+2*rowIndex, 0, xlsxSheet, parcelNumber);
            rowIndex++;
        }
    }

    /**
     * writes inflow or outflow value in a specific cell
     * @param rowIndex      Index of row
     * @param columnIndex   Index of column
     * @param xlsxSheet     excel sheet
     * @param value         value of inflow or outflow
     */
    private void writeValueIntoCell(Integer rowIndex,
                                    Integer columnIndex,
                                    XSSFSheet xlsxSheet,
                                    Integer value){

        Row row = xlsxSheet.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        cell.setCellValue(value);

    }


    /**
     * writes all inflows and outflows into parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    private void writeAllInflowsAndOutFlowsIntoParcelTable (List<Integer> orderedListOfOldParcelNumbers,
                                                            List<Integer> orderedListOfNewParcelNumbers,
                                                            DataExtractionParcel dataExtractionParcel,
                                                            XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write all inflows and outflows of each parcel into parcel table.");


        for (int oldParcel : orderedListOfOldParcelNumbers) {
            for (int newParcel : orderedListOfNewParcelNumbers) {

                Integer area = getAreaOfFlowBetweenOldAndNewParcel(oldParcel, newParcel, dataExtractionParcel);

                if (area != null) {
                    writeInflowAndOutflowOfOneParcelPair(oldParcel, newParcel, area, xlsxSheet);
                }
            }
        }
    }

    /**
     * gets area flow between two parcels
     * @param oldParcel             number of old parcel
     * @param newParcel             number of new parcel
     * @param dataExtractionParcel  get-methods for container
     * @return                      area flow
     */
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

    /**
     * Writes area flow into parcel table
     * @param oldParcelNumber   number of old parcel
     * @param newParcelNumber   number of new parcel
     * @param area              area flow
     * @param xlsxSheet         excel sheet
     */
    @Override
    public void writeInflowAndOutflowOfOneParcelPair(int oldParcelNumber,
                                                     int newParcelNumber,
                                                     int area,
                                                     XSSFSheet xlsxSheet) {

        Integer indexOldParcelNumber;
        Integer indexNewParcelNumber;

        indexOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        indexNewParcelNumber = getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        String errorMessage = "Either the old parcel " + oldParcelNumber + " or the new parcel " + newParcelNumber +
                " could not be found in the excel.";

        if (indexNewParcelNumber==null || indexOldParcelNumber==null){
            LOGGER.log(Level.SEVERE, errorMessage);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, errorMessage);
        } else {
            writeValueIntoCell(indexNewParcelNumber, indexOldParcelNumber, xlsxSheet, area);
        }
    }

    /**
     * Gets column index of a specific parcel
     * @param ParcelNumber  number of parcel
     * @param rowNumber     index of row
     * @param xlsxSheet     excel sheet
     * @return              index of column
     */
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
            LOGGER.log(Level.SEVERE, "Could not find Parcel " + ParcelNumber);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
                    ParcelNumber);
        }
    }

    /**
     * Gets row index of a specific parcel
     * @param newParcelNumber   number of new parcel
     * @param xlsxSheet         excel sheet
     * @return                  index of row
     */
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
            LOGGER.log(Level.SEVERE, "Could not finde parcel " +  newParcelNumber);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
                    newParcelNumber);
        } else {
            return indexNewParcelNumber;
        }
    }


    /**
     * Writes all rounding differences of the parcel table into parcel table
     * @param orderedListOfOldParcelNumbers List of numbers of old parcels
     * @param orderedListOfNewParcelNumbers List of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    private void writeAllRoundingDifferenceIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                           List<Integer> orderedListOfNewParcelNumbers,
                                                           DataExtractionParcel dataExtractionParcel,
                                                           XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the rounding difference for each parcel into parcel table.");

        for (int oldParcel : orderedListOfOldParcelNumbers) {

            Integer roundingDifference =  dataExtractionParcel.getRoundingDifference(oldParcel);

            if (roundingDifference != null && roundingDifference != 0) {

                writeRoundingDifference(oldParcel, -roundingDifference, orderedListOfNewParcelNumbers.size(), xlsxSheet);
            }
        }
    }

    /**
     * Writes a rounding difference from one parcel into parcel table
     * @param oldParcelNumber       number of old parcel
     * @param roundingDifference    value of rounding difference
     * @param numberOfNewParcels    amount of new parcels
     * @param xlsxSheet             excel sheet
     */
    @Override
    public void writeRoundingDifference(int oldParcelNumber,
                                        int roundingDifference,
                                        int numberOfNewParcels,
                                        XSSFSheet xlsxSheet) {

        Integer columnOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        Integer rowOldParcelNumber = 5 + 2 * numberOfNewParcels - 1;

        String errorMessage = "The old parcel "+ oldParcelNumber + " could not be found in the excel.";

        if (columnOldParcelNumber==null){
            LOGGER.log(Level.SEVERE, errorMessage);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, errorMessage);
        } else {
            writeValueIntoCell(rowOldParcelNumber, columnOldParcelNumber, xlsxSheet, roundingDifference);
        }
    }


    /**
     * Writes the sum of all rounding differences of parcel table into sum cell
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    private void writeSumOfRoundingDifferenceIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                             List<Integer> orderedListOfNewParcelNumbers,
                                                             DataExtractionParcel dataExtractionParcel,
                                                             XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the sum of all rounding differences into parcel table");

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {
            writeSumOfRoundingDifference(orderedListOfNewParcelNumbers.size(),
                    orderedListOfOldParcelNumbers.size(),
                    calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel),
                    xlsxSheet);
        }
    }

    /**
     * Gets all new areas of parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @return                              list of new areas
     */
    private List<Integer> getAllNewAreas (List<Integer> orderedListOfNewParcelNumbers,
                                          DataExtractionParcel dataExtractionParcel) {

        List<Integer> newAreaList = new ArrayList<>();

        for (int newParcel : orderedListOfNewParcelNumbers) {
            newAreaList.add(dataExtractionParcel.getNewArea(newParcel));
        }

        return newAreaList;
    }

    /**
     * Calculates sum of all rounding differences
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param dataExtractionParcel          get-methods for container
     * @return                              sum of all rounding differences
     */
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

    /**
     * Writes sum of rounding difference into specific cell
     * @param NumberOfNewParcels        amount of new parcels
     * @param NumberOfOldParcels        amount of old parcels
     * @param roundingDifferenceSum     value of rounding difference
     * @param xlsxSheet                 excel sheet
     */
    @Override
    public void writeSumOfRoundingDifference (int NumberOfNewParcels,
                                              int NumberOfOldParcels,
                                              int roundingDifferenceSum,
                                              XSSFSheet xlsxSheet){

        int rowNumber = 5 + 2 * NumberOfNewParcels - 1;
        int columnNumber = NumberOfOldParcels + 1;

        if (roundingDifferenceSum != 0) {
            writeValueIntoCell(rowNumber, columnNumber, xlsxSheet, roundingDifferenceSum);
        }
    }

    /**
     * Writes all new ares into parcel table
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    private void writeAllNewAreasIntoParcelTable(List<Integer> orderedListOfNewParcelNumbers,
                                                 DataExtractionParcel dataExtractionParcel,
                                                 XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write for each parcel the new area into parcel table.");

        for (int newParcel : orderedListOfNewParcelNumbers){

            int newArea = dataExtractionParcel.getNewArea(newParcel);

            writeNewArea(newParcel, newArea, xlsxSheet);
        }
    }

    /**
     * writes the new area of a parcel into parcel table
     * @param newParcelNumber   number of new parcel
     * @param area              value of area
     * @param xlsxSheet         excel sheet
     */
    @Override
    public void writeNewArea(int newParcelNumber,
                             int area,
                             XSSFSheet xlsxSheet) {

        Integer rowNewParcelNumber = getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        try {
            Row row = xlsxSheet.getRow(rowNewParcelNumber);
            Integer columnNewParcelNumber = row.getLastCellNum()-1;
            writeValueIntoCell(rowNewParcelNumber, columnNewParcelNumber, xlsxSheet, area);
        } catch (Exception e){
            LOGGER.log(Level.SEVERE,"Last row could not be found");
            throw new Avgbs2MtabException("Could not find last row");
        }
    }

    /**
     * gets all old areas and writes them into a hashmap
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param dataExtractionParcel          get-methods for parcel container
     * @return                              hashmap with all old areas
     */
    private HashMap<Integer, Integer> getAllOldAreas(List<Integer> orderedListOfNewParcelNumbers,
                                                     List<Integer> orderedListOfOldParcelNumbers,
                                                     DataExtractionParcel dataExtractionParcel) {

        LOGGER.log(Level.FINER, "Calculating for each old parcel the old area.");

        HashMap<Integer, Integer> oldAreaHashMap = new HashMap<>();
        Integer oldArea;
        Integer area;
        Integer roundingDifference;

        for (int oldParcel : orderedListOfOldParcelNumbers) {
            oldArea = 0;
            for (int newParcel : orderedListOfNewParcelNumbers) {

                area = getAreaOfFlowBetweenOldAndNewParcel(oldParcel, newParcel, dataExtractionParcel);

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
                LOGGER.log(Level.SEVERE,"Area of old parcel must not be null");
                throw new Avgbs2MtabException("Area of old parcel must not be null");
            }

        }

        return oldAreaHashMap;

    }

    /**
     * Writes all old areas into parcel table
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param oldAreaHashMap                hashmap with areas of old parcels
     * @param xlsxSheet                     excel sheet
     */
    private void writeAllOldAreasIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                                 List<Integer> orderedListOfNewParcelNumbers,
                                                 HashMap<Integer, Integer> oldAreaHashMap,
                                                 XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER,"Write for each old parcel the old area into parcel table.");

        for (int oldParcel : orderedListOfOldParcelNumbers) {
            Integer oldArea = oldAreaHashMap.get(oldParcel);

            writeOldArea(oldParcel, oldArea, orderedListOfNewParcelNumbers.size(), xlsxSheet);
        }
    }

    /**
     * Writes area of a old parcel into parcel table
     * @param oldParcelNumber       number of old parcel
     * @param oldArea               area of old parcel
     * @param numberOfnewParcels    amount of new parcels
     * @param xlsxSheet             excel sheet
     */
    @Override
    public void writeOldArea(int oldParcelNumber,
                             int oldArea,
                             int numberOfnewParcels,
                             XSSFSheet xlsxSheet){

        Integer columnOldParcelNumber = getColumnIndexOfParcelInTable(oldParcelNumber, 2, xlsxSheet);

        Integer rowOldParcelArea = 6 + 2 * numberOfnewParcels -1;

        String errorMessage = "The old parcel " + oldParcelNumber + " could not be found in the excel.";

        if (columnOldParcelNumber != null){
            writeValueIntoCell(rowOldParcelArea, columnOldParcelNumber, xlsxSheet, oldArea);
        } else {
            LOGGER.log(Level.SEVERE, errorMessage);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, errorMessage);
        }
    }


    /**
     * gets the sum of all area (old or new) and writes it into parcel table
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param oldAreaHashMap                hashmap with all old areas
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    private void writeAreaSumIntoParcelTable(List<Integer> orderedListOfOldParcelNumbers,
                                             List<Integer> orderedListOfNewParcelNumbers,
                                             HashMap<Integer, Integer> oldAreaHashMap,
                                             DataExtractionParcel dataExtractionParcel,
                                             XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the sum of the areas into parcel table.");

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {

            writeAreaSum(oldAreaHashMap, getAllNewAreas(orderedListOfNewParcelNumbers, dataExtractionParcel),
                    calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel), xlsxSheet);
        }
    }

    /**
     * Calculates the sum of all areas (old or new)
     * @param oldAreas              hashmap with all old areas
     * @param newAreas              list with all new areas
     * @param roundingDifference    sum of all rounding differences
     * @param xlsxSheet             excel sheet
     */
    @Override
    public void writeAreaSum(HashMap<Integer, Integer> oldAreas,
                             List<Integer> newAreas,
                             int roundingDifference,
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
            LOGGER.log(Level.SEVERE, "The sum of the old areas is not equal to the sum of the new areas.");
            throw new Avgbs2MtabException("The sum of the old areas is not equal to the sum of the new areas.");
        }
    }


    /**
     * Fills all values (parcel and dpr number, area values, rounding differences) into dpr table
     * @param filePath                  path, where the excel file should be saved to
     * @param workbook                  excel workbook
     * @param dataExtractionDPR         get-Methods for dpr-container
     * @param metadataOfParcelMutation  get-Methods for dpr metadata
     */
    public void fillValuesIntoDPRTable (String filePath, XSSFWorkbook workbook,
                                        DataExtractionDPR dataExtractionDPR,
                                        MetadataOfParcelMutation metadataOfParcelMutation) {

        try {
            OutputStream ExcelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<Integer> orderedListOfParcelNumbers = dataExtractionDPR.getParcelsAffectedByDPRs();
            List<Integer> orderedListOfDPRs = dataExtractionDPR.getNewDPRs();


            Integer numberOfNewParcelsInParcelTable = metadataOfParcelMutation.getNumberOfNewParcels();

            writeParcelsAndDPRsIntoTable(orderedListOfParcelNumbers, orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    xlsxSheet);


            writeAllFlowsIntoDPRTable(orderedListOfParcelNumbers, orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    dataExtractionDPR, xlsxSheet);


            writeAllRoundingDifferencesIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    dataExtractionDPR, xlsxSheet);


            writeAllNewAreasIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable, dataExtractionDPR,
                    xlsxSheet);


            workbook.write(ExcelFile);
            ExcelFile.close();

        } catch (IOException e){
            LOGGER.log(Level.SEVERE, "Could not write values into dpr table : " + e.getMessage());
            throw new RuntimeException(e);
        }
    }

    /**
     * writes all number of parcels and dprs into dpr table
     * @param orderedListOfParcelNumbers        list of numbers of parcels
     * @param orderedListOfDPRs                 list of number of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param xlsxSheet                         excel sheet
     */
    private void writeParcelsAndDPRsIntoTable(List<Integer> orderedListOfParcelNumbers,
                                              List<Integer> orderedListOfDPRs,
                                              int numberOfNewParcelsInParcelTable,
                                              XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the numbers of dprs and parcels into dpr table");

        writeParcelsAffectedByDPRsInTemplate(orderedListOfParcelNumbers, numberOfNewParcelsInParcelTable, xlsxSheet);
        writeDPRsInTemplate(orderedListOfDPRs, numberOfNewParcelsInParcelTable, xlsxSheet);
    }

    /**
     * Writes number of parcels into dpr table
     * @param orderedListOfParcelNumbersAffectedByDPRs  list of numbers of parcel
     * @param newParcelNumber                           amount of new parcels in parcel table
     * @param xlsxSheet                                 excel sheet
     */
    @Override
    public void writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                     int newParcelNumber,
                                                     XSSFSheet xlsxSheet) {

        Integer column = 1;

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);


        for (Integer parcelNumber : orderedListOfParcelNumbersAffectedByDPRs){

            writeValueIntoCell(indexOfParcelRow, column, xlsxSheet, parcelNumber);

            column++;
        }
    }

    /**
     * Calculates index of parcel row in dpr table
     * @param newParcelNumber   amount of new parcels
     * @param constant          constant value
     * @return                  index of parcel row
     */
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

    /**
     * Writes all dpr into dpr table
     * @param orderedListOfDPRs list of numbers of dprs
     * @param newParcelNumber   amount of new parcels
     * @param xlsxSheet         excel sheet
     */
    @Override
    public void writeDPRsInTemplate(List<Integer> orderedListOfDPRs,
                                    int newParcelNumber,
                                    XSSFSheet xlsxSheet) {

        Integer rowIndex = calculateIndexOfParcelRow(newParcelNumber, 12);

        for (Integer dpr : orderedListOfDPRs){

            writeValueIntoCell(rowIndex, 0, xlsxSheet, dpr.toString());

            rowIndex++;
            rowIndex++;
        }
    }

    /**
     * Rewrites the number of the given dpr and writes it into a specific cell
     * @param rowIndex      index of row
     * @param columnIndex   index of column
     * @param xlsxSheet     excel sheet
     * @param value         number of dpr as a string
     */
    private void writeValueIntoCell(Integer rowIndex,
                                    Integer columnIndex,
                                    XSSFSheet xlsxSheet,
                                    String value){

        XSSFRow row = xlsxSheet.getRow(rowIndex);
        XSSFCell cell =row.getCell(columnIndex);
        cell.setCellValue("(" + value + ")");

    }

    /**
     * Writes all flows between parcels and dprs into dpr table
     * @param orderedListOfParcelNumbers        list of numbers of parcels
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    private void writeAllFlowsIntoDPRTable(List<Integer> orderedListOfParcelNumbers,
                                           List<Integer> orderedListOfDPRs,
                                           int numberOfNewParcelsInParcelTable,
                                           DataExtractionDPR dataExtractionDPR,
                                           XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write all flows of area into dpr table");

        for (int parcel : orderedListOfParcelNumbers) {
            for (int dpr : orderedListOfDPRs) {

                Integer area = dataExtractionDPR.getAddedAreaDPR(parcel, dpr);

                writeDPRInflowAndOutflows(parcel, dpr, area, numberOfNewParcelsInParcelTable, xlsxSheet);
            }
        }
    }

    /**
     * Writes area flow between one parcel and one dpr into dpr table
     * @param parcelNumberAffectedByDPR number of parcel
     * @param dpr                       number of dpr
     * @param area                      value of area
     * @param newParcelNumber           amount of new parcels
     * @param xlsxSheet                 excel sheet
     */
    @Override
    public void writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                          int dpr,
                                          int area,
                                          int newParcelNumber,
                                          XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        int indexParcel = getColumnIndexOfParcelInTable(parcelNumberAffectedByDPR, indexOfParcelRow, xlsxSheet);

        int indexDPR = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        writeValueIntoCell(indexDPR, indexParcel, xlsxSheet, area);
    }

    /**
     * gets row index of dpr in dpr table
     * @param indexOfParcelRow  index of row with parcel numbers
     * @param dpr               number of dpr
     * @param xlsxSheet         excel sheet
     * @return                  row index of dpr
     */
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
            LOGGER.log(Level.SEVERE,"Could not finde DPR " + dpr + " in DPR-Table.");
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not finde DPR " +
                    dpr + " in DPR-Table.");
        }
    }

    /**
     * extracts the dpr number from a string
     * @param cell  excel cell
     * @return      number of dpr
     */
    private Integer getDPRNumberFromCell(Cell cell){
        String dprString = cell.getStringCellValue();
        int dprStringLength = dprString.length();

        return Integer.parseInt(dprString.substring(1, (dprStringLength-1)));
    }

    /**
     * Writes all rounding differences into dpr table
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-Methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    private void writeAllRoundingDifferencesIntoDPRTable(List<Integer> orderedListOfDPRs,
                                                         int numberOfNewParcelsInParcelTable,
                                                         DataExtractionDPR dataExtractionDPR,
                                                         XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write rounding difference for each dpr into dpr table");

        for (int dpr : orderedListOfDPRs) {

            Integer roundingDifference = dataExtractionDPR.getRoundingDifferenceDPR(dpr);

            if (roundingDifference != null && roundingDifference != 0) {
                writeDPRRoundingDifference(dpr, -roundingDifference, numberOfNewParcelsInParcelTable, xlsxSheet);
            }
        }
    }

    /**
     * Write rounding difference for one dpr
     * @param dpr                   number of dpr
     * @param roundingDifference    rounding difference
     * @param newParcelNumber       amount of new parcels in parcel table
     * @param xlsxSheet             excel sheet
     */
    @Override
    public void  writeDPRRoundingDifference(int dpr,
                                            int roundingDifference,
                                            int newParcelNumber,
                                            XSSFSheet xlsxSheet) {

        Integer rowDPRNumber;
        Integer columnRoundingDifference;

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        rowDPRNumber = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        columnRoundingDifference = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-2;

        writeValueIntoCell(rowDPRNumber, columnRoundingDifference, xlsxSheet, roundingDifference);
    }

    /**
     * writes all new areas into dpr table
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    private void writeAllNewAreasIntoDPRTable(List<Integer> orderedListOfDPRs,
                                              int numberOfNewParcelsInParcelTable,
                                              DataExtractionDPR dataExtractionDPR,
                                              XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write for each dpr the new area into dpr table");

        for (int dpr : orderedListOfDPRs) {
            Integer newArea = dataExtractionDPR.getNewAreaDPR(dpr);
            writeNewDPRArea(dpr, newArea, numberOfNewParcelsInParcelTable, xlsxSheet);
        }
    }


    /**
     * writes the new area of one dpr into dpr table
     * @param dpr               number of dpr
     * @param area              value of area
     * @param newParcelNumber   amount of new parcels in parcel table
     * @param xlsxSheet         excel sheet
     */
    @Override
    public void writeNewDPRArea(int dpr,
                                int area,
                                int newParcelNumber,
                                XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow = calculateIndexOfParcelRow(newParcelNumber, 10);

        Integer rowDPRNumber = getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        Integer columnNewArea = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-1;

        writeNewDPRAreaValueIntoCell(rowDPRNumber, columnNewArea, area, xlsxSheet);
    }

    /**
     * Writes new area value from a specific dpr into dpr table
     * @param rowDPRNumber      index of row of specific dpr
     * @param columnNewArea     index of column with new areas
     * @param area              value of area
     * @param xlsxSheet         excel sheet
     */
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
