package ch.so.agi.avgbs2mtab.writeexcel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;


public interface WriteExcel  {

    public XSSFWorkbook writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet);

    public XSSFWorkbook writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet);

    public XSSFWorkbook writeInflowAndOutflowOfOneParcelPair(int oldParcelNumber,
                                                             int newParcelNumber,
                                                             int area,
                                                             XSSFWorkbook workbook,
                                                             XSSFSheet xlsxSheet);

    public XSSFWorkbook writeNewArea(int newParcelNumber,
                                     int area,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet);

    public XSSFWorkbook writeRoundingDifference(int oldParcelNumber,
                                                int roundingDifference,
                                                int numberOfNewParcels,
                                                XSSFWorkbook workbook,
                                                XSSFSheet xlsxSheet);

    public XSSFWorkbook writeSumOfRoundingDifference (int NumberOfNewParcels,
                                                      int NumberOfOldParcels,
                                                      int roundingDifferenceSum,
                                                      XSSFWorkbook workbook,
                                                      XSSFSheet xlsxSheet);


    public XSSFWorkbook writeOldArea(int oldParcelNumber,
                                     int oldArea,
                                     int numberOfNewParcels,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet);

    public XSSFWorkbook writeAreaSum(HashMap<Integer, Integer> oldAreas,
                                     List<Integer> newAreas,
                                     int roundingDifference,
                                     XSSFWorkbook workbook,
                                     XSSFSheet xlsxSheet);


    public XSSFWorkbook writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                             int newParcelNumber,
                                                             XSSFWorkbook workbook,
                                                             XSSFSheet xlsxSheet);

    public XSSFWorkbook writeDPRsInTemplate(List<Integer> orderedListOfDPRs,
                                            int newParcelNumber,
                                            XSSFWorkbook workbook,
                                            XSSFSheet xlsxSheet);

    public XSSFWorkbook writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                                  int dpr,
                                                  int area,
                                                  int newParcelNumber,
                                                  XSSFWorkbook workbook,
                                                  XSSFSheet xlsxSheet);

    public XSSFWorkbook writeNewDPRArea(int dpr,
                                        int area,
                                        int newParcelNumber,
                                        XSSFWorkbook workbook,
                                        XSSFSheet xlsxSheet);

    public XSSFWorkbook writeDPRRoundingDifference(int dpr,
                                                   int roundingDifference,
                                                   int newParcelNumber,
                                                   XSSFWorkbook workbook,
                                                   XSSFSheet xlsxSheet);



}
