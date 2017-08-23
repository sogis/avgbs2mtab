package ch.so.agi.avgbs2mtab.writeexcel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.List;


public interface WriteExcel {

    public XSSFWorkbook writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                                  String filePath,
                                                  XSSFWorkbook workbook);

    public XSSFWorkbook writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                              String filePath,
                                              XSSFWorkbook workbook);

    public XSSFWorkbook writeInflowAndOutflows(int oldParcelNumber,
                                           int newParcelNumber,
                                           int area,
                                           String filePath,
                                           XSSFWorkbook workbook);

    public XSSFWorkbook writeNewArea(int newParcelNumber,
                                 int area,
                                 String filePath,
                                 XSSFWorkbook workbook);

    public XSSFWorkbook writeRoundingDifference(int oldParcelNumber,
                                            int roundingDifference,
                                            int numberOfNewParcels,
                                            String filePath,
                                            XSSFWorkbook workbook);

    public XSSFWorkbook writeSumOfRoundingDifference (int NumberOfNewParcels,
                                                      int NumberOfOldParcels,
                                                      int roundingDifferenceSum,
                                                      String filePath,
                                                      XSSFWorkbook workbook);


    public XSSFWorkbook writeOldArea(int oldParcelNumber,
                                 int oldArea,
                                 int roundingDifference,
                                 int numberOfNewParcels,
                                 String filePath,
                                 XSSFWorkbook workbook);

    public XSSFWorkbook writeAreaSum(HashMap<Integer, Integer> oldAreas,
                                 List<Integer> newAreas,
                                 int roundingDifference,
                                 String filePath,
                                 XSSFWorkbook workbook);


    public XSSFWorkbook writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                         int newParcelNumber,
                                                         String filePath,
                                                         XSSFWorkbook workbook);

    public XSSFWorkbook writeDPRsInTemplate(List<Integer> orderedListOfDPRs,
                                        int newParcelNumber,
                                        String filePath,
                                        XSSFWorkbook workbook);

    public XSSFWorkbook writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                              int dpr,
                                              int area,
                                              int newParcelNumber,
                                              String filePath,
                                              XSSFWorkbook workbook);

    public XSSFWorkbook writeNewDPRArea(int dpr,
                                    int area,
                                    int newParcelNumber,
                                    String filePath,
                                    XSSFWorkbook workbook);

    public XSSFWorkbook writeDPRRoundingDifference(int dpr,
                                               int roundingDifference,
                                               int newParcelNumber,
                                               String filePath,
                                               XSSFWorkbook workbook);



}
