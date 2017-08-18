package ch.so.agi.avgbs2mtab.writeexcel;


import com.sun.xml.internal.ws.api.pipe.FiberContextSwitchInterceptor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;


public interface WriteExcel {

    public XSSFWorkbook writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                                  String filePath,
                                                  XSSFWorkbook workbook);

    public Workbook writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                              String filePath,
                                              XSSFWorkbook workbook);

    public Workbook writeInflowAndOutflows(int oldParcelNumber,
                                           int newParcelNumber,
                                           int area,
                                           String filePath,
                                           XSSFWorkbook workbook);

    public Workbook writeNewArea(int newParcelNumber,
                                 int area,
                                 String filePath,
                                 XSSFWorkbook workbook);

    public Workbook writeRoundingDifference(int oldParcelNumber,
                                            int roundingDifference,
                                            String filePath,
                                            XSSFWorkbook workbook);



    public Workbook writeOldArea(int oldParcelNumber,
                                 List<Integer> areaOutflowsOfOldParcelNumber,
                                 int roundingDifference,
                                 String filePath,
                                 XSSFWorkbook workbook);

    public Workbook writeAreaSum(List<Integer> oldAreas,
                                 List<Integer> newAreas,
                                 int roundingDifference,
                                 String filePath,
                                 XSSFWorkbook workbook);


    public Workbook writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                         String filePath,
                                                         XSSFWorkbook workbook);

    public Workbook writeDPRsInTemplate(List<String> orderedListOfDPRs,
                                        String filePath,
                                        XSSFWorkbook workbook);

    public Workbook writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                              int dpr,
                                              String filePath,
                                              XSSFWorkbook workbook);

    public Workbook writeNewDPRArea(int dpr,
                                    String filePath,
                                    XSSFWorkbook workbook);

    public Workbook writeDPRRoundingDifference(int dpr,
                                               String filePath,
                                               XSSFWorkbook workbook);



}
