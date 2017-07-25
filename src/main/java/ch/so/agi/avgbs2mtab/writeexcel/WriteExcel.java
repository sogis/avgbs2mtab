package ch.so.agi.avgbs2mtab.writeexcel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;


public interface WriteExcel {

    public Workbook writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers);

    public Workbook writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers);

    public Workbook writeInflowAndOutflows(int oldParcelNumber, int newParcelNumber);

    public Workbook writeNewArea(int newParcelNumber);

    public Workbook writeRoundingDifference(int oldParcelNumber);


    public Workbook writeParcelsAffectedByDPRInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs);

    public Workbook writeDPRInTemplate(List<String> orderedListOfDPR);

    public Workbook writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR, int dpr);

    public Workbook writeNewDPRArea(int dpr);

    public Workbook writeDPRRoundingDifference(int dpr);



}
