package ch.so.agi.avgbs2mtab.writeexcel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;


public interface WriteExcel {

    public Workbook writeOldPropertiesInTemplate(List<Integer> orderedListOfOldProperties);

    public Workbook writeNewPropertiesInTemplate(List<Integer> orderedListOfNewProperties);

    public Workbook writeInflowAndOutflows(int oldProperty, int newProperty);

    public Workbook writeNewArea(int newProperty);

    public Workbook writeOldArea(int oldProperty);

    public Workbook writeRoundingDifference(int oldProperty);

    public Workbook writePropertiesInTemplate(List<Integer> orderedListOfProperties);

    public Workbook writeDPRInTemplate(List<String> orderedListOfDPR);

    public Workbook writeDPRInflowAndOutflows(int property, int dpr);

    public Workbook writeNewDPRArea(int dpr);

    public Workbook writeDPRRoundingDifference(int dpr);



}
