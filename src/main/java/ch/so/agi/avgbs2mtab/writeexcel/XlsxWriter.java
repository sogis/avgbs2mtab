package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XlsxWriter {

    private DataExtractionDPR dataExtractionDPR;
    private DataExtractionParcel dataExtractionParcel;
    private MetadataOfDPRMutation metadataOfDPRMutation;
    private MetadataOfParcelMutation metadataOfParcelMutation;

    public XlsxWriter(DataExtractionParcel dataExtractionParcel, DataExtractionDPR dataExtractionDPR,
                      MetadataOfParcelMutation metadataOfParcelMutation, MetadataOfDPRMutation metadataOfDPRMutation) {
        this.dataExtractionDPR = dataExtractionDPR;
        this.dataExtractionParcel = dataExtractionParcel;
        this.metadataOfDPRMutation = metadataOfDPRMutation;
        this.metadataOfParcelMutation = metadataOfParcelMutation;

    }

    public void writeXlsx(String filePath) {

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        //todo: Info-log with message: "start generating and writing excel"
        //todo: Debug-log with message: "start generating excel-template"
        XSSFWorkbook workbook = xlsxTemplate.createExcelTemplate(filePath, metadataOfParcelMutation, metadataOfDPRMutation);
        //todo: Debug-log with message: "finished generating excel-template; start writing parcels into excel-template"
        workbook = excelData.fillValuesIntoParcelTable(filePath, workbook, dataExtractionParcel);
        //todo: Debug-log with message: "finished writing parcels into excel-template; start writing dprs into excel-templat"
        excelData.fillValuesIntoDPRTable(filePath, workbook, dataExtractionDPR, metadataOfParcelMutation);
        //todo: Debug-log with message: "finished writing dprs into excel-template"

        //todo: Info-log with message: "avgbs2mtab finished"
    }
}
