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

        XSSFWorkbook workbook = xlsxTemplate.createExcelTemplate(filePath, metadataOfParcelMutation, metadataOfDPRMutation);
        workbook = excelData.fillValuesIntoParcelTable(filePath, workbook, dataExtractionParcel);
        excelData.fillValuesIntoDPRTable(filePath, workbook, dataExtractionDPR, metadataOfParcelMutation);
    }
}
