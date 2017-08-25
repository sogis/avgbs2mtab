package ch.so.agi.avgbs2mtab.mutdat;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


import java.io.*;
import java.util.logging.Logger;

public class XLSXTemplate implements ExcelTemplate {

    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());



    public XSSFWorkbook createExcelTemplate (String filePath,MetadataOfParcelMutation metadataOfParcelMutation,
                                             MetadataOfDPRMutation metadataOfDPRMutation ) {



        Integer numberOfNewParcels = metadataOfParcelMutation.getNumberOfNewParcels();
        Integer numberOfOldParcels = metadataOfParcelMutation.getNumberOfOldParcels();

        Integer numberOfParcelsAffectedByDPRs = metadataOfDPRMutation.getNumberOfParcelsAffectedByDPRs();
        Integer numberOfDPRs = metadataOfDPRMutation.getNumberOfDPRs();

        return generateWorkbookTemplate(filePath, numberOfNewParcels, numberOfOldParcels,
                numberOfParcelsAffectedByDPRs, numberOfDPRs);
    }

    private XSSFWorkbook generateWorkbookTemplate(String filePath, int numberOfNewParcels, int numberOfOldParcels,
                                                  int numberOfParcelsAffectedByDPRs, int numberOfDPRs){

        LOGGER.log(java.util.logging.Level.FINER, "Start creating Excel-Workbook");
        XSSFWorkbook workbook = createWorkbook(filePath);
        LOGGER.log(java.util.logging.Level.FINER, "Finished creating Excel-Workbook; Start creating empty table" +
                "with parcels");
        workbook = createParcelTable(workbook, filePath, numberOfNewParcels, numberOfOldParcels,
                numberOfParcelsAffectedByDPRs);
        LOGGER.log(java.util.logging.Level.FINER, "Finished creating table with parcels; Start creating empty " +
                "table with dprs");
        workbook = createDPRTable(workbook, filePath, numberOfParcelsAffectedByDPRs, numberOfDPRs,
                numberOfNewParcels, numberOfOldParcels);
        LOGGER.log(java.util.logging.Level.FINER, "Finished creating table with dprs");

        return workbook;
    }


    @Override
    public XSSFWorkbook createWorkbook(String filePath) {

        try {
            OutputStream ExcelFile = new FileOutputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook();
            workbook.createSheet("Mutationstabelle");
            workbook.write(ExcelFile);

            return workbook;

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public XSSFWorkbook createParcelTable(XSSFWorkbook excelTemplate,String filePath, int newParcels, int oldParcels,
                                          int parcelsAffectedByDPR) {

        XSSFSheet sheet = excelTemplate.getSheet("Mutationstabelle");

        sheet = addMergedRegions(sheet, oldParcels);

        if (oldParcels == 0 || newParcels == 0){
            oldParcels = 1;
            newParcels = 1;
        }

        sheet = setCellSize(sheet, oldParcels, parcelsAffectedByDPR);


        for (int i = 0; i < (newParcels*2+5) + 1; i++){
            Row row =sheet.createRow(i);

            if (i==0) {

                stylingFirstParcelRow(row, oldParcels, excelTemplate);

            } else if (i==1) {

                stylingSecondParcelRow(row, oldParcels, excelTemplate);

            } else if (i==2) {

                stylingThirdParcelRow(row, oldParcels, excelTemplate);


            } else if (i==newParcels*2+5) {

                stylingLastParcelRow(row, oldParcels, excelTemplate);


            } else {

                stylingEveryOtherParcelRow(row, oldParcels, newParcels, i, excelTemplate);

            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            out.close();

        } catch (IOException e) {
            throw new RuntimeException(e);

        }

        return excelTemplate;

    }

    private XSSFSheet addMergedRegions(XSSFSheet sheet, int oldParcels) {
        if (oldParcels>1){
            sheet.addMergedRegion(new CellRangeAddress(0,0,1, oldParcels));
            sheet.addMergedRegion(new CellRangeAddress(1,1,1, oldParcels));
        }
        return sheet;
    }

    private XSSFSheet setCellSize(XSSFSheet sheet, int oldParcels, int parcelsAffectedByDPR){

        sheet = setDefaultCellSize(sheet);
        sheet = setColumnHeight(sheet, oldParcels, parcelsAffectedByDPR);

        return sheet;
    }

    private XSSFSheet setDefaultCellSize(XSSFSheet sheet){
        sheet.setDefaultRowHeight((short) 300);
        sheet.setDefaultColumnWidth((short) 18.43);

        return sheet;
    }

    private XSSFSheet setColumnHeight(XSSFSheet sheet, int oldParcels, int parcelsAffectedByDPR) {
        sheet.setColumnWidth(0,19*253);
        if (oldParcels >= parcelsAffectedByDPR) {
            sheet.setColumnWidth(oldParcels + 1, 14 * 253);
            sheet.setColumnWidth(oldParcels + 2, 14 * 253);
        }
        return sheet;

    }

    private void  stylingFirstParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate) {
        Cell cell;
        for (int c = 1; c <= oldParcels; c++){
            cell = row.createCell(c);
            cell.setCellValue("Alte Liegenschaften");

            String color = "lightGray";
            String border_bottom = "";
            String border_top = "thick";
            String border_left = "";
            String border_right = "";

            if (c==1) {
                if (oldParcels == 1) {
                    border_left ="thick";
                    border_right = "thick";
                } else if (oldParcels > 1) {
                    border_left = "thick";
                }
            } else if (c==oldParcels) {
                border_right = "thick";
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom,
                    border_top, border_left, border_right, 0, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight((short) 600);

    }

    private XSSFCellStyle getStyleForCell(String color, String border_bottom, String border_top, String border_left,
                                          String border_right, int indent, XSSFWorkbook excelTemplate ) {

        XSSFCellStyle style = excelTemplate.createCellStyle();

        XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217,217));

        XSSFFont font = excelTemplate.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setFontName("Arial");
        font.setItalic(false);

        style.setFont(font);


        if (color.equals("lightGray")){
            style.setFillForegroundColor(lightGray);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setWrapText(true);
        } else {
            style.setVerticalAlignment(VerticalAlignment.BOTTOM);
            style.setAlignment(HorizontalAlignment.RIGHT);
            style.setIndention((short) indent);
        }

        switch (border_bottom) {
            case "thick":
                style.setBorderBottom(BorderStyle.THICK);
                break;
            case "thin":
                style.setBorderBottom(BorderStyle.THIN);
                break;
            case "":
                style.setBorderBottom(BorderStyle.NONE);
        }

        switch (border_top) {
            case "thick":
                style.setBorderTop(BorderStyle.THICK);
                break;
            case "thin":
                style.setBorderTop(BorderStyle.THIN);
                break;
            case "":
                style.setBorderTop(BorderStyle.NONE);
        }

        switch (border_left) {
            case "thick":
                style.setBorderLeft(BorderStyle.THICK);
                break;
            case "thin":
                style.setBorderLeft(BorderStyle.THIN);
                break;
        }

        switch (border_right) {
            case "thick":
                style.setBorderRight(BorderStyle.THICK);
                break;
            case "thin":
                style.setBorderRight(BorderStyle.THIN);
                break;
        }

        return style;

    }

    private void stylingSecondParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate){
        Cell cell;
        for (int c = 0; c <= oldParcels +1; c++){
            cell = row.createCell(c);
            String color = "lightGray";
            String border_bottom = "thin";
            String border_top = "thin";
            String border_left = "thick";
            String border_right = "thick";

            if (c==0){
                border_top = "thick";
                cell.setCellValue("Neue Liegenschaften");
            } else if (c==1) {
                if (oldParcels > 1) {
                    border_right = "thin";
                }
                cell.setCellValue("Grundstück-Nr.");
            } else if (c==oldParcels){
                border_left = "thin";
            } else if (c<oldParcels) {
                border_left = "thin";
                border_right = "thin";
            } else if (c==oldParcels+1){
                border_top = "thick";
                cell.setCellValue("Neue Fläche");
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    0, excelTemplate);
            cell.setCellStyle(newStyle);
        }


        row.setHeight((short) 600);
    }

    private void stylingThirdParcelRow (Row row, int oldParcels, XSSFWorkbook excelTemplate){
        Cell cell;
        for (int c = 0; c <= oldParcels +1; c++){
            cell = row.createCell(c);

            String color = "";
            String border_bottom = "thick";
            String border_top = "thin";
            String border_left = "thick";
            String border_right = "thick";
            Integer indent = 2;


            if (c==0){
                color = "lightGray";
                indent = 0;
                cell.setCellValue("Grundstück-Nr.");
            } else if (c==1) {
                if (oldParcels > 1) {
                    border_right = "thin";
                }
            } else if (c==oldParcels){
                border_left = "thin";
            } else if (c<oldParcels) {
                border_left = "thin";
                border_right = "thin";
            } else if (c==oldParcels+1){
                color = "lightGray";
                indent = 0;
                cell.setCellValue("[m2]");
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight((short) 600);
    }

    private void stylingLastParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate){
        Cell cell;
        for (int c = 0; c <= oldParcels + 1; c++) {
            cell = row.createCell(c);

            String color = "";
            String border_bottom = "thick";
            String border_top = "thick";
            String border_left = "thick";
            String border_right = "thick";
            Integer indent = 2;

            if (c == 0) {
                color = "lightGray";
                indent = 0;
                cell.setCellValue("Alte Fläche [m2]");
            } else if (c == 1) {
                if (oldParcels > 1) {
                    border_right = "thin";
                }
            } else if (c == oldParcels) {
                border_left = "thin";
            } else if (c < oldParcels) {
                border_left = "thin";
                border_right = "thin";
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight((short) 600);
    }

    private void stylingEveryOtherParcelRow(Row row, int oldParcels, int newParcels, int i, XSSFWorkbook excelTemplate){
        Cell cell;
        String border_bottom;
        String border_top;

        if (i % 2 == 0){
            border_bottom = "thin";
            border_top = "";
        } else {
            border_bottom = "";
            border_top = "thin";
        }


        for (int c = 0; c <= oldParcels + 1; c++) {
            cell = row.createCell(c);

            String color = "";
            String border_left = "thick";
            String border_right = "thick";
            Integer indent = 2;

            if (c == 0) {
                if (i==newParcels*2+5-1){
                    border_bottom = "thick";
                    border_top = "";
                    indent = 0;
                    cell.setCellValue("Rundungsdifferenz");
                }
            } else if (c == 1) {
                if (oldParcels > 1) {
                    border_right = "thin";
                }
            } else if (c == oldParcels) {
                border_left = "thin";
            } else if (c < oldParcels) {
                border_left = "thin";
                border_right = "thin";
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }
    }




    @Override
    public XSSFWorkbook createDPRTable(XSSFWorkbook excelTemplate, String filePath, int parcels, int dpr,
                                       int newParcels, int oldParcels) {

        if (newParcels == 0 ){
            newParcels = 1;
        } else if (oldParcels == 0){
            oldParcels = 1;
        }

        if (dpr==0){
            dpr = 1;
            parcels = 1;
        } else if (parcels == 0){
            dpr = 1;
            parcels = 1;
        }

        int rowStartIndex = (9 + 2 * newParcels - 1);

        XSSFSheet sheet = excelTemplate.getSheet("Mutationstabelle");

        sheet = setColumnWidth(parcels, oldParcels, sheet);

        sheet = addMergedRegionsDPR(rowStartIndex, parcels, sheet);


        for (int i = rowStartIndex; i < (rowStartIndex + 3 + 2 * dpr); i++) {
            Row row = sheet.createRow(i);
            if (i == rowStartIndex) {

                stylingFirstDPRRow(row, parcels, excelTemplate);

            } else if (i == rowStartIndex + 1) {

                stylingSecondDPRRow(row, parcels, oldParcels, excelTemplate);

            } else if (i == rowStartIndex + 2) {

                stylingRowWithDPRNumber(row, parcels, excelTemplate);

            }  else {

                stylingEveryOtherDPRRow(row, i, rowStartIndex, parcels, dpr, excelTemplate);

            }
        }


        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return excelTemplate;
    }

    private XSSFSheet setColumnWidth(int parcels, int oldParcels, XSSFSheet sheet){
        if (parcels > oldParcels) {
            sheet.setColumnWidth(parcels+1,14*253);
            sheet.setColumnWidth(parcels+2,14*253);
        }
        return sheet;
    }

    private XSSFSheet addMergedRegionsDPR(int rowStartIndex, int parcels, XSSFSheet sheet){
        if (parcels>1){
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex, rowStartIndex,1, parcels));
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex + 1,rowStartIndex + 1,
                    1, parcels));
        }

        sheet.addMergedRegion((new CellRangeAddress(rowStartIndex + 1, rowStartIndex + 2,
                parcels + 1, parcels + 1)));

        return sheet;
    }

    private void stylingFirstDPRRow(Row row, int parcels, XSSFWorkbook excelTemplate){
        Cell cell;

        for (int c = 1; c <= parcels; c++) {
            cell = row.createCell(c);
            cell.setCellValue("Liegenschaften");

            String color = "lightGray";
            String border_bottom = "";
            String border_top = "thick";
            String border_left = "";
            String border_right = "";

            if (c == 1) {
                if (parcels == 1) {
                    border_left = "thick";
                    border_right = "thick";
                } else if (parcels > 1) {
                    border_left="thick";
                }
            } else if (c == parcels) {
                border_right = "thick";
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    0, excelTemplate);
            cell.setCellStyle(newStyle);

        }
        row.setHeight((short) 600);
    }

    private void stylingSecondDPRRow (Row row, int parcels, int oldParcels, XSSFWorkbook excelTemplate){
        Cell cell;

        for (int c = 0; c <= parcels + 2; c++){
            cell = row.createCell(c);

            String color = "lightGray";
            String border_bottom = "thin";
            String border_top = "thin";
            String border_left = "thin";
            String border_right = "thin";

            if (c==0){
                border_top = "thick";
                border_left = "thick";
                border_right = "thick";
                cell.setCellValue("Selbst. Recht");

            } else if (c==1) {
                border_left = "thick";
                cell.setCellValue("Grundstück-Nr.");
            } else if (c==parcels+1){
                border_bottom = "";
                border_top = "thick";
                border_right = "thick";
                if (parcels >= oldParcels) {
                    cell.setCellValue("Rundungs-differenz");
                } else {
                    cell.setCellValue("Rundungsdifferenz");
                }
            }else if (c==parcels + 2){
                border_top = "thick";
                border_left = "thick";
                border_right = "thick";
                cell.setCellValue("Selbst. Recht Fläche");
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    0, excelTemplate);
            if (c==0){
                newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            }
            cell.setCellStyle(newStyle);
        }

        row.setHeight((short) 600);
    }

    private void stylingRowWithDPRNumber(Row row, int parcels, XSSFWorkbook excelTemplate) {

        Cell cell;
        for (int c = 0; c <= parcels + 2; c++) {
            cell = row.createCell(c);

            String color = "lightGray";
            String border_bottom = "thick";
            String border_top = "thin";
            String border_left = "thin";
            String border_right = "thick";
            Integer indent = 0;


            if (c == 0) {
                border_left = "thick";
                cell.setCellValue("Grundstück-Nr.");

            } else if (c == 1) {
                color = "";
                border_left = "thick";
                border_right = "thin";
                indent = 2;
            } else if (c <= parcels && parcels != 1) {
                color = "";
                border_right = "thin";
                indent = 2;
            } else if (c == parcels + 1) {
                border_top = "";
            } else if (c == parcels + 2) {
                border_left = "thick";
                cell.setCellValue("[m2]");
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            if (c==0){
                newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            }
            cell.setCellStyle(newStyle);
        }

        row.setHeight((short) 600);
    }

    private void stylingEveryOtherDPRRow(Row row, int i, int rowStartIndex, int parcels, int dpr,
                                         XSSFWorkbook excelTemplate ) {
        String border_bottom;
        String border_top;
        Cell cell;

        if ((i-rowStartIndex) % 2 == 0){
            border_bottom = "thin";
            border_top = "";
        } else {
            border_bottom = "";
            border_top = "thin";
        }
        for (int c = 0; c <= parcels + 2; c++) {
            cell = row.createCell(c);

            String color = "";
            String border_left = "thick";
            String border_right = "thick";

            if (c == 0) {
                if (i==dpr * 2 + rowStartIndex + 2){
                    border_bottom = "thick";
                    border_top = "";
                    cell.setCellType(CellType.STRING);
                } else {
                    if (border_bottom.equals("thin")){
                        cell.setCellType(CellType.STRING);
                    }
                }
            } else if (c == 1) {
                if (i==dpr * 2 + rowStartIndex + 2){
                    border_bottom = "thick";
                    border_top = "";
                    border_right = "thin";
                } else {
                    border_right = "thin";
                }
            } else if (c <= parcels && parcels != 1) {
                if (i==dpr * 2 + rowStartIndex + 2){
                    border_bottom = "thick";
                    border_top = "";
                    border_left = "thin";
                    border_right = "thin";
                } else {
                    border_left = "thin";
                    border_right = "thin";
                }
            } else if (c == parcels + 1) {
                if (i==dpr * 2 + rowStartIndex + 2){
                    border_bottom = "thick";
                    border_top = "";
                    border_left = "thin";
                } else {
                    border_left = "thin";
                }
            } else if (c == parcels + 2) {
                if (i==dpr * 2 + rowStartIndex + 2){

                    border_bottom = "thick";
                    border_top = "";
                }
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    2, excelTemplate);
            cell.setCellStyle(newStyle);
        }
    }

}
