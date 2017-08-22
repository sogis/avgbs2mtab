package ch.so.agi.avgbs2mtab.mutdat;

import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class XLSXTemplate implements ExcelTemplate {

    @Override
    public XSSFWorkbook createWorkbook(String filePath) {

        int lastSlash = filePath.lastIndexOf("/");
        String pathWithoutFilename = filePath.substring(0, lastSlash+1);

        Path xlsxFilePath = Paths.get(pathWithoutFilename);

        if (Files.isWritable(xlsxFilePath)){
            try {
                OutputStream ExcelFile = new FileOutputStream(filePath);
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet xlsxSheet = workbook.createSheet("Mutationstabelle");
                workbook.write(ExcelFile);

                return workbook;

            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } else {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_NO_ACCESS_TO_FOLDER,
                    "Can not write in directory: " + filePath);
        }

    }

    @Override
    public XSSFWorkbook createParcelTable(XSSFWorkbook excelTemplate,String filePath, int newParcels, int oldParcels,
                                          int parcelsAffectedByDPR) {

        Sheet sheet = excelTemplate.getSheet("Mutationstabelle");

        sheet.setDefaultRowHeight((short) 300);
        sheet.setDefaultColumnWidth((short) 18.43);

        Cell cell;

        if (oldParcels>1){
            sheet.addMergedRegion(new CellRangeAddress(0,0,1,oldParcels));
            sheet.addMergedRegion(new CellRangeAddress(1,1,1,oldParcels));
        }

        if (oldParcels == 0 || newParcels == 0){
            oldParcels = 1;
            newParcels = 1;
        }


        sheet.setColumnWidth(0,19*253);
        if (oldParcels >= parcelsAffectedByDPR) {
            sheet.setColumnWidth(oldParcels + 1, 14 * 253);
            sheet.setColumnWidth(oldParcels + 2, 14 * 253);
        }

        for (int i = 0; i < (newParcels*2+5) + 1; i++){
            Row row =sheet.createRow(i);
            if (i==0) {
                for (int c = 1; c <= oldParcels; c++){
                    cell = row.createCell(c);
                    cell.setCellValue("Alte Liegenschaften");

                    if (c==1) {
                        if (oldParcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                    "thick", "thick", "thick", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (oldParcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                    "thick", "thick", "", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c==oldParcels && oldParcels!=1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                "thick", "", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                "thick", "", "", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    }
                }

                row.setHeight((short) 600);
            } else if (i==1) {
                for (int c = 0; c <= oldParcels +1; c++){
                    cell = row.createCell(c);

                    if (c==0){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thick", "thick", "thick", 0, excelTemplate);
                        newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Neue Liegenschaften");

                    } else if (c==1) {
                        if (oldParcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                    "thin", "thick", "thick", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (oldParcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                    "thin", "thick", "thin", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                        cell.setCellValue("Grundstück-Nr.");
                    } else if (c==oldParcels && oldParcels!=1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thin", "thin", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c<oldParcels) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thin", "thin", "thin", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c==oldParcels+1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thick", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Neue Fläche");
                    }
                }


                row.setHeight((short) 600);
            } else if (i==2) {
                for (int c = 0; c <= oldParcels +1; c++){
                    cell = row.createCell(c);

                    if (c==0){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "thin", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Grundstück-Nr.");
                    } else if (c==1) {
                        if (oldParcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "thin", "thick", "thick",2 , excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (oldParcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "thin", "thick", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c==oldParcels && oldParcels!=1){
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thin", "thin", "thick",2,  excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c<oldParcels) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thin", "thin", "thin", 2, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c==oldParcels+1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "thin", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("[m2]");
                    }
                }

                row.setHeight((short) 600);

            } else if (i==newParcels*2+5) {
                for (int c = 0; c <= oldParcels + 1; c++) {
                    cell = row.createCell(c);

                    if (c == 0) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "thick", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Alte Fläche [m2]");
                    } else if (c == 1) {
                        if (oldParcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "thick", "thick", "thick", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (oldParcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "thick", "thick", "thin", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c == oldParcels && oldParcels != 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thick", "thin", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c < oldParcels) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thick", "thin", "thin", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c == oldParcels + 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thick", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    }
                }

                row.setHeight((short) 600);

            } else {
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

                    if (c == 0) {
                        if (i==newParcels*2+5-1){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "", "thick", "thick", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                            cell.setCellValue("Rundungsdifferenz");


                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }

                    } else if (c == 1) {
                        if (oldParcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (oldParcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c == oldParcels && oldParcels != 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                border_top, "thin", "thick", 2, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c < oldParcels) {
                        XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                border_top, "thin", "thin", 2, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c == oldParcels + 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                border_top, "thick", "thick", 2, excelTemplate);
                        cell.setCellStyle(newStyle);
                    }
                }
            }
        }




        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            //out.close();
        } catch (FileNotFoundException e){

        } catch (IOException e) {

        }

        return (XSSFWorkbook) excelTemplate;

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
        }

        int rowStartIndex = (9 + 2 * newParcels - 1);

        Sheet sheet = excelTemplate.getSheet("Mutationstabelle");

        if (parcels > oldParcels) {
            sheet.setColumnWidth(parcels+1,14*253);
            sheet.setColumnWidth(parcels+2,14*253);
        }

        Cell cell;

        if (parcels>1){
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex, rowStartIndex,1, parcels));
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex + 1,rowStartIndex + 1,
                    1, parcels));
        }
        sheet.addMergedRegion((new CellRangeAddress(rowStartIndex + 1, rowStartIndex + 2,
                parcels + 1, parcels + 1)));


        for (int i = rowStartIndex; i < (rowStartIndex + 3 + 2 * dpr); i++) {
            Row row = sheet.createRow(i);
            if (i == rowStartIndex) {
                for (int c = 1; c <= parcels; c++) {
                    cell = row.createCell(c);
                    cell.setCellValue("Liegenschaften");
                    if (c == 1) {
                        if (parcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                    "thick", "thick", "thick", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (parcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                    "thick", "thick", "", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c == parcels && parcels != 1) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                "thick", "", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                "thick", "", "", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    }
                }
                row.setHeight((short) 600);


            } else if (i == rowStartIndex + 1) {
                for (int c = 0; c <= parcels + 2; c++){
                    cell = row.createCell(c);

                    if (c==0){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thick", "thick", "thick", 0, excelTemplate);
                        newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Selbst. Recht");

                    } else if (c==1) {
                        if (parcels == 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                    "thin", "thick", "thin", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        } else if (parcels > 1) {
                            XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                    "thin", "thick", "thin", 0, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                        cell.setCellValue("Grundstück-Nr.");
                    } else if (c==parcels && parcels!=1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thin", "thin", "thin", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c<parcels) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thin", "thin", "thin", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c==parcels+1){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "",
                                "thick", "thin", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        if (parcels >= oldParcels) {
                            cell.setCellValue("Rundungs-differenz");
                        } else {
                            cell.setCellValue("Rundungsdifferenz");
                        }
                    }else if (c==parcels + 2){
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thin",
                                "thick", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Selbst. Recht Fläche");
                    }
                }


                row.setHeight((short) 600);
            } else if (i == rowStartIndex + 2) {
                for (int c = 0; c <= parcels + 2; c++) {
                    cell = row.createCell(c);

                    if (c == 0) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "thin", "thick", "thick", 0, excelTemplate);
                        newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("Grundstück-Nr.");

                    } else if (c == 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick", "thin",
                                "thick", "thin", 2, excelTemplate);cell.setCellStyle(newStyle);
                    } else if (c <= parcels && parcels != 1) {
                        XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "thin", "thin", "thin", 2, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c == parcels + 1) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "", "thin", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                    } else if (c == parcels + 2) {
                        XSSFCellStyle newStyle = getStyleForCell("lightGray", "thick",
                                "thin", "thick", "thick", 0, excelTemplate);
                        cell.setCellStyle(newStyle);
                        cell.setCellValue("[m2]");
                    }
                }


                row.setHeight((short) 600);
            }  else {
                String border_bottom;
                String border_top;

                if ((i-rowStartIndex) % 2 == 0){
                    border_bottom = "thin";
                    border_top = "";
                } else {
                    border_bottom = "";
                    border_top = "thin";
                }
                for (int c = 0; c <= parcels + 2; c++) {
                    cell = row.createCell(c);

                    if (c == 0) {
                        if (i==dpr * 2 + rowStartIndex + 2){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                "", "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                            cell.setCellType(CellType.STRING);

                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                            if (border_bottom.equals("thin")){
                                cell.setCellType(CellType.STRING);
                            }
                        }

                    } else if (c == 1) {
                        if (i==dpr * 2 + rowStartIndex + 2){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "", "thick", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);

                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c <= parcels && parcels != 1) {
                        if (i==dpr * 2 + rowStartIndex + 2){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "", "thin", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);

                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thin", "thin", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c == parcels + 1) {
                        if (i==dpr * 2 + rowStartIndex + 2){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "", "thin", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);

                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thin", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    } else if (c == parcels + 2) {
                        if (i==dpr * 2 + rowStartIndex + 2){

                            cell = row.createCell(c);
                            XSSFCellStyle newStyle = getStyleForCell("", "thick",
                                    "", "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);

                        } else {
                            XSSFCellStyle newStyle = getStyleForCell("", border_bottom,
                                    border_top, "thick", "thick", 2, excelTemplate);
                            cell.setCellStyle(newStyle);
                        }
                    }
                }
            }
        }


        try {

            //todo: change path
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            //out.close();
        } catch (FileNotFoundException e){

        } catch (IOException e) {

        }

        return (XSSFWorkbook) excelTemplate;
    }


    private XSSFCellStyle getStyleForCell(String color, String border_bottom, String border_top, String border_left,
                                          String border_right, int indent, XSSFWorkbook excelTemplate ) {

        XSSFCellStyle style = (XSSFCellStyle) excelTemplate.createCellStyle();

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
}
