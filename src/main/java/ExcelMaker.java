import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelMaker {
    private static Map<String, CellStyle> styles;
    private static SXSSFSheet sheet;

    public static byte[] createExcel(ArrayList<String> headers, ArrayList<ArrayList<Object>> bodies /*ArrayList<ReportRow> reportRows*/) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        createStyles(workbook);

//        ArrayList<String> headers = Report1Row.headers;
        ArrayList<Object> body;

        int i;
        sheet = workbook.createSheet();
        SXSSFRow row = sheet.createRow(0);
        for (i = 0; i < headers.size(); i++) {
            SXSSFCell cell = row.createCell(i);
            cell.setCellStyle(styles.get("style2"));
            cell.setCellValue(headers.get(i));
        }

        for(int j = 0; j < bodies.size(); j++) {
            body = bodies.get(j);
            row = sheet.createRow(j + 1);
            for (i = 0; i < headers.size(); i++) {
                SXSSFCell cell = row.createCell(i);
                cell.setCellStyle(styles.get("style3"));
                cell.setCellValue(body.get(i).toString());
            }
        }

//        for(int j = 0; j < reportRows.size(); j++) {
//            body = reportRows.get(j).getBody();
//            row = sheet.createRow(j + 1);
//            for (i = 0; i < headers.size(); i++) {
//                SXSSFCell cell = row.createCell(i);
//                cell.setCellStyle(styles.get("style3"));
//                cell.setCellValue(body.get(i).toString());
//            }
//        }

        ByteArrayOutputStream fileOut = new ByteArrayOutputStream();
        try {
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return fileOut.toByteArray();
    }

    private static void createStyles(Workbook workbook) {
        styles = new HashMap<>();

        CellStyle headerCellStyle = workbook.createCellStyle();

        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 24);
        headerFont.setFontName("Arial");


        headerFont.setColor(IndexedColors.BLACK.getIndex());

        headerFont.setBold(true);
        headerFont.setItalic(false);
        headerCellStyle.setFont(headerFont);


        styles.put("style1", headerCellStyle);



        XSSFCellStyle titleCellStyle = (XSSFCellStyle) workbook.createCellStyle();
//        titleCellStyle.setWrapText(true);
        Font titleFont = workbook.createFont();
//        titleFont.setFontHeightInPoints((short) 12);
        titleFont.setFontName("Arial");
        titleFont.setColor(IndexedColors.BLACK.getIndex());
        titleFont.setBold(true);
//        titleCellStyle.setAlignment(HorizontalAlignment.FILL);
//        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(220, 230, 241)));
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderLeft(BorderStyle.MEDIUM);

        styles.put("style2", titleCellStyle);

        CellStyle dataStyle = workbook.createCellStyle();
        Font dataFont = workbook.createFont();
//        dataStyle.setWrapText(true);
//        dataFont.setFontHeightInPoints((short) 10);
        dataFont.setFontName("Arial");
        dataFont.setColor(IndexedColors.BLACK.getIndex());
//        dataStyle.setAlignment(HorizontalAlignment.FILL);
//        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setFont(dataFont);
        dataStyle.setBorderBottom(BorderStyle.MEDIUM);
        dataStyle.setBorderTop(BorderStyle.MEDIUM);
        dataStyle.setBorderRight(BorderStyle.MEDIUM);
        dataStyle.setBorderLeft(BorderStyle.MEDIUM);

        styles.put("style3", dataStyle);


        CellStyle newStyle = workbook.createCellStyle();
        newStyle.setWrapText(true);

        newStyle.setAlignment(HorizontalAlignment.CENTER);
        newStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font newFont = workbook.createFont();
        newFont.setFontHeightInPoints((short) 11);
        newFont.setFontName("Arial");
        newFont.setColor(IndexedColors.BLACK.getIndex());
        newStyle.setFont(newFont);
        styles.put("style4", newStyle);

        CellStyle newStyle1 = workbook.createCellStyle();
        newStyle1.setWrapText(true);
        newStyle1.setAlignment(HorizontalAlignment.CENTER);
        newStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        Font newFont1 = workbook.createFont();
        newFont1.setFontHeightInPoints((short) 18);
        newFont1.setFontName("Arial");
        newFont1.setColor(IndexedColors.BLACK.getIndex());
        newFont1.setBold(true);
        newStyle1.setFont(newFont1);
        styles.put("style5", newStyle1);

        CellStyle sumStyle = workbook.createCellStyle();
        sumStyle.setWrapText(true);
        sumStyle.setAlignment(HorizontalAlignment.CENTER);
        sumStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font sumFont = workbook.createFont();
        sumFont.setFontHeightInPoints((short) 12);
        sumFont.setFontName("Arial");
        sumFont.setColor(IndexedColors.BLACK.getIndex());
        sumFont.setBold(true);
        sumStyle.setFont(sumFont);
        sumStyle.setBorderTop(BorderStyle.MEDIUM);
        styles.put("style6", sumStyle);


        CellStyle newStyleDoubles = workbook.createCellStyle();
        newStyleDoubles.setWrapText(true);

        newStyleDoubles.setAlignment(HorizontalAlignment.CENTER);
        newStyleDoubles.setVerticalAlignment(VerticalAlignment.CENTER);
        Font newFontDoubles = workbook.createFont();
        newFontDoubles.setFontHeightInPoints((short) 11);
        newFontDoubles.setFontName("Arial");
        newFontDoubles.setColor(IndexedColors.BLACK.getIndex());
        newStyleDoubles.setFont(newFontDoubles);
        styles.put("style7", newStyleDoubles);
    }
}
