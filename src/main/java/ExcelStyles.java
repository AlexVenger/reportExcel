import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.HashMap;
import java.util.Map;

public class ExcelStyles {
    public static void createStyles(Workbook workbook, Map<String, CellStyle> styles) {

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
        Font titleFont = workbook.createFont();
        titleCellStyle.setWrapText(true);
        titleFont.setFontName("Arial");
        titleFont.setColor(IndexedColors.WHITE.getIndex());
        titleFont.setBold(true);
        titleCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(1, 42, 73)));
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderLeft(BorderStyle.MEDIUM);

        styles.put("style2", titleCellStyle);

        CellStyle dataStyle = workbook.createCellStyle();
        Font dataFont = workbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setColor(IndexedColors.BLACK.getIndex());
        dataStyle.setFont(dataFont);
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
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
