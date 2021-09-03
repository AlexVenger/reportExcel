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
    private static SXSSFSheet sheet;
    private static Map<String, CellStyle> styles = new HashMap<>();

    public static byte[] createExcel(List<String> headers, List<List<Object>> bodies) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        ExcelStyles.createStyles(workbook, styles);

        List<Object> body;

        sheet = workbook.createSheet();
        SXSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            SXSSFCell cell = row.createCell(i);
            cell.setCellStyle(styles.get("style2"));
            cell.setCellValue(headers.get(i));
        }

        for(int j = 0; j < bodies.size(); j++) {
            body = bodies.get(j);
            row = sheet.createRow(j + 1);
            for (int i = 0; i < headers.size(); i++) {
                SXSSFCell cell = row.createCell(i);
                cell.setCellStyle(styles.get("style3"));
                cell.setCellValue(body.get(i).toString());
            }
        }

        ByteArrayOutputStream fileOut = new ByteArrayOutputStream();
        try {
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return fileOut.toByteArray();
    }
}
