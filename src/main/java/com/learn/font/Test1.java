package com.learn.font;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;
import java.io.FileOutputStream;

/**
 * @author : Kristen
 * @date : 2022/1/29
 * @description : 设置单元格字体
 */
public class Test1 {
    public static void main(String[] args) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Fontstyle");
        XSSFRow row = spreadsheet.createRow(2);
        //Create a new font and alter it.
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 30);
        font.setFontName("IMPACT");
        font.setItalic(true);
        font.setColor(HSSFColor.BRIGHT_GREEN.index);
        //Set font into style
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        // Create a cell with a value and set style to it.
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("Font Style");
        cell.setCellStyle(style);
        FileOutputStream out = new FileOutputStream("/Users/kristen/dev-software/font-style.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("font-style.xlsx written successfully");
    }
}
