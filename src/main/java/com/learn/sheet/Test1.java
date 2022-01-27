package com.learn.sheet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * @author ：Kristen
 * @date ：2022/1/27
 * @description : 写入 xlsx
 */
public class Test1 {
    public static void main(String[] args) throws Exception {
        //Create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");
        //Create row object
        XSSFRow row;
        //This data needs to be written (Object[])
        Map<String, Object[]> map = new TreeMap<>();
        map.put("1", new Object[]{"EMP ID", "EMP NAME", "DESIGNATION"});
        map.put("2", new Object[]{"tp01", "Gopal", "Technical Manager"});
        map.put("3", new Object[]{"tp02", "Manisha", "Proof Reader"});
        map.put("4", new Object[]{"tp03", "Masthan", "Technical Writer"});
        map.put("5", new Object[]{"tp04", "Satish", "Technical Writer"});
        map.put("6", new Object[]{"tp05", "Krishna", "Technical Writer"});
        //Iterate over data and write to sheet
        Set<String> set = map.keySet();
        int rowid = 0;
        for (String key : set) {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = map.get(key);
            int id = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(id++);
                cell.setCellValue((String) obj);
            }
        }
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File("E:/Write-sheet.xlsx"));
        workbook.write(out);
        out.close();
        System.out.println("Write-sheet.xlsx written successfully");
    }
}
