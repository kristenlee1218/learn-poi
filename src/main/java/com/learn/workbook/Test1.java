package com.learn.workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author ：Kristen
 * @date ：2022/1/27
 * @description : CreateWorkBook
 */
public class Test1 {
    public static void main(String[] args) throws IOException {
        //Create Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(new File("E:/create-workbook.xlsx"));
        //write operation workbook using file out object
        workbook.write(out);
        out.close();
        System.out.println("create-workbook.xlsx written successfully");
    }
}
