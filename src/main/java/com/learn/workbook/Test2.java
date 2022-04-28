package com.learn.workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

/**
 * @author ：Kristen
 * @date ：2022/1/27
 * @description : OpenWorkBook
 */
public class Test2 {
    public static void main(String[] args) throws Exception {
        File file = new File("/Users/kristen/IdeaProjects/docu/openwork-book.xlsx");
        FileInputStream fip = new FileInputStream(file);
        //Get the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(fip);
        if (file.isFile() && file.exists()) {
            System.out.println("openwork-book.xlsx file open successfully.");
        } else {
            System.out.println("Error to open openwork-book.xlsx file.");
        }
    }
}
