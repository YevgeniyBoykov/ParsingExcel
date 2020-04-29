package com.devcolibri;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;



public class Main {

    public static void main(String... args) throws Exception {

        File xlsFile = new File("C:\\temp\\Book1.xlsx");

        FileInputStream fileIn = new FileInputStream(xlsFile);
        XSSFWorkbook book = new XSSFWorkbook(fileIn);


        XSSFSheet sheet = book.getSheetAt(0);
        System.out.println(sheet);

        for (Row row : sheet) {
            row.getLastCellNum();
            if (row.getLastCellNum() >= 2) {
                Cell cell = row.getCell(1);
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue() + "\t");
                        break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue() + "\t");
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue() + "\t");
                        break;
                    case FORMULA:
                        System.out.println(cell.getCellFormula() + "\t");
                        break;
                }

            }
        }
    }
}
