package com.devcolibri;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;


public class Main {

    public static void main(String... args) throws Exception {
        String txtFile = "C:\\temp\\2\\test.txt";
        File xlsFile = new File("C:\\temp\\2\\Book1.xlsx");

        FileInputStream fileIn = new FileInputStream(xlsFile);
        XSSFWorkbook book = new XSSFWorkbook(fileIn);
        ReadFile file = new ReadFile();
        String[] arrTxt = file.readFile(txtFile).split(";");
        for (String str : arrTxt)
            System.out.println(str);
//        System.out.println(arrTxt);

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
