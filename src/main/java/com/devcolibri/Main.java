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
//        for (String str : arrTxt)
//            System.out.println(str);
        for(int sheetNum = 0; sheetNum < book.getNumberOfSheets(); sheetNum++)
        {
            XSSFSheet sheet = book.getSheetAt(sheetNum);

            for (Row row : sheet) {
                String checkStringFromExcel;
                if (row.getLastCellNum() >= 2) {
                    Cell cell = row.getCell(1);
                    checkStringFromExcel = cell.getStringCellValue();
//                    System.out.println(checkStringFromExcel);
                    boolean isConsist = false;

                    for(int i = 0; i < arrTxt.length; i++){
                        if (checkStringFromExcel.equals(arrTxt[i])){
                            isConsist = true;
                        }
//                        switch (cell.getCellType()) {
//                            case STRING:
//                                checkStringFromExcel = cell.getStringCellValue();
//                                break;
//                            case NUMERIC:
//                                checkStringFromExcel = cell.getNumericCellValue();
//                                break;
//                            case BOOLEAN:
//                                System.out.println(cell.getBooleanCellValue() + "\t");
//                                break;
//                            case FORMULA:
//                                System.out.println(cell.getCellFormula() + "\t");
//                                break;
                    }
                    if (isConsist) {

                    } else{
                        System.out.println("Item " + checkStringFromExcel + " not found in " + book.getSheetName(sheetNum)+ "!");
                    }
                }
            }
        }
    }
}
