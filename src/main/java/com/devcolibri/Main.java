package com.devcolibri;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.regex.Pattern;


public class Main {

    public static void main(String... args) throws Exception {
        String txtFile = "C:\\temp\\2\\waitlist.txt";
        File xlsFile = new File("C:\\temp\\2\\CV_10_Java_Class_Usage.xlsx");
//        String regexp = "\"([^\"]*)\"";
//        Pattern pattern = Pattern.compile(regexp);

        FileInputStream fileIn = new FileInputStream(xlsFile);
        XSSFWorkbook book = new XSSFWorkbook(fileIn);

        ReadFile file = new ReadFile();
        String[] arrTxt = file.readFile(txtFile).split(";");
        System.out.println(arrTxt.length);

        for (int sheetNum = 0; sheetNum < book.getNumberOfSheets(); sheetNum++) {
            XSSFSheet sheet = book.getSheetAt(sheetNum);
            System.out.println(book.getSheetName(sheetNum));
            for (Row row : sheet) {
                if (row.getLastCellNum() == 2 && (row.getRowNum() > 0)) {
                    Cell cellDefinition = row.getCell(0);
                    Cell cellValue = row.getCell(1);
                    String checkStringFromExcel = cellValue.getStringCellValue();
                    checkStringFromExcel = checkStringFromExcel.substring(checkStringFromExcel.indexOf("(\"") + 2, checkStringFromExcel.indexOf("\")"));
//                    System.out.println(row.getRowNum() + ". " + checkStringFromExcel);
                    boolean isConsist = false;
                    int numberFoundPosition = 0;

                    for (int i = 0; i < arrTxt.length; i++) {
                        if (checkStringFromExcel.equals(arrTxt[i])) {
                            isConsist = true;
                            numberFoundPosition = i + 1;

                        }
                    }

                    if (isConsist) {
//                        System.out.println("Element " + checkStringFromExcel+ " found in " + numberFoundPosition + " position in txt file!");
                    } else {
                        System.out.println("Item " + checkStringFromExcel + " not found in " + txtFile + "!");
                        System.out.println(row.getRowNum());
                        System.out.println(cellDefinition);
                        System.out.println(checkStringFromExcel);
                    }
                }
            }
        }
    }
}