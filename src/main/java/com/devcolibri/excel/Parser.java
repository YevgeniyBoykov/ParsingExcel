//package com.devcolibri.excel;
//
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.util.Iterator;
//
//public class Parser {
//
//    public static String parse(XSSFWorkbook wb) {
//
//        String result = "";
//
//        XSSFSheet sheet = wb.getSheetAt(0);
//        Iterator<Row> it = sheet.iterator();
////        while (it.hasNext()) {
//            Row row = it.next();
//            Iterator<Cell> cells = row.iterator();
//            while (cells.hasNext()) {
//                Cell cell = cells.next();
//                int cellType = cell.getCellType();
//                switch (cellType) {
//                    case Cell.CELL_TYPE_STRING:
//                        System.out.println(cell.getCellType());
//                        result += cell.getStringCellValue() + "=";
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.println(cellType);
//                        result += "[" + cell.getNumericCellValue() + "]";
//                        break;
//
//                    case Cell.CELL_TYPE_FORMULA:
//                        System.out.println(cellType);
//                        result += "[" + cell.getNumericCellValue() + "]";
//                        break;
//                    default:
//                        System.out.println(cellType);
//                        result += "|";
//                        break;
//                }
//            }
//            result += "\n";
////        }
//
//        return result;
//    }
//
//}
