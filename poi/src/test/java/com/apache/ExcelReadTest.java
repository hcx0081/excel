package com.apache;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;

/**
 * @Description:
 */
public class ExcelReadTest {
    
    String PATH = "D:\\IDEA\\IntelliJ IDEA 2021.3.3\\Workspace\\excel\\poi\\";
    
    @Test
    public void test03Excel() throws IOException {
        // 1、获取一个工作簿（03版本）
        FileInputStream fileInputStream = new FileInputStream(PATH + "03版MyTable.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2、获取一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        
        // 获取第一行
        Row row1 = sheet.getRow(0); // (1,x)
        // 获取单元格
        Cell cell11 = row1.getCell(0);
        System.out.println(cell11.getStringCellValue());
        
        Cell cell12 = row1.getCell(1);
        System.out.println(cell12.getStringCellValue());
        
        Cell cell13 = row1.getCell(2);
        System.out.println(cell13.getStringCellValue());
        
        // 获取第二行
        Row row2 = sheet.getRow(1); // (2,x)
        // 获取单元格
        Cell cell21 = row2.getCell(0);
        System.out.println(cell21.getStringCellValue());
        
        Cell cell22 = row2.getCell(1);
        System.out.println(cell22.getNumericCellValue());
        
        Cell cell23 = row2.getCell(2);
        System.out.println(cell23.getLocalDateTimeCellValue().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        
        // 关闭流
        fileInputStream.close();
        System.out.println("03版本Excel获取完成");
    }
    
    @Test
    public void test03Excel2() throws IOException {
        // 1、获取一个工作簿（03版本）
        FileInputStream fileInputStream = new FileInputStream(PATH + "03版MyTable.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2、获取一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        
        /* 获取第一行，即表头 */
        Row rowTitle = sheet.getRow(0); // (1,x)
        if (rowTitle != null) {
            // 获取表头列数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    // 获取表头列数据
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
        }
        
        // 空行
        System.out.println();
        
        /* 获取表中的内容 */
        int rowCount = sheet.getPhysicalNumberOfRows();
        // 注意此处从1开始
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                // 获取每行列数
                int cellCount = row.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null) {
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + " | ");
                                break;
                            
                            // 数字（分为日期和普通数字）
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    System.out.print(cell.getLocalDateTimeCellValue().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")) + " | ");
                                } else {
                                    System.out.print(cell.getNumericCellValue() + " | ");
                                }
                                break;
                        }
                    }
                }
            }
        }
        
        // 关闭流
        fileInputStream.close();
        System.out.println("\n03版本Excel获取完成");
    }
    
    @Test
    public void test03ExcelWithFormula() throws IOException {
        // 1、获取一个工作簿（03版本）
        FileInputStream fileInputStream = new FileInputStream(PATH + "Formula.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2、获取一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        
        /* 获取第一行，即表头 */
        Row rowTitle = sheet.getRow(0); // (1,x)
        if (rowTitle != null) {
            // 获取表头列数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    // 获取表头列数据
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
        }
        
        // 空行
        System.out.println();
        
        /* 获取表中的内容 */
        int rowCount = sheet.getPhysicalNumberOfRows();
        // 注意此处从1开始
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                // 获取每行列数
                int cellCount = row.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null) {
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + " | ");
                                break;
                            
                            // 数字（分为日期和普通数字）
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    System.out.print(cell.getLocalDateTimeCellValue().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")) + " | ");
                                } else {
                                    System.out.print(cell.getNumericCellValue() + " | ");
                                }
                                break;
                            
                            case FORMULA:
                                HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
                                System.out.print(formulaEvaluator.evaluate(cell).formatAsString() + " | ");
                        }
                    }
                }
                System.out.println();
            }
        }
        
        // 关闭流
        fileInputStream.close();
        System.out.println("\n03版本Excel获取完成");
    }
}
