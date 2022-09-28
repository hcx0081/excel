package com.apache;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;

/**
 * @Description:
 */
public class ExcelWriteTest {
    
    String PATH = "D:\\IDEA\\IntelliJ IDEA 2021.3.3\\Workspace\\excel\\poi\\";
    
    @Test
    public void test03Excel() throws IOException {
        // 1、创建一个工作簿（03版本）
        Workbook workbook = new HSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet();
        
        // 创建第一行
        Row row1 = sheet.createRow(0); // (1,x)
        // 创建单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("姓名");
        
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("年龄");
        
        Cell cell13 = row1.createCell(2);
        cell13.setCellValue("创建时间");
        
        // 创建第二行
        Row row2 = sheet.createRow(1); // (2,x)
        // 创建单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("张三");
        
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(20);
        
        Cell cell23 = row2.createCell(2);
        CellStyle cell23Style = workbook.createCellStyle();
        // 设置日期时间格式
        cell23Style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        // 设置列宽（以字符宽度的256分之一为单位）
        sheet.setColumnWidth(2, 25 * 256);
        cell23.setCellStyle(cell23Style);
        cell23.setCellValue(LocalDateTime.now());
        
        
        /* 生成一张表 */
        /* 03版本Excel后缀为xls */
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "03版MyTable.xls");
        // 写出流
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        System.out.println("03版本Excel生成完成");
    }
    
    @Test
    public void test07Excel() throws IOException {
        // 1、创建一个工作簿（07版本）
        Workbook workbook = new XSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet();
        
        // 创建第一行
        Row row1 = sheet.createRow(0); // (1,x)
        // 创建单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("姓名");
        
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("年龄");
        
        Cell cell13 = row1.createCell(2);
        cell13.setCellValue("创建时间");
        
        // 创建第二行
        Row row2 = sheet.createRow(1); // (2,x)
        // 创建单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("张三");
        
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(20);
        
        Cell cell23 = row2.createCell(2);
        CellStyle cell23Style = workbook.createCellStyle();
        // 设置日期时间格式
        cell23Style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        // 设置列宽（以字符宽度的256分之一为单位）
        sheet.setColumnWidth(2, 25 * 256);
        cell23.setCellStyle(cell23Style);
        cell23.setCellValue(LocalDateTime.now());
        
        
        /* 生成一张表 */
        /* 07版本Excel后缀为xlsx */
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "07版MyTable.xlsx");
        // 写出流
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        System.out.println("07版本Excel生成完成");
    }
    
    @Test
    public void test03ExcelBigData() throws IOException {
        /* 开始时间 */
        long begin = System.currentTimeMillis();
        
        // 1、创建一个工作簿（03版本）
        Workbook workbook = new HSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet();
        
        for (int row = 0; row < 65536; row++) {
            Row sheetRow = sheet.createRow(row);
            Cell rowCell = sheetRow.createCell(0);
            rowCell.setCellValue("hello");
        }
        
        /* 生成一张表 */
        /* 03版本Excel后缀为xls */
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "03版MyBigTable.xls");
        // 写出流
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        
        /* 结束时间 */
        long end = System.currentTimeMillis();
        System.out.println("总耗时：" + (double) (end - begin) / 1000);
        System.out.println("03版本Excel生成完成");
    }
    
    @Test
    public void test07ExcelBigData() throws IOException {
        /* 开始时间 */
        long begin = System.currentTimeMillis();
        
        // 1、创建一个工作簿（07版本）
        Workbook workbook = new XSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet();
        
        for (int row = 0; row < 65536; row++) {
            Row sheetRow = sheet.createRow(row);
            Cell rowCell = sheetRow.createCell(0);
            rowCell.setCellValue("hello");
        }
        
        /* 生成一张表 */
        /* 07版本Excel后缀为xlsx */
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "07版MyBigTable.xlsx");
        // 写出流
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        
        /* 结束时间 */
        long end = System.currentTimeMillis();
        System.out.println("总耗时：" + (double) (end - begin) / 1000);
        System.out.println("07版本Excel生成完成");
    }
    
    @Test
    public void test07ExcelBigDataSuper() throws IOException {
        /* 开始时间 */
        long begin = System.currentTimeMillis();
        
        // 1、创建一个工作簿（07版本）
        Workbook workbook = new SXSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet();
        
        for (int row = 0; row < 65536; row++) {
            Row sheetRow = sheet.createRow(row);
            Cell rowCell = sheetRow.createCell(0);
            rowCell.setCellValue("hello");
        }
        
        /* 生成一张表 */
        /* 07版本Excel后缀为xlsx */
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "07版MyBigTableS.xlsx");
        // 写出流
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        
        /* 使用SXSSFWorkbook进行大数据量写操作后记得清除临时文件 */
        ((SXSSFWorkbook) workbook).dispose();
        
        /* 结束时间 */
        long end = System.currentTimeMillis();
        System.out.println("总耗时：" + (double) (end - begin) / 1000);
        System.out.println("07版本Excel生成完成");
    }
}
