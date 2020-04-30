package com.it.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;


import java.io.FileOutputStream;

/**
 * @author gxp
 */
public class ExcelWriteTest {

    String PATH = "D:/idea_workspace/poi/poi/";

    /**
     * 测试03
     * @throws Exception
     */
    @Test
    public void testWrite03() throws Exception {
        // 1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        // 2.创建一个工作表
        Sheet sheet = workbook.createSheet("用户统计表");
        // 3.创建一个行 （1,1）
        Row row1 = sheet.createRow(0);
        // 4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("用户数");
        // (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(10);

        // 第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 生成一张表（IO流），03版本使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "用户统计表.xls");
        // 输出
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        System.out.println("用户表03生成完毕");
    }

    /**
     * 测试07
     * @throws Exception
     */
    @Test
    public void testWrite07() throws Exception {
        // 1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        // 2.创建一个工作表
        Sheet sheet = workbook.createSheet("用户统计表2");
        // 3.创建一个行 （1,1）
        Row row1 = sheet.createRow(0);
        // 4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("用户数");
        // (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(10);

        // 第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 生成一张表（IO流），07版本使用xlsx结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "用户统计表07.xlsx");
        // 输出
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        System.out.println("用户表07生成完毕");
    }

    /**
     * 测试03，最多只能输入65536行，否则报错
     * @throws Exception
     */
    @Test
    public void testWrite03BigData() throws Exception{
        long begin = System.currentTimeMillis();

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        long end = System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }

    /**
     * 测试07，不限制行数
     * @throws Exception
     */
    @Test
    public void testWrite07BigData() throws Exception{
        long begin = System.currentTimeMillis();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigData.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        long end = System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }

    /**
     * 测试07，使用缓存 SXSSFWorkbook
     * @throws Exception
     */
    @Test
    public void testWrite07BigDataS() throws Exception{
        long begin = System.currentTimeMillis();

        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        // 清除临时文件！
        ((SXSSFWorkbook) workbook).dispose();
        long end = System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }
}
