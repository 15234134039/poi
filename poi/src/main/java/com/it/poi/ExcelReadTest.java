package com.it.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

/**
 * @author gxp
 */
public class ExcelReadTest {

    String PATH = "D:/idea_workspace/poi/poi/";

    /**
     * 读取03
     * @throws Exception
     */
    @Test
    public void testRead03() throws Exception {
        // 获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "用户统计表.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        // 得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 得到行
        Row row = sheet.getRow(0);
        // 得到列
        Cell cell = row.getCell(1);

        // 读取值的时候，一定要注意类型
        // getStringCellValue() 字符串类型
        // getNumericCellValue() 数字类型
        System.out.println(cell.getNumericCellValue());
        inputStream.close();
    }

    /**
     * 读取07
     * @throws Exception
     */
    @Test
    public void testRead07() throws Exception {
        // 获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "用户统计表07.xlsx");
        Workbook workbook = new XSSFWorkbook(inputStream);
        // 得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 得到行
        Row row = sheet.getRow(0);
        // 得到列
        Cell cell = row.getCell(1);

        // 读取值的时候，一定要注意类型
        // getStringCellValue() 字符串类型
        // getNumericCellValue() 数字类型
        System.out.println(cell.getNumericCellValue());
        inputStream.close();
    }

    /**
     * 读取不同类型数据
     * @throws Exception
     */
    @Test
    public void testCellType() throws Exception {
        FileInputStream inputStream = new FileInputStream(PATH + "明细表.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        // 获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }

        // 获取表中的内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    //System.out.print("[" + (rowNum + 1) + "-" + (cellNum + 1) + "]");
                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING:         // 字符串
                                System.out.print("[String]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:        // 布尔
                                System.out.print("[Boolean]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:      // 空
                                System.out.print("[Blank]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:        // 数字
                                System.out.print("[Numeric]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {  // 日期
                                    System.out.print("[日期]");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd HH:ss:mm");
                                } else {
                                    // 不是日期格式，防止数字过长
                                    System.out.print("[转换为字符串输出]");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("[数据类型错误]");
                                break;
                        }
                        System.out.println(cellValue);

                    }

                }

            }

        }
        inputStream.close();

    }

    /**
     * 读取计算公式
     * @throws Exception
     */
    @Test
    public void testFormula() throws Exception {
        FileInputStream inputStream = new FileInputStream(PATH + "公式表.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);
        // 拿到计算公式
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        // 输出单元格内容
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_FORMULA:   //公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }






    }
}
