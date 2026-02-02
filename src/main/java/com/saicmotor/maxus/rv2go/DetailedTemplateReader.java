package com.saicmotor.maxus.rv2go;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * 详细Excel文件读取工具 - 用于读取模板文件
 */
public class DetailedTemplateReader {

    /**
     * 读取Excel文件的所有行数据
     * @param filePath Excel文件路径
     * @throws IOException 文件读取异常
     */
    public static void readAllRows(String filePath) throws IOException {
        System.out.println("========================================");
        System.out.println("文件: " + filePath);
        System.out.println("========================================");

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 获取所有sheet名称
            int numberOfSheets = workbook.getNumberOfSheets();
            System.out.println("Sheet数量: " + numberOfSheets);
            System.out.println();

            // 遍历每个sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                System.out.println("【Sheet " + (i + 1) + "】: " + sheetName);
                System.out.println("----------------------------------------");

                // 获取实际行数（排除空行）
                int rowCount = sheet.getPhysicalNumberOfRows();
                System.out.println("总行数: " + rowCount);
                System.out.println();

                // 读取所有行
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        System.out.print("第 " + (rowIndex + 1) + " 行: ");
                        StringBuilder rowData = new StringBuilder();

                        int lastCol = row.getLastCellNum();
                        for (int colIndex = 0; colIndex < lastCol; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (rowData.length() > 0) {
                                rowData.append(" | ");
                            }
                            String cellValue = getCellValueAsString(cell);
                            // 显示完整内容，不截断
                            if (cellValue.isEmpty()) {
                                rowData.append("(空)");
                            } else {
                                // 替换换行符和制表符以便显示
                                cellValue = cellValue.replace("\n", "\\n").replace("\t", "\\t");
                                rowData.append(cellValue);
                            }
                        }

                        System.out.println(rowData.toString());
                    }
                }

                System.out.println();
            }
        }
    }

    /**
     * 获取单元格的字符串值
     * @param cell 单元格
     * @return 单元格的字符串值
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // 处理数字，避免科学计数法
                    double numValue = cell.getNumericCellValue();
                    if (numValue == (long) numValue) {
                        return String.valueOf((long) numValue);
                    } else {
                        return String.valueOf(numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    /**
     * 主方法 - 读取指定的Excel文件
     */
    public static void main(String[] args) {
        String filePath = "/Users/fanxiang/Documents/hy/电子票模版.xlsx";

        try {
            readAllRows(filePath);
        } catch (IOException e) {
            System.err.println("读取文件失败: " + filePath);
            e.printStackTrace();
        }
    }
}
