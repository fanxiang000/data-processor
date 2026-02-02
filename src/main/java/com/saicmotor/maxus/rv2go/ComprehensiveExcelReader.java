package com.saicmotor.maxus.rv2go;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 综合Excel文件读取工具
 * 用于读取Excel文件的完整结构信息，包括Sheet、表头、数据示例和所有行数据
 */
public class ComprehensiveExcelReader {

    /**
     * 读取Excel文件的完整结构
     * @param filePath Excel文件路径
     * @param showAllRows 是否显示所有行数据
     * @throws IOException 文件读取异常
     */
    public static void readExcelStructure(String filePath, boolean showAllRows) throws IOException {
        System.out.println("\n");
        System.out.println("######################################################################");
        System.out.println("# 文件: " + filePath);
        System.out.println("######################################################################");

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 获取所有sheet名称
            int numberOfSheets = workbook.getNumberOfSheets();
            System.out.println("# Sheet数量: " + numberOfSheets);
            System.out.println("######################################################################");

            // 遍历每个sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                System.out.println("\n【Sheet " + (i + 1) + "】: " + sheetName);
                System.out.println("----------------------------------------------------------------------");

                // 获取实际行数（排除空行）
                int rowCount = sheet.getPhysicalNumberOfRows();
                System.out.println("总行数: " + rowCount);
                System.out.println();

                if (rowCount > 0) {
                    // 查找表头行（第一个包含多个非空单元格的行）
                    int headerRowIndex = findHeaderRow(sheet, rowCount);
                    List<String> headers = new ArrayList<>();

                    if (headerRowIndex >= 0) {
                        Row headerRow = sheet.getRow(headerRowIndex);
                        System.out.println("【表头列名】(第 " + (headerRowIndex + 1) + " 行):");
                        for (Cell cell : headerRow) {
                            String cellValue = getCellValueAsString(cell);
                            headers.add(cellValue.isEmpty() ? "(空)" : cellValue);
                        }

                        // 打印表头
                        for (int j = 0; j < headers.size(); j++) {
                            System.out.println("  列 " + (j + 1) + ": " + headers.get(j));
                        }
                        System.out.println();
                    }

                    // 显示数据
                    if (showAllRows) {
                        System.out.println("【所有行数据】:");
                    } else {
                        System.out.println("【前5行数据示例】:");
                    }

                    int displayCount = showAllRows ? rowCount : Math.min(5, rowCount);

                    for (int rowIndex = 0; rowIndex < displayCount; rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        if (row != null) {
                            System.out.print("  第 " + (rowIndex + 1) + " 行: ");
                            StringBuilder rowData = new StringBuilder();

                            int lastCol = row.getLastCellNum();
                            for (int colIndex = 0; colIndex < lastCol; colIndex++) {
                                Cell cell = row.getCell(colIndex);
                                if (rowData.length() > 0) {
                                    rowData.append(" | ");
                                }
                                String cellValue = getCellValueAsString(cell);
                                // 限制每个单元格显示长度
                                if (cellValue.length() > 50) {
                                    cellValue = cellValue.substring(0, 47) + "...";
                                }
                                if (cellValue.isEmpty()) {
                                    cellValue = "(空)";
                                } else {
                                    // 替换换行符和制表符以便显示
                                    cellValue = cellValue.replace("\n", "\\n").replace("\t", "\\t");
                                }
                                rowData.append(cellValue);
                            }

                            System.out.println(rowData.toString());
                        }
                    }
                } else {
                    System.out.println("  (空Sheet)");
                }
            }
        }
    }

    /**
     * 查找表头行（第一个包含至少3个非空单元格的行）
     * @param sheet Sheet对象
     * @param rowCount 总行数
     * @return 表头行索引，找不到返回-1
     */
    private static int findHeaderRow(Sheet sheet, int rowCount) {
        for (int i = 0; i < Math.min(20, rowCount); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int nonEmptyCount = 0;
                for (Cell cell : row) {
                    String value = getCellValueAsString(cell);
                    if (!value.isEmpty()) {
                        nonEmptyCount++;
                    }
                }
                // 如果有至少3个非空单元格，认为是表头行
                if (nonEmptyCount >= 3) {
                    return i;
                }
            }
        }
        return -1;
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
     * 主方法 - 读取指定的两个Excel文件
     */
    public static void main(String[] args) {
        String[] filePaths = {
            "/Users/fanxiang/Documents/hy/电子票模版.xlsx",
            "/Users/fanxiang/Documents/hy/东南国资出库电子票.xlsx"
        };

        for (String filePath : filePaths) {
            try {
                // 读取完整结构，包括所有行数据
                readExcelStructure(filePath, true);
            } catch (IOException e) {
                System.err.println("读取文件失败: " + filePath);
                e.printStackTrace();
            }
            System.out.println();
            System.out.println();
        }
    }
}
