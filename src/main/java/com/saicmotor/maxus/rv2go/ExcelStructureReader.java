package com.saicmotor.maxus.rv2go;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel文件结构读取工具
 * 用于读取Excel文件的Sheet信息、表头和数据示例
 */
public class ExcelStructureReader {

    /**
     * 读取Excel文件的完整结构
     * @param filePath Excel文件路径
     * @throws IOException 文件读取异常
     */
    public static void readExcelStructure(String filePath) throws IOException {
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
                System.out.println("数据行数: " + rowCount);

                if (rowCount > 0) {
                    // 查找第一个非空行作为表头
                    int headerRowIndex = -1;
                    for (int j = 0; j < Math.min(10, rowCount); j++) {
                        Row row = sheet.getRow(j);
                        if (row != null && hasNonEmptyCells(row)) {
                            headerRowIndex = j;
                            break;
                        }
                    }

                    if (headerRowIndex >= 0) {
                        Row headerRow = sheet.getRow(headerRowIndex);
                        List<String> headers = new ArrayList<>();
                        System.out.println("\n【表头列名】(第 " + (headerRowIndex + 1) + " 行):");
                        for (Cell cell : headerRow) {
                            String cellValue = getCellValueAsString(cell);
                            headers.add(cellValue.isEmpty() ? "(空)" : cellValue);
                        }

                        // 打印表头
                        for (int j = 0; j < headers.size(); j++) {
                            System.out.println("  列 " + (j + 1) + ": " + headers.get(j));
                        }
                    }

                    // 读取前5行数据示例
                    System.out.println("\n【前5行数据示例】:");
                    int dataRowCount = Math.min(5, rowCount);

                    for (int rowIndex = 0; rowIndex < dataRowCount; rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        if (row != null) {
                            System.out.print("  第 " + (rowIndex + 1) + " 行: ");
                            StringBuilder rowData = new StringBuilder();

                            for (Cell cell : row) {
                                if (rowData.length() > 0) {
                                    rowData.append(" | ");
                                }
                                String cellValue = getCellValueAsString(cell);
                                // 限制每个单元格显示长度
                                if (cellValue.length() > 30) {
                                    cellValue = cellValue.substring(0, 27) + "...";
                                }
                                rowData.append(cellValue);
                            }

                            System.out.println(rowData.toString());
                        }
                    }
                } else {
                    System.out.println("  (空Sheet)");
                }

                System.out.println();
                System.out.println();
            }
        }
    }

    /**
     * 检查行是否有非空单元格
     * @param row 行
     * @return 如果行中至少有一个非空单元格则返回true
     */
    private static boolean hasNonEmptyCells(Row row) {
        if (row == null) {
            return false;
        }
        for (Cell cell : row) {
            String value = getCellValueAsString(cell);
            if (!value.isEmpty()) {
                return true;
            }
        }
        return false;
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
                readExcelStructure(filePath);
            } catch (IOException e) {
                System.err.println("读取文件失败: " + filePath);
                e.printStackTrace();
            }
            System.out.println();
            System.out.println();
        }
    }
}
