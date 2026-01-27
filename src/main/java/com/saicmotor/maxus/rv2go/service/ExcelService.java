package com.saicmotor.maxus.rv2go.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Excel 处理服务
 * 提供 Excel 文件读取、合并等功能
 */
public class ExcelService {

    /**
     * 合并两个 Excel 文件
     * 将表 2 中指定的列合并到表 1，基于关联列进行匹配
     *
     * @param file1         表 1（主表）
     * @param file2         表 2（合并表）
     * @param joinKeys      关联列名数组
     * @param columnsToMerge 要从表 2 合并的列名数组
     * @param outputFile    输出文件
     * @return 是否成功
     */
    public boolean mergeExcelFiles(File file1, File file2, String[] joinKeys,
                                   String[] columnsToMerge, File outputFile) {
        return mergeExcelFilesWithExclude(file1, file2, null, joinKeys, columnsToMerge, null, outputFile);
    }

    /**
     * 合并两个 Excel 文件，并排除表1中与表3匹配的数据
     * 将表 2 中指定的列合并到表 1，基于关联列进行匹配
     * 表1中与表3匹配的行将被排除
     *
     * @param file1         表 1（主表）
     * @param file2         表 2（合并表）
     * @param file3         表 3（排除表，可为null）
     * @param joinKeys      表1表2关联列名数组
     * @param columnsToMerge 要从表 2 合并的列名数组
     * @param excludeKeys   表1表3排除关联列名数组（可为null）
     * @param outputFile    输出文件
     * @return 是否成功
     */
    public boolean mergeExcelFilesWithExclude(File file1, File file2, File file3,
                                              String[] joinKeys, String[] columnsToMerge,
                                              String[] excludeKeys, File outputFile) {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        Workbook workbook3 = null;
        Workbook outputWorkbook = null;

        try {
            // 读取 Excel 文件
            workbook1 = readWorkbook(file1);
            workbook2 = readWorkbook(file2);

            if (workbook1 == null || workbook2 == null) {
                System.err.println("无法读取 Excel 文件");
                return false;
            }

            // 如果有表3，也读取表3
            boolean enableExclude = file3 != null && excludeKeys != null && excludeKeys.length > 0;
            if (enableExclude) {
                workbook3 = readWorkbook(file3);
                if (workbook3 == null) {
                    System.err.println("无法读取表 3 文件");
                    return false;
                }
            }

            // 获取第一个工作表
            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);
            Sheet sheet3 = null;
            Row header3 = null;
            Map<String, Integer> columnMap3 = null;

            // 读取表头
            Row header1 = sheet1.getRow(0);
            Row header2 = sheet2.getRow(0);

            if (header1 == null || header2 == null) {
                System.err.println("Excel 文件没有表头");
                return false;
            }

            // 获取列名映射
            Map<String, Integer> columnMap1 = getColumnMapping(header1);
            Map<String, Integer> columnMap2 = getColumnMapping(header2);

            // 如果启用表3排除，读取并验证表3
            if (enableExclude) {
                sheet3 = workbook3.getSheetAt(0);
                header3 = sheet3.getRow(0);
                if (header3 == null) {
                    System.err.println("表 3 文件没有表头");
                    return false;
                }
                columnMap3 = getColumnMapping(header3);

                // 验证排除关联列是否存在
                for (String key : excludeKeys) {
                    if (!columnMap1.containsKey(key) || !columnMap3.containsKey(key)) {
                        System.err.println("排除关联列 '" + key + "' 在表1或表3中不存在");
                        return false;
                    }
                }
            }

            // 验证关联列是否存在
            for (String key : joinKeys) {
                if (!columnMap1.containsKey(key) || !columnMap2.containsKey(key)) {
                    System.err.println("关联列 '" + key + "' 在某个表中不存在");
                    return false;
                }
            }

            // 验证要合并的列是否存在
            for (String col : columnsToMerge) {
                if (!columnMap2.containsKey(col)) {
                    System.err.println("要合并的列 '" + col + "' 在表 2 中不存在");
                    return false;
                }
            }

            // 创建输出工作簿
            outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("MergedData");

            // 创建输出表头
            int outputColCount = createOutputHeader(header1, header2, outputSheet,
                    columnMap1, columnsToMerge);

            // 构建表 2 的索引（基于关联列）
            Map<String, Map<String, Object>> sheet2Index = buildSheet2Index(sheet2, joinKeys, columnMap2);

            // 构建表 3 的排除索引（基于排除关联列）
            Set<String> sheet3ExcludeKeys = null;
            if (enableExclude) {
                sheet3ExcludeKeys = buildSheet3ExcludeIndex(sheet3, excludeKeys, columnMap3);
            }

            // 合并数据（排除表3中存在的数据）
            mergeDataWithExclude(sheet1, outputSheet, joinKeys, columnsToMerge, columnMap1,
                    sheet2Index, excludeKeys, sheet3ExcludeKeys);

            // 自动调整列宽
            for (int i = 0; i < outputColCount; i++) {
                outputSheet.autoSizeColumn(i);
            }

            // 写入输出文件
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }

            return true;

        } catch (Exception e) {
            System.err.println("合并 Excel 文件时出错: " + e.getMessage());
            e.printStackTrace();
            return false;
        } finally {
            closeQuietly(workbook1);
            closeQuietly(workbook2);
            closeQuietly(workbook3);
            closeQuietly(outputWorkbook);
        }
    }

    /**
     * 读取 Excel 工作簿
     */
    private Workbook readWorkbook(File file) {
        try (FileInputStream fis = new FileInputStream(file)) {
            String fileName = file.getName().toLowerCase();
            if (fileName.endsWith(".xlsx")) {
                return new XSSFWorkbook(fis);
            } else if (fileName.endsWith(".xls")) {
                return WorkbookFactory.create(fis);
            }
        } catch (IOException e) {
            System.err.println("读取文件失败: " + file.getAbsolutePath() + " - " + e.getMessage());
        }
        return null;
    }

    /**
     * 获取列名到列索引的映射
     */
    private Map<String, Integer> getColumnMapping(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();
        for (Cell cell : headerRow) {
            String columnName = getCellValueAsString(cell);
            if (columnName != null && !columnName.isEmpty()) {
                columnMap.put(columnName, cell.getColumnIndex());
            }
        }
        return columnMap;
    }

    /**
     * 创建输出表头
     */
    private int createOutputHeader(Row header1, Row header2, Sheet outputSheet,
                                   Map<String, Integer> columnMap1, String[] columnsToMerge) {
        Row outputHeader = outputSheet.createRow(0);
        int colIndex = 0;

        // 添加表 1 的所有列
        for (Cell cell : header1) {
            String columnName = getCellValueAsString(cell);
            Cell newCell = outputHeader.createCell(colIndex++);
            newCell.setCellValue(columnName);
        }

        // 添加表 2 中要合并的列（跳过已存在的列）
        Set<String> existingColumns = new HashSet<>(columnMap1.keySet());
        for (String colName : columnsToMerge) {
            if (!existingColumns.contains(colName)) {
                Cell newCell = outputHeader.createCell(colIndex++);
                newCell.setCellValue(colName);
                existingColumns.add(colName);
            }
        }

        return colIndex;
    }

    /**
     * 构建表 2 的索引，用于快速查找
     * 返回：关联键值 -> {列名 -> 单元格值}
     */
    private Map<String, Map<String, Object>> buildSheet2Index(Sheet sheet2, String[] joinKeys,
                                                               Map<String, Integer> columnMap2) {
        Map<String, Map<String, Object>> index = new HashMap<>();

        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row row = sheet2.getRow(i);
            if (row == null) continue;

            // 构建关联键
            StringBuilder keyBuilder = new StringBuilder();
            for (String joinKey : joinKeys) {
                int colIndex = columnMap2.get(joinKey);
                Cell cell = row.getCell(colIndex);
                String value = getCellValueAsString(cell);
                keyBuilder.append(value).append("|||");
            }
            String key = keyBuilder.toString();

            // 存储该行的所有列数据
            Map<String, Object> rowData = new HashMap<>();
            for (Cell cell : row) {
                String colName = getColumnName(sheet2.getRow(0), cell.getColumnIndex());
                if (colName != null) {
                    rowData.put(colName, getCellValue(cell));
                }
            }
            index.put(key, rowData);
        }

        return index;
    }

    /**
     * 合并数据到输出表（带排除功能）
     */
    private void mergeDataWithExclude(Sheet sheet1, Sheet outputSheet, String[] joinKeys,
                                      String[] columnsToMerge, Map<String, Integer> columnMap1,
                                      Map<String, Map<String, Object>> sheet2Index,
                                      String[] excludeKeys, Set<String> sheet3ExcludeKeys) {
        Row header1 = sheet1.getRow(0);
        Row outputHeader = outputSheet.getRow(0);

        // 获取输出表头中各列的索引
        Map<String, Integer> outputColumnMap = new HashMap<>();
        for (Cell cell : outputHeader) {
            outputColumnMap.put(getCellValueAsString(cell), cell.getColumnIndex());
        }

        int outputRowIndex = 1;  // 从第1行开始，第0行是表头
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 == null) continue;

            // 构建排除键（用于检查是否在表3中）
            String excludeKey = null;
            if (excludeKeys != null && sheet3ExcludeKeys != null) {
                StringBuilder excludeKeyBuilder = new StringBuilder();
                for (String key : excludeKeys) {
                    int colIndex = columnMap1.get(key);
                    Cell cell = row1.getCell(colIndex);
                    String value = getCellValueAsString(cell);
                    excludeKeyBuilder.append(value).append("|||");
                }
                excludeKey = excludeKeyBuilder.toString();
            }

            // 如果在表3中找到匹配，跳过此行
            if (excludeKey != null && sheet3ExcludeKeys.contains(excludeKey)) {
                continue;
            }

            Row outputRow = outputSheet.createRow(outputRowIndex++);

            // 复制表 1 的数据
            for (Cell cell : row1) {
                Cell newCell = outputRow.createCell(cell.getColumnIndex());
                copyCellValue(cell, newCell);
            }

            // 构建关联键
            StringBuilder keyBuilder = new StringBuilder();
            for (String joinKey : joinKeys) {
                int colIndex = columnMap1.get(joinKey);
                Cell cell = row1.getCell(colIndex);
                String value = getCellValueAsString(cell);
                keyBuilder.append(value).append("|||");
            }
            String key = keyBuilder.toString();

            // 从表 2 查找匹配的数据并合并
            Map<String, Object> matchedRow = sheet2Index.get(key);
            if (matchedRow != null) {
                for (String colName : columnsToMerge) {
                    Integer outputColIndex = outputColumnMap.get(colName);
                    if (outputColIndex != null && matchedRow.containsKey(colName)) {
                        Cell newCell = outputRow.createCell(outputColIndex);
                        setCellValue(newCell, matchedRow.get(colName));
                    }
                }
            }
        }
    }

    /**
     * 获取单元格值作为字符串
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        Object value = getCellValue(cell);
        return value != null ? value.toString() : "";
    }

    /**
     * 获取单元格值
     */
    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return null;
            default:
                return null;
        }
    }

    /**
     * 设置单元格值
     */
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
            // 设置日期格式
            CellStyle style = cell.getSheet().getWorkbook().createCellStyle();
            style.setDataFormat((short) 0x16); // yyyy-mm-dd 格式
            cell.setCellStyle(style);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 复制单元格值（保留格式）
     */
    private void copyCellValue(Cell source, Cell target) {
        if (source == null) {
            target.setBlank();
            return;
        }

        // 复制样式（包括日期格式）
        CellStyle sourceStyle = source.getCellStyle();
        if (sourceStyle != null) {
            target.setCellStyle(sourceStyle);
        }

        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                target.setCellValue(source.getNumericCellValue());
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellFormula(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                target.setBlank();
        }
    }

    /**
     * 构建表 3 的排除索引，用于快速查找要排除的行
     * 返回：要排除的关联键值集合
     */
    private Set<String> buildSheet3ExcludeIndex(Sheet sheet3, String[] excludeKeys,
                                                Map<String, Integer> columnMap3) {
        Set<String> excludeIndex = new HashSet<>();

        for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
            Row row = sheet3.getRow(i);
            if (row == null) continue;

            // 构建排除关联键
            StringBuilder keyBuilder = new StringBuilder();
            for (String excludeKey : excludeKeys) {
                int colIndex = columnMap3.get(excludeKey);
                Cell cell = row.getCell(colIndex);
                String value = getCellValueAsString(cell);
                keyBuilder.append(value).append("|||");
            }
            String key = keyBuilder.toString();
            excludeIndex.add(key);
        }

        return excludeIndex;
    }

    /**
     * 获取列名
     */
    private String getColumnName(Row headerRow, int columnIndex) {
        if (headerRow == null) return null;
        Cell cell = headerRow.getCell(columnIndex);
        return cell != null ? getCellValueAsString(cell) : null;
    }

    /**
     * 关闭工作簿，忽略异常
     */
    private void closeQuietly(Workbook workbook) {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
                // Ignore
            }
        }
    }
}
