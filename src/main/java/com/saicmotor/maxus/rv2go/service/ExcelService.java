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
     * 列运算规则
     */
    public static class ColumnCalculation {
        public String targetColumn;  // 目标列
        public String column1;       // 第一列
        public String operator;      // 运算符：add, subtract, multiply, divide
        public String column2;       // 第二列

        public ColumnCalculation(String targetColumn, String column1, String operator, String column2) {
            this.targetColumn = targetColumn;
            this.column1 = column1;
            this.operator = operator;
            this.column2 = column2;
        }
    }

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
        return mergeExcelFilesWithFilter(file1, file2, joinKeys, columnsToMerge, null, null, outputFile);
    }

    /**
     * 合并两个 Excel 文件，支持空值过滤
     * 将表 2 中指定的列合并到表 1，基于关联列进行匹配
     *
     * @param file1            表 1（主表）
     * @param file2            表 2（合并表）
     * @param joinKeys         关联列名数组
     * @param columnsToMerge   要从表 2 合并的列名数组
     * @param filterEmptyColumns 表1中需要检查空值的列名数组（可为null）
     * @param outputFile       输出文件
     * @return 是否成功
     */
    public boolean mergeExcelFilesWithFilter(File file1, File file2, String[] joinKeys,
                                             String[] columnsToMerge, String[] filterEmptyColumns,
                                             Map<String, String> subtractMap,
                                             File outputFile) {
        return mergeExcelFilesWithExclude(file1, file2, null, joinKeys, columnsToMerge, null, null, filterEmptyColumns, subtractMap, outputFile);
    }

    /**
     * 合并两个 Excel 文件，支持多组关联列（带回退逻辑）
     * 将表 2 中指定的列合并到表 1，尝试使用第一组关联列匹配，如果失败则尝试下一组
     *
     * @param file1            表 1（主表）
     * @param file2            表 2（合并表）
     * @param joinKeyGroups    关联列组列表，格式: "表1列1,表1列2=表2列1,表2列2"
     * @param columnsToMerge   要从表 2 合并的列名数组
     * @param filterEmptyColumns 表1中需要检查空值的列名数组（可为null）
     * @param outputFile       输出文件
     * @param highlightMatches 是否高亮标记匹配的行
     * @param columnCalculations 列运算规则列表（可为null）
     * @return 是否成功
     */
    public boolean mergeExcelFilesWithMultipleJoinGroups(File file1, File file2,
                                                        List<String> joinKeyGroups,
                                                        String[] columnsToMerge,
                                                        String[] filterEmptyColumns,
                                                        File outputFile,
                                                        boolean highlightMatches,
                                                        List<ColumnCalculation> columnCalculations) {
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        Workbook outputWorkbook = null;

        try {
            // 读取工作簿
            workbook1 = readWorkbook(file1);
            workbook2 = readWorkbook(file2);
            if (workbook1 == null || workbook2 == null) {
                System.err.println("无法读取 Excel 文件");
                return false;
            }

            outputWorkbook = workbook1.getClass().newInstance();
            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet();

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

            // 复制表1的表头
            Row outputHeader = outputSheet.createRow(0);
            int nextColIndex = 0;
            for (Cell cell : header1) {
                Cell newCell = outputHeader.createCell(nextColIndex++);
                copyCellValue(cell, newCell);
            }

            // 添加表2中要合并的列（跳过已存在的列）
            Set<String> existingColumns = new HashSet<>();
            for (Cell cell : header1) {
                String colName = getCellValueAsString(cell);
                if (colName != null && !colName.isEmpty()) {
                    existingColumns.add(colName);
                }
            }
            for (String colName : columnsToMerge) {
                if (!existingColumns.contains(colName)) {
                    Cell newCell = outputHeader.createCell(nextColIndex++);
                    newCell.setCellValue(colName);
                    existingColumns.add(colName);
                }
            }

            // 获取输出表头中各列的索引
            Map<String, Integer> outputColumnMap = new HashMap<>();
            for (Cell cell : outputHeader) {
                outputColumnMap.put(getCellValueAsString(cell), cell.getColumnIndex());
            }

            // 验证关联列组
            List<JoinKeyGroup> parsedGroups = parseJoinKeyGroups(joinKeyGroups, columnMap1, columnMap2);
            if (parsedGroups.isEmpty()) {
                System.err.println("没有有效的关联列组");
                return false;
            }

            // 验证要合并的列是否存在
            for (String col : columnsToMerge) {
                if (!columnMap2.containsKey(col)) {
                    System.err.println("要合并的列 '" + col + "' 在表 2 中不存在");
                    return false;
                }
            }

            // 验证空值过滤列是否存在
            if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                for (String col : filterEmptyColumns) {
                    if (!columnMap1.containsKey(col)) {
                        System.err.println("空值检查列 '" + col + "' 在表 1 中不存在");
                        return false;
                    }
                }
            }

            // 为所有关联列组构建表2的索引（包含行号信息）
            List<Map<String, Integer>> sheet2KeyToRowMap = new ArrayList<>();
            List<Map<Integer, Map<String, Object>>> sheet2RowDataMap = new ArrayList<>();
            for (JoinKeyGroup group : parsedGroups) {
                String[] table2KeysArray = group.table2Keys.toArray(new String[0]);
                Map<String, Integer> keyToRow = new HashMap<>();
                Map<Integer, Map<String, Object>> rowData = new HashMap<>();

                for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
                    Row row = sheet2.getRow(i);
                    if (row == null) continue;

                    // 构建关联键
                    StringBuilder keyBuilder = new StringBuilder();
                    for (String key : table2KeysArray) {
                        Integer colIndex = columnMap2.get(key);
                        if (colIndex != null) {
                            Cell cell = row.getCell(colIndex);
                            String value = getCellValueAsString(cell);
                            keyBuilder.append(value).append("|||");
                        }
                    }
                    String key = keyBuilder.toString();

                    // 存储键到行号的映射（如果重复键，保留第一个）
                    if (!keyToRow.containsKey(key)) {
                        keyToRow.put(key, i);
                    }

                    // 存储行数据
                    Map<String, Object> cellData = new HashMap<>();
                    for (Cell cell : row) {
                        String colName = getColumnName(header2, cell.getColumnIndex());
                        if (colName != null) {
                            cellData.put(colName, getCellValue(cell));
                        }
                    }
                    rowData.put(i, cellData);
                }

                sheet2KeyToRowMap.add(keyToRow);
                sheet2RowDataMap.add(rowData);
            }

            // 两阶段匹配：记录已使用的表2行号和已匹配的表1行号
            Set<Integer> usedSheet2Rows = new HashSet<>();  // 已使用的表2行号
            Set<Integer> matchedSheet1Rows = new HashSet<>();  // 已匹配的表1行号
            Map<Integer, Map<String, Object>> sheet1Matches = new HashMap<>();  // 表1行号 -> 匹配的表2数据

            // 第一阶段：用组1匹配
            if (parsedGroups.size() > 0) {
                JoinKeyGroup group = parsedGroups.get(0);
                Map<String, Integer> keyToRow = sheet2KeyToRowMap.get(0);

                for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                    Row row1 = sheet1.getRow(i);
                    if (row1 == null) continue;

                    // 检查空值过滤
                    if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                        boolean hasEmptyValue = false;
                        for (String colName : filterEmptyColumns) {
                            Integer colIndex = columnMap1.get(colName);
                            if (colIndex != null) {
                                Cell cell = row1.getCell(colIndex);
                                String value = getCellValueAsString(cell);
                                if (value == null || value.trim().isEmpty()) {
                                    hasEmptyValue = true;
                                    break;
                                }
                            }
                        }
                        if (hasEmptyValue) {
                            continue;
                        }
                    }

                    // 构建表1的关联键
                    StringBuilder keyBuilder = new StringBuilder();
                    for (String key : group.table1Keys) {
                        Integer colIndex = columnMap1.get(key);
                        if (colIndex != null) {
                            Cell cell = row1.getCell(colIndex);
                            String value = getCellValueAsString(cell);
                            keyBuilder.append(value).append("|||");
                        }
                    }
                    String key = keyBuilder.toString();

                    // 查找匹配的表2数据
                    Integer sheet2Row = keyToRow.get(key);
                    if (sheet2Row != null && !usedSheet2Rows.contains(sheet2Row)) {
                        matchedSheet1Rows.add(i);
                        usedSheet2Rows.add(sheet2Row);
                        sheet1Matches.put(i, sheet2RowDataMap.get(0).get(sheet2Row));
                    }
                }
            }

            // 第二阶段：用组2匹配未匹配的表1行（排除已使用的表2行）
            if (parsedGroups.size() > 1) {
                JoinKeyGroup group = parsedGroups.get(1);
                Map<String, Integer> keyToRow = sheet2KeyToRowMap.get(1);

                for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                    // 跳过已匹配的表1行
                    if (matchedSheet1Rows.contains(i)) {
                        continue;
                    }

                    Row row1 = sheet1.getRow(i);
                    if (row1 == null) continue;

                    // 检查空值过滤
                    if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                        boolean hasEmptyValue = false;
                        for (String colName : filterEmptyColumns) {
                            Integer colIndex = columnMap1.get(colName);
                            if (colIndex != null) {
                                Cell cell = row1.getCell(colIndex);
                                String value = getCellValueAsString(cell);
                                if (value == null || value.trim().isEmpty()) {
                                    hasEmptyValue = true;
                                    break;
                                }
                            }
                        }
                        if (hasEmptyValue) {
                            continue;
                        }
                    }

                    // 构建表1的关联键
                    StringBuilder keyBuilder = new StringBuilder();
                    for (String key : group.table1Keys) {
                        Integer colIndex = columnMap1.get(key);
                        if (colIndex != null) {
                            Cell cell = row1.getCell(colIndex);
                            String value = getCellValueAsString(cell);
                            keyBuilder.append(value).append("|||");
                        }
                    }
                    String key = keyBuilder.toString();

                    // 查找匹配的表2数据（排除已使用的行）
                    Integer sheet2Row = keyToRow.get(key);
                    if (sheet2Row != null && !usedSheet2Rows.contains(sheet2Row)) {
                        matchedSheet1Rows.add(i);
                        usedSheet2Rows.add(sheet2Row);
                        sheet1Matches.put(i, sheet2RowDataMap.get(1).get(sheet2Row));
                    }
                }
            }

            // 输出结果
            int outputRowIndex = 1;  // 从第1行开始，第0行是表头
            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row1 = sheet1.getRow(i);
                if (row1 == null) continue;

                // 检查空值过滤
                if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                    boolean hasEmptyValue = false;
                    for (String colName : filterEmptyColumns) {
                        Integer colIndex = columnMap1.get(colName);
                        if (colIndex != null) {
                            Cell cell = row1.getCell(colIndex);
                            String value = getCellValueAsString(cell);
                            if (value == null || value.trim().isEmpty()) {
                                hasEmptyValue = true;
                                break;
                            }
                        }
                    }
                    if (hasEmptyValue) {
                        continue;
                    }
                }

                // 复制表1的数据
                Row outputRow = outputSheet.createRow(outputRowIndex++);
                for (Cell cell : row1) {
                    Cell newCell = outputRow.createCell(cell.getColumnIndex());
                    copyCellValue(cell, newCell);
                }

                // 如果找到匹配，设置背景色并合并数据
                Map<String, Object> matchedRow2 = sheet1Matches.get(i);
                if (matchedRow2 != null) {
                    // 根据参数决定是否设置整行背景色为橙色
                    if (highlightMatches) {
                        for (Cell cell : outputRow) {
                            CellStyle newStyle = outputWorkbook.createCellStyle();
                            CellStyle originalStyle = cell.getCellStyle();
                            cloneStyle(newStyle, originalStyle);
                            newStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                            newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell.setCellStyle(newStyle);
                        }
                    }

                    // 合并表2的数据
                    for (String colName : columnsToMerge) {
                        if (outputColumnMap.containsKey(colName)) {
                            Object value = matchedRow2.get(colName);
                            Cell targetCell = outputRow.createCell(outputColumnMap.get(colName));
                            setCellValue(targetCell, value);
                        }
                    }
                }
            }

            // 应用列运算
            if (columnCalculations != null && !columnCalculations.isEmpty()) {
                applyColumnCalculations(outputSheet, outputColumnMap, columnCalculations);
            }

            // 自动调整列宽
            for (int i = 0; i < outputSheet.getRow(0).getLastCellNum(); i++) {
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
            closeQuietly(outputWorkbook);
        }
    }

    /**
     * 解析关联列组
     */
    private List<JoinKeyGroup> parseJoinKeyGroups(List<String> joinKeyGroups,
                                                   Map<String, Integer> columnMap1,
                                                   Map<String, Integer> columnMap2) {
        List<JoinKeyGroup> result = new ArrayList<>();
        for (String groupStr : joinKeyGroups) {
            String[] parts = groupStr.split("=");
            if (parts.length != 2) continue;

            String[] table1Cols = parts[0].split("[,，]");
            String[] table2Cols = parts[1].split("[,，]");

            // 去除空格并验证列存在
            List<String> t1Keys = new ArrayList<>();
            List<String> t2Keys = new ArrayList<>();

            boolean valid = true;
            for (String col : table1Cols) {
                String trimmed = col.trim();
                if (!trimmed.isEmpty() && columnMap1.containsKey(trimmed)) {
                    t1Keys.add(trimmed);
                } else if (!trimmed.isEmpty()) {
                    System.err.println("表1中不存在列: " + trimmed);
                    valid = false;
                    break;
                }
            }

            for (String col : table2Cols) {
                String trimmed = col.trim();
                if (!trimmed.isEmpty() && columnMap2.containsKey(trimmed)) {
                    t2Keys.add(trimmed);
                } else if (!trimmed.isEmpty()) {
                    System.err.println("表2中不存在列: " + trimmed);
                    valid = false;
                    break;
                }
            }

            if (valid && !t1Keys.isEmpty() && !t2Keys.isEmpty() && t1Keys.size() == t2Keys.size()) {
                result.add(new JoinKeyGroup(t1Keys, t2Keys));
            }
        }
        return result;
    }

    /**
     * 关联列组内部类
     */
    private static class JoinKeyGroup {
        List<String> table1Keys;
        List<String> table2Keys;

        JoinKeyGroup(List<String> table1Keys, List<String> table2Keys) {
            this.table1Keys = table1Keys;
            this.table2Keys = table2Keys;
        }
    }

    /**
     * 合并两个 Excel 文件，并排除表1中与表3匹配的数据
     * 将表 2 中指定的列合并到表 1，基于关联列进行匹配
     * 表1中与表3匹配的行将被排除（支持两层条件：先匹配条件1，不匹配则尝试条件2）
     *
     * @param file1            表 1（主表）
     * @param file2            表 2（合并表）
     * @param file3            表 3（排除表，可为null）
     * @param joinKeys         表1表2关联列名数组
     * @param columnsToMerge   要从表 2 合并的列名数组
     * @param excludeKeys      表1表3排除关联列名数组（条件1，可为null）
     * @param excludeKeys2     表1表3排除关联列名数组（条件2，可为null）
     * @param filterEmptyColumns 表1中需要检查空值的列名数组（可为null）
     * @param subtractMap      减法运算映射：表1列名 -> 表3列名（可为null）
     * @param outputFile       输出文件
     * @return 是否成功
     */
    public boolean mergeExcelFilesWithExclude(File file1, File file2, File file3,
                                              String[] joinKeys, String[] columnsToMerge,
                                              String[] excludeKeys, String[] excludeKeys2,
                                              String[] filterEmptyColumns,
                                              Map<String, String> subtractMap,
                                              File outputFile) {
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

                // 验证排除关联列（条件1）是否存在
                for (String keyPair : excludeKeys) {
                    // 解析格式: "表1列名=表3列名"
                    String[] parts = keyPair.split("=");
                    if (parts.length != 2) {
                        System.err.println("排除关联列1 '" + keyPair + "' 格式错误，应为 '表1列名=表3列名'");
                        return false;
                    }
                    String col1 = parts[0].trim();
                    String col3 = parts[1].trim();
                    if (!columnMap1.containsKey(col1)) {
                        System.err.println("排除关联列1 中的表1列名 '" + col1 + "' 不存在");
                        return false;
                    }
                    if (!columnMap3.containsKey(col3)) {
                        System.err.println("排除关联列1 中的表3列名 '" + col3 + "' 不存在");
                        return false;
                    }
                }

                // 验证排除关联列（条件2）是否存在
                if (excludeKeys2 != null && excludeKeys2.length > 0) {
                    for (String keyPair : excludeKeys2) {
                        // 解析格式: "表1列名=表3列名"
                        String[] parts = keyPair.split("=");
                        if (parts.length != 2) {
                            System.err.println("排除关联列2 '" + keyPair + "' 格式错误，应为 '表1列名=表3列名'");
                            return false;
                        }
                        String col1 = parts[0].trim();
                        String col3 = parts[1].trim();
                        if (!columnMap1.containsKey(col1)) {
                            System.err.println("排除关联列2 中的表1列名 '" + col1 + "' 不存在");
                            return false;
                        }
                        if (!columnMap3.containsKey(col3)) {
                            System.err.println("排除关联列2 中的表3列名 '" + col3 + "' 不存在");
                            return false;
                        }
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

            // 验证空值过滤列是否存在
            if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                for (String col : filterEmptyColumns) {
                    if (!columnMap1.containsKey(col)) {
                        System.err.println("空值检查列 '" + col + "' 在表 1 中不存在");
                        return false;
                    }
                }
            }

            // 验证减法运算列是否存在
            if (subtractMap != null && !subtractMap.isEmpty()) {
                for (Map.Entry<String, String> entry : subtractMap.entrySet()) {
                    String col1 = entry.getKey();   // 表1的列
                    String col3 = entry.getValue();  // 表3的列
                    if (!columnMap1.containsKey(col1)) {
                        System.err.println("减法运算列 '" + col1 + "' 在表 1 中不存在");
                        return false;
                    }
                    if (enableExclude && !columnMap3.containsKey(col3)) {
                        System.err.println("减法运算列 '" + col3 + "' 在表 3 中不存在");
                        return false;
                    }
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

            // 构建表 3 的排除索引（基于排除关联列 - 两层条件）
            Set<String> sheet3ExcludeKeys1 = null;
            Set<String> sheet3ExcludeKeys2 = null;
            Map<String, Map<String, Object>> sheet3DataIndex = null;
            if (enableExclude) {
                sheet3ExcludeKeys1 = buildSheet3ExcludeIndex(sheet3, excludeKeys, columnMap3);
                // 构建条件2的索引
                if (excludeKeys2 != null && excludeKeys2.length > 0) {
                    sheet3ExcludeKeys2 = buildSheet3ExcludeIndex(sheet3, excludeKeys2, columnMap3);
                }
                // 如果需要减法运算，构建表3的数据索引
                if (subtractMap != null && !subtractMap.isEmpty()) {
                    sheet3DataIndex = buildSheet3DataIndex(sheet3, excludeKeys, excludeKeys2, columnMap3);
                }
            }

            // 合并数据（排除表3中存在的数据，过滤空值行，执行减法运算）
            mergeDataWithExclude(sheet1, outputSheet, joinKeys, columnsToMerge, columnMap1,
                    sheet2Index, excludeKeys, sheet3ExcludeKeys1, excludeKeys2, sheet3ExcludeKeys2,
                    filterEmptyColumns, subtractMap, sheet3DataIndex);

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
     * 读取 Excel 文件的表头列名
     *
     * @param file Excel 文件
     * @return 表头列名列表
     */
    public List<String> readExcelHeaders(File file) {
        List<String> headers = new ArrayList<>();
        Workbook workbook = null;
        try {
            workbook = readWorkbook(file);
            if (workbook == null) {
                return headers;
            }

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    String columnName = getCellValueAsString(cell);
                    if (columnName != null && !columnName.isEmpty()) {
                        headers.add(columnName);
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("读取表头失败: " + e.getMessage());
        } finally {
            closeQuietly(workbook);
        }
        return headers;
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
     * 合并数据到输出表（带两层条件排除、空值过滤和减法运算）
     */
    private void mergeDataWithExclude(Sheet sheet1, Sheet outputSheet, String[] joinKeys,
                                      String[] columnsToMerge, Map<String, Integer> columnMap1,
                                      Map<String, Map<String, Object>> sheet2Index,
                                      String[] excludeKeys, Set<String> sheet3ExcludeKeys1,
                                      String[] excludeKeys2, Set<String> sheet3ExcludeKeys2,
                                      String[] filterEmptyColumns,
                                      Map<String, String> subtractMap,
                                      Map<String, Map<String, Object>> sheet3DataIndex) {
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

            // 检查空值过滤（如果指定列为空，跳过此行）
            if (filterEmptyColumns != null && filterEmptyColumns.length > 0) {
                boolean hasEmptyValue = false;
                for (String colName : filterEmptyColumns) {
                    Integer colIndex = columnMap1.get(colName);
                    if (colIndex != null) {
                        Cell cell = row1.getCell(colIndex);
                        String value = getCellValueAsString(cell);
                        if (value == null || value.trim().isEmpty()) {
                            hasEmptyValue = true;
                            break;
                        }
                    }
                }
                if (hasEmptyValue) {
                    continue;  // 跳过有空值的行
                }
            }

            // 两层条件排除和匹配：先检查条件1，如果条件1不匹配，再检查条件2
            boolean shouldExclude = false;
            Map<String, Object> matchedRow3 = null;

            // 条件1：检查是否在表3中
            if (excludeKeys != null && sheet3ExcludeKeys1 != null && sheet3DataIndex != null) {
                StringBuilder excludeKeyBuilder = new StringBuilder();
                for (String keyPair : excludeKeys) {
                    // 解析格式: "表1列名=表3列名"，只取表1列名
                    String[] parts = keyPair.split("=");
                    if (parts.length != 2) continue;
                    String col1Name = parts[0].trim();  // 获取表1的列名

                    int colIndex = columnMap1.get(col1Name);
                    Cell cell = row1.getCell(colIndex);
                    String value = getCellValueAsString(cell);
                    excludeKeyBuilder.append(value).append("|||");
                }
                String excludeKey = excludeKeyBuilder.toString();
                if (sheet3ExcludeKeys1.contains(excludeKey)) {
                    shouldExclude = true;
                    matchedRow3 = sheet3DataIndex.get(excludeKey);
                }
            }

            // 条件2：如果条件1不匹配，且有条件2，则检查条件2
            if (!shouldExclude && excludeKeys2 != null && sheet3ExcludeKeys2 != null && sheet3DataIndex != null) {
                StringBuilder excludeKeyBuilder = new StringBuilder();
                for (String keyPair : excludeKeys2) {
                    // 解析格式: "表1列名=表3列名"，只取表1列名
                    String[] parts = keyPair.split("=");
                    if (parts.length != 2) continue;
                    String col1Name = parts[0].trim();  // 获取表1的列名

                    int colIndex = columnMap1.get(col1Name);
                    Cell cell = row1.getCell(colIndex);
                    String value = getCellValueAsString(cell);
                    excludeKeyBuilder.append(value).append("|||");
                }
                String excludeKey = excludeKeyBuilder.toString();
                if (sheet3ExcludeKeys2.contains(excludeKey)) {
                    shouldExclude = true;
                    matchedRow3 = sheet3DataIndex.get(excludeKey);
                }
            }

            // 如果需要减法运算但没有匹配到表3的数据，跳过此行（不进行减法）
            boolean needSubtract = subtractMap != null && !subtractMap.isEmpty();
            if (needSubtract && matchedRow3 == null) {
                // 没有匹配到表3，不执行减法，直接复制原数据
            } else if (shouldExclude && needSubtract) {
                // 匹配到表3且需要减法：不排除，而是执行减法运算
                shouldExclude = false;
            }

            // 如果确实需要排除（且不需要减法，或减法已处理），跳过此行
            if (shouldExclude) {
                continue;
            }

            Row outputRow = outputSheet.createRow(outputRowIndex++);

            // 复制表 1 的数据
            for (Cell cell : row1) {
                Cell newCell = outputRow.createCell(cell.getColumnIndex());
                copyCellValue(cell, newCell);
            }

            // 执行减法运算
            if (needSubtract && matchedRow3 != null) {
                for (Map.Entry<String, String> entry : subtractMap.entrySet()) {
                    String col1 = entry.getKey();   // 表1的列
                    String col3 = entry.getValue();  // 表3的列

                    Integer colIndex1 = columnMap1.get(col1);
                    Integer outputColIndex = outputColumnMap.get(col1);

                    if (colIndex1 != null && outputColIndex != null && matchedRow3.containsKey(col3)) {
                        Object val1 = getCellValue(row1.getCell(colIndex1));
                        Object val3 = matchedRow3.get(col3);

                        // 执行减法
                        Double result = performSubtraction(val1, val3);
                        if (result != null) {
                            Cell newCell = outputRow.getCell(outputColIndex);
                            if (newCell == null) {
                                newCell = outputRow.createCell(outputColIndex);
                            }
                            newCell.setCellValue(result);
                        }
                    }
                }
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
     * 执行减法运算
     */
    private Double performSubtraction(Object val1, Object val3) {
        if (val1 == null || val3 == null) {
            return null;
        }

        double num1 = 0;
        double num3 = 0;

        // 解析表1的值
        if (val1 instanceof Number) {
            num1 = ((Number) val1).doubleValue();
        } else if (val1 instanceof String) {
            try {
                num1 = Double.parseDouble(((String) val1).trim());
            } catch (NumberFormatException e) {
                return null;
            }
        } else {
            return null;
        }

        // 解析表3的值
        if (val3 instanceof Number) {
            num3 = ((Number) val3).doubleValue();
        } else if (val3 instanceof String) {
            try {
                num3 = Double.parseDouble(((String) val3).trim());
            } catch (NumberFormatException e) {
                return null;
            }
        } else {
            return null;
        }

        return num1 - num3;
    }

    /**
     * 构建表 3 的数据索引，用于减法运算
     * excludeKeys1/2 格式: "表1列名=表3列名"
     * 返回：关联键值 -> {列名 -> 单元格值}
     */
    private Map<String, Map<String, Object>> buildSheet3DataIndex(Sheet sheet3, String[] excludeKeys1,
                                                                 String[] excludeKeys2,
                                                                 Map<String, Integer> columnMap3) {
        Map<String, Map<String, Object>> index = new HashMap<>();

        for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
            Row row = sheet3.getRow(i);
            if (row == null) continue;

            // 优先使用条件1构建键
            StringBuilder keyBuilder = new StringBuilder();
            String[] keysToUse = (excludeKeys1 != null && excludeKeys1.length > 0) ? excludeKeys1 : excludeKeys2;

            for (String keyPair : keysToUse) {
                // 解析格式: "表1列名=表3列名"，只取表3列名
                String[] parts = keyPair.split("=");
                if (parts.length != 2) continue;
                String col3Name = parts[1].trim();  // 获取表3的列名

                Integer colIndex = columnMap3.get(col3Name);
                if (colIndex == null) continue;
                Cell cell = row.getCell(colIndex);
                String value = getCellValueAsString(cell);
                keyBuilder.append(value).append("|||");
            }
            String key = keyBuilder.toString();

            // 存储该行的所有列数据
            Map<String, Object> rowData = new HashMap<>();
            for (Cell cell : row) {
                String colName = getColumnName(sheet3.getRow(0), cell.getColumnIndex());
                if (colName != null) {
                    rowData.put(colName, getCellValue(cell));
                }
            }
            index.put(key, rowData);
        }

        return index;
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
            DataFormat dataFormat = cell.getSheet().getWorkbook().createDataFormat();
            style.setDataFormat(dataFormat.getFormat("yyyy-mm-dd"));
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

        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                double numericValue = source.getNumericCellValue();
                target.setCellValue(numericValue);
                // 如果是日期格式，在目标工作簿中创建日期样式
                if (DateUtil.isCellDateFormatted(source)) {
                    CellStyle dateStyle = target.getSheet().getWorkbook().createCellStyle();
                    // 使用内置的日期格式 yyyy-mm-dd
                    DataFormat dataFormat = target.getSheet().getWorkbook().createDataFormat();
                    dateStyle.setDataFormat(dataFormat.getFormat("yyyy-mm-dd"));
                    target.setCellStyle(dateStyle);
                }
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
     * excludeKeys 格式: "表1列名=表3列名"
     * 返回：要排除的关联键值集合
     */
    private Set<String> buildSheet3ExcludeIndex(Sheet sheet3, String[] excludeKeys,
                                                Map<String, Integer> columnMap3) {
        Set<String> excludeIndex = new HashSet<>();

        for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
            Row row = sheet3.getRow(i);
            if (row == null) continue;

            // 构建排除关联键（只使用表3的列名）
            StringBuilder keyBuilder = new StringBuilder();
            for (String keyPair : excludeKeys) {
                // 解析格式: "表1列名=表3列名"，只取表3列名
                String[] parts = keyPair.split("=");
                if (parts.length != 2) {
                    System.err.println("排除关联列格式错误: " + keyPair);
                    continue;
                }
                String col3Name = parts[1].trim();  // 获取表3的列名

                Integer colIndex = columnMap3.get(col3Name);
                if (colIndex == null) {
                    System.err.println("表3中不存在列: " + col3Name);
                    continue;
                }
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
    /**
     * 克隆单元格样式
     */
    private void cloneStyle(CellStyle newStyle, CellStyle originalStyle) {
        newStyle.setAlignment(originalStyle.getAlignment());
        newStyle.setBorderBottom(originalStyle.getBorderBottom());
        newStyle.setBorderLeft(originalStyle.getBorderLeft());
        newStyle.setBorderRight(originalStyle.getBorderRight());
        newStyle.setBorderTop(originalStyle.getBorderTop());
        newStyle.setBottomBorderColor(originalStyle.getBottomBorderColor());
        newStyle.setFillBackgroundColor(originalStyle.getFillBackgroundColor());
        newStyle.setDataFormat(originalStyle.getDataFormat());
        newStyle.setHidden(originalStyle.getHidden());
        newStyle.setLocked(originalStyle.getLocked());
        newStyle.setRotation(originalStyle.getRotation());
        newStyle.setTopBorderColor(originalStyle.getTopBorderColor());
        newStyle.setVerticalAlignment(originalStyle.getVerticalAlignment());
        newStyle.setWrapText(originalStyle.getWrapText());
    }

    /**
     * 应用列运算
     */
    private void applyColumnCalculations(Sheet outputSheet, Map<String, Integer> outputColumnMap,
                                        List<ColumnCalculation> columnCalculations) {
        Row headerRow = outputSheet.getRow(0);
        if (headerRow == null) return;

        // 确保目标列存在，如果不存在则添加到表头
        Set<String> existingColumns = new HashSet<>();
        for (Cell cell : headerRow) {
            String colName = getCellValueAsString(cell);
            if (colName != null && !colName.isEmpty()) {
                existingColumns.add(colName);
            }
        }

        for (ColumnCalculation calc : columnCalculations) {
            if (!existingColumns.contains(calc.targetColumn)) {
                // 添加新列到表头
                int newColIndex = headerRow.getLastCellNum();
                Cell newHeaderCell = headerRow.createCell(newColIndex);
                newHeaderCell.setCellValue(calc.targetColumn);
                outputColumnMap.put(calc.targetColumn, newColIndex);
                existingColumns.add(calc.targetColumn);
            }
        }

        // 对每一行应用运算
        for (int i = 1; i <= outputSheet.getLastRowNum(); i++) {
            Row row = outputSheet.getRow(i);
            if (row == null) continue;

            for (ColumnCalculation calc : columnCalculations) {
                Integer col1Index = outputColumnMap.get(calc.column1);
                Integer col2Index = outputColumnMap.get(calc.column2);
                Integer targetIndex = outputColumnMap.get(calc.targetColumn);

                if (col1Index != null && col2Index != null && targetIndex != null) {
                    Double result = performColumnCalculation(row, col1Index, col2Index, calc.operator);
                    if (result != null) {
                        Cell targetCell = row.getCell(targetIndex);
                        if (targetCell == null) {
                            targetCell = row.createCell(targetIndex);
                        }
                        targetCell.setCellValue(result);
                    }
                }
            }
        }
    }

    /**
     * 执行列运算
     */
    private Double performColumnCalculation(Row row, int col1Index, int col2Index, String operator) {
        Cell cell1 = row.getCell(col1Index);
        Cell cell2 = row.getCell(col2Index);

        if (cell1 == null || cell2 == null) {
            return null;
        }

        Double val1 = getNumericValue(cell1);
        Double val2 = getNumericValue(cell2);

        if (val1 == null || val2 == null) {
            return null;
        }

        switch (operator) {
            case "add":
                return val1 + val2;
            case "subtract":
                return val1 - val2;
            case "multiply":
                return val1 * val2;
            case "divide":
                if (val2 == 0.0) {
                    return null;  // 避免除以零
                }
                return val1 / val2;
            default:
                return null;
        }
    }

    /**
     * 获取单元格的数值
     */
    private Double getNumericValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                try {
                    String str = cell.getStringCellValue().trim();
                    if (str.isEmpty()) {
                        return null;
                    }
                    return Double.parseDouble(str);
                } catch (NumberFormatException e) {
                    return null;
                }
            case BLANK:
                return null;
            default:
                return null;
        }
    }

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
