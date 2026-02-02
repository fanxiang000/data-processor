package com.saicmotor.maxus.rv2go.ui;

import com.saicmotor.maxus.rv2go.service.ExcelService;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.*;
import java.util.List;
import java.util.prefs.Preferences;

import static java.awt.Desktop.getDesktop;

/**
 * Excel 合并功能面板
 * 功能：上传两个 Excel，关联列名匹配，将表 2 中的指定列合并到表 1
 */
public class ExcelMergePanel extends JPanel {
    private JTextField file1Field;
    private JTextField file2Field;
    private JComboBox<String> sheet1Combo;  // 表1 sheet页选择下拉框
    private JComboBox<String> sheet2Combo;  // 表2 sheet页选择下拉框
    private JButton btnSelectColumns1;  // 表1列选择按钮
    private JButton btnSelectColumns2;  // 表2列选择按钮
    private JTextField joinKeysField;
    private JComboBox<String> joinKeysCombo1;  // 表1关联列下拉框
    private JComboBox<String> joinKeysCombo2;  // 表2关联列下拉框
    private JButton btnAddJoinKeyGroup;  // 添加关联列组按钮
    private JTable joinKeyGroupsTable;  // 关联列组表格
    private JoinKeyGroupsTableModel joinKeyGroupsTableModel;  // 关联列组表格模型
    private JTextField columnsToMergeField;
    private JComboBox<String> columnsToMergeCombo;  // 表2合并列下拉框
    private JButton btnAddMergeColumn;  // 添加合并列按钮
    private JList<String> mergeColumnsList;  // 合并列列表
    private DefaultListModel<String> mergeColumnsModel;  // 合并列列表模型
    private JCheckBox highlightMatchesCheckBox;  // 高亮匹配行复选框
    private JTextField headerRow1Field;  // 表1表头开始行
    private JTextField headerRow2Field;  // 表2表头开始行
    private JTextArea logArea;
    private JButton executeButton;

    private File selectedFile1;
    private File selectedFile2;

    // 存储选中的sheet页索引
    private int selectedSheetIndex1 = 0;  // 表1选中的sheet索引，默认0
    private int selectedSheetIndex2 = 0;  // 表2选中的sheet索引，默认0

    // 存储表头开始行（从0开始）
    private int headerRow1 = 0;  // 表1表头开始行，默认0
    private int headerRow2 = 0;  // 表2表头开始行，默认0

    // 存储各表的列名
    private List<String> columns1;
    private List<String> columns2;

    // 存储关联列组：每个元素是一个字符串 "表1列1,表1列2=表2列1,表2列2"
    private List<String> joinKeyGroups;
    // 存储要合并的列名
    private List<String> columnsToMergeList;

    private final ExcelService excelService;
    private final Preferences prefs;
    private File lastOutputFile;  // 保存最后输出的文件路径

    public ExcelMergePanel() {
        this.excelService = new ExcelService();
        this.prefs = Preferences.userNodeForPackage(ExcelMergePanel.class);
        this.columns1 = new ArrayList<>();
        this.columns2 = new ArrayList<>();
        this.joinKeyGroups = new ArrayList<>();
        this.columnsToMergeList = new ArrayList<>();
        initComponents();
    }

    private void initComponents() {
        setLayout(new BorderLayout(15, 15));
        setBorder(BorderFactory.createEmptyBorder(15, 20, 15, 20));
        setBackground(new Color(245, 247, 250));

        // 顶部标题区域
        JPanel titlePanel = new JPanel(new BorderLayout());
        titlePanel.setBackground(new Color(245, 247, 250));
        titlePanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 10, 5));

        JPanel titleLeft = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        titleLeft.setBackground(new Color(245, 247, 250));

        JLabel titleLabel = new JLabel("Excel 合并工具");
        titleLabel.setFont(new Font("微软雅黑", Font.BOLD, 20));
        titleLabel.setForeground(new Color(44, 62, 80));
        titleLeft.add(titleLabel);

        JLabel descLabel = new JLabel("  - 将两个 Excel 文件按关联列合并");
        descLabel.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        descLabel.setForeground(new Color(127, 140, 141));
        titleLeft.add(descLabel);

        titlePanel.add(titleLeft, BorderLayout.WEST);
        add(titlePanel, BorderLayout.NORTH);

        // 中间内容面板（可滚动）
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
        contentPanel.setBackground(new Color(245, 247, 250));

        // 文件选择区域
        JPanel filePanel = createFileSelectionPanel();
        contentPanel.add(filePanel);
        contentPanel.add(Box.createVerticalStrut(12));

        // 配置区域
        JPanel configPanel = createConfigPanel();
        contentPanel.add(configPanel);
        contentPanel.add(Box.createVerticalStrut(12));

        // 执行按钮
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.setBackground(new Color(245, 247, 250));
        executeButton = new JButton("▶ 执行合并");
        executeButton.setFont(new Font("微软雅黑", Font.BOLD, 15));
        executeButton.setPreferredSize(new Dimension(160, 48));
        executeButton.setFocusPainted(false);
        executeButton.setContentAreaFilled(false);
        executeButton.setOpaque(true);
        executeButton.setBackground(new Color(0, 150, 136));
        executeButton.setForeground(Color.WHITE);
        executeButton.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(0, 120, 110), 2),
            BorderFactory.createEmptyBorder(12, 24, 12, 24)
        ));
        executeButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
        // 悬停效果
        executeButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) {
                executeButton.setBackground(new Color(0, 120, 110));
            }
            public void mouseExited(java.awt.event.MouseEvent e) {
                executeButton.setBackground(new Color(0, 150, 136));
            }
        });
        executeButton.addActionListener(this::executeMerge);
        buttonPanel.add(executeButton);
        contentPanel.add(buttonPanel);

        JScrollPane contentScrollPane = new JScrollPane(contentPanel);
        contentScrollPane.setBorder(null);
        contentScrollPane.setOpaque(false);
        contentScrollPane.getViewport().setOpaque(false);
        contentScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        add(contentScrollPane, BorderLayout.CENTER);

        // 底部日志区域
        JPanel logPanel = new JPanel(new BorderLayout(5, 5));
        logPanel.setBackground(new Color(245, 247, 250));
        logPanel.setBorder(BorderFactory.createTitledBorder(
            BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
            "执行日志",
            javax.swing.border.TitledBorder.LEFT,
            javax.swing.border.TitledBorder.TOP,
            new Font("微软雅黑", Font.BOLD, 12),
            new Color(52, 73, 94)
        ));
        logArea = new JTextArea(8, 50);
        logArea.setEditable(false);
        logArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
        logArea.setBackground(new Color(253, 254, 255));
        logArea.setForeground(new Color(52, 73, 94));
        logArea.setMargin(new Insets(8, 10, 8, 10));
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logScrollPane.setBorder(BorderFactory.createEmptyBorder());
        logScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        logPanel.add(logScrollPane, BorderLayout.CENTER);
        add(logPanel, BorderLayout.SOUTH);
    }

    private JPanel createFileSelectionPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(new Color(255, 255, 255));
        panel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(210, 215, 220), 1),
            BorderFactory.createEmptyBorder(15, 15, 15, 15)
        ));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(6, 8, 6, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        Font labelFont = new Font("微软雅黑", Font.PLAIN, 13);
        Color labelColor = new Color(52, 73, 94);

        // 表 1 文件选择
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel label1 = new JLabel("表 1（主表）:");
        label1.setFont(labelFont);
        label1.setForeground(labelColor);
        panel.add(label1, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        file1Field = new JTextField(30);
        file1Field.setEditable(false);
        file1Field.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        file1Field.setPreferredSize(new Dimension(0, 32));
        file1Field.setBackground(new Color(250, 252, 255));
        file1Field.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        panel.add(file1Field, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse1 = createActionButton("浏览...");
        btnBrowse1.addActionListener(e -> selectFile(1));
        panel.add(btnBrowse1, gbc);

        // 表 1 sheet 页选择（新行）
        gbc.gridx = 1;
        gbc.gridy = 2;
        gbc.weightx = 1.0;
        sheet1Combo = new JComboBox<>();
        sheet1Combo.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        sheet1Combo.setPreferredSize(new Dimension(0, 32));
        sheet1Combo.setBackground(new Color(250, 252, 255));
        sheet1Combo.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        sheet1Combo.addActionListener(e -> onSheet1Changed());
        panel.add(sheet1Combo, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        headerRow1Field = new JTextField("0");
        headerRow1Field.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        headerRow1Field.setPreferredSize(new Dimension(50, 32));
        headerRow1Field.setBackground(new Color(250, 252, 255));
        headerRow1Field.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        headerRow1Field.setToolTipText("表头开始的行号（从0开始，0表示第1行）");
        panel.add(headerRow1Field, gbc);

        gbc.gridx = 0;
        gbc.weightx = 0;
        JLabel sheet1Label = new JLabel("Sheet页:");
        sheet1Label.setFont(labelFont);
        sheet1Label.setForeground(labelColor);
        panel.add(sheet1Label, gbc);

        // 表 2 文件选择
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.weightx = 0;
        JLabel label2 = new JLabel("表 2（合并表）:");
        label2.setFont(labelFont);
        label2.setForeground(labelColor);
        panel.add(label2, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        file2Field = new JTextField(30);
        file2Field.setEditable(false);
        file2Field.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        file2Field.setPreferredSize(new Dimension(0, 32));
        file2Field.setBackground(new Color(250, 252, 255));
        file2Field.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        panel.add(file2Field, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse2 = createActionButton("浏览...");
        btnBrowse2.addActionListener(e -> selectFile(2));
        panel.add(btnBrowse2, gbc);

        // 表 2 sheet 页选择（新行）
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.weightx = 0;
        JLabel sheet2Label = new JLabel("Sheet页:");
        sheet2Label.setFont(labelFont);
        sheet2Label.setForeground(labelColor);
        panel.add(sheet2Label, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        sheet2Combo = new JComboBox<>();
        sheet2Combo.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        sheet2Combo.setPreferredSize(new Dimension(0, 32));
        sheet2Combo.setBackground(new Color(250, 252, 255));
        sheet2Combo.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        sheet2Combo.addActionListener(e -> onSheet2Changed());
        panel.add(sheet2Combo, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        headerRow2Field = new JTextField("0");
        headerRow2Field.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        headerRow2Field.setPreferredSize(new Dimension(50, 32));
        headerRow2Field.setBackground(new Color(250, 252, 255));
        headerRow2Field.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        headerRow2Field.setToolTipText("表头开始的行号（从0开始，0表示第1行）");
        panel.add(headerRow2Field, gbc);

        return panel;
    }

    private JButton createActionButton(String text) {
        JButton button = new JButton(text);
        button.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        button.setPreferredSize(new Dimension(75, 32));
        button.setFocusPainted(false);
        button.setBackground(new Color(240, 244, 248));
        button.setForeground(new Color(52, 73, 94));
        button.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
            BorderFactory.createEmptyBorder(6, 12, 6, 12)
        ));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        return button;
    }

    private JPanel createConfigPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(new Color(255, 255, 255));
        panel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(210, 215, 220), 1),
            BorderFactory.createEmptyBorder(15, 15, 15, 15)
        ));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(6, 8, 6, 8);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.anchor = GridBagConstraints.WEST;

        Font labelFont = new Font("微软雅黑", Font.PLAIN, 13);
        Color labelColor = new Color(52, 73, 94);
        Font fieldFont = new Font("微软雅黑", Font.PLAIN, 13);

        // 第一行：关联列配置标题和添加按钮
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel label1 = new JLabel("关联列:");
        label1.setFont(labelFont);
        label1.setForeground(labelColor);
        panel.add(label1, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        panel.add(Box.createHorizontalGlue(), gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnAddJoinKey = createActionButton("添加");
        btnAddJoinKey.setPreferredSize(new Dimension(60, 32));
        btnAddJoinKey.addActionListener(e -> addJoinKeyGroup());
        panel.add(btnAddJoinKey, gbc);

        // 第二行：关联列表格
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        joinKeyGroupsTableModel = new JoinKeyGroupsTableModel();
        joinKeyGroupsTable = new JTable(joinKeyGroupsTableModel);
        joinKeyGroupsTable.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        joinKeyGroupsTable.setRowHeight(40);
        joinKeyGroupsTable.getTableHeader().setFont(new Font("微软雅黑", Font.BOLD, 12));
        joinKeyGroupsTable.getTableHeader().setForeground(new Color(52, 73, 94));
        joinKeyGroupsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        joinKeyGroupsTable.setGridColor(new Color(235, 240, 245));
        joinKeyGroupsTable.setAutoCreateRowSorter(true);

        JScrollPane tableScrollPane = new JScrollPane(joinKeyGroupsTable);
        tableScrollPane.setPreferredSize(new Dimension(0, 120));
        tableScrollPane.setMaximumSize(new Dimension(Integer.MAX_VALUE, 120));
        tableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        tableScrollPane.setBorder(BorderFactory.createLineBorder(new Color(220, 225, 230), 1));
        panel.add(tableScrollPane, gbc);
        gbc.gridwidth = 1;

        // 第三行：删除关联列按钮
        gbc.gridx = 2;
        gbc.gridy = 2;
        gbc.weightx = 0;
        JButton btnRemoveJoinKey = createActionButton("删除");
        btnRemoveJoinKey.setPreferredSize(new Dimension(60, 28));
        btnRemoveJoinKey.addActionListener(e -> removeJoinKeyGroup());
        panel.add(btnRemoveJoinKey, gbc);

        // 第四行：合并列配置
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.weightx = 0;
        JLabel label2 = new JLabel("合并列:");
        label2.setFont(labelFont);
        label2.setForeground(labelColor);
        panel.add(label2, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        columnsToMergeCombo = new JComboBox<>();
        columnsToMergeCombo.setFont(fieldFont);
        columnsToMergeCombo.setPreferredSize(new Dimension(0, 32));
        columnsToMergeCombo.addItem("--选择表2列--");
        panel.add(columnsToMergeCombo, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        btnAddMergeColumn = createActionButton("添加");
        btnAddMergeColumn.setPreferredSize(new Dimension(60, 32));
        btnAddMergeColumn.addActionListener(e -> addMergeColumn());
        panel.add(btnAddMergeColumn, gbc);

        // 第五行：合并列列表
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        mergeColumnsModel = new DefaultListModel<>();
        mergeColumnsList = new JList<>(mergeColumnsModel);
        mergeColumnsList.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        mergeColumnsList.setVisibleRowCount(6);
        mergeColumnsList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        mergeColumnsList.setBackground(new Color(253, 254, 255));
        mergeColumnsList.setBorder(BorderFactory.createEmptyBorder(8, 10, 8, 10));
        JScrollPane mergeListScrollPane = new JScrollPane(mergeColumnsList);
        mergeListScrollPane.setPreferredSize(new Dimension(0, 100));
        mergeListScrollPane.setMaximumSize(new Dimension(Integer.MAX_VALUE, 100));
        mergeListScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        mergeListScrollPane.setBorder(BorderFactory.createLineBorder(new Color(220, 225, 230), 1));
        panel.add(mergeListScrollPane, gbc);
        gbc.gridwidth = 1;

        // 第六行：删除合并列按钮
        gbc.gridx = 2;
        gbc.gridy = 5;
        gbc.weightx = 0;
        JButton btnRemoveMerge = createActionButton("删除");
        btnRemoveMerge.setPreferredSize(new Dimension(60, 28));
        btnRemoveMerge.addActionListener(e -> removeMergeColumn());
        panel.add(btnRemoveMerge, gbc);

        // 第七行：高亮匹配行复选框
        gbc.gridx = 0;
        gbc.gridy = 6;
        gbc.weightx = 0;
        gbc.gridwidth = 3;
        highlightMatchesCheckBox = new JCheckBox("高亮标记匹配的行（橙色背景）");
        highlightMatchesCheckBox.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        highlightMatchesCheckBox.setForeground(labelColor);
        highlightMatchesCheckBox.setBackground(new Color(255, 255, 255));
        highlightMatchesCheckBox.setSelected(true);  // 默认选中
        panel.add(highlightMatchesCheckBox, gbc);
        gbc.gridwidth = 1;

        // 隐藏的字段（保留用于兼容）
        joinKeysField = new JTextField();
        joinKeysField.setVisible(false);
        columnsToMergeField = new JTextField();
        columnsToMergeField.setVisible(false);
        btnAddJoinKeyGroup = new JButton();  // 保留用于兼容
        joinKeysCombo1 = new JComboBox<>();  // 保留用于兼容
        joinKeysCombo2 = new JComboBox<>();  // 保留用于兼容

        return panel;
    }

    private void selectFile(int fileNumber) {
        JFileChooser fileChooser = new JFileChooser();

        // 恢复上次打开的目录
        String lastPath = prefs.get("lastDirectory", null);
        if (lastPath != null) {
            File lastDir = new File(lastPath);
            if (lastDir.exists() && lastDir.isDirectory()) {
                fileChooser.setCurrentDirectory(lastDir);
            }
        }

        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                return f.isDirectory() || f.getName().toLowerCase().endsWith(".xlsx")
                        || f.getName().toLowerCase().endsWith(".xls");
            }

            @Override
            public String getDescription() {
                return "Excel 文件 (*.xlsx, *.xls)";
            }
        });

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();

            // 保存当前目录到 Preferences
            File currentDir = fileChooser.getCurrentDirectory();
            if (currentDir != null) {
                prefs.put("lastDirectory", currentDir.getAbsolutePath());
            }

            // 在后台线程读取sheet页列表和表头
            final File file = selectedFile;
            new SwingWorker<Object[], Void>() {
                @Override
                protected Object[] doInBackground() {
                    List<String> sheetNames = excelService.readSheetNames(file);
                    List<String> headers = excelService.readExcelHeaders(file, 0);  // 默认读取第一个sheet
                    return new Object[]{sheetNames, headers};
                }

                @Override
                protected void done() {
                    try {
                        Object[] result = get();
                        List<String> sheetNames = (List<String>) result[0];
                        List<String> headers = (List<String>) result[1];

                        if (fileNumber == 1) {
                            selectedFile1 = file;
                            file1Field.setText(file.getAbsolutePath());

                            // 更新sheet页下拉框
                            sheet1Combo.removeAllItems();
                            if (sheetNames != null && !sheetNames.isEmpty()) {
                                for (int i = 0; i < sheetNames.size(); i++) {
                                    sheet1Combo.addItem((i + 1) + ". " + sheetNames.get(i));
                                }
                                selectedSheetIndex1 = 0;
                            }

                            columns1 = headers != null ? headers : new ArrayList<>();
                            // 更新关联列下拉框
                            updateComboBoxOptions();
                        } else if (fileNumber == 2) {
                            selectedFile2 = file;
                            file2Field.setText(file.getAbsolutePath());

                            // 更新sheet页下拉框
                            sheet2Combo.removeAllItems();
                            if (sheetNames != null && !sheetNames.isEmpty()) {
                                for (int i = 0; i < sheetNames.size(); i++) {
                                    sheet2Combo.addItem((i + 1) + ". " + sheetNames.get(i));
                                }
                                selectedSheetIndex2 = 0;
                            }

                            columns2 = headers != null ? headers : new ArrayList<>();
                            // 更新关联列下拉框
                            updateComboBoxOptions();
                        }
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }.execute();
        }
    }

    private void executeMerge(ActionEvent e) {
        // 验证输入
        if (selectedFile1 == null || selectedFile2 == null) {
            JOptionPane.showMessageDialog(this, "请选择表 1 和表 2 的 Excel 文件", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 验证关联列组
        if (joinKeyGroups.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请至少添加一个关联列组", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 验证合并列
        if (columnsToMergeList.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请至少添加一个要合并的列", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 读取表头开始行
        try {
            headerRow1 = Integer.parseInt(headerRow1Field.getText().trim());
            headerRow2 = Integer.parseInt(headerRow2Field.getText().trim());
            if (headerRow1 < 0 || headerRow2 < 0) {
                JOptionPane.showMessageDialog(this, "表头开始行不能为负数", "错误", JOptionPane.ERROR_MESSAGE);
                return;
            }
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "表头开始行必须是有效的整数", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 禁用按钮
        executeButton.setEnabled(false);
        logArea.setText("开始执行合并操作...\n");

        // 保存为 final 变量供内部类使用
        final File file1 = selectedFile1;
        final File file2 = selectedFile2;
        final List<String> joinKeyGroupsList = new ArrayList<>(joinKeyGroups);
        final List<String> columnsToMergeListFinal = new ArrayList<>(columnsToMergeList);
        final boolean highlightMatches = highlightMatchesCheckBox.isSelected();
        final int headerRow1Final = headerRow1;
        final int headerRow2Final = headerRow2;

        // 选择输出文件路径（在后台线程之前）
        String outputFileName = "merged_" + file1.getName();
        File outputFile = chooseSaveFile(outputFileName);
        if (outputFile == null) {
            executeButton.setEnabled(true);
            logArea.append("用户取消了保存操作\n");
            return;
        }
        final File finalOutputFile = outputFile;

        // 在后台线程执行
        new SwingWorker<Boolean, String>() {
            private Exception caughtException = null;

            @Override
            protected Boolean doInBackground() {
                try {
                String[] columnsArray = columnsToMergeListFinal.toArray(new String[0]);

                publish("正在读取表 1: " + file1.getName());
                publish("正在读取表 2: " + file2.getName());

                // 显示关联列组信息
                for (int i = 0; i < joinKeyGroupsList.size(); i++) {
                    String[] parts = joinKeyGroupsList.get(i).split("=");
                    if (parts.length == 2) {
                        publish("关联列组" + (i + 1) + ": 表1[" + parts[0] + "] ⇔ 表2[" + parts[1] + "]");
                    }
                }
                publish("合并列: " + String.join(", ", columnsArray));
                publish("表1表头行: " + (headerRow1Final + 1));
                publish("表2表头行: " + (headerRow2Final + 1));

                publish("输出文件: " + finalOutputFile.getAbsolutePath());
                publish("开始执行合并...");
                boolean success = excelService.mergeExcelFilesWithMultipleJoinGroups(
                        file1,
                        file2,
                        selectedSheetIndex1,
                        selectedSheetIndex2,
                        headerRow1Final,
                        headerRow2Final,
                        joinKeyGroupsList,
                        columnsArray,
                        null,
                        finalOutputFile,
                        highlightMatches,
                        null
                );

                if (success) {
                    lastOutputFile = finalOutputFile;  // 保存输出文件路径
                    publish("合并完成！");
                } else {
                    publish("合并失败，请检查日志");
                }

                return success;
                } catch (Exception ex) {
                    caughtException = ex;
                    publish("错误: " + ex.getClass().getSimpleName() + " - " + ex.getMessage());
                    // 打印堆栈跟踪到日志
                    java.io.StringWriter sw = new java.io.StringWriter();
                    java.io.PrintWriter pw = new java.io.PrintWriter(sw);
                    ex.printStackTrace(pw);
                    publish("堆栈跟踪:\n" + sw.toString());
                    return false;
                }
            }

            @Override
            protected void process(List<String> chunks) {
                for (String message : chunks) {
                    logArea.append(message + "\n");
                }
            }

            @Override
            protected void done() {
                executeButton.setEnabled(true);

                // 首先检查是否在 doInBackground 中捕获了异常
                if (caughtException != null) {
                    logArea.append("\n=== 执行失败 ===\n");
                    logArea.append("错误类型: " + caughtException.getClass().getName() + "\n");
                    logArea.append("错误信息: " + caughtException.getMessage() + "\n");
                    JOptionPane.showMessageDialog(ExcelMergePanel.this,
                            "执行失败: " + caughtException.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                try {
                    if (get()) {
                        // 显示成功对话框，提供打开文件夹选项
                        Object[] options = {"确定", "打开文件夹"};
                        int choice = JOptionPane.showOptionDialog(
                                ExcelMergePanel.this,
                                "Excel 合并成功！\n输出文件: " + lastOutputFile.getName(),
                                "成功",
                                JOptionPane.YES_NO_OPTION,
                                JOptionPane.INFORMATION_MESSAGE,
                                null,
                                options,
                                options[0]
                        );

                        // 如果选择"打开文件夹"
                        if (choice == 1) {
                            openFileLocation(lastOutputFile);
                        }
                    } else {
                        logArea.append("\n=== 合并失败 ===\n");
                        logArea.append("返回值为 false，请检查配置参数\n");
                        JOptionPane.showMessageDialog(ExcelMergePanel.this,
                                "Excel 合并失败，请检查输入配置和日志", "失败", JOptionPane.ERROR_MESSAGE);
                    }
                } catch (Exception ex) {
                    logArea.append("\n=== 系统错误 ===\n");
                    logArea.append("错误: " + ex.getMessage() + "\n");
                    for (StackTraceElement element : ex.getStackTrace()) {
                        logArea.append("    " + element.toString() + "\n");
                    }
                    JOptionPane.showMessageDialog(ExcelMergePanel.this,
                            "执行过程中发生错误: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    /**
     * 添加关联列组（支持多对选择）
     */
    private void addJoinKeyGroup() {
        // 检查列是否已加载
        if (columns1 == null || columns1.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请先选择表1文件", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        if (columns2 == null || columns2.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请先选择表2文件", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // 使用多对选择对话框
        SinglePairSelectorDialog dialog = new SinglePairSelectorDialog(
            (JFrame) SwingUtilities.getWindowAncestor(this),
            "添加关联列",
            columns1,
            columns2
        );
        dialog.setVisible(true);

        List<Pair> pairs = dialog.getSelectedPairs();
        if (pairs == null || pairs.isEmpty()) {
            return;  // 用户取消
        }

        // 将所有配对组合成一个组
        if (!pairs.isEmpty()) {
            StringBuilder table1Cols = new StringBuilder();
            StringBuilder table2Cols = new StringBuilder();
            for (int i = 0; i < pairs.size(); i++) {
                Pair pair = pairs.get(i);
                if (i > 0) {
                    table1Cols.append(",");
                    table2Cols.append(",");
                }
                table1Cols.append(pair.column1);
                table2Cols.append(pair.column2);
            }
            String groupStr = table1Cols.toString() + "=" + table2Cols.toString();
            joinKeyGroups.add(groupStr);
        }

        // 更新列表显示
        updateJoinKeyGroupsList();
    }

    /**
     * 选择保存文件路径
     */
    private File chooseSaveFile(String defaultFileName) {
        JFileChooser fileChooser = new JFileChooser();

        // 恢复上次打开的目录
        String lastPath = prefs.get("lastDirectory", null);
        if (lastPath != null) {
            File lastDir = new File(lastPath);
            if (lastDir.exists() && lastDir.isDirectory()) {
                fileChooser.setCurrentDirectory(lastDir);
            }
        }

        fileChooser.setDialogTitle("选择保存位置");
        fileChooser.setSelectedFile(new File(defaultFileName));

        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                return f.isDirectory() || f.getName().toLowerCase().endsWith(".xlsx")
                        || f.getName().toLowerCase().endsWith(".xls");
            }

            @Override
            public String getDescription() {
                return "Excel 文件 (*.xlsx, *.xls)";
            }
        });

        int result = fileChooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();

            // 保存当前目录到 Preferences
            File currentDir = fileChooser.getCurrentDirectory();
            if (currentDir != null) {
                prefs.put("lastDirectory", currentDir.getAbsolutePath());
            }

            // 确保文件有扩展名
            String fileName = selectedFile.getName();
            if (!fileName.toLowerCase().endsWith(".xlsx") && !fileName.toLowerCase().endsWith(".xls")) {
                selectedFile = new File(selectedFile.getParentFile(), fileName + ".xlsx");
            }

            return selectedFile;
        }

        return null;  // 用户取消
    }

    /**
     * 列对
     */
    private static class Pair {
        String column1;
        String column2;

        Pair(String column1, String column2) {
            this.column1 = column1;
            this.column2 = column2;
        }

        @Override
        public String toString() {
            return column1 + " ⇔ " + column2;
        }
    }

    /**
     * 逐对选择列对话框
     */
    private class SinglePairSelectorDialog extends JDialog {
        private JComboBox<String> table1Combo;
        private JComboBox<String> table2Combo;
        private JTable pairsTable;
        private PairsTableModel pairsTableModel;
        private JButton addButton;
        private JButton removeButton;
        private JButton okButton;
        private JButton cancelButton;
        private List<Pair> selectedPairs;

        public SinglePairSelectorDialog(JFrame parent, String title, List<String> columns1, List<String> columns2) {
            super(parent, title, true);
            this.selectedPairs = new ArrayList<>();
            initComponents(columns1, columns2);
            setDefaultCloseOperation(DISPOSE_ON_CLOSE);
            pack();
            setLocationRelativeTo(parent);
        }

        private void initComponents(List<String> columns1, List<String> columns2) {
            setLayout(new BorderLayout(10, 10));
            ((JComponent) getContentPane()).setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

            // 顶部说明
            JLabel hintLabel = new JLabel("添加关联列组（可添加多个配对，将组合为一个组）：");
            hintLabel.setFont(new Font("微软雅黑", Font.BOLD, 13));
            hintLabel.setForeground(new Color(52, 73, 94));
            add(hintLabel, BorderLayout.NORTH);

            // 中间主面板
            JPanel mainPanel = new JPanel(new BorderLayout(10, 10));

            // 已选择的配对列表
            pairsTableModel = new PairsTableModel();
            pairsTable = new JTable(pairsTableModel);
            pairsTable.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            pairsTable.setRowHeight(28);
            pairsTable.getTableHeader().setFont(new Font("微软雅黑", Font.BOLD, 12));
            pairsTable.getTableHeader().setForeground(new Color(52, 73, 94));
            pairsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
            pairsTable.setGridColor(new Color(235, 240, 245));
            JScrollPane tableScrollPane = new JScrollPane(pairsTable);
            tableScrollPane.setPreferredSize(new Dimension(500, 150));
            tableScrollPane.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
                "已选择的配对",
                javax.swing.border.TitledBorder.LEFT,
                javax.swing.border.TitledBorder.TOP,
                new Font("微软雅黑", Font.BOLD, 12),
                new Color(52, 73, 94)
            ));
            mainPanel.add(tableScrollPane, BorderLayout.CENTER);

            // 当前选择区
            JPanel selectionPanel = new JPanel(new GridBagLayout());
            selectionPanel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
                "添加新配对",
                javax.swing.border.TitledBorder.LEFT,
                javax.swing.border.TitledBorder.TOP,
                new Font("微软雅黑", Font.BOLD, 12),
                new Color(52, 73, 94)
            ));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(8, 10, 8, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.anchor = GridBagConstraints.CENTER;

            // 表1列下拉框
            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.weightx = 1.0;
            JLabel label1 = new JLabel("表1列:");
            label1.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            selectionPanel.add(label1, gbc);

            gbc.gridy = 1;
            table1Combo = new JComboBox<>();
            table1Combo.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            table1Combo.setPreferredSize(new Dimension(200, 28));
            for (String col : columns1) {
                table1Combo.addItem(col);
            }
            selectionPanel.add(table1Combo, gbc);

            // 中间符号
            gbc.gridx = 1;
            gbc.gridy = 0;
            gbc.gridheight = 2;
            gbc.weightx = 0;
            JLabel arrowLabel = new JLabel("⇔");
            arrowLabel.setFont(new Font("微软雅黑", Font.BOLD, 20));
            arrowLabel.setForeground(new Color(52, 152, 219));
            selectionPanel.add(arrowLabel, gbc);

            // 表2列下拉框
            gbc.gridx = 2;
            gbc.gridy = 0;
            gbc.gridheight = 1;
            gbc.weightx = 1.0;
            JLabel label2 = new JLabel("表2列:");
            label2.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            selectionPanel.add(label2, gbc);

            gbc.gridy = 1;
            table2Combo = new JComboBox<>();
            table2Combo.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            table2Combo.setPreferredSize(new Dimension(200, 28));
            for (String col : columns2) {
                table2Combo.addItem(col);
            }
            selectionPanel.add(table2Combo, gbc);

            // 添加按钮
            gbc.gridx = 3;
            gbc.gridy = 0;
            gbc.gridheight = 2;
            gbc.weightx = 0;
            addButton = new JButton("+ 添加配对");
            addButton.setFont(new Font("微软雅黑", Font.BOLD, 13));
            addButton.setPreferredSize(new Dimension(100, 36));
            addButton.setFocusPainted(false);
            addButton.setContentAreaFilled(false);
            addButton.setOpaque(true);
            addButton.setBackground(new Color(46, 204, 113));
            addButton.setForeground(Color.WHITE);
            addButton.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(39, 174, 96), 1),
                BorderFactory.createEmptyBorder(8, 16, 8, 16)
            ));
            addButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
            // 悬停效果
            addButton.addMouseListener(new java.awt.event.MouseAdapter() {
                public void mouseEntered(java.awt.event.MouseEvent e) {
                    addButton.setBackground(new Color(39, 174, 96));
                }
                public void mouseExited(java.awt.event.MouseEvent e) {
                    addButton.setBackground(new Color(46, 204, 113));
                }
            });
            addButton.addActionListener(e -> addPair());
            selectionPanel.add(addButton, gbc);

            mainPanel.add(selectionPanel, BorderLayout.SOUTH);

            add(mainPanel, BorderLayout.CENTER);

            // 底部按钮
            JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));

            removeButton = new JButton("删除选中");
            removeButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            removeButton.addActionListener(e -> removeSelectedPair());

            okButton = new JButton("确定");
            okButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            okButton.setPreferredSize(new Dimension(70, 28));
            okButton.addActionListener(e -> confirmSelection());

            cancelButton = new JButton("取消");
            cancelButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            cancelButton.setPreferredSize(new Dimension(70, 28));
            cancelButton.addActionListener(e -> {
                selectedPairs = null;
                dispose();
            });

            bottomPanel.add(removeButton);
            bottomPanel.add(Box.createHorizontalStrut(10));
            bottomPanel.add(okButton);
            bottomPanel.add(cancelButton);
            add(bottomPanel, BorderLayout.SOUTH);

            getRootPane().setDefaultButton(addButton);
        }

        private void addPair() {
            Object col1 = table1Combo.getSelectedItem();
            Object col2 = table2Combo.getSelectedItem();

            if (col1 == null || col2 == null) {
                JOptionPane.showMessageDialog(this, "请选择表1列和表2列", "提示", JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            String col1Str = col1.toString();
            String col2Str = col2.toString();

            // 检查是否已存在
            for (Pair p : selectedPairs) {
                if (p.column1.equals(col1Str) && p.column2.equals(col2Str)) {
                    JOptionPane.showMessageDialog(this, "该配对已存在", "提示", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }
            }

            selectedPairs.add(new Pair(col1Str, col2Str));
            pairsTableModel.setData(selectedPairs);

            // 自动选中新添加的行
            pairsTable.setRowSelectionInterval(selectedPairs.size() - 1, selectedPairs.size() - 1);

            // 焦点返回到添加按钮，方便继续添加
            addButton.requestFocus();
        }

        private void removeSelectedPair() {
            int selectedRow = pairsTable.getSelectedRow();
            if (selectedRow >= 0) {
                selectedPairs.remove(selectedRow);
                pairsTableModel.setData(selectedPairs);
            } else {
                JOptionPane.showMessageDialog(this, "请先选择要删除的配对", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        }

        private void confirmSelection() {
            if (selectedPairs.isEmpty()) {
                JOptionPane.showMessageDialog(this, "请至少添加一个配对", "提示", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            dispose();
        }

        public List<Pair> getSelectedPairs() {
            return selectedPairs.isEmpty() ? null : selectedPairs;
        }

        /**
         * 配对表格模型
         */
        private class PairsTableModel extends AbstractTableModel {
            private List<Pair> data;
            private final String[] columnNames = {"序号", "表1列", "表2列"};

            public PairsTableModel() {
                this.data = new ArrayList<>();
            }

            public void setData(List<Pair> data) {
                this.data = new ArrayList<>(data);
                fireTableDataChanged();
            }

            @Override
            public int getRowCount() {
                return data.size();
            }

            @Override
            public int getColumnCount() {
                return columnNames.length;
            }

            @Override
            public String getColumnName(int column) {
                return columnNames[column];
            }

            @Override
            public Object getValueAt(int rowIndex, int columnIndex) {
                if (rowIndex >= 0 && rowIndex < data.size()) {
                    Pair pair = data.get(rowIndex);
                    switch (columnIndex) {
                        case 0:
                            return rowIndex + 1;
                        case 1:
                            return pair.column1;
                        case 2:
                            return pair.column2;
                    }
                }
                return "";
            }
        }
    }


    /**
     * 删除选中的关联列组
     */
    private void removeJoinKeyGroup() {
        int selectedIndex = joinKeyGroupsTable.getSelectedRow();
        if (selectedIndex >= 0) {
            joinKeyGroups.remove(selectedIndex);
            updateJoinKeyGroupsList();
        } else {
            JOptionPane.showMessageDialog(this, "请先选择要删除的关联列组", "提示", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * 从下拉框获取选中的列（支持手动输入的列）
     */
    private Set<String> getSelectedColumnsFromCombo(JComboBox<String> combo, List<String> availableColumns) {
        Set<String> result = new HashSet<>();
        Object selected = combo.getSelectedItem();

        if (selected != null && !selected.toString().startsWith("--")) {
            String selectedStr = selected.toString();
            // 如果是手动输入的列名（不在可用列表中），直接添加
            if (availableColumns.contains(selectedStr)) {
                result.add(selectedStr);
            } else {
                // 尝试按逗号分割（支持手动输入多个列名）
                String[] parts = selectedStr.split("[,，]");
                for (String part : parts) {
                    String trimmed = part.trim();
                    if (!trimmed.isEmpty()) {
                        result.add(trimmed);
                    }
                }
            }
        }

        return result;
    }

    /**
     * 更新关联列组表格显示
     */
    private void updateJoinKeyGroupsList() {
        joinKeyGroupsTableModel.setData(joinKeyGroups);
        // 自动调整列宽
        autoResizeTableColumns(joinKeyGroupsTable);
    }

    /**
     * 自动调整表格列宽
     */
    private void autoResizeTableColumns(JTable table) {
        final int margin = 10;
        for (int column = 0; column < table.getColumnCount(); column++) {
            int width = 50; // 最小宽度
            for (int row = 0; row < table.getRowCount(); row++) {
                javax.swing.table.TableCellRenderer renderer = table.getCellRenderer(row, column);
                Component comp = table.prepareRenderer(renderer, row, column);
                width = Math.max(width, comp.getPreferredSize().width + margin);
            }
            table.getColumnModel().getColumn(column).setPreferredWidth(width);
        }
    }

    /**
     * 更新下拉框选项
     */
    private void updateComboBoxOptions() {
        // 更新表1关联列下拉框
        joinKeysCombo1.removeAllItems();
        joinKeysCombo1.addItem("--选择表1列--");
        if (columns1 != null && !columns1.isEmpty()) {
            for (String col : columns1) {
                joinKeysCombo1.addItem(col);
            }
        }

        // 更新表2关联列下拉框
        joinKeysCombo2.removeAllItems();
        joinKeysCombo2.addItem("--选择表2列--");
        if (columns2 != null && !columns2.isEmpty()) {
            for (String col : columns2) {
                joinKeysCombo2.addItem(col);
            }
        }

        // 更新表2合并列下拉框
        columnsToMergeCombo.removeAllItems();
        columnsToMergeCombo.addItem("--选择表2列--");
        if (columns2 != null && !columns2.isEmpty()) {
            for (String col : columns2) {
                columnsToMergeCombo.addItem(col);
            }
        }
    }

    /**
     * 添加要合并的列
     */
    private void addMergeColumn() {
        Object selected = columnsToMergeCombo.getSelectedItem();
        if (selected == null || selected.toString().startsWith("--")) {
            JOptionPane.showMessageDialog(this, "请选择一个表2的列", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        String colName = selected.toString();
        if (columnsToMergeList.contains(colName)) {
            JOptionPane.showMessageDialog(this, "该列已添加", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        columnsToMergeList.add(colName);
        updateMergeColumnsList();

        // 重置下拉框选择
        columnsToMergeCombo.setSelectedIndex(0);
    }

    /**
     * 删除选中的合并列
     */
    private void removeMergeColumn() {
        int selectedIndex = mergeColumnsList.getSelectedIndex();
        if (selectedIndex >= 0) {
            columnsToMergeList.remove(selectedIndex);
            updateMergeColumnsList();
        } else {
            JOptionPane.showMessageDialog(this, "请先选择要删除的列", "提示", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * 更新合并列列表显示
     */
    private void updateMergeColumnsList() {
        mergeColumnsModel.clear();
        for (String col : columnsToMergeList) {
            mergeColumnsModel.addElement(col);
        }
    }

    /**
     * 更新表2合并列文本框（保留用于兼容）
     */
    private void updateColumnsToMergeFromCombo() {
        Object selected = columnsToMergeCombo.getSelectedItem();
        if (selected != null && !selected.toString().startsWith("--")) {
            appendToField(columnsToMergeField, selected.toString());
        }
    }

    /**
     * 打开文件所在文件夹
     */
    private void openFileLocation(File file) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop desktop = getDesktop();
                if (desktop.isSupported(Desktop.Action.OPEN)) {
                    // 获取父目录
                    File parentDir = file.getParentFile();
                    if (parentDir != null && parentDir.exists()) {
                        desktop.open(parentDir);
                    } else {
                        JOptionPane.showMessageDialog(this,
                                "无法打开文件夹，路径不存在", "错误", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "打开文件夹失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * 表1 sheet页选择变化处理
     */
    private void onSheet1Changed() {
        if (sheet1Combo.getItemCount() == 0 || selectedFile1 == null) {
            return;
        }

        int selectedIndex = sheet1Combo.getSelectedIndex();
        if (selectedIndex < 0) {
            return;
        }

        selectedSheetIndex1 = selectedIndex;

        // 在后台线程重新读取列名
        final File file = selectedFile1;
        final int sheetIndex = selectedIndex;
        new SwingWorker<List<String>, Void>() {
            @Override
            protected List<String> doInBackground() {
                return excelService.readExcelHeaders(file, sheetIndex);
            }

            @Override
            protected void done() {
                try {
                    List<String> headers = get();
                    columns1 = headers != null ? headers : new ArrayList<>();
                    // 更新关联列下拉框
                    updateComboBoxOptions();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        }.execute();
    }

    /**
     * 表2 sheet页选择变化处理
     */
    private void onSheet2Changed() {
        if (sheet2Combo.getItemCount() == 0 || selectedFile2 == null) {
            return;
        }

        int selectedIndex = sheet2Combo.getSelectedIndex();
        if (selectedIndex < 0) {
            return;
        }

        selectedSheetIndex2 = selectedIndex;

        // 在后台线程重新读取列名
        final File file = selectedFile2;
        final int sheetIndex = selectedIndex;
        new SwingWorker<List<String>, Void>() {
            @Override
            protected List<String> doInBackground() {
                return excelService.readExcelHeaders(file, sheetIndex);
            }

            @Override
            protected void done() {
                try {
                    List<String> headers = get();
                    columns2 = headers != null ? headers : new ArrayList<>();
                    // 清空合并列列表和关联列组（因为列名可能变化了）
                    columnsToMergeList.clear();
                    updateMergeColumnsList();
                    joinKeyGroups.clear();
                    updateJoinKeyGroupsList();
                    // 更新关联列下拉框
                    updateComboBoxOptions();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        }.execute();
    }

    /**
     * 选择列名（从文件选择对话框点击）
     */
    private void selectColumns(int fileNumber) {
        List<String> columns = null;
        String tableName = "";

        switch (fileNumber) {
            case 1:
                columns = columns1;
                tableName = "表1";
                break;
            case 2:
                columns = columns2;
                tableName = "表2";
                break;
            default:
                return;
        }

        if (columns == null || columns.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                "请先选择 " + tableName + " 文件", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        ColumnSelectorDialog dialog = new ColumnSelectorDialog(
            (JFrame) SwingUtilities.getWindowAncestor(this),
            "选择 " + tableName + " 的列名",
            columns
        );

        dialog.setVisible(true);
        List<String> selected = dialog.getSelectedColumns();

        if (selected != null && !selected.isEmpty()) {
            String newText = String.join(", ", selected);
            // 根据fileNumber决定添加到哪个输入框
            if (fileNumber == 1) {
                appendToField(joinKeysField, newText);
            } else if (fileNumber == 2) {
                appendToField(columnsToMergeField, newText);
            }
        }
    }

    /**
     * 为配置字段选择列名
     */
    private void selectColumnsForField(String fieldName, int fileNumber) {
        List<String> columns = null;
        JTextField targetField = null;

        switch (fieldName) {
            case "关联列名":
                if (fileNumber == 1) {
                    columns = columns1;
                    targetField = joinKeysField;
                } else if (fileNumber == 2) {
                    columns = columns2;
                    targetField = joinKeysField;
                }
                break;
            case "表2合并列":
                columns = columns2;
                targetField = columnsToMergeField;
                break;
        }

        if (columns == null || columns.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                "请先选择对应的 Excel 文件", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        ColumnSelectorDialog dialog = new ColumnSelectorDialog(
            (JFrame) SwingUtilities.getWindowAncestor(this),
            "选择列名",
            columns
        );

        dialog.setVisible(true);
        List<String> selected = dialog.getSelectedColumns();

        if (selected != null && !selected.isEmpty()) {
            appendToField(targetField, String.join(", ", selected));
        }
    }

    /**
     * 将文本追加到输入框
     */
    private void appendToField(JTextField field, String text) {
        String current = field.getText().trim();
        if (!current.isEmpty()) {
            text = current + ", " + text;
        }
        field.setText(text);
    }

    /**
     * 列选择对话框（双列表穿梭框）
     */
    private static class ColumnSelectorDialog extends JDialog {
        private JList<String> availableList;      // 左边：可用列
        private JList<String> selectedList;       // 右边：已选列
        private DefaultListModel<String> availableModel;
        private DefaultListModel<String> selectedModel;
        private JButton btnAdd;
        private JButton btnRemove;
        private JButton btnAddAll;
        private JButton btnRemoveAll;
        private JButton okButton;
        private List<String> selectedColumns;

        public ColumnSelectorDialog(JFrame parent, String title, List<String> columns) {
            super(parent, title, true);
            this.selectedColumns = new ArrayList<>();
            initComponents(columns);
            setDefaultCloseOperation(DISPOSE_ON_CLOSE);
            pack();
            setLocationRelativeTo(parent);
        }

        private void initComponents(List<String> columns) {
            setLayout(new BorderLayout(10, 10));
            ((JComponent) getContentPane()).setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

            // 顶部说明
            JPanel hintPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
            JLabel hintLabel = new JLabel("选择需要的列名：");
            hintLabel.setFont(new Font("微软雅黑", Font.BOLD, 13));
            hintLabel.setForeground(new Color(52, 73, 94));
            hintPanel.add(hintLabel);
            add(hintPanel, BorderLayout.NORTH);

            // 中间主面板（左右列表）
            JPanel mainPanel = new JPanel(new GridBagLayout());
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(5, 5, 5, 5);
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 1.0;

            Font listFont = new Font("微软雅黑", Font.PLAIN, 13);

            // 左边：可用列
            JPanel leftPanel = createListPanel("可用列", listFont);
            availableModel = new DefaultListModel<>();
            for (String col : columns) {
                availableModel.addElement(col);
            }
            availableList = new JList<>(availableModel);
            availableList.setFont(listFont);
            availableList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            JScrollPane leftScroll = new JScrollPane(availableList);
            leftScroll.setPreferredSize(new Dimension(200, 300));
            leftPanel.add(leftScroll, BorderLayout.CENTER);

            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.weightx = 1.0;
            mainPanel.add(leftPanel, gbc);

            // 中间：移动按钮
            JPanel buttonPanel = new JPanel(new GridLayout(4, 1, 5, 5));
            buttonPanel.setBorder(BorderFactory.createEmptyBorder(30, 10, 30, 10));

            btnAdd = createMoveButton("▶ 添加");
            btnAdd.addActionListener(e -> moveSelected(availableList, availableModel, selectedModel));

            btnRemove = createMoveButton("◀ 移除");
            btnRemove.addActionListener(e -> moveSelected(selectedList, selectedModel, availableModel));

            btnAddAll = createMoveButton("▶▶ 全部");
            btnAddAll.addActionListener(e -> moveAll(availableModel, selectedModel));

            btnRemoveAll = createMoveButton("◀◀ 全部");
            btnRemoveAll.addActionListener(e -> moveAll(selectedModel, availableModel));

            buttonPanel.add(btnAdd);
            buttonPanel.add(btnRemove);
            buttonPanel.add(btnAddAll);
            buttonPanel.add(btnRemoveAll);

            gbc.gridx = 1;
            gbc.gridy = 0;
            gbc.weightx = 0;
            mainPanel.add(buttonPanel, gbc);

            // 右边：已选列
            JPanel rightPanel = createListPanel("已选列", listFont);
            selectedModel = new DefaultListModel<>();
            selectedList = new JList<>(selectedModel);
            selectedList.setFont(listFont);
            selectedList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            JScrollPane rightScroll = new JScrollPane(selectedList);
            rightScroll.setPreferredSize(new Dimension(200, 300));
            rightPanel.add(rightScroll, BorderLayout.CENTER);

            gbc.gridx = 2;
            gbc.gridy = 0;
            gbc.weightx = 1.0;
            mainPanel.add(rightPanel, gbc);

            add(mainPanel, BorderLayout.CENTER);

            // 底部按钮
            JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            okButton = new JButton("确定");
            okButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            okButton.setPreferredSize(new Dimension(70, 28));
            okButton.addActionListener(e -> {
                // 将已选列转换为列表
                selectedColumns = new ArrayList<>();
                for (int i = 0; i < selectedModel.size(); i++) {
                    selectedColumns.add(selectedModel.getElementAt(i));
                }
                if (selectedColumns.isEmpty()) {
                    selectedColumns = null;
                }
                dispose();
            });

            JButton cancelButton = new JButton("取消");
            cancelButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            cancelButton.setPreferredSize(new Dimension(70, 28));
            cancelButton.addActionListener(e -> {
                selectedColumns = null;
                dispose();
            });

            bottomPanel.add(okButton);
            bottomPanel.add(cancelButton);
            add(bottomPanel, BorderLayout.SOUTH);

            getRootPane().setDefaultButton(okButton);
        }

        /**
         * 创建列表面板
         */
        private JPanel createListPanel(String title, Font font) {
            JPanel panel = new JPanel(new BorderLayout());
            panel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
                title,
                javax.swing.border.TitledBorder.LEFT,
                javax.swing.border.TitledBorder.TOP,
                new Font("微软雅黑", Font.BOLD, 12),
                new Color(52, 73, 94)
            ));
            panel.setBackground(new Color(255, 255, 255));
            return panel;
        }

        /**
         * 创建移动按钮
         */
        private JButton createMoveButton(String text) {
            JButton btn = new JButton(text);
            btn.setFont(new Font("微软雅黑", Font.PLAIN, 11));
            btn.setPreferredSize(new Dimension(80, 30));
            btn.setFocusPainted(false);
            btn.setBackground(new Color(240, 244, 248));
            btn.setForeground(new Color(52, 73, 94));
            btn.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
                BorderFactory.createEmptyBorder(5, 10, 5, 10)
            ));
            btn.setCursor(new Cursor(Cursor.HAND_CURSOR));
            // 悬停效果
            btn.addMouseListener(new java.awt.event.MouseAdapter() {
                public void mouseEntered(java.awt.event.MouseEvent e) {
                    btn.setBackground(new Color(220, 230, 240));
                }
                public void mouseExited(java.awt.event.MouseEvent e) {
                    btn.setBackground(new Color(240, 244, 248));
                }
            });
            return btn;
        }

        /**
         * 移动选中的项目
         */
        private void moveSelected(JList<String> sourceList, DefaultListModel<String> sourceModel,
                                  DefaultListModel<String> targetModel) {
            int[] selectedIndices = sourceList.getSelectedIndices();
            if (selectedIndices.length == 0) {
                return;
            }

            // 从后往前删除，避免索引变化
            List<String> toMove = new ArrayList<>();
            for (int i = selectedIndices.length - 1; i >= 0; i--) {
                String item = sourceModel.remove(selectedIndices[i]);
                toMove.add(0, item);
            }

            // 添加到目标列表
            for (String item : toMove) {
                targetModel.addElement(item);
            }
        }

        /**
         * 移动所有项目
         */
        private void moveAll(DefaultListModel<String> sourceModel, DefaultListModel<String> targetModel) {
            while (sourceModel.size() > 0) {
                String item = sourceModel.remove(0);
                targetModel.addElement(item);
            }
        }

        public List<String> getSelectedColumns() {
            return selectedColumns;
        }
    }

    /**
     * 关联列组表格模型
     */
    private static class JoinKeyGroupsTableModel extends AbstractTableModel {
        private List<String> data;
        private final String[] columnNames = {"序号", "表1列", "表2列"};

        public JoinKeyGroupsTableModel() {
            this.data = new ArrayList<>();
        }

        public void setData(List<String> data) {
            this.data = new ArrayList<>(data);
            fireTableDataChanged();
        }

        @Override
        public int getRowCount() {
            return data.size();
        }

        @Override
        public int getColumnCount() {
            return columnNames.length;
        }

        @Override
        public String getColumnName(int column) {
            return columnNames[column];
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            if (rowIndex >= 0 && rowIndex < data.size()) {
                String group = data.get(rowIndex);
                String[] parts = group.split("=");
                if (parts.length == 2) {
                    switch (columnIndex) {
                        case 0:
                            return "组" + (rowIndex + 1);
                        case 1:
                            // 将列名用逗号分隔显示
                            return parts[0].replace(",", ", ");
                        case 2:
                            // 将列名用逗号分隔显示
                            return parts[1].replace(",", ", ");
                    }
                }
            }
            return "";
        }
    }

}
