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
    private JTable columnCalcTable;  // 列运算表格
    private ColumnCalcTableModel columnCalcTableModel;  // 列运算表格模型
    private JTextArea logArea;
    private JButton executeButton;

    private File selectedFile1;
    private File selectedFile2;

    // 存储各表的列名
    private List<String> columns1;
    private List<String> columns2;

    // 存储关联列组：每个元素是一个字符串 "表1列1,表1列2=表2列1,表2列2"
    private List<String> joinKeyGroups;
    // 存储要合并的列名
    private List<String> columnsToMergeList;
    // 存储列运算规则：目标列 = 列1 运算符 列2
    private List<ColumnCalculation> columnCalculations;

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
        this.columnCalculations = new ArrayList<>();
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

        gbc.gridx = 3;
        gbc.weightx = 0;
        btnSelectColumns1 = createActionButton("选择列");
        btnSelectColumns1.setEnabled(false);
        btnSelectColumns1.addActionListener(e -> selectColumns(1));
        panel.add(btnSelectColumns1, gbc);

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

        gbc.gridx = 3;
        gbc.weightx = 0;
        btnSelectColumns2 = createActionButton("选择列");
        btnSelectColumns2.setEnabled(false);
        btnSelectColumns2.addActionListener(e -> selectColumns(2));
        panel.add(btnSelectColumns2, gbc);

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

        // 第八行：列运算配置标题
        gbc.gridx = 0;
        gbc.gridy = 7;
        gbc.weightx = 0;
        JLabel label3 = new JLabel("列运算:");
        label3.setFont(labelFont);
        label3.setForeground(labelColor);
        panel.add(label3, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        panel.add(Box.createHorizontalGlue(), gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnAddCalc = createActionButton("添加运算");
        btnAddCalc.setPreferredSize(new Dimension(80, 32));
        btnAddCalc.addActionListener(e -> addColumnCalculation());
        panel.add(btnAddCalc, gbc);

        // 第九行：列运算表格
        gbc.gridx = 0;
        gbc.gridy = 8;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        columnCalcTableModel = new ColumnCalcTableModel();
        columnCalcTable = new JTable(columnCalcTableModel);
        columnCalcTable.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        columnCalcTable.setRowHeight(28);
        columnCalcTable.getTableHeader().setFont(new Font("微软雅黑", Font.BOLD, 12));
        columnCalcTable.getTableHeader().setForeground(new Color(52, 73, 94));
        columnCalcTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        columnCalcTable.setGridColor(new Color(235, 240, 245));
        JScrollPane calcTableScrollPane = new JScrollPane(columnCalcTable);
        calcTableScrollPane.setPreferredSize(new Dimension(0, 80));
        calcTableScrollPane.setMaximumSize(new Dimension(Integer.MAX_VALUE, 80));
        calcTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        calcTableScrollPane.setBorder(BorderFactory.createLineBorder(new Color(220, 225, 230), 1));
        panel.add(calcTableScrollPane, gbc);
        gbc.gridwidth = 1;

        // 第十行：删除列运算按钮
        gbc.gridx = 2;
        gbc.gridy = 9;
        gbc.weightx = 0;
        JButton btnRemoveCalc = createActionButton("删除");
        btnRemoveCalc.setPreferredSize(new Dimension(60, 28));
        btnRemoveCalc.addActionListener(e -> removeColumnCalculation());
        panel.add(btnRemoveCalc, gbc);

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

            // 在后台线程读取表头
            final File file = selectedFile;
            new SwingWorker<List<String>, Void>() {
                @Override
                protected List<String> doInBackground() {
                    return excelService.readExcelHeaders(file);
                }

                @Override
                protected void done() {
                    try {
                        List<String> headers = get();
                        if (fileNumber == 1) {
                        selectedFile1 = file;
                        file1Field.setText(file.getAbsolutePath());
                        columns1 = headers != null ? headers : new ArrayList<>();
                        // 启用"选择列"按钮
                        if (btnSelectColumns1 != null) {
                            btnSelectColumns1.setEnabled(true);
                        }
                        // 更新关联列下拉框
                        updateComboBoxOptions();
                    } else if (fileNumber == 2) {
                        selectedFile2 = file;
                        file2Field.setText(file.getAbsolutePath());
                        columns2 = headers != null ? headers : new ArrayList<>();
                        // 启用"选择列"按钮
                        if (btnSelectColumns2 != null) {
                            btnSelectColumns2.setEnabled(true);
                        }
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

        // 禁用按钮
        executeButton.setEnabled(false);
        logArea.setText("开始执行合并操作...\n");

        // 保存为 final 变量供内部类使用
        final File file1 = selectedFile1;
        final File file2 = selectedFile2;
        final List<String> joinKeyGroupsList = new ArrayList<>(joinKeyGroups);
        final List<String> columnsToMergeListFinal = new ArrayList<>(columnsToMergeList);
        final boolean highlightMatches = highlightMatchesCheckBox.isSelected();

        // 转换列运算规则为Service格式
        final List<ExcelService.ColumnCalculation> columnCalculationsService = new ArrayList<>();
        for (ColumnCalculation calc : columnCalculations) {
            columnCalculationsService.add(new ExcelService.ColumnCalculation(
                calc.targetColumn, calc.column1, calc.operator, calc.column2
            ));
        }

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

                // 生成输出文件路径
                String outputPath = file1.getParent();
                String outputFileName = "merged_" + file1.getName();
                final File outputFile = new File(outputPath, outputFileName);

                publish("开始执行合并...");
                boolean success = excelService.mergeExcelFilesWithMultipleJoinGroups(
                        file1,
                        file2,
                        joinKeyGroupsList,
                        columnsArray,
                        null,
                        outputFile,
                        highlightMatches,
                        columnCalculationsService
                );

                if (success) {
                    lastOutputFile = outputFile;  // 保存输出文件路径
                    publish("合并完成！");
                    publish("输出文件: " + outputFile.getAbsolutePath());
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
     * 列选择对话框
     */
    private static class ColumnSelectorDialog extends JDialog {
        private JList<String> columnList;
        private JButton okButton;
        private JButton selectAllButton;
        private JButton clearButton;
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

            // 顶部说明
            JLabel hintLabel = new JLabel("请选择列名（支持多选）：");
            hintLabel.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            add(hintLabel, BorderLayout.NORTH);

            // 中间列表
            columnList = new JList<>(columns.toArray(new String[0]));
            columnList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            columnList.setFont(new Font("微软雅黑", Font.PLAIN, 13));
            JScrollPane scrollPane = new JScrollPane(columnList);
            add(scrollPane, BorderLayout.CENTER);

            // 底部按钮
            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            selectAllButton = new JButton("全选");
            selectAllButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            selectAllButton.addActionListener(e -> columnList.setSelectionInterval(0, columnList.getModel().getSize() - 1));

            clearButton = new JButton("清空");
            clearButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            clearButton.addActionListener(e -> columnList.clearSelection());

            okButton = new JButton("确定");
            okButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            okButton.addActionListener(e -> {
                selectedColumns = columnList.getSelectedValuesList();
                if (selectedColumns.isEmpty()) {
                    selectedColumns = null;
                }
                dispose();
            });

            JButton cancelButton = new JButton("取消");
            cancelButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            cancelButton.addActionListener(e -> {
                selectedColumns = null;
                dispose();
            });

            buttonPanel.add(selectAllButton);
            buttonPanel.add(clearButton);
            buttonPanel.add(Box.createHorizontalStrut(10));
            buttonPanel.add(okButton);
            buttonPanel.add(cancelButton);
            add(buttonPanel, BorderLayout.SOUTH);

            getRootPane().setDefaultButton(okButton);
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

    /**
     * 列运算规则
     */
    private static class ColumnCalculation {
        String targetColumn;  // 目标列
        String column1;       // 第一列
        String operator;      // 运算符：+, -, *, /
        String column2;       // 第二列

        ColumnCalculation(String targetColumn, String column1, String operator, String column2) {
            this.targetColumn = targetColumn;
            this.column1 = column1;
            this.operator = operator;
            this.column2 = column2;
        }

        @Override
        public String toString() {
            String opSymbol;
            switch (operator) {
                case "add": opSymbol = "+"; break;
                case "subtract": opSymbol = "-"; break;
                case "multiply": opSymbol = "×"; break;
                case "divide": opSymbol = "÷"; break;
                default: opSymbol = operator;
            }
            return targetColumn + " = " + column1 + " " + opSymbol + " " + column2;
        }
    }

    /**
     * 列运算表格模型
     */
    private static class ColumnCalcTableModel extends AbstractTableModel {
        private List<ColumnCalculation> data;
        private final String[] columnNames = {"目标列", "列1", "运算", "列2"};

        public ColumnCalcTableModel() {
            this.data = new ArrayList<>();
        }

        public void setData(List<ColumnCalculation> data) {
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
                ColumnCalculation calc = data.get(rowIndex);
                switch (columnIndex) {
                    case 0:
                        return calc.targetColumn;
                    case 1:
                        return calc.column1;
                    case 2:
                        switch (calc.operator) {
                            case "add": return "+";
                            case "subtract": return "-";
                            case "multiply": return "×";
                            case "divide": return "÷";
                            default: return calc.operator;
                        }
                    case 3:
                        return calc.column2;
                }
            }
            return "";
        }
    }

    /**
     * 添加列运算
     */
    private void addColumnCalculation() {
        // 获取所有可用的列（表1的列 + 合并后的表2的列）
        List<String> availableColumns = new ArrayList<>();
        if (columns1 != null) {
            availableColumns.addAll(columns1);
        }
        if (columns2 != null) {
            for (String col : columnsToMergeList) {
                if (!availableColumns.contains(col)) {
                    availableColumns.add(col);
                }
            }
        }

        if (availableColumns.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请先选择文件和合并列", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        ColumnCalcDialog dialog = new ColumnCalcDialog(
            (JFrame) SwingUtilities.getWindowAncestor(this),
            "添加列运算",
            availableColumns
        );
        dialog.setVisible(true);

        ColumnCalculation calc = dialog.getCalculation();
        if (calc != null) {
            columnCalculations.add(calc);
            columnCalcTableModel.setData(columnCalculations);
        }
    }

    /**
     * 删除选中的列运算
     */
    private void removeColumnCalculation() {
        int selectedIndex = columnCalcTable.getSelectedRow();
        if (selectedIndex >= 0) {
            columnCalculations.remove(selectedIndex);
            columnCalcTableModel.setData(columnCalculations);
        } else {
            JOptionPane.showMessageDialog(this, "请先选择要删除的列运算", "提示", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * 列运算配置对话框
     */
    private class ColumnCalcDialog extends JDialog {
        private JComboBox<String> targetColumnCombo;
        private JComboBox<String> column1Combo;
        private JComboBox<String> operatorCombo;
        private JComboBox<String> column2Combo;
        private JButton okButton;
        private JButton cancelButton;
        private ColumnCalculation calculation;

        private final String[] OPERATORS = {"+", "-", "×", "÷"};

        public ColumnCalcDialog(JFrame parent, String title, List<String> availableColumns) {
            super(parent, title, true);
            initComponents(availableColumns);
            setDefaultCloseOperation(DISPOSE_ON_CLOSE);
            pack();
            setLocationRelativeTo(parent);
        }

        private void initComponents(List<String> availableColumns) {
            setLayout(new BorderLayout(10, 10));
            ((JComponent) getContentPane()).setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

            // 顶部说明
            JLabel hintLabel = new JLabel("设置列运算规则：目标列 = 列1 运算符 列2");
            hintLabel.setFont(new Font("微软雅黑", Font.BOLD, 13));
            hintLabel.setForeground(new Color(52, 73, 94));
            add(hintLabel, BorderLayout.NORTH);

            // 中间配置面板
            JPanel configPanel = new JPanel(new GridBagLayout());
            configPanel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(new Color(200, 205, 210), 1),
                "运算配置",
                javax.swing.border.TitledBorder.LEFT,
                javax.swing.border.TitledBorder.TOP,
                new Font("微软雅黑", Font.BOLD, 12),
                new Color(52, 73, 94)
            ));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(8, 10, 8, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.anchor = GridBagConstraints.WEST;

            Font labelFont = new Font("微软雅黑", Font.PLAIN, 13);
            Font comboFont = new Font("微软雅黑", Font.PLAIN, 13);

            // 目标列
            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.weightx = 0;
            JLabel targetLabel = new JLabel("目标列:");
            targetLabel.setFont(labelFont);
            configPanel.add(targetLabel, gbc);

            gbc.gridx = 1;
            gbc.weightx = 1.0;
            targetColumnCombo = new JComboBox<>();
            targetColumnCombo.setFont(comboFont);
            targetColumnCombo.setPreferredSize(new Dimension(200, 28));
            targetColumnCombo.setEditable(true);
            for (String col : availableColumns) {
                targetColumnCombo.addItem(col);
            }
            configPanel.add(targetColumnCombo, gbc);

            // 列1
            gbc.gridx = 0;
            gbc.gridy = 1;
            gbc.weightx = 0;
            JLabel col1Label = new JLabel("列1:");
            col1Label.setFont(labelFont);
            configPanel.add(col1Label, gbc);

            gbc.gridx = 1;
            gbc.weightx = 1.0;
            column1Combo = new JComboBox<>();
            column1Combo.setFont(comboFont);
            column1Combo.setPreferredSize(new Dimension(200, 28));
            for (String col : availableColumns) {
                column1Combo.addItem(col);
            }
            configPanel.add(column1Combo, gbc);

            // 运算符
            gbc.gridx = 0;
            gbc.gridy = 2;
            gbc.weightx = 0;
            JLabel opLabel = new JLabel("运算符:");
            opLabel.setFont(labelFont);
            configPanel.add(opLabel, gbc);

            gbc.gridx = 1;
            gbc.weightx = 1.0;
            operatorCombo = new JComboBox<>(OPERATORS);
            operatorCombo.setFont(comboFont);
            operatorCombo.setPreferredSize(new Dimension(200, 28));
            configPanel.add(operatorCombo, gbc);

            // 列2
            gbc.gridx = 0;
            gbc.gridy = 3;
            gbc.weightx = 0;
            JLabel col2Label = new JLabel("列2:");
            col2Label.setFont(labelFont);
            configPanel.add(col2Label, gbc);

            gbc.gridx = 1;
            gbc.weightx = 1.0;
            column2Combo = new JComboBox<>();
            column2Combo.setFont(comboFont);
            column2Combo.setPreferredSize(new Dimension(200, 28));
            for (String col : availableColumns) {
                column2Combo.addItem(col);
            }
            configPanel.add(column2Combo, gbc);

            add(configPanel, BorderLayout.CENTER);

            // 底部按钮
            JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            okButton = new JButton("确定");
            okButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            okButton.setPreferredSize(new Dimension(70, 28));
            okButton.addActionListener(e -> confirmCalculation());
            cancelButton = new JButton("取消");
            cancelButton.setFont(new Font("微软雅黑", Font.PLAIN, 12));
            cancelButton.setPreferredSize(new Dimension(70, 28));
            cancelButton.addActionListener(e -> {
                calculation = null;
                dispose();
            });
            bottomPanel.add(okButton);
            bottomPanel.add(cancelButton);
            add(bottomPanel, BorderLayout.SOUTH);

            getRootPane().setDefaultButton(okButton);
        }

        private void confirmCalculation() {
            Object target = targetColumnCombo.getSelectedItem();
            Object col1 = column1Combo.getSelectedItem();
            Object op = operatorCombo.getSelectedItem();
            Object col2 = column2Combo.getSelectedItem();

            if (target == null || target.toString().trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "请输入或选择目标列", "提示", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            if (col1 == null || col1.toString().trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "请选择列1", "提示", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            if (col2 == null || col2.toString().trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "请选择列2", "提示", JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            String targetCol = target.toString().trim();
            String col1Str = col1.toString().trim();
            String opStr = op.toString();
            String col2Str = col2.toString().trim();

            // 转换运算符
            String operator;
            switch (opStr) {
                case "+": operator = "add"; break;
                case "-": operator = "subtract"; break;
                case "×": operator = "multiply"; break;
                case "÷": operator = "divide"; break;
                default: operator = "add";
            }

            calculation = new ColumnCalculation(targetCol, col1Str, operator, col2Str);
            dispose();
        }

        public ColumnCalculation getCalculation() {
            return calculation;
        }
    }
}
