package com.saicmotor.maxus.rv2go.ui;

import com.saicmotor.maxus.rv2go.service.ExcelService;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.*;
import java.util.List;
import java.util.prefs.Preferences;

import static java.awt.Desktop.getDesktop;

/**
 * 数据清洗功能面板
 * 功能：导入 Excel，选择需要的列后导出
 */
public class DataCleanPanel extends JPanel {
    private JTextField fileField;
    private JButton btnSelectColumns;
    private JList<String> selectedColumnsList;
    private DefaultListModel<String> selectedColumnsModel;
    private JTextArea logArea;
    private JButton executeButton;

    private File selectedFile;
    private List<String> columns;
    private List<String> selectedColumns;

    private final ExcelService excelService;
    private final Preferences prefs;
    private File lastOutputFile;

    public DataCleanPanel() {
        this.excelService = new ExcelService();
        this.prefs = Preferences.userNodeForPackage(DataCleanPanel.class);
        this.columns = new ArrayList<>();
        this.selectedColumns = new ArrayList<>();
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

        JLabel titleLabel = new JLabel("数据清洗工具");
        titleLabel.setFont(new Font("微软雅黑", Font.BOLD, 20));
        titleLabel.setForeground(new Color(44, 62, 80));
        titleLeft.add(titleLabel);

        JLabel descLabel = new JLabel("  - 导入 Excel，选择需要的列后导出");
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

        // 列选择区域
        JPanel columnPanel = createColumnSelectionPanel();
        contentPanel.add(columnPanel);
        contentPanel.add(Box.createVerticalStrut(12));

        // 执行按钮
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.setBackground(new Color(245, 247, 250));
        executeButton = new JButton("▶ 执行导出");
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
        executeButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) {
                executeButton.setBackground(new Color(0, 120, 110));
            }
            public void mouseExited(java.awt.event.MouseEvent e) {
                executeButton.setBackground(new Color(0, 150, 136));
            }
        });
        executeButton.addActionListener(this::executeExport);
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

        // 文件选择
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel label = new JLabel("选择文件:");
        label.setFont(labelFont);
        label.setForeground(labelColor);
        panel.add(label, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        fileField = new JTextField(30);
        fileField.setEditable(false);
        fileField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        fileField.setPreferredSize(new Dimension(0, 32));
        fileField.setBackground(new Color(250, 252, 255));
        fileField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        panel.add(fileField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse = createActionButton("浏览...");
        btnBrowse.addActionListener(e -> selectFile());
        panel.add(btnBrowse, gbc);

        gbc.gridx = 3;
        gbc.weightx = 0;
        btnSelectColumns = createActionButton("选择列");
        btnSelectColumns.setEnabled(false);
        btnSelectColumns.addActionListener(e -> selectColumns());
        panel.add(btnSelectColumns, gbc);

        return panel;
    }

    private JPanel createColumnSelectionPanel() {
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

        // 标签
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel label = new JLabel("已选列:");
        label.setFont(labelFont);
        label.setForeground(labelColor);
        panel.add(label, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        panel.add(Box.createHorizontalGlue(), gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnClear = createActionButton("清空");
        btnClear.setPreferredSize(new Dimension(60, 28));
        btnClear.addActionListener(e -> clearSelectedColumns());
        panel.add(btnClear, gbc);

        // 列表
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        selectedColumnsModel = new DefaultListModel<>();
        selectedColumnsList = new JList<>(selectedColumnsModel);
        selectedColumnsList.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        selectedColumnsList.setVisibleRowCount(8);
        selectedColumnsList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        selectedColumnsList.setBackground(new Color(253, 254, 255));
        selectedColumnsList.setBorder(BorderFactory.createEmptyBorder(8, 10, 8, 10));
        JScrollPane listScrollPane = new JScrollPane(selectedColumnsList);
        listScrollPane.setPreferredSize(new Dimension(0, 150));
        listScrollPane.setMaximumSize(new Dimension(Integer.MAX_VALUE, 150));
        listScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        listScrollPane.setBorder(BorderFactory.createLineBorder(new Color(220, 225, 230), 1));
        panel.add(listScrollPane, gbc);
        gbc.gridwidth = 1;

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

    private void selectFile() {
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
            File file = fileChooser.getSelectedFile();

            // 保存当前目录到 Preferences
            File currentDir = fileChooser.getCurrentDirectory();
            if (currentDir != null) {
                prefs.put("lastDirectory", currentDir.getAbsolutePath());
            }

            // 在后台线程读取表头
            new SwingWorker<List<String>, Void>() {
                @Override
                protected List<String> doInBackground() {
                    return excelService.readExcelHeaders(file);
                }

                @Override
                protected void done() {
                    try {
                        List<String> headers = get();
                        selectedFile = file;
                        fileField.setText(file.getAbsolutePath());
                        columns = headers != null ? headers : new ArrayList<>();
                        btnSelectColumns.setEnabled(true);
                        // 清空之前选择的列
                        selectedColumns.clear();
                        updateSelectedColumnsList();
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(DataCleanPanel.this,
                            "读取文件失败: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }.execute();
        }
    }

    private void selectColumns() {
        if (columns == null || columns.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                "请先选择文件", "提示", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        ColumnSelectorDialog dialog = new ColumnSelectorDialog(
            (JFrame) SwingUtilities.getWindowAncestor(this),
            "选择需要导出的列",
            columns
        );

        dialog.setVisible(true);
        List<String> selected = dialog.getSelectedColumns();

        if (selected != null && !selected.isEmpty()) {
            selectedColumns.clear();
            selectedColumns.addAll(selected);
            updateSelectedColumnsList();
        }
    }

    private void clearSelectedColumns() {
        selectedColumns.clear();
        updateSelectedColumnsList();
    }

    private void updateSelectedColumnsList() {
        selectedColumnsModel.clear();
        for (String col : selectedColumns) {
            selectedColumnsModel.addElement(col);
        }
    }

    private void executeExport(ActionEvent e) {
        // 验证输入
        if (selectedFile == null) {
            JOptionPane.showMessageDialog(this, "请选择 Excel 文件", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (selectedColumns.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请选择需要导出的列", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 禁用按钮
        executeButton.setEnabled(false);
        logArea.setText("开始执行导出操作...\n");

        // 保存为 final 变量供内部类使用
        final File inputFile = selectedFile;
        final List<String> columnsToExport = new ArrayList<>(selectedColumns);

        // 在后台线程执行
        new SwingWorker<Boolean, String>() {
            private Exception caughtException = null;

            @Override
            protected Boolean doInBackground() {
                try {
                    publish("正在读取文件: " + inputFile.getName());
                    publish("导出列: " + String.join(", ", columnsToExport));

                    // 生成输出文件路径
                    String outputPath = inputFile.getParent();
                    String outputFileName = "cleaned_" + inputFile.getName();
                    final File outputFile = new File(outputPath, outputFileName);

                    publish("开始执行导出...");
                    boolean success = excelService.exportSelectedColumns(
                            inputFile,
                            columnsToExport,
                            outputFile
                    );

                    if (success) {
                        lastOutputFile = outputFile;
                        publish("导出完成！");
                        publish("输出文件: " + outputFile.getAbsolutePath());
                    } else {
                        publish("导出失败，请检查日志");
                    }

                    return success;
                } catch (Exception ex) {
                    caughtException = ex;
                    publish("错误: " + ex.getClass().getSimpleName() + " - " + ex.getMessage());
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

                if (caughtException != null) {
                    logArea.append("\n=== 执行失败 ===\n");
                    logArea.append("错误类型: " + caughtException.getClass().getName() + "\n");
                    logArea.append("错误信息: " + caughtException.getMessage() + "\n");
                    JOptionPane.showMessageDialog(DataCleanPanel.this,
                            "执行失败: " + caughtException.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                try {
                    if (get()) {
                        Object[] options = {"确定", "打开文件夹"};
                        int choice = JOptionPane.showOptionDialog(
                                DataCleanPanel.this,
                                "数据导出成功！\n输出文件: " + lastOutputFile.getName(),
                                "成功",
                                JOptionPane.YES_NO_OPTION,
                                JOptionPane.INFORMATION_MESSAGE,
                                null,
                                options,
                                options[0]
                        );

                        if (choice == 1) {
                            openFileLocation(lastOutputFile);
                        }
                    } else {
                        logArea.append("\n=== 导出失败 ===\n");
                        logArea.append("返回值为 false，请检查配置参数\n");
                        JOptionPane.showMessageDialog(DataCleanPanel.this,
                                "数据导出失败，请检查输入配置和日志", "失败", JOptionPane.ERROR_MESSAGE);
                    }
                } catch (Exception ex) {
                    logArea.append("\n=== 系统错误 ===\n");
                    logArea.append("错误: " + ex.getMessage() + "\n");
                    for (StackTraceElement element : ex.getStackTrace()) {
                        logArea.append("    " + element.toString() + "\n");
                    }
                    JOptionPane.showMessageDialog(DataCleanPanel.this,
                            "执行过程中发生错误: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    private void openFileLocation(File file) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop desktop = getDesktop();
                if (desktop.isSupported(Desktop.Action.OPEN)) {
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
            JLabel hintLabel = new JLabel("选择需要导出的列：");
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
}
