package com.saicmotor.maxus.rv2go.ui;

import com.saicmotor.maxus.rv2go.service.ExcelService;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.List;

/**
 * Excel 合并功能面板
 * 功能：上传两个 Excel，关联列名匹配，将表 2 中的指定列合并到表 1
 */
public class ExcelMergePanel extends JPanel {
    private JTextField file1Field;
    private JTextField file2Field;
    private JTextField joinKeysField;
    private JTextField columnsToMergeField;
    private JTextArea logArea;
    private JButton executeButton;

    private File selectedFile1;
    private File selectedFile2;

    private final ExcelService excelService;

    public ExcelMergePanel() {
        this.excelService = new ExcelService();
        initComponents();
    }

    private void initComponents() {
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 顶部标题
        JPanel titlePanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JLabel titleLabel = new JLabel("Excel 合并工具");
        titleLabel.setFont(new Font("微软雅黑", Font.BOLD, 18));
        titlePanel.add(titleLabel);
        add(titlePanel, BorderLayout.NORTH);

        // 中间内容面板
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));

        // 文件选择区域
        JPanel filePanel = createFileSelectionPanel();
        contentPanel.add(filePanel);
        contentPanel.add(Box.createVerticalStrut(10));

        // 配置区域
        JPanel configPanel = createConfigPanel();
        contentPanel.add(configPanel);
        contentPanel.add(Box.createVerticalStrut(10));

        // 执行按钮
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        executeButton = new JButton("执行合并");
        executeButton.setFont(new Font("微软雅黑", Font.PLAIN, 14));
        executeButton.addActionListener(this::executeMerge);
        buttonPanel.add(executeButton);
        contentPanel.add(buttonPanel);

        add(contentPanel, BorderLayout.CENTER);

        // 底部日志区域
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(BorderFactory.createTitledBorder("执行日志"));
        logArea = new JTextArea(8, 50);
        logArea.setEditable(false);
        logArea.setFont(new Font("微软雅黑", Font.PLAIN, 12));
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logPanel.add(logScrollPane, BorderLayout.CENTER);
        add(logPanel, BorderLayout.SOUTH);
    }

    private JPanel createFileSelectionPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder("文件选择"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // 表 1 文件选择
        gbc.gridx = 0;
        gbc.gridy = 0;
        panel.add(new JLabel("表 1（主表）:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        file1Field = new JTextField(30);
        file1Field.setEditable(false);
        panel.add(file1Field, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse1 = new JButton("浏览...");
        btnBrowse1.addActionListener(e -> selectFile(1));
        panel.add(btnBrowse1, gbc);

        // 表 2 文件选择
        gbc.gridx = 0;
        gbc.gridy = 1;
        panel.add(new JLabel("表 2（合并表）:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        file2Field = new JTextField(30);
        file2Field.setEditable(false);
        panel.add(file2Field, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse2 = new JButton("浏览...");
        btnBrowse2.addActionListener(e -> selectFile(2));
        panel.add(btnBrowse2, gbc);

        return panel;
    }

    private JPanel createConfigPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder("合并配置"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.anchor = GridBagConstraints.WEST;

        // 关联列名
        gbc.gridx = 0;
        gbc.gridy = 0;
        panel.add(new JLabel("关联列名:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        joinKeysField = new JTextField(30);
        joinKeysField.setToolTipText("输入用于匹配的列名，多个列名用逗号分隔");
        panel.add(joinKeysField, gbc);

        // 要合并的列名
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.weightx = 0;
        panel.add(new JLabel("表2合并列:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        columnsToMergeField = new JTextField(30);
        columnsToMergeField.setToolTipText("输入要从表2合并到表1的列名，多个列名用逗号分隔");
        panel.add(columnsToMergeField, gbc);

        // 说明标签
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        JLabel hintLabel = new JLabel("<html><i>说明：关联列名是两张表中用于匹配数据的列；表2合并列是要从表2复制到表1的列</i></html>");
        hintLabel.setFont(new Font("微软雅黑", Font.PLAIN, 11));
        panel.add(hintLabel, gbc);

        return panel;
    }

    private void selectFile(int fileNumber) {
        JFileChooser fileChooser = new JFileChooser();
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
            if (fileNumber == 1) {
                selectedFile1 = selectedFile;
                file1Field.setText(selectedFile.getAbsolutePath());
            } else {
                selectedFile2 = selectedFile;
                file2Field.setText(selectedFile.getAbsolutePath());
            }
        }
    }

    private void executeMerge(ActionEvent e) {
        // 验证输入
        if (selectedFile1 == null || selectedFile2 == null) {
            JOptionPane.showMessageDialog(this, "请选择两个 Excel 文件", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        String joinKeys = joinKeysField.getText().trim();
        String columnsToMerge = columnsToMergeField.getText().trim();

        if (joinKeys.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请输入关联列名", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (columnsToMerge.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请输入要合并的列名", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 禁用按钮
        executeButton.setEnabled(false);
        logArea.setText("开始执行合并操作...\n");

        // 在后台线程执行
        new SwingWorker<Boolean, String>() {
            @Override
            protected Boolean doInBackground() throws Exception {
                String[] joinKeyArray = joinKeys.split("[,，]");
                String[] columnsArray = columnsToMerge.split("[,，]");

                // 去除空格
                for (int i = 0; i < joinKeyArray.length; i++) {
                    joinKeyArray[i] = joinKeyArray[i].trim();
                }
                for (int i = 0; i < columnsArray.length; i++) {
                    columnsArray[i] = columnsArray[i].trim();
                }

                publish("正在读取表 1: " + selectedFile1.getName());
                publish("正在读取表 2: " + selectedFile2.getName());
                publish("关联列: " + String.join(", ", joinKeyArray));
                publish("合并列: " + String.join(", ", columnsArray));

                // 生成输出文件路径
                String outputPath = selectedFile1.getParent();
                String outputFileName = "merged_" + selectedFile1.getName();
                File outputFile = new File(outputPath, outputFileName);

                publish("开始执行合并...");
                boolean success = excelService.mergeExcelFiles(
                        selectedFile1,
                        selectedFile2,
                        joinKeyArray,
                        columnsArray,
                        outputFile
                );

                if (success) {
                    publish("合并完成！");
                    publish("输出文件: " + outputFile.getAbsolutePath());
                } else {
                    publish("合并失败，请检查日志");
                }

                return success;
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
                try {
                    if (get()) {
                        JOptionPane.showMessageDialog(ExcelMergePanel.this,
                                "Excel 合并成功！", "成功", JOptionPane.INFORMATION_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(ExcelMergePanel.this,
                                "Excel 合并失败，请检查输入配置", "失败", JOptionPane.ERROR_MESSAGE);
                    }
                } catch (Exception ex) {
                    logArea.append("\n错误: " + ex.getMessage());
                    JOptionPane.showMessageDialog(ExcelMergePanel.this,
                            "执行过程中发生错误: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }
}
