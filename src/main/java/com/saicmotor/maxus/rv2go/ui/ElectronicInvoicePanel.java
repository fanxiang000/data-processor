package com.saicmotor.maxus.rv2go.ui;

import com.saicmotor.maxus.rv2go.service.ExcelService;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.util.List;
import java.util.prefs.Preferences;

import static java.awt.Desktop.getDesktop;

/**
 * 电子发票功能面板
 * 功能：将出库电子票数据合并到电子票模版
 * - 如果数据不存在则新增
 * - 如果数据存在则更新数量
 */
public class ElectronicInvoicePanel extends JPanel {
    private JTextField templateFileField;
    private JTextField outboundFileField;
    private JTextField templateHeaderRowField;
    private JTextField outboundHeaderRowField;
    private JTextField taxClassificationField;  // 商品和服务税收分类编码
    private JTextArea logArea;
    private JButton executeButton;

    private File selectedTemplateFile;
    private File selectedOutboundFile;

    private final ExcelService excelService;
    private final Preferences prefs;
    private File lastOutputFile;

    public ElectronicInvoicePanel() {
        this.excelService = new ExcelService();
        this.prefs = Preferences.userNodeForPackage(ElectronicInvoicePanel.class);
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

        JLabel titleLabel = new JLabel("电子发票处理工具");
        titleLabel.setFont(new Font("微软雅黑", Font.BOLD, 20));
        titleLabel.setForeground(new Color(44, 62, 80));
        titlePanel.add(titleLabel, BorderLayout.WEST);
        add(titlePanel, BorderLayout.NORTH);

        // 中间内容面板（可滚动）
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
        contentPanel.setBackground(new Color(245, 247, 250));

        // 文件选择区域
        JPanel filePanel = createFileSelectionPanel();
        contentPanel.add(filePanel);
        contentPanel.add(Box.createVerticalStrut(12));

        // 执行按钮
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.setBackground(new Color(245, 247, 250));
        executeButton = new JButton("▶ 执行处理");
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
        executeButton.addActionListener(this::executeProcess);
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

        // 模板文件选择
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel label1 = new JLabel("电子票模版:");
        label1.setFont(labelFont);
        label1.setForeground(labelColor);
        panel.add(label1, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        templateFileField = new JTextField(30);
        templateFileField.setEditable(false);
        templateFileField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        templateFileField.setPreferredSize(new Dimension(0, 32));
        templateFileField.setBackground(new Color(250, 252, 255));
        templateFileField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        panel.add(templateFileField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse1 = createActionButton("浏览...");
        btnBrowse1.addActionListener(e -> selectTemplateFile());
        panel.add(btnBrowse1, gbc);

        // 模板表头行
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.weightx = 0;
        JLabel label2 = new JLabel("模版表头行:");
        label2.setFont(labelFont);
        label2.setForeground(labelColor);
        panel.add(label2, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        templateHeaderRowField = new JTextField("3");
        templateHeaderRowField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        templateHeaderRowField.setPreferredSize(new Dimension(0, 32));
        templateHeaderRowField.setBackground(new Color(250, 252, 255));
        templateHeaderRowField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        templateHeaderRowField.setToolTipText("表头所在的行号（从1开始，模版默认第3行，填3）");
        panel.add(templateHeaderRowField, gbc);

        // 出库文件选择
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.weightx = 0;
        JLabel label3 = new JLabel("出库电子票:");
        label3.setFont(labelFont);
        label3.setForeground(labelColor);
        panel.add(label3, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        outboundFileField = new JTextField(30);
        outboundFileField.setEditable(false);
        outboundFileField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        outboundFileField.setPreferredSize(new Dimension(0, 32));
        outboundFileField.setBackground(new Color(250, 252, 255));
        outboundFileField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        panel.add(outboundFileField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JButton btnBrowse2 = createActionButton("浏览...");
        btnBrowse2.addActionListener(e -> selectOutboundFile());
        panel.add(btnBrowse2, gbc);

        // 出库表头行
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.weightx = 0;
        JLabel label4 = new JLabel("出库表头行:");
        label4.setFont(labelFont);
        label4.setForeground(labelColor);
        panel.add(label4, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        outboundHeaderRowField = new JTextField("10");
        outboundHeaderRowField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        outboundHeaderRowField.setPreferredSize(new Dimension(0, 32));
        outboundHeaderRowField.setBackground(new Color(250, 252, 255));
        outboundHeaderRowField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        outboundHeaderRowField.setToolTipText("表头所在的行号（从1开始，出库默认第10行，填10）");
        panel.add(outboundHeaderRowField, gbc);

        // 商品和服务税收分类编码
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.weightx = 0;
        JLabel label5 = new JLabel("税收分类编码:");
        label5.setFont(labelFont);
        label5.setForeground(labelColor);
        panel.add(label5, gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1.0;
        taxClassificationField = new JTextField();
        taxClassificationField.setFont(new Font("微软雅黑", Font.PLAIN, 13));
        taxClassificationField.setPreferredSize(new Dimension(0, 32));
        taxClassificationField.setBackground(new Color(250, 252, 255));
        taxClassificationField.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(6, 10, 6, 10)
        ));
        taxClassificationField.setToolTipText("填写后，所有新增行的'商品和服务税收分类编码'列都会使用此值");
        panel.add(taxClassificationField, gbc);
        gbc.gridwidth = 1;

        // 说明文字
        gbc.gridx = 0;
        gbc.gridy = 5;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        JTextArea hintArea = new JTextArea();
        hintArea.setText("处理说明：\n" +
                "• 如果出库票的商品在模板中不存在，则新增一行\n" +
                "• 如果出库票的商品在模板中已存在，则累加商品数量\n" +
                "• 数据映射：商品名称→项目名称，数量→商品数量，定价→商品单价\n" +
                "• 默认值：单位=册，税率=0，优惠政策类型=免税，金额=单价×数量");
        hintArea.setEditable(false);
        hintArea.setFont(new Font("微软雅黑", Font.PLAIN, 12));
        hintArea.setBackground(new Color(253, 254, 255));
        hintArea.setForeground(new Color(100, 110, 120));
        hintArea.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        panel.add(hintArea, gbc);
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

    private void selectTemplateFile() {
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

            selectedTemplateFile = selectedFile;
            templateFileField.setText(selectedFile.getAbsolutePath());
        }
    }

    private void selectOutboundFile() {
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

            selectedOutboundFile = selectedFile;
            outboundFileField.setText(selectedFile.getAbsolutePath());
        }
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

    private void executeProcess(ActionEvent e) {
        // 验证输入
        if (selectedTemplateFile == null || selectedOutboundFile == null) {
            JOptionPane.showMessageDialog(this, "请选择电子票模版和出库电子票文件", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 读取表头行（从1开始，转换为从0开始的索引）
        int templateHeaderRow, outboundHeaderRow;
        try {
            templateHeaderRow = Integer.parseInt(templateHeaderRowField.getText().trim()) - 1;
            outboundHeaderRow = Integer.parseInt(outboundHeaderRowField.getText().trim()) - 1;
            if (templateHeaderRow < 0 || outboundHeaderRow < 0) {
                JOptionPane.showMessageDialog(this, "表头行必须是大于0的整数", "错误", JOptionPane.ERROR_MESSAGE);
                return;
            }
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "表头行必须是有效的整数", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 禁用按钮
        executeButton.setEnabled(false);
        logArea.setText("开始执行电子发票处理...\n");

        // 保存为 final 变量供内部类使用
        final File templateFile = selectedTemplateFile;
        final File outboundFile = selectedOutboundFile;
        final int templateHeaderRowFinal = templateHeaderRow;
        final int outboundHeaderRowFinal = outboundHeaderRow;
        final String taxClassification = taxClassificationField.getText().trim();

        // 选择输出文件路径（在后台线程之前）
        String outputFileName = "电子票_已处理_" + System.currentTimeMillis() + ".xlsx";
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
                    publish("正在读取电子票模版: " + templateFile.getName());
                    publish("正在读取出库电子票: " + outboundFile.getName());
                    publish("模版表头行: " + (templateHeaderRowFinal + 1));
                    publish("出库表头行: " + (outboundHeaderRowFinal + 1));
                    if (!taxClassification.isEmpty()) {
                        publish("税收分类编码: " + taxClassification);
                    }

                    publish("输出文件: " + finalOutputFile.getAbsolutePath());
                    publish("开始执行处理...");
                    boolean success = excelService.processElectronicInvoice(
                            templateFile,
                            outboundFile,
                            finalOutputFile,
                            templateHeaderRowFinal,
                            outboundHeaderRowFinal,
                            taxClassification
                    );

                    if (success) {
                        lastOutputFile = finalOutputFile;
                        publish("处理完成！");
                    } else {
                        publish("处理失败，请检查日志");
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
                    JOptionPane.showMessageDialog(ElectronicInvoicePanel.this,
                            "执行失败: " + caughtException.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                try {
                    if (get()) {
                        // 显示成功对话框，提供打开文件夹选项
                        Object[] options = {"确定", "打开文件夹"};
                        int choice = JOptionPane.showOptionDialog(
                                ElectronicInvoicePanel.this,
                                "电子发票处理成功！\n输出文件: " + lastOutputFile.getName(),
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
                        logArea.append("\n=== 处理失败 ===\n");
                        logArea.append("返回值为 false，请检查配置参数\n");
                        JOptionPane.showMessageDialog(ElectronicInvoicePanel.this,
                                "电子发票处理失败，请检查输入配置和日志", "失败", JOptionPane.ERROR_MESSAGE);
                    }
                } catch (Exception ex) {
                    logArea.append("\n=== 系统错误 ===\n");
                    logArea.append("错误: " + ex.getMessage() + "\n");
                    for (StackTraceElement element : ex.getStackTrace()) {
                        logArea.append("    " + element.toString() + "\n");
                    }
                    JOptionPane.showMessageDialog(ElectronicInvoicePanel.this,
                            "执行过程中发生错误: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
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
}
