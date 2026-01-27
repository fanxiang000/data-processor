package com.saicmotor.maxus.rv2go.ui;

import javax.swing.*;
import java.awt.*;

/**
 * 主窗口 - 左右布局
 * 左侧：功能按钮面板
 * 右侧：主要操作界面
 */
public class MainWindow extends JFrame {
    private JPanel leftPanel;
    private JPanel rightPanel;
    private CardLayout cardLayout;

    public MainWindow() {
        setTitle("数据处理工具");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 700);
        setLocationRelativeTo(null);

        initComponents();
    }

    private void initComponents() {
        // 使用 JSplitPane 分割面板
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        splitPane.setDividerLocation(200);
        splitPane.setResizeWeight(0.2);

        // 左侧面板 - 功能按钮
        leftPanel = createLeftPanel();
        splitPane.setLeftComponent(leftPanel);

        // 右侧面板 - 主要操作区域
        rightPanel = new JPanel();
        cardLayout = new CardLayout();
        rightPanel.setLayout(cardLayout);
        splitPane.setRightComponent(rightPanel);

        // 添加各个功能面板
        addFunctionPanels();

        add(splitPane, BorderLayout.CENTER);
    }

    private JPanel createLeftPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createTitledBorder("功能菜单"));
        panel.setBackground(new Color(240, 240, 245));

        JButton btnExcelMerge = createFunctionButton("Excel 合并", "excelMerge");
        JButton btnDataClean = createFunctionButton("数据清洗", "dataClean");
        JButton btnDataConvert = createFunctionButton("数据转换", "dataConvert");
        JButton btnDataExport = createFunctionButton("数据导出", "dataExport");

        panel.add(Box.createVerticalStrut(10));
        panel.add(btnExcelMerge);
        panel.add(Box.createVerticalStrut(10));
        panel.add(btnDataClean);
        panel.add(Box.createVerticalStrut(10));
        panel.add(btnDataConvert);
        panel.add(Box.createVerticalStrut(10));
        panel.add(btnDataExport);
        panel.add(Box.createVerticalGlue());

        return panel;
    }

    private JButton createFunctionButton(String text, final String action) {
        JButton button = new JButton(text);
        button.setMaximumSize(new Dimension(180, 40));
        button.setAlignmentX(Component.CENTER_ALIGNMENT);
        button.setFont(new Font("微软雅黑", Font.PLAIN, 14));
        button.addActionListener(e -> cardLayout.show(rightPanel, action));
        return button;
    }

    private void addFunctionPanels() {
        // Excel 合并面板
        rightPanel.add(new ExcelMergePanel(), "excelMerge");

        // 其他功能面板占位符
        rightPanel.add(new PlaceholderPanel("数据清洗功能"), "dataClean");
        rightPanel.add(new PlaceholderPanel("数据转换功能"), "dataConvert");
        rightPanel.add(new PlaceholderPanel("数据导出功能"), "dataExport");

        // 默认显示 Excel 合并面板
        cardLayout.show(rightPanel, "excelMerge");
    }

    /**
     * 占位符面板，用于未实现的功能
     */
    private static class PlaceholderPanel extends JPanel {
        public PlaceholderPanel(String text) {
            setLayout(new GridBagLayout());
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 10, 10, 10);

            JLabel label = new JLabel(text, SwingConstants.CENTER);
            label.setFont(new Font("微软雅黑", Font.BOLD, 18));
            add(label, gbc);
        }
    }
}
