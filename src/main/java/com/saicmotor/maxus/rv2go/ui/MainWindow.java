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
    private JButton selectedButton;  // 当前选中的按钮

    public MainWindow() {
        setTitle("数据处理工具");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1400, 900);
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(1200, 800));

        initComponents();
    }

    private void initComponents() {
        // 使用 JSplitPane 分割面板
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        splitPane.setDividerLocation(220);
        splitPane.setResizeWeight(0.0);  // 左侧固定宽度，右侧可调整

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

        // 延迟设置初始激活按钮状态
        SwingUtilities.invokeLater(() -> {
            for (java.awt.Component comp : leftPanel.getComponents()) {
                if (comp instanceof JButton) {
                    JButton btn = (JButton) comp;
                    if ("Excel 合并".equals(btn.getText())) {
                        selectedButton = btn;
                        btn.setBackground(new Color(52, 152, 219));
                        btn.setForeground(new Color(255, 255, 255));
                        btn.setBorder(BorderFactory.createCompoundBorder(
                            BorderFactory.createMatteBorder(1, 3, 1, 1, new Color(52, 152, 219)),
                            BorderFactory.createEmptyBorder(10, 15, 10, 15)
                        ));
                        break;
                    }
                }
            }
        });
    }

    private JPanel createLeftPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));
        panel.setBackground(new Color(250, 252, 255));

        // 标题区域
        JPanel titleWrapper = new JPanel();
        titleWrapper.setLayout(new BoxLayout(titleWrapper, BoxLayout.Y_AXIS));
        titleWrapper.setBackground(new Color(250, 252, 255));
        titleWrapper.setAlignmentX(Component.LEFT_ALIGNMENT);
        titleWrapper.setMaximumSize(new Dimension(190, 80));

        JLabel titleLabel = new JLabel("数据处理工具");
        titleLabel.setFont(new Font("微软雅黑", Font.BOLD, 18));
        titleLabel.setForeground(new Color(44, 62, 80));
        titleLabel.setAlignmentX(Component.LEFT_ALIGNMENT);

        JLabel subtitleLabel = new JLabel("Data Processor");
        subtitleLabel.setFont(new Font("微软雅黑", Font.PLAIN, 11));
        subtitleLabel.setForeground(new Color(127, 140, 141));
        subtitleLabel.setAlignmentX(Component.LEFT_ALIGNMENT);

        titleWrapper.add(titleLabel);
        titleWrapper.add(Box.createVerticalStrut(5));
        titleWrapper.add(subtitleLabel);
        titleWrapper.add(Box.createVerticalStrut(15));

        panel.add(titleWrapper);
        panel.add(Box.createVerticalStrut(10));

        // 功能按钮
        JButton btnExcelMerge = createFunctionButton("Excel 合并", "excelMerge", true);
        JButton btnDataClean = createFunctionButton("数据清洗", "dataClean", false);
        JButton btnDataConvert = createFunctionButton("数据转换", "dataConvert", false);
        JButton btnDataExport = createFunctionButton("数据导出", "dataExport", false);

        panel.add(btnExcelMerge);
        panel.add(Box.createVerticalStrut(8));
        panel.add(btnDataClean);
        panel.add(Box.createVerticalStrut(8));
        panel.add(btnDataConvert);
        panel.add(Box.createVerticalStrut(8));
        panel.add(btnDataExport);
        panel.add(Box.createVerticalGlue());

        return panel;
    }

    private JButton createFunctionButton(String text, final String action, boolean isFirst) {
        JButton button = new JButton(text);
        button.setMaximumSize(new Dimension(190, 45));
        button.setMinimumSize(new Dimension(190, 45));
        button.setPreferredSize(new Dimension(190, 45));
        button.setAlignmentX(Component.LEFT_ALIGNMENT);
        button.setFont(new Font("微软雅黑", Font.PLAIN, 14));

        // 设置按钮样式
        button.setBackground(new Color(255, 255, 255));
        button.setForeground(new Color(52, 73, 94));
        button.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
            BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));
        button.setFocusPainted(false);
        button.setContentAreaFilled(false);
        button.setOpaque(true);
        button.setHorizontalAlignment(SwingConstants.LEFT);

        // 鼠标悬停效果
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) {
                // 如果不是当前选中的按钮，才应用悬停效果
                if (button != selectedButton) {
                    button.setBackground(new Color(236, 240, 247));
                    button.setBorder(BorderFactory.createCompoundBorder(
                        BorderFactory.createMatteBorder(1, 3, 1, 1, new Color(52, 152, 219)),
                        BorderFactory.createEmptyBorder(10, 15, 10, 15)
                    ));
                }
            }
            public void mouseExited(java.awt.event.MouseEvent e) {
                // 如果不是当前选中的按钮，才重置样式
                if (button.isEnabled() && button != selectedButton) {
                    button.setBackground(new Color(255, 255, 255));
                    button.setBorder(BorderFactory.createCompoundBorder(
                        BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
                        BorderFactory.createEmptyBorder(10, 15, 10, 15)
                    ));
                }
            }
        });

        button.addActionListener(e -> {
            // 重置所有按钮样式
            resetAllButtons();
            // 设置当前按钮为激活状态
            selectedButton = button;
            button.setBackground(new Color(52, 152, 219));
            button.setForeground(new Color(255, 255, 255));
            button.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(1, 3, 1, 1, new Color(52, 152, 219)),
                BorderFactory.createEmptyBorder(10, 15, 10, 15)
            ));
            cardLayout.show(rightPanel, action);
        });

        return button;
    }

    private void resetAllButtons() {
        // 遍历左侧面板的所有组件，重置按钮样式
        for (java.awt.Component comp : leftPanel.getComponents()) {
            if (comp instanceof JButton) {
                JButton btn = (JButton) comp;
                btn.setBackground(new Color(255, 255, 255));
                btn.setForeground(new Color(52, 73, 94));
                btn.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(220, 225, 230), 1),
                    BorderFactory.createEmptyBorder(10, 15, 10, 15)
                ));
            }
        }
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
