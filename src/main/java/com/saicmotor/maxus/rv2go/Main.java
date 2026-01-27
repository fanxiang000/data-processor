package com.saicmotor.maxus.rv2go;

import com.saicmotor.maxus.rv2go.ui.MainWindow;

import javax.swing.*;

/**
 * 应用程序入口
 */
public class Main {
    public static void main(String[] args) {
        // 使用事件调度线程启动 GUI
        SwingUtilities.invokeLater(() -> {
            try {
                // 设置系统外观
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }

            MainWindow mainWindow = new MainWindow();
            mainWindow.setVisible(true);
        });
    }
}
