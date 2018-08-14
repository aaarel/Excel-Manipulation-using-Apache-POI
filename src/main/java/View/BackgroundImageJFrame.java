package View;


import Controller.SmartShipApplication;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.IOException;

/**
 * Created by Ariel Peretz for Smartship
 * This is the main View for the application
 */


public class BackgroundImageJFrame extends JFrame {
    private MyPanel contentPane;

    public static void main(String[] args) {
        Runnable runnable = new Runnable() {
            @Override
            public void run() {
                new BackgroundImageJFrame().displayGui();
            }
        };
        EventQueue.invokeLater(runnable);
    }

    private void displayGui() {
        JFrame frame = new JFrame("Main Program Window");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        final JTextField jTextField = new JTextField("Enter Fuel surcharge", 14);
        final JButton button = new JButton("Click here to start");
        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                button.setText("Running...");
                String fuel = jTextField.getText();
                jTextField.setText("");
                new SmartShipApplication().main(new String[]{fuel});
            }
        });
        contentPane = new MyPanel();
        frame.setContentPane(contentPane);
        frame.setLocationByPlatform(true);
        frame.setVisible(true);
        frame.setTitle("Smartship Customer Invoice System");
        frame.pack();
        frame.add(jTextField);
        frame.add(button);
    }

    private class MyPanel extends JPanel {
        private BufferedImage image;

        public MyPanel() {
            try {
                image = ImageIO.read(BackgroundImageJFrame.class.getResource("../Smartship intro WEB_11.png"));
            } catch (IOException ioe) {
                ioe.printStackTrace();
            }
        }

        @Override
        public Dimension getPreferredSize() {
            return image == null ? new Dimension(960, 640) : new Dimension(image.getWidth(), image.getHeight());
        }

        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            g.drawImage(image, 0, 0, this);
        }
    }
}
