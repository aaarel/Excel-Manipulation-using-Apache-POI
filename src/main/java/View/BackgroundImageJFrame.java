package View;

import Controller.SmartShipApplication;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

/**
 * Created by Ariel Peretz for Smartship
 * This is the main View for the application
 */

//a Jframe class with specific properties to contain the Frames for the gui
public class BackgroundImageJFrame extends JFrame {
    private MyPanel contentPane;

    //main method of execution (entry point for gui App)
    public static void main(String[] args) {
        Runnable runnable = new Runnable() {
            @Override
            public void run() {
                new BackgroundImageJFrame().displayGui();
            }
        };
        EventQueue.invokeLater(runnable);
    }

    //method to paint the gui to the display using Jframe
    private void displayGui() {
        final JFrame frame = new JFrame("Main Program Window");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setLayout(new FlowLayout());

        final JTextField jTextField = new JTextField("הכנס היטל דלק כאן", 14);

        final JButton button = new JButton("Click here to start");
        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                final String fuel = jTextField.getText();

                jTextField.setText("");
                button.setText("פועל..."); //TODO why doesn't this show up ?
                frame.repaint();
                boolean success = new SmartShipApplication().applicationFlow(fuel);
                try {
                    TimeUnit.SECONDS.sleep(5);
                } catch (InterruptedException interruptedException) {
                    interruptedException.printStackTrace();
                }
                if (success) {
                    button.setText("בוצע");
                } else {
                    button.setText("קרתה תקלה - בדוק קובץ שגיאות (exceptions and logs)");
                }
            }
        });
        contentPane = new MyPanel();
        frame.setContentPane(contentPane);
        frame.setLocationByPlatform(true);
        frame.setVisible(true);
        frame.setTitle("(Smartship) סמארטשיפ מערכת ניהול חשבוניות");
        frame.pack();
        frame.add(jTextField);
        frame.add(button);
    }

    //a panel of an image with specific properties to be used in a Jframe
    private class MyPanel extends JPanel {
        private BufferedImage image;

        public MyPanel() {
            try {
                image = ImageIO.read(BackgroundImageJFrame.class.getResource("/Smartship intro WEB_11.png"));
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
