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
 */

//TODO main class here

//TODO add text fields for %


public class BackgroundImageJFrame extends JFrame {
    private MyPanel contentPane;

    public BackgroundImageJFrame() {

//
//		setTitle("Smartship (inc.) Model.Customer Invoice System");
//		//setSize(530, 401);
//		setSize(new Dimension(540, 480));
//		setName("Smartship img");
//		setLocationRelativeTo(null);
//		setDefaultCloseOperation(EXIT_ON_CLOSE);
//		setVisible(true);
//
//		//	One way
//		setLayout(new BorderLayout());
//		JLabel background = new JLabel(imageIcon);
//		add(background, BorderLayout.CENTER);
//		background.setLayout(new FlowLayout());
//		//l1 = new JLabel("Here is a button");
//		b1 = new JButton("Start Here");
//		//background.add(l1);
//		background.add(b1, BorderLayout.AFTER_LAST_LINE);

///* Another way
//		setLayout(new BorderLayout());
//		setContentPane(new JLabel(new ImageIcon("C:\\Users\\Computer\\Downloads\\colorful design.png")));
//		setLayout(new FlowLayout());
//		l1=new JLabel("Here is a button");
//		b1=new JButton("I am a button");
//		add(l1);
//		add(b1);
//		// Just for refresh :) Not optional!
//		setSize(399,399);
//		setSize(400,400);
//	}
//
//	*/
    }

    //JLabel l1;
    //ImageIcon imageIcon = new ImageIcon("../ship.png");

    public static void main(String[] args) {
        Runnable runnable = new Runnable() {
            @Override
            public void run() {
                new BackgroundImageJFrame().displayGUI();
            }
        };
        EventQueue.invokeLater(runnable);
    }

    private void displayGUI() {
        JFrame frame = new JFrame("Image Example");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        contentPane = new MyPanel();

        frame.setContentPane(contentPane);
        frame.pack();
        frame.setLocationByPlatform(true);
        frame.setVisible(true);
        final JButton button = new JButton("Start Here");
        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                button.setText("Done");
                new SmartShipApplication().test();
                //new SmartShipApplication().main(new String[]{"", ""});

            }
        });
        frame.add(button);
        frame.setTitle("Smartship (inc.) Model.Customer Invoice System");
        frame.setSize(new Dimension(506, 400));


    }

    private class MyPanel extends JPanel {

        private BufferedImage image;

        public MyPanel() {
            try {
                image = ImageIO.read(BackgroundImageJFrame.class.getResource("../ship.png"));
            } catch (IOException ioe) {
                ioe.printStackTrace();
            }
        }

        @Override
        public Dimension getPreferredSize() {
            return image == null ? new Dimension(400, 300) : new Dimension(image.getWidth(), image.getHeight());
        }

        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            g.drawImage(image, 0, 0, this);
        }
    }
}
