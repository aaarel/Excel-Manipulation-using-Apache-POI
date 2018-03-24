package View;

import javax.swing.*;
import java.awt.*;

/**
 * Created by ARIELPE on 6/12/2017.
 */
public class BackgroundImageJFrame extends JFrame {
	JButton b1;
	//JLabel l1;
	ImageIcon imageIcon = new ImageIcon("C:\\ship.png");

	public BackgroundImageJFrame() {
		setTitle("Smartship (inc.) Model.Customer Invoice System");
		//setSize(530, 401);
		setSize(new Dimension(540, 480));
		setName("Smartship img");
		setLocationRelativeTo(null);
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		setVisible(true);

		//	One way
		setLayout(new BorderLayout());
		JLabel background = new JLabel(imageIcon);
		add(background, BorderLayout.CENTER);
		background.setLayout(new FlowLayout());
		//l1 = new JLabel("Here is a button");
		b1 = new JButton("Start Here");
		//background.add(l1);
		background.add(b1, BorderLayout.AFTER_LAST_LINE);

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

	public static void main(String args[]) {
		new BackgroundImageJFrame();
	}

}
