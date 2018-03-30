package View; /**
 * Created by ARIELPE on 6/12/2017.
 */
//Imports are listed in full to show what's being used
//could just import javax.swing.* and java.awt.* etc..

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class GuiApp1 {

	//Note: Typically the main method will be in a
	//separate class. As this is a simple one class
	//example it's all in the one class.
	public static void main(String[] args) {

		new GuiApp1();
	}

	public GuiApp1() {
		JFrame guiFrame = new JFrame();

		//background img
		ImageIcon imageIcon = new ImageIcon("C:\\ship.png");
		//setLayout(new BorderLayout());
		guiFrame.setLayout(new BorderLayout());
		JLabel background = new JLabel(imageIcon);
		guiFrame.add(background, BorderLayout.CENTER);


		//make sure the program exits when the frame closes
		guiFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		guiFrame.setTitle("Smartship Customer Invoice System");
		guiFrame.setSize(new Dimension(525, 445));
		guiFrame.setName("Smartship img");
		guiFrame.setVisible(true);
		//setLocationRelativeTo(null);


		//This will center the JFrame in the middle of the screen
		guiFrame.setLocationRelativeTo(null);

		//Options for the JComboBox
		String[] fruitOptions = {"Apple", "Apricot", "Banana"
				, "Cherry", "Date", "Kiwi", "Orange", "Pear", "Strawberry"};

		//Options for the JList
		String[] vegOptions = {"Asparagus", "Beans", "Broccoli", "Cabbage"
				, "Carrot", "Celery", "Cucumber", "Leek", "Mushroom"
				, "Pepper", "Radish", "Shallot", "Spinach", "Swede"
				, "Turnip"};

		//The first JPanel contains a JLabel and JCombobox
		final JPanel comboPanel = new JPanel();
		JLabel comboLbl = new JLabel("Fruits:");
		JComboBox fruits = new JComboBox(fruitOptions);

		comboPanel.add(comboLbl);
		comboPanel.add(fruits);

		//Create the second JPanel. Add a JLabel and JList and
		//make use the JPanel is not visible.
		final JPanel listPanel = new JPanel();
		listPanel.setVisible(false);
		final JLabel listLbl = new JLabel("Vegetables:");
		JList vegs = new JList(vegOptions);
		vegs.setLayoutOrientation(JList.HORIZONTAL_WRAP);

		listPanel.add(listLbl);
		listPanel.add(vegs);

		JButton vegFruitBut = new JButton("Fruit or Veg");

		//The ActionListener class is used to handle the
		//event that happens when the user clicks the button.
		//As there is not a lot that needs to happen we can
		//define an anonymous inner class to make the code simpler.
		vegFruitBut.addActionListener(new ActionListener() {
			//@Override
			public void actionPerformed(ActionEvent event) {
				//When the fruit of veg button is pressed
				//the setVisible value of the listPanel and
				//comboPanel is switched from true to
				//value or vice versa.
				listPanel.setVisible(!listPanel.isVisible());
				comboPanel.setVisible(!comboPanel.isVisible());

			}
		});

		//when button is clicked run application
		JButton runProgramBut = new JButton("Start Here");
		//runProgramBut.setPreferredSize(new Dimension(100,5));
		runProgramBut.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String[] vegOptions = {"Asparagus", "Beans", "Broccoli", "Cabbage"};
				new TestApp1().main(vegOptions);
				listPanel.setVisible(!listPanel.isVisible());
				comboPanel.setVisible(!comboPanel.isVisible());
			}
		});

		//The JFrame uses the BorderLayout layout manager.
		//Put the two JPanels and JButton in different areas.
		//guiFrame.add(comboPanel, BorderLayout.NORTH);
		//guiFrame.add(listPanel, BorderLayout.CENTER);
		//guiFrame.add(vegFruitBut, BorderLayout.WEST);
		guiFrame.add(runProgramBut, BorderLayout.SOUTH);

		//make sure the JFrame is visible
		guiFrame.setVisible(true);
	}

	class TestApp1 {
		public void main(String[] args) {
			System.out.println("app Controller.TestApp1 running");
			for (String str: args){
				System.out.println("Str: " + str);
			}
		}
	}

}