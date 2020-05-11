import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.util.Collections;
import java.util.List;
import java.util.TimerTask;

import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.DefaultListModel;
import javax.swing.FocusManager;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.ListModel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.BadLocationException;

public class DynamicList extends JComponent implements ActionListener, FocusListener, KeyListener {

	private String[] defaultValues = { "a", "abcdefgaaa", "ab", "abc", "abcd", "abcde", "abcdef", "absdefg", "abcdefgh",
			"abcdefghi", "abcdefghij", "abcdefghijk", "abcdefghijkl", "abcdefghijklm", };
	private DefaultListModel<String> listeAffiche = new DefaultListModel<String>();
	private DefaultListModel<String> listeModel = new DefaultListModel<String>();
	private JList jList = createJList();
	private JScrollPane scrollPane;
	private JTextField textField = createTextField();
	private JButton button;

	public DynamicList() {
		super();
		super.setLayout(new BorderLayout());

		this.button = new JButton(new ImageIcon("src/main/resources/Arrow.jpg"));
		button.setSize(20, 20);
		button.addActionListener(this);
		scrollPane = new JScrollPane(jList);
		textField.addMouseListener(new MouseListener() {

			@Override
			public void mouseClicked(MouseEvent e) {
				scrollPane.setVisible(true);
			}

			public void mousePressed(MouseEvent e) {
			}

			public void mouseReleased(MouseEvent e) {
			}

			public void mouseEntered(MouseEvent e) {
				scrollPane.setVisible(true);
			}

			public void mouseExited(MouseEvent e) {
				try {
					if (!FocusManager.getCurrentManager().getFocusOwner().equals(textField))
						scrollPane.setVisible(false);
				} catch (NullPointerException e1) {
					scrollPane.setVisible(false);
					System.out.println(e1.getMessage());
				}
			}
		});
		textField.addFocusListener(this);
		scrollPane.addFocusListener(this);
		textField.addKeyListener(this);
		jList.addMouseListener(new MouseListener() {

			private int eventCnt = 0;
			java.util.Timer timer = new java.util.Timer("doubleClickTimer", false);

			@Override
			public void mouseClicked(final MouseEvent e) {
				eventCnt = e.getClickCount();
				if (e.getClickCount() == 1) {
					timer.schedule(new TimerTask() {
						@Override
						public void run() {
							if (eventCnt == 1) {
								System.err.println("You did a single click.");
							} else if (eventCnt > 1) {
								System.err.println("You did a double click.");
								textField.setText((String) jList.getSelectedValue());
								scrollPane.setVisible(false);
							}
							eventCnt = 0;
						}
					}, 200);
				}
			}

			public void mousePressed(MouseEvent e) {
			}

			public void mouseReleased(MouseEvent e) {
			}

			public void mouseEntered(MouseEvent e) {
			}

			public void mouseExited(MouseEvent e) {
			}

		});
		super.add(scrollPane, BorderLayout.CENTER);
		JPanel temp = new JPanel();
		temp.setLayout(new BoxLayout(temp, BoxLayout.LINE_AXIS));
		temp.add(textField);
		temp.add(button);
		super.add(temp, BorderLayout.NORTH);
	}

	public int getItemCount() {
		return this.listeModel.getSize();
	}

	public String getItemAt(int index) {
		return listeModel.get(index);
	}

	public void addDynamicList(JPanel panel) {
		panel.add(new JScrollPane(jList));
		panel.add(createTextField(), BorderLayout.PAGE_END);
	}

	private JTextField createTextField() {
		final JTextField field = new JTextField(15);
		field.getDocument().addDocumentListener(new DocumentListener() {
			@Override
			public void insertUpdate(DocumentEvent e) {
				filter();
			}

			@Override
			public void removeUpdate(DocumentEvent e) {
				filter();
			}

			@Override
			public void changedUpdate(DocumentEvent e) {
			}

			private void filter() {
				String filter = field.getText();
				filterModel((DefaultListModel<String>) jList.getModel(), filter);
				int i = 0;
				boolean test = false;
				while (i < jList.getModel().getSize() && !test) {
					test = ((String) jList.getModel().getElementAt(i)).toUpperCase().contains(filter.toUpperCase());
					jList.setSelectedIndex(i);
				}
			}
		});
		return field;
	}

	private JList createJList() {
		JList list = new JList(listeAffiche);
		list.setVisibleRowCount(6);
		return list;
	}
	public void initAffiche() {
		for(int i =0;i<listeModel.size();i++) {
			if(!listeAffiche.contains(listeModel.get(i))) {
				listeAffiche.addElement(listeModel.get(i));
			}
		}
	}
	
	
	public void filterModel(DefaultListModel<String> model, String filter) {
		for (int i = 0; i < listeModel.getSize(); i++) {
			if (!listeModel.get(i).toUpperCase().contains(filter.toUpperCase())) {
				if (model.contains(listeModel.get(i))) {
					model.removeElement(listeModel.get(i));
				}
			} else {
				if (!model.contains(listeModel.get(i))) {
					model.addElement(listeModel.get(i));
				}
				if(listeModel.get(i).toUpperCase().equals(filter.toUpperCase())) {
					initAffiche();
					jList.ensureIndexIsVisible(jList.getSelectedIndex());
				}
			}
		}
	}

	public String getSelectedItem() {
		return textField.getText();
	}

	public boolean isInList(String test) {
		for (int i = 0; i < jList.getModel().getSize(); i++) {
			if (jList.getModel().getElementAt(i).equals(test))
				return true;
		}
		return false;
	}

	public void setSelectedItem(String item) {
		initAffiche();
		System.out.println("In selected Item");
		for (int i = 0; i < listeAffiche.getSize(); i++) {
			System.out.println(((String)listeAffiche.getElementAt(i)));
			System.out.println(item);
			if (((String)listeAffiche.getElementAt(i)).equals(item)) {
				System.out.println("Item Found");
				jList.setSelectedIndex(i);
				textField.setText((String) jList.getSelectedValue());
				return;
			}
		}
	}

	public void addItem(String item) {

		listeAffiche.addElement(item);
		listeModel.addElement(item);
		if (listeModel.getSize() == 1)
			textField.setText(listeAffiche.get(0));
	}

	public void removeItem(String item) {
		listeAffiche.removeElement(item);
		listeModel.removeElement(item);
	}

	public void hideScrollPane() {
		this.scrollPane.setVisible(false);
	}

	public static void main(String[] args) {
		DynamicList test = new DynamicList();
		test.setBounds(10, 10, 150, 27);
		JFrame frame = new JFrame();
		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.add(Box.createRigidArea(new Dimension(10, 10)));
		panel.add(test);
		panel.add(Box.createRigidArea(new Dimension(10, 10)));
		DynamicList test2 = new DynamicList();
		test2.setBounds(10, 100, 200, 137);
		test2.addItem("a");
		test2.addItem("ab");
		test2.addItem("abc");
		test2.addItem("gef");
		test2.addItem("abcd");
		test2.addItem("abcde");
		test2.addItem("abcdef");
		test2.addItem("abcdefg");
		test2.addItem("abc");
		test2.addItem("abc");
		test2.addItem("abc");
		test2.removeItem("gef");
		panel.add(test2);
		frame.setContentPane(panel);
		frame.setSize(500, 500);
		frame.setLocationRelativeTo(null);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setVisible(true);
		test2.hideScrollPane();

		String doub = "8,99";
		String doub2 = "60,00";
		String doub3 = "24.99";
		System.out.println(doub3);
		System.out.println(doub3.replace(".",","));
	}

	public void actionPerformed(ActionEvent e) {
		if (!this.scrollPane.isVisible())
			scrollPane.setVisible(true);
		else
			this.scrollPane.setVisible(false);
	}

	@Override
	public void focusGained(FocusEvent e) {
		scrollPane.setVisible(true);
	}

	@Override
	public void focusLost(FocusEvent e) {
		scrollPane.setVisible(false);

	}

	@Override
	public void keyTyped(KeyEvent e) {
	}

	public void keyPressed(KeyEvent e) {
		if (e.getKeyCode() == 38 && jList.getSelectedIndex() > 0) {
			jList.setSelectedIndex(jList.getSelectedIndex() - 1);
			jList.ensureIndexIsVisible(jList.getSelectedIndex());

		}
		if (e.getKeyCode() == 40 && jList.getSelectedIndex() < jList.getModel().getSize() - 1) {
			jList.setSelectedIndex(jList.getSelectedIndex() + 1);
			jList.ensureIndexIsVisible(jList.getSelectedIndex());
		}
		if (e.getKeyCode() == 10) {
			textField.setText((String) jList.getSelectedValue());
			System.out.println("10)");
			scrollPane.setVisible(false);
		}
		if (e.getKeyCode() == 27) {
			textField.transferFocus();
			scrollPane.setVisible(false);
		}
		System.out.println(e.getKeyCode());
	}

	public void keyReleased(KeyEvent e) {
	}
}