
import java.awt.BorderLayout;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;

public class ChangeIdSheet implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		System.out.println("change excel id");
		JFrame excelRange = new JFrame();
		ImageIcon img = new ImageIcon("src/main/resources/AEIR logo.png"); // va chercher dans les ressources l'image
		// qui servira d'icone a la fenetre
		excelRange.setIconImage(img.getImage()); // applique cette icone
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("idSheets");
		} catch (IOException e1) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e1.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		String[] tabTitre = new String[sheetTemp.get(0).size()];
		for (int i = 0; i < sheetTemp.get(0).size(); i++) {
			tabTitre[i] = (String) sheetTemp.get(0).get(i);
		}
		String[][] tabContent = new String[sheetTemp.size() + 1][sheetTemp.get(0).size()];
		for (int i = 1; i < sheetTemp.size(); i++) {
			for (int j = 0; j < sheetTemp.get(0).size(); j++) {
				tabContent[i - 1][j] = (String) sheetTemp.get(i).get(j);
			}
		}
		JTable jtable = new JTable(tabContent, tabTitre);
		jtable.setRowHeight(22); // parametre la hauteur des cellules a 22 pixels
		Font f = new Font("Calibri", Font.PLAIN, 16); // cree une police
		jtable.setFont(f); // applique la police au tableau
		jtable.setEnabled(true);
		JScrollPane scrollpane = new JScrollPane(jtable);
		JPanel tabpanel = new JPanel();
		tabpanel.setLayout(new BorderLayout());
		tabpanel.add(scrollpane, BorderLayout.CENTER);

		excelRange.setContentPane(tabpanel);
		excelRange.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		excelRange.addWindowListener(new WindowAdapter() { // fonction qui permet de confirmer la
			public void windowClosing(WindowEvent we) { // fermeture de la fenetre, utilisee pendant
				String ObjButtons[] = { "Yes", "No", "Annuler" }; // les ecritures mais pas dans le menu
				int PromptResult = JOptionPane.showOptionDialog(null,
						"Voulez-vous enregistrer les changements effectués ?", "Confirmation de fermeture",
						JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons, ObjButtons[1]);
				if (PromptResult == JOptionPane.YES_OPTION) {
					try {
						System.out.println(tabContent.toString());
						File idSheet = new File("src/main/resources/Data/idSheets.txt");
						idSheet.delete();
						idSheet = new File("src/main/resources/Data/idSheets.txt");
						BufferedWriter writer = new BufferedWriter(new FileWriter(idSheet));
						writer.write("Club;;Id sheet debit;;\n");
						for (int i = 0; i < jtable.getRowCount(); i++) {
							System.out.println(i + "   " + (String) jtable.getValueAt(i, 0)+ (String) jtable.getValueAt(i, 1)+"try");
							if (!(((String) jtable.getValueAt(i, 0)) == null || ((String) jtable.getValueAt(i, 0)).isEmpty()
									|| ((String) jtable.getValueAt(i, 1)) == null
									|| ((String) jtable.getValueAt(i, 1)).isEmpty())) {
								writer.write(jtable.getValueAt(i, 0) + ";;" + jtable.getValueAt(i, 1) + ";;\n");
							}
						}
						writer.close();
					} catch (IOException e) {
						JOptionPane jop = new JOptionPane();
						jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
					}

					excelRange.setVisible(false);
					excelRange.dispose();
				} else if (PromptResult == JOptionPane.NO_OPTION) {
					excelRange.setVisible(false);
					excelRange.dispose();
				}
			}
		});
		excelRange.setLocationRelativeTo(null);
		excelRange.setBounds(0, 0, 1000, 600);
		excelRange.setResizable(false);
		excelRange.setVisible(true);

	}

}
