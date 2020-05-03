
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

public class ChangeExcelRange implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		System.out.println("change excel id");
		JFrame excelRange = new JFrame();
		ImageIcon img = new ImageIcon("src/main/resources/AEIR logo.png"); // va chercher dans les ressources l'image
		// qui servira d'icone a la fenetre
		excelRange.setIconImage(img.getImage()); // applique cette icone
		excelRange.setBounds(0, 0, 1200, 300);
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("sheetsRange");
		} catch (IOException e1) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e1.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		String[] tabTitre = new String[sheetTemp.get(0).size()];
		for (int i = 0; i < sheetTemp.get(0).size(); i++) {
			tabTitre[i] = (String) sheetTemp.get(0).get(i);
		}
		String[][] tabContent = new String[sheetTemp.size() - 1][sheetTemp.get(0).size()];
		for (int i = 1; i < sheetTemp.size(); i++) {
			for (int j = 0; j < sheetTemp.get(0).size(); j++) {
				tabContent[i - 1][j] = (String) sheetTemp.get(i).get(j);
			}
		}
		JTable jtable = new JTable(tabContent, tabTitre);
		jtable.setRowHeight(50); // parametre la hauteur des cellules a 22 pixels
		Font f = new Font("Calibri", Font.PLAIN, 16); // cree une police
		jtable.setFont(f); // applique la police au tableau
		jtable.setEnabled(true);
		JScrollPane scrollpane = new JScrollPane(jtable);

		excelRange.setContentPane(scrollpane);
		excelRange.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		excelRange.addWindowListener(new WindowAdapter() { // fonction qui permet de confirmer la
			public void windowClosing(WindowEvent we) { // fermeture de la fenetre, utilisee pendant
				String ObjButtons[] = { "Yes", "No", "Annuler" }; // les ecritures mais pas dans le menu
				int PromptResult = JOptionPane.showOptionDialog(null,
						"Voulez-vous enregistrer les changements effectués ?", "Confirmation de fermeture",
						JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons, ObjButtons[1]);
				if (PromptResult == JOptionPane.YES_OPTION) {
					boolean test = true;
					if (!jtable.getValueAt(0, 0).equals("VIR") || !jtable.getValueAt(1, 0).equals("CB")
							|| !jtable.getValueAt(2, 0).equals("CHQ")) {
						test = false;
						JOptionPane PromptResul = new JOptionPane();
						PromptResul.showMessageDialog(null,
								"Les champs dans la premiere colonne ne peuvent pas etre changes.\n"
										+ "Ils doivent etre dans l'ordre : \"VIR\", \"CB\", \"CHQ\"",
								"Erreur code mode de Paiement", JOptionPane.ERROR_MESSAGE);
					}
					boolean test2 = true;
					for (int i = 0; i < 3; i++) {
						String temp = (String) jtable.getValueAt(i, 1);
						for (int j = 0; j < 4; j = j + 3) {
							if (65 > temp.charAt(temp.length() - 1 - j) || temp.charAt(temp.length() - 1 - j) > 90) {
								test2 = false;
							}
						}
						if (temp.charAt(temp.length() - 2) != ':') {
							test2 = false;
						}
						if (temp.charAt(temp.length() - 5) != '!')
							test2 = false;
						if (temp.charAt(temp.length() - 3) < 49 || temp.charAt(temp.length() - 3) > 57)
							test2 = false;
					}
					if (!test2) {
						test = false;
						JOptionPane PromptResul = new JOptionPane();
						PromptResul.showMessageDialog(null, "Au moins une des plages a ete mal rentree(s).\n"
								+ "Elles doivent etre de la forme : \"Virements!B4:N\" avec \"Virements\" le nom de l'onglet dans le sheets, \"B\" la premiere colonne ou il y a des donnees, \"4\"la lign ou il y a les titres avec la ligne suivante qui contient deja les donnees, \"N\" la derniere colonne ou il y a des donnees.\n"
								+ "PS : Le numero de ligne doit etre compris entre 1 et 9 sinon cela occasionnera des problemes dans le code.\n"
								+ "Les colonnes doivent aussi etre mises en majuscule.\n"
								+ "De plus la derniere colonne doit etre une colonne contenant forcement une valeur lorsqu'il y a une ecriture sur la ligne. A la creation de ce logiciel nous utilisions une colonne qui prenait la valeur \"1\" si il y avait des choses ecrites sur la ligne",
								"Erreur format plage de l'excel", JOptionPane.ERROR_MESSAGE);
					}
					test2 = true;
					for (int i = 0; i < jtable.getRowCount(); i++) {
						for (int j = 2; j < jtable.getColumnCount(); j++) {
							if (((String) jtable.getValueAt(i, j)).length() > 1
									|| (int) Character
											.toUpperCase(((String) jtable.getValueAt(i, j)).charAt(0)) > (int) Character
													.toUpperCase(((String) jtable.getValueAt(i, 1))
															.charAt(((String) jtable.getValueAt(i, 1)).length() - 1))
									|| (int) Character
											.toUpperCase(((String) jtable.getValueAt(i, j)).charAt(0)) < (int) Character
													.toUpperCase(((String) jtable.getValueAt(i, 1))
															.charAt(((String) jtable.getValueAt(i, 1)).length() - 4)))
								test2 = false;
						}
						
					}
					if (!test2) {
						test=false;
						JOptionPane PromptResul = new JOptionPane();
						PromptResul.showMessageDialog(null,
								"Au moins un des noms de colonne est mal ecrit.\nLes noms de colonnes doivent etre compris entre les deux colonnes delimitant la plage et ne doivent contenir qu'un caractere.",
								"Erreur nom de colonne", JOptionPane.ERROR_MESSAGE);
					}
					if (test) {
						try {
							System.out.println(tabContent.toString());

							File idSheet = new File("src/main/resources/Data/sheetsRange.txt");
							idSheet.delete();
							idSheet = new File("src/main/resources/Data/sheetsRange.txt");
							BufferedWriter writer = new BufferedWriter(new FileWriter(idSheet));
							writer.write(
									"Feuille;;Plage de donnees;;colonne date Facture;;colonne du libelle;;colonne du montant;;colonne du nom du club;;colonne de la date de Paiement;;colonne du numero de cheque;;colonne de la date de comptabilisation;;\n");
							for (int i = 0; i < jtable.getRowCount(); i++) {
								writer.write(jtable.getValueAt(i, 0) + ";;"+jtable.getValueAt(i, 1) + ";;");
								for (int j = 2; j < jtable.getColumnCount(); j++) {
									writer.write(((String)jtable.getValueAt(i, j)).toUpperCase() + ";;");
								}
								writer.write("\n");
							}
							writer.close();
						} catch (IOException e) {
							JOptionPane jop = new JOptionPane();
							jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
						}

						excelRange.setVisible(false);
						excelRange.dispose();
						
						SheetsQuickstart.updateSheetsRange();
						SheetsQuickstart.updateSheetsColumn();
					}
				} else if (PromptResult == JOptionPane.NO_OPTION) {
					excelRange.setVisible(false);
					excelRange.dispose();
				}
			}
		});
		excelRange.setLocationRelativeTo(null);
		scrollpane.setSize(excelRange.getSize());
		excelRange.setResizable(true);
		excelRange.setVisible(true);

	}

}
