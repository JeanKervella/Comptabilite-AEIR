import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public class CommencerEcritures implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		String club = (String) ((JComboBox) SheetsQuickstart.components.get(0)).getSelectedItem();
		String flux = (String) ((JComboBox) SheetsQuickstart.components.get(1)).getSelectedItem();// TODO voir pour
																									// fiches de caisse
																									// et
		// credit
		String modePaie = (String) ((JComboBox) SheetsQuickstart.components.get(2)).getSelectedItem();
		modePaie = SheetsQuickstart.getRange(modePaie);
		SheetsQuickstart.frame.setTitle("Logiciel de comptabilité AEIR - " + club + " - " + flux);
		SheetsQuickstart.frame.setVisible(true);
		int moPaie = 0;
		if (modePaie.equals("VIR"))
			moPaie = 0;
		if (modePaie.equals("CB"))
			moPaie = 1;
		if (modePaie.equals("CHQ"))
			moPaie = 2;
		try {
			List<List<Object>> sheet = SheetsQuickstart.getData("idSheets");

			String idSheetWanted = null;
			String temp = null;
			for (int i = 0; i < sheet.size(); i++) {
				temp = (String) sheet.get(i).get(0);
				if (temp.equals(club))
					idSheetWanted = (String) sheet.get(i).get(1);
			}

			sheet = SheetsQuickstart.getSheet(idSheetWanted, modePaie);
			SheetsQuickstart.tab = new String[sheet.size()][sheet.get(0).size()];

			for (int i = 0; i < sheet.size(); i++) {
				for (int j = 0; j < sheet.get(0).size(); j++) {
					if (!sheet.get(i).get(j).equals(null))
						SheetsQuickstart.tab[i][j] = (String) sheet.get(i).get(j);
				}
			}
		} catch (IOException exception) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (GeneralSecurityException exception2) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, exception2.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		SheetsQuickstart.numeroLigne = 1;
		while (SheetsQuickstart.numeroLigne < SheetsQuickstart.tab.length
				&& !SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][6]]
						.equals("")
				&& SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][4]].equals("")) {
			SheetsQuickstart.numeroLigne++;
		}
		if (SheetsQuickstart.numeroLigne >= SheetsQuickstart.tab.length) {
			JOptionPane.showMessageDialog(null,
					"Il semblerait que cet excel soit vide ou ait deja totalement ete rentre en comptabilite",
					"Erreur : Excel vide", JOptionPane.INFORMATION_MESSAGE);
		} else {
			
			SheetsQuickstart.frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
			SheetsQuickstart.frame.addWindowListener(SheetsQuickstart.fermetureFenetre);
			String[][] tempTab = { SheetsQuickstart.tab[SheetsQuickstart.numeroLigne] };
			SheetsQuickstart.tabAffiche = new JTable(tempTab, SheetsQuickstart.tab[0]);
			SheetsQuickstart.tabLayout(1800, 66, 44);
			SheetsQuickstart.selected = new String[SheetsQuickstart.tab.length][13];
			for (int i = 0; i < SheetsQuickstart.selected.length; i++) {
				for (int j = 0; j < SheetsQuickstart.selected[0].length; j++) {
					SheetsQuickstart.selected[i][j] = ".";
				}
			}
			SheetsQuickstart.initLists();
			SheetsQuickstart.container = new JPanel();
			SheetsQuickstart.container.setLayout(null);

			SheetsQuickstart.frame.setSize(1800, 500); // formate la taille de la fenetre
			SheetsQuickstart.frame.setContentPane(SheetsQuickstart.container); // indique que contaier est le contenu de
																				// la
																				// fenetre

			SheetsQuickstart.container.add(SheetsQuickstart.tabPanel);
			
			// JCOMBOBOX
			// 1er JComboBox
			List<List<Object>> sheetTemp;
			String compare = "";
			if (flux.equals("Debit"))
				compare = "A";
			if (flux.equals("Credit"))
				compare = "V";

			SheetsQuickstart.addBouton("journalFacture", "dynamic", 200, 55, 100, 20); // .get(0)
			SheetsQuickstart.addLibelle("journal Facture", 110, 55, 100, 20);
			try {
				sheetTemp = SheetsQuickstart.getData("journaux ventes et achats");
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(0)).addItem((String) sheetTemp.get(i).get(0));
					if (club.equals((String) sheetTemp.get(i).get(3)) && compare.equals(sheetTemp.get(i).get(2)))
						((DynamicList) SheetsQuickstart.components.get(0))
								.setSelectedItem((String) sheetTemp.get(i).get(0));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}

			SheetsQuickstart.addBouton("date facture", "jtf", 600, 50, 70, 30); // .get(1)
			SheetsQuickstart.addLibelle("date Facture", 520, 55, 150, 20);
			SheetsQuickstart.addBouton("libelle", "jtf", 1000, 50, 200, 30); // get(2)
			SheetsQuickstart.addLibelle("libelle", 960, 55, 150, 20);
			SheetsQuickstart.addBouton("montant", "jtf", 1400, 50, 70, 30); // get(3)
			SheetsQuickstart.addLibelle("Montant", 1350, 55, 150, 20);

			SheetsQuickstart.addBouton("compte Facture", "dynamic", 200, 105, 280, 20); // get(4)
			SheetsQuickstart.addLibelle("compte Facture", 105, 105, 180, 20);
			SheetsQuickstart.addBouton("compte Fournisseur", "dynamic", 600, 105, 280, 20); // get(5)
			SheetsQuickstart.addLibelle("compte Fournisseur", 480, 105, 180, 20);
			SheetsQuickstart.addBouton("code analytique", "dynamic", 1000, 105, 230, 27); // get(6)
			SheetsQuickstart.addLibelle("code analytique", 905, 105, 150, 20);
			SheetsQuickstart.addBouton("code journal banque", "dynamic", 1400, 105, 100, 27); // get(7)
			SheetsQuickstart.addLibelle("code journal banque", 1280, 105, 150, 20);

			SheetsQuickstart.addBouton("date Paiement", "jtf", 200, 150, 70, 30); // get(8)
			SheetsQuickstart.addLibelle("date Paiement", 110, 155, 150, 20);
			SheetsQuickstart.addBouton("compte Banque", "dynamic", 600, 155, 220, 20); // get(9)
			SheetsQuickstart.addLibelle("compte Banque", 500, 155, 150, 20);
			SheetsQuickstart.addBouton("mode paiement", "combo", 1000, 155, 60, 27); // get(10)
			SheetsQuickstart.addLibelle("mode Paiement", 905, 155, 150, 20);
			if (SheetsQuickstart.getModePaiement(modePaie).equals("CHQ")) {
				SheetsQuickstart.addBouton("numero de cheque", "jtf", 1400, 150, 70, 30); // get(11)
				SheetsQuickstart.addLibelle("numero de cheque", 1285, 150, 150, 20);
			} else
				SheetsQuickstart.addBouton("nothing", "nothing", 0, 0, 0, 0); // get(11)

			SheetsQuickstart.addBouton("Annuler tout", "button", 20, 200, 200, 50); // get(12)
			SheetsQuickstart.addBouton("Enregistrer ce qui a deja ete fait", "button", 240, 200, 230, 50); // get(13)
			SheetsQuickstart.addBouton("Annuler la derniere ecriture", "button", 920, 200, 200, 50); // get(14)
			SheetsQuickstart.addBouton("Sauter cette ecriture", "button", 1140, 200, 200, 50); // get(15)
			SheetsQuickstart.addBouton("Enregistrer et suivante", "button", 1360, 200, 200, 50); // get(16)
			SheetsQuickstart.components.get(16).requestFocus();

			((JButton) SheetsQuickstart.components.get(14)).addActionListener(new EcriturePrecedente());
			((JButton) SheetsQuickstart.components.get(15)).addActionListener(new SauterEcriture());
			((JButton) SheetsQuickstart.components.get(16)).addActionListener(new EcritureSuivante());

			try {
				sheetTemp = SheetsQuickstart.getData("comptes de charge");
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(4)).addItem((String) sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}

			// 3eme JComboBox
			try {
				sheetTemp = SheetsQuickstart.getData("comptes fournisseur");
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(5)).addItem((String) sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}
			// 4eme JComboBox
			try {
				sheetTemp = SheetsQuickstart.getData("codes analytiques");
				String codeAnal = SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][3]];
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(6)).addItem((String) sheetTemp.get(i).get(1));
					if (sheetTemp.get(i).get(2).equals(codeAnal))
						((DynamicList) SheetsQuickstart.components.get(6))
								.setSelectedItem((String) sheetTemp.get(i).get(1));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}

			// 5eme JComboBox
			try {
				sheetTemp = SheetsQuickstart.getData("journaux de banques");
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(7)).addItem((String) sheetTemp.get(i).get(0));
					if (club.equals((String) sheetTemp.get(i).get(3)))
						((DynamicList) SheetsQuickstart.components.get(7))
								.setSelectedItem((String) sheetTemp.get(i).get(0));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}
			// 6eme JComboBox
			try {
				sheetTemp = SheetsQuickstart.getData("comptes bancaires");
				for (int i = 1; i < sheetTemp.size(); i++) {
					((DynamicList) SheetsQuickstart.components.get(9)).addItem((String) sheetTemp.get(i).get(2));
					if (club.equals((String) sheetTemp.get(i).get(3)))
						((DynamicList) SheetsQuickstart.components.get(9))
								.setSelectedItem((String) sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}

			// 7eme JComboBox
			((JComboBox) SheetsQuickstart.components.get(10)).addItem("VIR");
			((JComboBox) SheetsQuickstart.components.get(10)).addItem("CHQ");
			((JComboBox) SheetsQuickstart.components.get(10)).addItem("CB");
			((JComboBox) SheetsQuickstart.components.get(10))
					.setSelectedItem(SheetsQuickstart.getModePaiement(modePaie));
			// JCOMBOBOX

			// JTEXTFIELD
			// 1er JTextField
			SheetsQuickstart.components.get(1).setForeground(Color.BLUE); // pose la couleur en fond

			// 2eme JTextField
			SheetsQuickstart.components.get(2).setForeground(Color.BLUE);

			SheetsQuickstart.components.get(3).setForeground(Color.BLUE);

			((JTextField) SheetsQuickstart.components.get(8)).setForeground(Color.BLUE);

			if (SheetsQuickstart.components.get(11).getClass().equals("class javax.swing.JTextField")) {
				((JTextField) SheetsQuickstart.components.get(11)).setForeground(Color.BLUE);
			}

			SheetsQuickstart.updateEcriture();
			// JTEXTFIELD


			// SheetsQuickstart.container.setBackground(Color.red);

			
			System.out.println(SheetsQuickstart.numeroLigne);
			SheetsQuickstart.hideAllScrollPane();

		}

	}

}
