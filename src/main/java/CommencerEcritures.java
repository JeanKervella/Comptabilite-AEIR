import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public class CommencerEcritures implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		String club = (String) SheetsQuickstart.combos.get(0).getSelectedItem();
		String flux = (String) SheetsQuickstart.combos.get(1).getSelectedItem();// TODO voir pour fiches de caisse et
																				// credit
		String modePaie = (String) SheetsQuickstart.combos.get(2).getSelectedItem();
		modePaie = SheetsQuickstart.getRange(modePaie);
		SheetsQuickstart.frame.setTitle("Logiciel de comptabilité AEIR - " + club + " - " + flux);
		SheetsQuickstart.frame.setVisible(true);
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
			System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
		} catch (GeneralSecurityException exception2) {
			System.out.println("merde y'a une erreur mais je sais pas ou - GeneralSecurityException");
		}
		SheetsQuickstart.numeroLigne = 1;
		while (SheetsQuickstart.numeroLigne < SheetsQuickstart.tab.length
				&& !SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][11].equals("")) {
			SheetsQuickstart.numeroLigne++;
		}
		if (SheetsQuickstart.numeroLigne >= SheetsQuickstart.tab.length) {
			JOptionPane.showMessageDialog(null,
					"Il semblerait que cet excel soit vide ou ait deja totalement ete rentre en comptabilite",
					"Erreur : Excel vide", JOptionPane.INFORMATION_MESSAGE);
		} else {
			// confirmation de la fermeture de la fenetre - code trouve sur internet
			//System.out.println(SheetsQuickstart.toStringTab());
			SheetsQuickstart.frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
			SheetsQuickstart.frame.addWindowListener(SheetsQuickstart.fermetureFenetre);// fermeture fenetre - fin
			String[][] tempTab = { SheetsQuickstart.tab[SheetsQuickstart.numeroLigne] };
			SheetsQuickstart.tabAffiche = new JTable(tempTab, SheetsQuickstart.tab[0]);
			SheetsQuickstart.selected = new String[SheetsQuickstart.tab.length][13];
			for (int i=0;i<SheetsQuickstart.selected.length;i++) {
				for (int j = 0;j<SheetsQuickstart.selected[0].length;j++) {
					SheetsQuickstart.selected[i][j] = ".";
				}
			}
			SheetsQuickstart.initLists();;

			// JCOMBOBOX
			// 1er JComboBox
			List<List<Object>> sheetTemp;
			String compare="";
			if (flux.equals("Debit"))
				compare = "A";
			if (flux.equals("Credit"))
				compare = "V";
			JComboBox journalFacture = new JComboBox(); // cree un JComboBox ou bouton qui propose une liste deroulante
			journalFacture.setPreferredSize(new Dimension(100, 20)); // pose la taille
			try {
				sheetTemp = SheetsQuickstart.getData("journaux ventes et achats");
				for (int i = 1; i < sheetTemp.size(); i++) {
					journalFacture.addItem((String) sheetTemp.get(i).get(0));
					System.out.println(i);
					if (club.equals((String) sheetTemp.get(i).get(3)) && compare.equals(sheetTemp.get(i).get(2)))
						journalFacture.setSelectedItem((String) sheetTemp.get(i).get(0));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			}
			AutoCompletion.enable(journalFacture); // permet de chercher dans la liste en appelant la classe
													// AutoCompletion
			// 2eme JComboBox
			JComboBox compteFacture = new JComboBox();
			compteFacture.setPreferredSize(new Dimension(300, 20));
			try {
				sheetTemp = SheetsQuickstart.getData("comptes de charge");
				for (int i = 1; i < sheetTemp.size(); i++) {
					compteFacture.addItem((String) sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			}
			
			AutoCompletion.enable(compteFacture);
			// 3eme JComboBox
			JComboBox compteFournisseur = new JComboBox();
			compteFournisseur.setPreferredSize(new Dimension(300, 20));
			try {
				sheetTemp = SheetsQuickstart.getData("comptes fournisseur");
				for (int i = 1; i < sheetTemp.size(); i++) {
					compteFournisseur.addItem((String) sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			}
			AutoCompletion.enable(compteFournisseur);
			// 4eme JComboBox
			JComboBox codeAnalytique = new JComboBox();
			codeAnalytique.setPreferredSize(new Dimension(150, 20));
			try {
				sheetTemp = SheetsQuickstart.getData("codes analytiques");
				String codeAnal = SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][0];
				for (int i = 1; i < sheetTemp.size(); i++) {
					codeAnalytique.addItem((String) sheetTemp.get(i).get(1));
					if (sheetTemp.get(i).get(2).equals(codeAnal))
						codeAnalytique.setSelectedItem(sheetTemp.get(i).get(1));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			} 

			AutoCompletion.enable(codeAnalytique);
			// 5eme JComboBox
			JComboBox journalBanque = new JComboBox();
			journalBanque.setPreferredSize(new Dimension(100, 20));
			try {
				sheetTemp = SheetsQuickstart.getData("journaux de banques");
				for (int i = 1; i < sheetTemp.size(); i++) {
					journalBanque.addItem((String) sheetTemp.get(i).get(0));
					if (club.equals((String) sheetTemp.get(i).get(3)))
						journalBanque.setSelectedItem((String) sheetTemp.get(i).get(0));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			}
			AutoCompletion.enable(journalBanque);
			// 6eme JComboBox
			JComboBox compteBanque = new JComboBox();
			compteBanque.setPreferredSize(new Dimension(270, 20));
			try {
				sheetTemp = SheetsQuickstart.getData("comptes bancaires");
				for (int i = 1; i < sheetTemp.size(); i++) {
					compteBanque.addItem((String) sheetTemp.get(i).get(2));
					if(club.equals((String) sheetTemp.get(i).get(3))) compteBanque.setSelectedItem(sheetTemp.get(i).get(2));
				}
			} catch (IOException exception) {
				System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
			}
			AutoCompletion.enable(compteBanque);
			// 7eme JComboBox
			JComboBox modePaiement = new JComboBox();
			modePaiement.setPreferredSize(new Dimension(70, 20));
			modePaiement.addItem("VIR");
			modePaiement.addItem("CHQ");
			modePaiement.addItem("CB");
			modePaiement.setSelectedItem(SheetsQuickstart.getModePaiement(modePaie));
			AutoCompletion.enable(modePaiement);
			// rentrer tous ces JComboBox dans combos
			SheetsQuickstart.combos.add(journalFacture);
			SheetsQuickstart.combos.add(compteFacture);
			SheetsQuickstart.combos.add(compteFournisseur);
			SheetsQuickstart.combos.add(codeAnalytique);
			SheetsQuickstart.combos.add(journalBanque);
			SheetsQuickstart.combos.add(compteBanque);
			SheetsQuickstart.combos.add(modePaiement);
			// JCOMBOBOX

			// JTEXTFIELD
			// 1er JTextField
			JTextField dateFacture = new JTextField(); // cree un JTextField, ou un champs a remplir
			dateFacture.setPreferredSize(new Dimension(80, 30)); // pose sa taille
			dateFacture.setForeground(Color.BLUE); // pose la couleur en fond
			dateFacture.setText(SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][7]); // pose la valeur par defaut du
																						// champs
			// 2eme JTextField
			JTextField libelle = new JTextField();
			libelle.setPreferredSize(new Dimension(300, 30));
			libelle.setForeground(Color.BLUE);
			libelle.setText(SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][5]);
			// 3eme JTextField
			JTextField montant = new JTextField();
			montant.setPreferredSize(new Dimension(70, 30));
			montant.setForeground(Color.BLUE);
			montant.setText(SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][3]);
			// 4eme JTextField
			JTextField datePaiement = new JTextField();
			datePaiement.setPreferredSize(new Dimension(80, 30));
			datePaiement.setForeground(Color.BLUE);
			datePaiement.setText(SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][10]);
			// 5eme JTextField
			JTextField numeroCheque = new JTextField();
			numeroCheque.setPreferredSize(new Dimension(60, 30));
			numeroCheque.setForeground(Color.BLUE);
			numeroCheque.setText(SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][2]);
			// rentre tous ces JTextFields dans jtfs
			SheetsQuickstart.jtfs.add(dateFacture);
			SheetsQuickstart.jtfs.add(libelle);
			SheetsQuickstart.jtfs.add(montant);
			SheetsQuickstart.jtfs.add(datePaiement);
			SheetsQuickstart.jtfs.add(numeroCheque); // TODO ajouter que si cheque
			// JTEXTFIELD
			// LABEL
			// creation des labels
			JLabel label1 = new JLabel("code journal facture"); // cree un label ou "titre"
			JLabel label2 = new JLabel("date facture");
			JLabel label3 = new JLabel("libelle");
			JLabel label4 = new JLabel("montant");
			JLabel label5 = new JLabel("compte facture");
			JLabel label6 = new JLabel("compte fournisseur");
			JLabel label7 = new JLabel("code analytique");
			JLabel label8 = new JLabel("code journal banque");
			JLabel label9 = new JLabel("date paiement");
			JLabel label10 = new JLabel("compte banque");
			JLabel label11 = new JLabel("mode de paiement");
			JLabel label12 = new JLabel("numero de cheque");
			// rentre tous les labels dans labels
			SheetsQuickstart.labels.add(label1);
			SheetsQuickstart.labels.add(label2);
			SheetsQuickstart.labels.add(label3);
			SheetsQuickstart.labels.add(label4);
			SheetsQuickstart.labels.add(label5);
			SheetsQuickstart.labels.add(label6);
			SheetsQuickstart.labels.add(label7);
			SheetsQuickstart.labels.add(label8);
			SheetsQuickstart.labels.add(label9);
			SheetsQuickstart.labels.add(label10);
			SheetsQuickstart.labels.add(label11);
			SheetsQuickstart.labels.add(label12);
			// LABEL

			// BUTTON
			JButton bouton = new JButton("Annuler tout");
			JButton bouton2 = new JButton("Enregistrer ce qui a deja ete fait");
			JButton bouton3 = new JButton("Sauter cette ecriture");
			JButton bouton4 = new JButton("Enregistrer et suivante");
			JButton bouton5 = new JButton("Annuler la derniere ecriture");
			bouton4.addActionListener(new EcritureSuivante());
			bouton5.addActionListener(new EcriturePrecedente());
			bouton3.addActionListener(new SauterEcriture());
			
			SheetsQuickstart.frame.getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER,0),"clickButton");

			SheetsQuickstart.frame.getRootPane().getActionMap().put("clickButton",new AbstractAction(){
					        public void actionPerformed(ActionEvent ae)
					        {
					    bouton4.doClick();
					        }
					    });
			
			SheetsQuickstart.boutons.add(bouton);
			SheetsQuickstart.boutons.add(bouton2);
			SheetsQuickstart.boutons.add(bouton3);
			SheetsQuickstart.boutons.add(bouton4);
			SheetsQuickstart.boutons.add(bouton5);
			bouton.setBounds(50,180,150,40);
			bouton2.setBounds(300,180,300,40);
			bouton3.setBounds(1250,180,220,40);
			bouton4.setBounds(1500,180,170,40);
			bouton5.setBounds(1070,180,150,40);
			
			// BUTTON
			
			SheetsQuickstart.tabAffiche.setEnabled(false); // empeche que les champs soient modifiable une fois la
															// fenetre
															// ouverte
			SheetsQuickstart.tabAffiche.setRowHeight(22);
			Font f = new Font("Calibri", Font.PLAIN, 16); // par exemple
			SheetsQuickstart.tabAffiche.setFont(f);
			SheetsQuickstart.tabPanel=new JPanel();
			JScrollPane scrollPane = new JScrollPane(SheetsQuickstart.tabAffiche);
			scrollPane.setBounds(0, 0, 1700, 44);
			SheetsQuickstart.tabPanel.add(scrollPane);
			SheetsQuickstart.tabPanel.setLayout(null);
			SheetsQuickstart.tabPanel.setBounds(0,0,1700,44);

			// cree un JPanel correspondant a chaque bouton
			JPanel first = new JPanel(); // cree un JPanel
			first.add(SheetsQuickstart.labels.get(0)); // ajoute le label correspondant en appelant labels
			first.add(SheetsQuickstart.combos.get(0)); // ajoute le JComboBox correspondant en appelant combos
			first.setBounds(37,50,350,35);
			JPanel second = new JPanel();
			second.add(SheetsQuickstart.labels.get(1));
			second.add(SheetsQuickstart.jtfs.get(0));
			second.setBounds(462,50,350,35);
			JPanel third = new JPanel();
			third.add(SheetsQuickstart.labels.get(2));
			third.add(SheetsQuickstart.jtfs.get(1));
			third.setBounds(887,50,350,35);

			// tout pareil pour b3, b4, b5, b6
			JPanel fourth = new JPanel();
			fourth.add(SheetsQuickstart.labels.get(3));
			fourth.add(SheetsQuickstart.jtfs.get(2));
			fourth.setBounds(1312,50,350,35);
			JPanel fifth = new JPanel();
			fifth.add(SheetsQuickstart.labels.get(4));
			fifth.add(SheetsQuickstart.combos.get(1));
			fifth.setBounds(37,90,450,35);
			JPanel sixth = new JPanel();
			sixth.add(SheetsQuickstart.labels.get(5));
			sixth.add(SheetsQuickstart.combos.get(2));
			sixth.setBounds(462,90,450,35);
			JPanel seventh = new JPanel();
			seventh.add(SheetsQuickstart.labels.get(6));
			seventh.add(SheetsQuickstart.combos.get(3));
			seventh.setBounds(887,90,350,35);
			JPanel eighth = new JPanel();
			eighth.add(SheetsQuickstart.labels.get(7));
			eighth.add(SheetsQuickstart.combos.get(4));
			eighth.setBounds(1312,90,350,35);
			JPanel nineth = new JPanel();
			nineth.add(SheetsQuickstart.labels.get(8));
			nineth.add(SheetsQuickstart.jtfs.get(3));
			nineth.setBounds(37,130,350,35);
			JPanel tenth = new JPanel();
			tenth.add(SheetsQuickstart.labels.get(9));
			tenth.add(SheetsQuickstart.combos.get(5));
			tenth.setBounds(462,130,420,35);
			JPanel eleventh = new JPanel();
			eleventh.add(SheetsQuickstart.labels.get(10));
			eleventh.add(SheetsQuickstart.combos.get(6));
			eleventh.setBounds(887,130,350,35);
			JPanel twelfth = new JPanel();
			twelfth.add(SheetsQuickstart.labels.get(11));
			twelfth.add(SheetsQuickstart.jtfs.get(4)); // TODO ajouter que si cheque
			twelfth.setBounds(1312,130,350,35);

			SheetsQuickstart.container = new JPanel();
			SheetsQuickstart.container.setLayout(null); // formate
																													// le
																													// container
																													// (un
																													// autre
																													// JPanel)
																													// pour
																													// qu'il
																													// liste
																													// ses
																													// JPanels
																													// en
																													// colonne
			SheetsQuickstart.container.add(SheetsQuickstart.tabPanel); // ajoute b1 dans container
			SheetsQuickstart.container.add(first); // ...
			SheetsQuickstart.container.add(second);
			SheetsQuickstart.container.add(third);
			SheetsQuickstart.container.add(fourth);
			SheetsQuickstart.container.add(fifth);
			SheetsQuickstart.container.add(sixth);
			SheetsQuickstart.container.add(seventh);
			SheetsQuickstart.container.add(eighth);
			SheetsQuickstart.container.add(nineth);
			SheetsQuickstart.container.add(tenth);
			SheetsQuickstart.container.add(eleventh);
			SheetsQuickstart.container.add(twelfth);
			SheetsQuickstart.container.add(bouton);
			SheetsQuickstart.container.add(bouton2);
			SheetsQuickstart.container.add(bouton3);
			SheetsQuickstart.container.add(bouton4);
			SheetsQuickstart.container.add(bouton5);
			// AJOUT DE TOUS LES BOUTONS AU CONTAINER

			// SheetsQuickstart.container.setBackground(Color.red);

			SheetsQuickstart.frame.setSize(1700, 300); // formate la taille de la fenetre
			SheetsQuickstart.frame.setContentPane(SheetsQuickstart.container); // indique que contaier est le contenu de
																				// la
																				// fenetre
			System.out.println(SheetsQuickstart.numeroLigne);
		}

	}

}
