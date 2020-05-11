import java.awt.Color;
import java.awt.Image;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTable;

public class firstThread extends Thread {
	public firstThread() {
		super();
	}
	
	public firstThread(String name) {
		super(name);
	}
	
	public void run() {
		try {

			SheetsQuickstart.updateSheetsColumn();
			SheetsQuickstart.updateSheetsRange();
			System.out.println(SheetsQuickstart.toStringColumns());
			Color transparent = new Color(255, 255, 255, 160); // cree une couleur transparente qui permettra par la suite
																// de voir
			// le fond d'ecran au travers

			if (SheetsQuickstart.frame.getWindowListeners().length != 0) // si la fermeture de la fenetre demandait une confirmation, on
														// l'enleve
				SheetsQuickstart.frame.removeWindowListener(SheetsQuickstart.fermetureFenetre);
			SheetsQuickstart.frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // quand t'appuies sur la croix en haut a droite ca ferme
																	// la fenetre
			SheetsQuickstart.frame.setLocationRelativeTo(null); // met la fenetre au centre

			SheetsQuickstart.initLists(); // remet toutes les listes de boutons a zero

			ImageIcon img = new ImageIcon("src/main/resources/AEIR logo.png"); // va chercher dans les ressources l'image
																				// qui servira d'icone a la fenetre
			SheetsQuickstart.frame.setIconImage(img.getImage()); // applique cette icone

			String[][] tabInfo = new String[SheetsQuickstart.ecritureRestante.length - 1][SheetsQuickstart.ecritureRestante[0].length];
			for (int i = 0; i < tabInfo.length; i++) {
				tabInfo[i] = SheetsQuickstart.ecritureRestante[i + 1];
			}
			SheetsQuickstart.tabAffiche = new JTable(tabInfo, SheetsQuickstart.ecritureRestante[0]); // remplis tabAffiche avec les titres et
																					// le tableau
			// crees ci-dessus

			SheetsQuickstart.tabAffiche.setRowHeight(22); // parametre la hauteur des cellules a 22 pixels
			SheetsQuickstart.tabAffiche.setBackground(transparent);

			JMenuBar menuBar = new JMenuBar(); // cree une barre des taches
			JMenu donnees = new JMenu("Données Sage"); // cree un menu "donnees"
			JMenu excel = new JMenu("excels"); // cree un menu "excels"
			JButton excelBouton = new JButton("Plages des excels");
			excelBouton.addActionListener(new ChangeExcelRange());
			JButton excelId = new JButton("Id des Excels");
			excelId.addActionListener(new ChangeIdSheet());
			excel.add(excelBouton);
			excel.add(excelId);
			menuBar.add(donnees); // ajoute a la barre des taches un onglet "Donnees"
			SheetsQuickstart.frame.setJMenuBar(menuBar); // ajoute la barre des taches a la fenetre
			menuBar.add(excel); // y ajoute un onglet "excels"

			SheetsQuickstart.container = new JPanel(); // remet le container, qui correspond a tout le contenu de la
														// fenetre, a zero
			SheetsQuickstart.container.setBackground(Color.white); // met le fond du container en blanc
			SheetsQuickstart.container.setLayout(null); // le layout est ce qui gere le placement des composants dans le
														// container
			// on le met sur null pour pouvoir choisir precisement la place de chaque
			// composant
			int tabSize = SheetsQuickstart.tabAffiche.getRowHeight() * SheetsQuickstart.ecritureRestante.length;
			if (tabSize > 650)
				tabSize = 650;
			SheetsQuickstart.tabLayout(1400, tabSize, 22);

			SheetsQuickstart.container.add(SheetsQuickstart.tabPanel); // ajoute le tableau au container
			SheetsQuickstart.addBouton("club", "combo", 120, SheetsQuickstart.tabPanel.getHeight() + 45, 260, 20);
			SheetsQuickstart.addBouton("flux", "combo", 120, SheetsQuickstart.tabPanel.getHeight() + 135, 260, 20);
			SheetsQuickstart.addBouton("modePaiement", "combo", 120, SheetsQuickstart.tabPanel.getHeight() + 225, 260, 20);

			List<List<Object>> sheet = SheetsQuickstart.getData("idSheets"); // recupere les clubs correspondants aux excels
			for (int i = 1; i < sheet.size(); i++) { // ajoute tous les noms de clubs au bouton
				((JComboBox) SheetsQuickstart.components.get(0)).addItem(sheet.get(i).get(0));
			}

			// AutoCompletion.enable(club); // permet de rechercher dans le bouton la valeur
			// qu'on cherche
			((JComboBox) SheetsQuickstart.components.get(1)).addItem("Fiche de caisse - pas encore codé"); // TODO
			((JComboBox) SheetsQuickstart.components.get(1)).addItem("Debit"); // ajoute "debit" au bouton des flux
			((JComboBox) SheetsQuickstart.components.get(1)).addItem("Credit - pas encore codé"); // TODO
			((JComboBox) SheetsQuickstart.components.get(1)).setSelectedItem("Debit");
			((JComboBox) SheetsQuickstart.components.get(2)).addItem("VIR");
			((JComboBox) SheetsQuickstart.components.get(2)).addItem("CHQ");
			((JComboBox) SheetsQuickstart.components.get(2)).addItem("CB");

			SheetsQuickstart.addTitledBorder("Flux choisi", 100, SheetsQuickstart.tabPanel.getHeight() + 110, 300, 60);
			SheetsQuickstart.addTitledBorder("Excel choisi", 100, SheetsQuickstart.tabPanel.getHeight() + 20, 300, 60);
			SheetsQuickstart.addTitledBorder("Mode de Paiement choisi", 100, SheetsQuickstart.tabPanel.getHeight() + 200, 300, 60);

			SheetsQuickstart.addBouton("Commencer les ecritures comptables", "button", 700, SheetsQuickstart.tabPanel.getHeight() + 115, 400,
					50);
			((JButton) SheetsQuickstart.components.get(3)).addActionListener(new CommencerEcritures()); // ajoute une action liee au bouton,
																						// voir "CommencerEcriture"

			SheetsQuickstart.frame.setLayout(null);
			SheetsQuickstart.frame.setResizable(false); // empeche de redimensionner la fenetre
			SheetsQuickstart.frame.setBounds(0, 0, 1400, SheetsQuickstart.tabPanel.getHeight() + 350);
			SheetsQuickstart.container.setBounds(SheetsQuickstart.frame.getX(), SheetsQuickstart.frame.getY(), SheetsQuickstart.frame.getWidth(), SheetsQuickstart.frame.getHeight());

			ImageIcon img2 = new ImageIcon("src/main/resources/feu artifice.jpg"); // va chercher l'image dans les
			// ressources
			Image imgTemp = img2.getImage();
			imgTemp = SheetsQuickstart.getScaledImage(imgTemp, SheetsQuickstart.frame.getWidth(), SheetsQuickstart.frame.getHeight());
			img2 = new ImageIcon(imgTemp);
			JLabel background = new JLabel(img2, JLabel.CENTER);// cree un JPanel pour creer le fond d'ecran
			background.setBounds(0, 0, SheetsQuickstart.frame.getWidth(), SheetsQuickstart.frame.getHeight());
			SheetsQuickstart.container.add(background);

			SheetsQuickstart.frame.setContentPane(SheetsQuickstart.container); // met le container en contenu de la fenetre

			SheetsQuickstart.frame.setVisible(true); // rend l'affichage de la fenetre possible
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		}
	}
}
