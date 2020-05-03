import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.JTextComponent;

import java.awt.Color;
import java.awt.Component;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.HeadlessException;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.awt.image.BufferedImage;
import java.io.*;
import java.security.GeneralSecurityException;
import java.util.*;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

public class SheetsQuickstart {
	// TRUCS COPIES SUR INTERNET QUE JE COMPRENDS PAS TROP
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String TOKENS_DIRECTORY_PATH = "tokens";
	// TRUCS COPIES SUR INTERNET QUE JE COMPRENDS PAS TROP

	// ATTRIBUTS
	public static JFrame frame = new JFrame("Logiciel de Comptabilite AEIR"); // fenetre du logiciel
	public static JPanel container = new JPanel(); // represente le contenu de l'ecran
	public static List<JComponent> components = new ArrayList<JComponent>(); // liste de tous les boutons, textfields...
	public static List<JLabel> labels = new ArrayList<JLabel>(); // liste de tous les noms des boutons
	public static JTable tabAffiche; // tableau "graphique" reprenant les valeurs sur le Drive, utile pour changer la
										// taille des cellule ou changer la police
	public static String[][] tab; // tableau de string reprenant les valeurs du Drive
	public static String[][] selected; // tableau de string contenant les ecritures validees par l'utilisateur, ce
										// tableau est de la meme taille que "tab" et a pour valeur par defaut "." dans
										// toutes ses cases
	public static int numeroLigne = 1; // entier qui correspond a la ligne dans laquelle on est dans "tab"
	public static JPanel tabPanel = new JPanel(); // tableau associe a "tab" et "tabAffiche" qui sera celui reellement
													// affiche a l'ecran
	public static WindowAdapter fermetureFenetre = (new WindowAdapter() { // fonction qui permet de confirmer la
		public void windowClosing(WindowEvent we) { // fermeture de la fenetre, utilisee pendant
			String ObjButtons[] = { "Yes", "No" }; // les ecritures mais pas dans le menu
			int PromptResult = JOptionPane.showOptionDialog(null,
					"Vous êtes sûrs de vouloir quitter?\nTout le travail effectué pourrait être perdu",
					"Confirmation de fermeture", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null,
					ObjButtons, ObjButtons[1]);
			if (PromptResult == JOptionPane.YES_OPTION) {
				System.exit(0);
			}
		}
	});
	public static String[][] sheetsRange; // 1ere ligne pour les VIR, 2eme ligne pour la CB, 3eme ligne pour les CHQ
	public static int[][] sheetsColumn; // 1ere ligne pour les VIR, 2eme ligne pour la CB, 3eme ligne pour les CHQ
	// ATTRIBUTS

	// toString()
	/**
	 * toString qui permet de visualiser "selected" sous la forme d'un string
	 * 
	 * @return "selected" sous une forme de String
	 */
	public static String toStringSelected() {
		String res = "";
		for (int i = 0; i < selected.length; i++) {
			for (int j = 0; j < selected[0].length; j++) {
				res += selected[i][j] + " ";
			}
			res += "\n";
		}
		return res;
	}

	/**
	 * Methode qui permet de redimensionner une image, ici on l'utilisera
	 * essentiellement pour le fond d'ecran
	 * 
	 * @param srcImg image a redimensionner
	 * @param w      largeure voulue
	 * @param h      hauteur voulue
	 * @return l'image redimensionnee
	 */
	public static Image getScaledImage(Image srcImg, int w, int h) {
		BufferedImage resizedImg = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2 = resizedImg.createGraphics();

		g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
		g2.drawImage(srcImg, 0, 0, w, h, null);
		g2.dispose();

		return resizedImg;
	}

	/**
	 * Methode qui permet de creer un bouton, de l'ajouter a la liste de boutons
	 * correspondante et de l'ajouter au container
	 * 
	 * @param name   nom du bouton
	 * @param type   "combo", "jtf", "button", "label", "dynamic", "nothing" sont
	 *               les seuls champs acceptes "nothing" sera utilise pour ne pas
	 *               affiche le textfield "numero de cheque" si l'on est en virement
	 *               ou en cb, ouais c'est la meilleure solution que j'ai trouve et
	 *               alors ?
	 * @param x      l'abscisse du coin en haut a gauche du bouton
	 * @param y      l'ordonnee du coin en haut a gauche du bouton
	 * @param width  la largeur du bouton
	 * @param height la hauteur du bouton
	 */
	public static void addBouton(String name, String type, int x, int y, int width, int height) {
		JComponent temp;
		if (type.equals("combo")) {
			temp = new JComboBox();
		} else if (type.equals("jtf")) {
			temp = new JTextField();
		} else if (type.equals("button")) {
			temp = new JButton(name);
		} else if (type.equals("label")) {
			temp = new JLabel(name);
		} else if (type.equals("dynamic")) {
			temp = new DynamicList();
			height = 137;
		} else if (type.equals("nothing")) {
			temp = new JSeparator(); // au moins jsuis sur de pas utiliser de JSeparator en vrai donc ca posera pas
										// de probleme
			components.add(temp);
			return;
		} else {
			System.err.println("\n\nLe type de bouton indique n'es pas valable\n\n");
			return;
		}
		temp.setBounds(x, y, width, height);
		if (type.equals("label")) {
			labels.add((JLabel) temp);
		} else
			components.add(temp);
		container.add(temp);
	}

	public static boolean selectedCorrect() {
		DynamicList temp = new DynamicList();
		JTextField temp2 = new JTextField();
		JSeparator temp3 = new JSeparator();

		for (int i = 0; i < components.size(); i++) {
			boolean test = false;

			// Verifie que si le champ "numero de cheque" n'est pas affiche le mode de
			// paiemet selectionne n'est pas "CHQ"
			if (components.get(11).getClass().equals(temp3.getClass())
					&& ((JComboBox) components.get(10)).getSelectedItem().equals("CHQ")) {
				JOptionPane PromptResult = new JOptionPane();
				PromptResult.showMessageDialog(null,
						"Vous ne pouvez pas sélectionner \"CHQ\" comme mode de paiement.\nIl faudrait le numero de cheque avec",
						"Erreur mode de Paiement", JOptionPane.ERROR_MESSAGE);
				return false;
			}

			// Verifie que si le composant est une DynamicList le texte selectionne soit
			// bien dans la liste
			if (components.get(i).getClass().equals(temp.getClass())) {
				for (int j = 0; j < ((DynamicList) components.get(i)).getItemCount(); j++) {
					if (((DynamicList) components.get(i)).getItemAt(j)
							.equals(((DynamicList) components.get(i)).getSelectedItem()))
						test = true;
				}
				if (!test) {
					String ObjButtons[] = { "Yes", "No" };
					int PromptResult = JOptionPane.showOptionDialog(null, // TODO ouvrir une autre fenetre pour creer le
																			// nvx compte
							"Il y a au moins un des champs de liste deroulante ou l'item selectionne n'est pas dans la liste.\n Voulez-vous vraiment continuer ?\nLe logiciel Sage creera un nouveau compte/nouveau journal",
							"Confirmation de changement d'ecriture", JOptionPane.DEFAULT_OPTION,
							JOptionPane.WARNING_MESSAGE, null, ObjButtons, ObjButtons[1]);
					if (PromptResult != JOptionPane.YES_OPTION) {
						return false;
					}
				}
			}
			test = false;

			// Verifie tous les JTextField
			if (components.get(i).getClass().equals(temp2.getClass())) {

				// Verifie le format des dates
				if (i == 1 || i == 8) {
					try {
						if (!((JTextField) components.get(i)).getText(10, 1).equals("\n"))
							test = true;
					} catch (BadLocationException e1) {

					}
					try {
						Integer.parseInt(((JTextField) components.get(i)).getText(0, 2));
						Integer.parseInt(((JTextField) components.get(i)).getText(3, 2));
						Integer.parseInt(((JTextField) components.get(i)).getText(6, 4));
						if (!((JTextField) components.get(i)).getText(2, 1).equals("/"))
							test = true;
						if (!((JTextField) components.get(i)).getText(5, 1).equals("/"))
							test = true;
					} catch (BadLocationException e1) {
						test = true;
					} catch (NumberFormatException e1) {
						test = true;
					}
					if (test) {
						JOptionPane PromptResult = new JOptionPane();
						PromptResult.showMessageDialog(null,
								"Au moins une des dates est mal ecrite\nLe format est \"jj/mm/yyyy\"",
								"Erreur format date", JOptionPane.ERROR_MESSAGE);
						return false;
					}
				} else if (((JTextField) components.get(i)).getText().equals("")) {
					JOptionPane PromptResul = new JOptionPane();
					PromptResul.showMessageDialog(null, "Au moins un des champs a remplir est vide",
							"Erreur champs a remplir", JOptionPane.ERROR_MESSAGE);
					return false;
				}
			}
		}
		return true;

	}

	public static void updateSheetsRange() {
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("sheetsRange");
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		sheetsRange = new String[3][2];
		for (int i = 0; i < 3; i++) {
			for (int j = 0; j < 2; j++) {
				sheetsRange[i][j] = (String) sheetTemp.get(i + 1).get(j);
			}
		}
	}

	public static void updateSheetsColumn() {
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("sheetsRange");
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		sheetsColumn = new int[sheetTemp.size() - 1][sheetTemp.get(0).size() - 2];
		char column;
		char firstColumn;
		for (int i = 0; i < sheetTemp.size() - 1; i++) {
			firstColumn = ((String) sheetTemp.get(i + 1).get(1))
					.charAt(((String) sheetTemp.get(i + 1).get(1)).length() - 4);
			for (int j = 0; j < sheetTemp.get(0).size() - 2; j++) {
				column = ((String) sheetTemp.get(i + 1).get(j + 2)).charAt(0);
				sheetsColumn[i][j] = column - firstColumn;
				System.out.println(column +"   "+ firstColumn);
			}
		}
	}
	
	public static String toStringColumns() {
		String res="       ";
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("sheetsRange");
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		System.out.println(sheetTemp.size());
		for(int i = 2;i<sheetTemp.get(0).size();i++) {
			res+=sheetTemp.get(0).get(i)+"  ";
		}
		res+="\n";
		for(int i = 0;i<sheetsColumn.length;i++) {
			if(i==0) res+="VIR :    ";
			if(i==1)res+="CB :     ";
			if(i==2)res+="CHQ :    ";
			for(int j = 0;j<sheetsColumn[0].length;j++) {
				res+= sheetsColumn[i][j]+"                         ";
			}
			res+="\n";
		}
		return res;
	}

	public static void hideAllScrollPane() {
		DynamicList temp = new DynamicList();
		for (int i = 0; i < components.size(); i++) {
			System.out.println(components.get(i).getClass());
			if (components.get(i).getClass().equals(temp.getClass())) {
				((DynamicList) components.get(i)).hideScrollPane();
				System.out.println("DynamicList");
			}
		}
	}

	/**
	 * Methode qui permet d'ajouter un libelle sur la fenetre
	 * 
	 * @param name         nom affiche à l'ecran
	 * @param y            le y du coin en haut a gauche
	 * @param xRightCorner le x qui correspond au x du cote droit du libelle
	 */
	public static void addLibelle(String name, int x, int y, int width, int height) {
		JLabel temp = new JLabel(name);
		temp.setBounds(x, y, width, height);
		labels.add(temp);
		container.add(temp);
	}

	public static void addTitledBorder(String name, int x, int y, int width, int height) {
		Font f = new Font("Calibri", Font.PLAIN | Font.BOLD, 20);
		Color transparent = new Color(255, 255, 255, 160);
		JPanel pan = new JPanel(); // cree un JPanel qui va nous permettre de customiser le bouton club
		pan.setLayout(null);
		pan.setBounds(x, y, width, height);
		pan.setBackground(transparent); // met la couleur "transparent" en fond de fluxPan
		TitledBorder borderTitle = new TitledBorder("Flux choisi"); // cree une bordure avec un titre pour le bouton
		borderTitle.setTitleFont(f); // applique la police de toute a l'heure
		pan.setBorder(borderTitle); // applique la bordure au bouton customise
		container.add(pan);
	}

	/**
	 * toString qui permet de visualiser "tab" sous la forme d'un string
	 * 
	 * @return "tab" sous une forme de String
	 */
	public static String toStringTab() {
		String res = "";
		for (int i = 0; i < tab.length; i++) {
			for (int j = 0; j < tab[0].length; j++) {
				res += tab[i][j] + " ";
			}
			res += "\n\n";
		}
		res += tab.length;
		return res;
	}
	// toString()

	/**
	 * Global instance of the scopes required by this quickstart. If modifying these
	 * scopes, delete your previously saved tokens/ folder.
	 */
	private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
	private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @param HTTP_TRANSPORT The network HTTP Transport.
	 * @return An authorized Credential object.
	 * @throws IOException If the credentials.json file cannot be found.
	 */
	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
		// Load client secrets.
		InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
		if (in == null) {
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES)
						.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
						.setAccessType("offline").build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}

	/**
	 * Permet d'obtenir la plage associee au mode de paiement
	 * 
	 * @param paiement un string de 2 ou 3 lettres majuscules correspond a un des 3
	 *                 modes de paiement : "VIR", "CHQ", "CB"
	 * @return la plage correspondante a prendre dans l'excel
	 */
	public static String getRange(String paiement) {
		for (int i=0;i<sheetsRange.length;i++) {
			if(sheetsRange[i][0].equals(paiement)) return sheetsRange[i][1];
		}
		return null;
	}

	/**
	 * Permet d'obtenir le mode de paiement associe a la plage de l'excel
	 * 
	 * @param range la plage a selectionner dans l'excel
	 * @return le mode de paiement associe, "VIR", "CHQ", "CB"
	 */
	public static String getModePaiement(String range) {
	for (int i=0;i<sheetsRange.length;i++) {
		if(sheetsRange[i][1].equals(range)) return sheetsRange[i][0];
	}
	return null;
	}

	/**
	 * Methode qui permet de remettre toutes les listes de boutons a zero
	 */
	public static void initLists() {
		components = new ArrayList<JComponent>(); // remet la liste des composants a zero
		labels = new ArrayList<JLabel>(); // remet la liste des labels (noms des boutons) a zero
	}

	/**
	 * Methode qui cree la fenetre du menu. Avec un tableau qui recapitule toutes
	 * les ecritures a rentrer en compta Ce tableau n'est pas encore mis en place a
	 * cause de la connexion avec les serveurs de google qui est limitee Cette
	 * fenetre a des bouton permettant de choisir quel excel comptabiliser, si l'on
	 * veut comptabiliser les debits, les credits ou les fiches de caisse Et un
	 * bouton pour commencer
	 * 
	 * @throws IOException
	 * @throws GeneralSecurityException
	 */
	public static void init() throws IOException, GeneralSecurityException {
		updateSheetsColumn();
		updateSheetsRange();
		System.out.println(toStringColumns());
		Color transparent = new Color(255, 255, 255, 160); // cree une couleur transparente qui permettra par la suite
															// de voir
		// le fond d'ecran au travers

		if (frame.getWindowListeners().length != 0) // si la fermeture de la fenetre demandait une confirmation, on
													// l'enleve
			frame.removeWindowListener(fermetureFenetre);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // quand t'appuies sur la croix en haut a droite ca ferme
																// la fenetre
		frame.setLocationRelativeTo(null); // met la fenetre au centre

		SheetsQuickstart.initLists(); // remet toutes les listes de boutons a zero

		ImageIcon img = new ImageIcon("src/main/resources/AEIR logo.png"); // va chercher dans les ressources l'image
																			// qui servira d'icone a la fenetre
		frame.setIconImage(img.getImage()); // applique cette icone

		String[] tabTitle = { "Club", "fiches de caisse", "debit - VIR", "debit - CHQ", "debit - CB", "credit - VIR",
				"credit - CHQ", "credit - CB" }; // tableau des titres des colonnes du tableau dans le menu
		List<List<Object>> sheetTemp = SheetsQuickstart.getData("idSheets"); // va chercher dans les ressources les ID
																				// des excels
		String[][] tabInfo = new String[sheetTemp.size()][8]; // cree un tableau de 8 colonnes et dont le nombre de
																// lignes correspond au nombre d'excels differents

		/*
		 * for (int i = 0; i < sheetTemp.size(); i++) { tabInfo[i][0] = (String)
		 * sheetTemp.get(i).get(0); tabInfo[i][1] = "0"; int j = 0; List<List<Object>>
		 * sT = SheetsQuickstart.getSheet((String) sheetTemp.get(i).get(1),
		 * "Virements!B4:M"); String test1; String test2; for (int k = 1; k < sT.size();
		 * k++) { test1 = (String) sT.get(k).get(9); test2 = (String) sT.get(k).get(10);
		 * if (!test1.equals("") && test2.equals("")) j++; } tabInfo[i][2] =
		 * Integer.toString(j); j = 0; sT = SheetsQuickstart.getSheet((String)
		 * sheetTemp.get(i).get(1), "Cheques!B4:M"); for (int k = 1; k < sT.size(); k++)
		 * { test1 = (String) sT.get(k).get(9); test2 = (String) sT.get(k).get(10); if
		 * (!test1.equals("") && test2.equals("")) j++; } tabInfo[i][3] =
		 * Integer.toString(j); j = 0; sT = SheetsQuickstart.getSheet((String)
		 * sheetTemp.get(i).get(1), "Carte bancaire!B4:M"); for (int k = 1; k <
		 * sT.size(); k++) { test1 = (String) sT.get(k).get(9); test2 = (String)
		 * sT.get(k).get(10); if (!test1.equals("") && test2.equals("")) j++; }
		 * tabInfo[i][4] = Integer.toString(j); tabInfo[i][5] = "0"; tabInfo[i][6] =
		 * "0"; tabInfo[i][7] = "0";
		 * 
		 * }
		 */

		SheetsQuickstart.tabAffiche = new JTable(tabInfo, tabTitle); // remplis tabAffiche avec les titres et le tableau
																		// crees ci-dessus
		tabAffiche.setRowHeight(22); // parametre la hauteur des cellules a 22 pixels
		tabAffiche.setBackground(transparent);

		// ca normalement ca devait mettre les ecritures au centre des cases mais ca
		// marche pas
		((JLabel) tabAffiche.getDefaultRenderer(Number.class)).setHorizontalTextPosition(JLabel.CENTER);

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
		frame.setJMenuBar(menuBar); // ajoute la barre des taches a la fenetre
		menuBar.add(excel); // y ajoute un onglet "excels"

		SheetsQuickstart.container = new JPanel(); // remet le container, qui correspond a tout le contenu de la
													// fenetre, a zero
		SheetsQuickstart.container.setBackground(Color.white); // met le fond du container en blanc
		SheetsQuickstart.container.setLayout(null); // le layout est ce qui gere le placement des composants dans le
													// container
		// on le met sur null pour pouvoir choisir precisement la place de chaque
		// composant

		int size = 22 + tabInfo.length * 22;
		if (size > 650)
			size = 650;
		SheetsQuickstart.tabLayout(1400, size, 22);

		SheetsQuickstart.container.add(tabPanel); // ajoute le tableau au container
		SheetsQuickstart.addBouton("club", "combo", 120, tabPanel.getHeight() + 45, 260, 20);
		SheetsQuickstart.addBouton("flux", "combo", 120, tabPanel.getHeight() + 135, 260, 20);
		SheetsQuickstart.addBouton("modePaiement", "combo", 120, tabPanel.getHeight() + 225, 260, 20);

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

		SheetsQuickstart.addTitledBorder("Flux choisi", 100, tabPanel.getHeight() + 110, 300, 60);
		SheetsQuickstart.addTitledBorder("Excel choisi", 100, tabPanel.getHeight() + 20, 300, 60);
		SheetsQuickstart.addTitledBorder("Mode de Paiement choisi", 100, tabPanel.getHeight() + 200, 300, 60);

		SheetsQuickstart.addBouton("Commencer les ecritures comptables", "button", 700, tabPanel.getHeight() + 115, 400,
				50);
		((JButton) components.get(3)).addActionListener(new CommencerEcritures()); // ajoute une action liee au bouton,
																					// voir "CommencerEcriture"

		frame.setLayout(null);
		frame.setResizable(false); // empeche de redimensionner la fenetre
		frame.setBounds(0, 0, 1400, tabPanel.getHeight() + 350);
		container.setBounds(frame.getX(), frame.getY(), frame.getWidth(), frame.getHeight());

		ImageIcon img2 = new ImageIcon("src/main/resources/feu artifice.jpg"); // va chercher l'image dans les
		// ressources
		Image imgTemp = img2.getImage();
		imgTemp = SheetsQuickstart.getScaledImage(imgTemp, frame.getWidth(), frame.getHeight());
		img2 = new ImageIcon(imgTemp);
		JLabel background = new JLabel(img2, JLabel.CENTER);// cree un JPanel pour creer le fond d'ecran
		background.setBounds(0, 0, frame.getWidth(), frame.getHeight());
		container.add(background);

		frame.setContentPane(container); // met le container en contenu de la fenetre

		frame.setVisible(true); // rend l'affichage de la fenetre possible

	}

	/**
	 * Met en forme le tableau apres que tabAffiche ait ete mis a jour
	 */
	/**
	 * @param widht  largeur du tableau
	 * @param height hauteur du tableau
	 */
	public static void tabLayout(int widht, int height, int rowHeight) {
		Color transparent = new Color(255, 255, 255, 160);
		tabAffiche.setEnabled(false); // empeche que les champs soient modifiable une fois la
		// fenetre
		// ouverte
		tabAffiche.setRowHeight(rowHeight); // parametre la hauteur des cellules a 22 pixels
		Font f = new Font("Calibri", Font.PLAIN, 16); // cree une police
		tabAffiche.setFont(f); // applique la police au tableau
		tabPanel = new JPanel(); // initialise le tabPanel
		JScrollPane scrollPane = new JScrollPane(SheetsQuickstart.tabAffiche);
		scrollPane.setBackground(transparent);
		scrollPane.setBounds(0, 0, widht, height);
		tabPanel.setLayout(null);
		tabPanel.setBounds(0, 0, widht, height);

		tabPanel.add(scrollPane);
	}

	/**
	 * Une fois les ecritures commencees Permet de changer le tableau et les champs
	 * des boutons lorsque l'on change d'ecriture
	 */
	public static void updateEcriture() {
		for (int i = 0; i<container.getComponentCount();i++) {
			if (container.getComponent(i).equals(tabPanel)) container.remove(i);
		}
		String[][] tempTab = { SheetsQuickstart.tab[SheetsQuickstart.numeroLigne] }; // cree le nouveau tableau
		tabAffiche = new JTable(tempTab, SheetsQuickstart.tab[0]);

		SheetsQuickstart.tabLayout(1800, 66, 44);

		container.add(SheetsQuickstart.tabPanel);// remet le tableau dans la fenetre une fois qu'il est bien modifie

		int y =0;
		JTextField temp = new JTextField();
		if(((JComboBox) components.get(10)).getSelectedItem().equals("CB")) y=1;
		if(((JComboBox) components.get(10)).getSelectedItem().equals("CHQ")) y=2;
		((JTextField) components.get(1)).setText(tempTab[0][sheetsColumn[y][0]]); // tempTab[0][0]
		((JTextField) components.get(2)).setText(tempTab[0][sheetsColumn[y][1]]);
		((JTextField) components.get(3)).setText(tempTab[0][sheetsColumn[y][2]]);
		((JTextField) components.get(8)).setText(tempTab[0][sheetsColumn[y][4]]);
		if (components.get(11).getClass().equals(temp.getClass()))
			((JTextField) components.get(11)).setText(tempTab[0][sheetsColumn[y][5]]);
		try {
			List<List<Object>> sheetTemp = SheetsQuickstart.getData("codes analytiques");
			String codeAnal = tempTab[0][sheetsColumn[y][3]];
			for (int i = 1; i < sheetTemp.size(); i++) {
				if (sheetTemp.get(i).get(2).equals(codeAnal)) {
				}
				((DynamicList) SheetsQuickstart.components.get(6)).setSelectedItem((String) sheetTemp.get(i).get(1));
			}
		} catch (IOException exception) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
	}

	/**
	 * Fonction qui permet d'aller chercher dans les ressources les donnees
	 * associees a celles de Sage
	 * 
	 * @param data nom du fichier texte correspondant
	 * @return les donnees sous la forme d'une double liste dont la premier liste
	 *         est les titres des colonnes
	 * @throws IOException
	 */
	public static List<List<Object>> getData(String data) throws IOException {
		List<List<Object>> res = new ArrayList();
		BufferedReader reader = new BufferedReader(
				new FileReader(new File("src/main/resources/Data/" + data + ".txt"))); // va chercher le fichier
																						// concerne
		String read = reader.readLine(); // lis le fichier ligne par ligne
		int j = 0;
		while (read != null) { // read vaut null quand il n'y a plus de ligne a lire
			j = 0;
			List<Object> temp = new ArrayList();
			String[] splitted = read.split(";;"); // separe le texte avec ";;"
			read = reader.readLine(); // read lit la ligne suivante
			while (j < splitted.length) { // parcours splitted pour ajouter au resultat tous les champs
				if (splitted[j] != "") {
					temp.add(splitted[j]);
				} else
					temp.add(" ");
				j++;
			}
			res.add(temp);
		}
		return res;

	}

	/**
	 * Verifie si un string est dans une liste ou non
	 * 
	 * @param listToCheck
	 * @param test        string a tester dans la liste
	 * @return true si le test est bien dans la liste, faux si non
	 */
	public static boolean isInList(List<String> listToCheck, String test) {
		boolean res = false;
		for (int i = 0; i < listToCheck.size(); i++) {
			if (((String) listToCheck.get(i)).equals(test))
				res = true;

		}
		return res;
	}

	/**
	 * Methode qui cree le fichier texte une fois les ecritures terminees Ce fichier
	 * texte pourra ensuite etre importe sur sage
	 * 
	 * @throws IOException
	 */
	public static void createText() throws IOException {
		List<List<String>> compare = new ArrayList(); // des qu'un fournisseur, un journal aura ete selectionne il est
														// ajoute dans cette list de list pour qu'il n'y est pas de
														// doublon
		compare.add(new ArrayList()); // correspond aux tiers deja marques
		compare.add(new ArrayList()); // correspond aux journaux ventes et achats deja marques
		compare.add(new ArrayList()); // correspond aux journaux de banques deja marques
		compare.add(new ArrayList()); // correspond aux mode de paiements deja marques

		List<List<Object>> sageData = SheetsQuickstart.getData("codes analytiques");
		File mouvement = new File("src/main/resources/textes/mouvement.txt");
		File tiers = new File("src/main/resources/textes/tiers.txt");
		File journal = new File("src/main/resources/textes/journal.txt");
		File modePaiement = new File("src/main/resources/textes/modePaiement.txt");
		BufferedWriter writer = new BufferedWriter(new FileWriter(mouvement));
		writer.write("##Transfert\n" + "##Section	Dos\n" + "EUR\n" + "##Section	Mvt\n");
		int compteur = 1;
		for (int i = 0; i < selected.length; i++) {
			if (selected[i][0].equals("true")) {
				String codeAnal = ".";
				String[] compteFournisseur = { ".", "." };
				int j = 0;
				sageData = SheetsQuickstart.getData("codes analytiques");
				while (codeAnal.equals(".")) {
					if (sageData.get(j).get(1).equals(selected[i][7])) {
						codeAnal = (String) sageData.get(j).get(0);
					}
					j++;
				}
				j = 0;
				sageData = SheetsQuickstart.getData("comptes fournisseur");
				while (compteFournisseur[0].equals(".")) {
					if (sageData.get(j).get(2).equals(selected[i][6])) {
						compteFournisseur[0] = (String) sageData.get(j).get(0);
						compteFournisseur[1] = (String) sageData.get(j).get(1);
					}
					j++;
				}
				writer.write("\"" + compteur + "\"	\"" + selected[i][1] + "\"	\"" + selected[i][2] + "\"	\""
						+ selected[i][5].substring(0, 7) + "\"	\"" + selected[i][5].substring(11) + "\"	\""
						+ selected[i][4] + "\"	D	B	\"" + selected[i][3] + "\"		\"10\"			\"" + codeAnal
						+ "\"	\"" + selected[i][7] + "\"\r\n");
				writer.write("\"" + compteur + "\"	\"" + selected[i][1] + "\"	\"" + selected[i][2] + "\"	\""
						+ compteFournisseur[0] + "\"	\"" + compteFournisseur[1] + "\"	\"" + selected[i][4]
						+ "\"	C	B	\"" + selected[i][3] + "\"		\"10\"		\"" + selected[i][2] + "\"\r\n");
				compteur++;
				writer.write("\"" + compteur + "\"	\"" + selected[i][8] + "\"	\"" + selected[i][9] + "\"	\""
						+ compteFournisseur[0] + "\"	\"" + compteFournisseur[1] + "\"	\"" + selected[i][4]
						+ "\"	D	B	\"" + selected[i][3] + "\"	");
				if (selected[i][11].equals("CHQ"))
					writer.write("\"" + selected[i][12] + "\"");
				writer.write("	\"6\"	\"" + selected[i][11] + "\"	\"" + selected[i][9] + "\"	\"" + codeAnal
						+ "\"	\"" + selected[i][7] + "\"\n");
				writer.write("\"" + compteur + "\"	\"" + selected[i][8] + "\"	\"" + selected[i][9] + "\"	\""
						+ selected[i][10].substring(0, 7) + "\"	\"" + selected[i][10].substring(11) + "\"	\""
						+ selected[i][4] + "\"	C	B	\"" + selected[i][3] + "\"		\"6\"	\"" + selected[i][11]
						+ "\"	\"" + selected[i][9] + "\"\n");
			}
			compteur++;
		}
		writer.close();
		writer = new BufferedWriter(new FileWriter(tiers));
		writer.write("##Section	Tiers\r\n");
		sageData = SheetsQuickstart.getData("comptes fournisseur");
		int cursor = 0; // curseur qui permettra de se deplacer dans les donnees de sage
		for (int i = 0; i < selected.length; i++) {
			cursor = 0;
			if (selected[i][0].equals("true") && !SheetsQuickstart.isInList(compare.get(0), selected[i][6])) {
				while (!selected[i][6].equals(sageData.get(cursor).get(2)) && cursor != sageData.size()) {
					cursor++;
				}
				compare.get(0).add(selected[i][6]);
			}

		}
		Collections.sort(compare.get(0));
		for (int i = 0; i < compare.get(0).size(); i++) {
			cursor = 0;
			while (!compare.get(0).get(i).equals(sageData.get(cursor).get(2)) && cursor != sageData.size()) {
				cursor++;
			}
			writer.write("\"" + sageData.get(cursor).get(0) + "\"	\"" + sageData.get(cursor).get(1) + "\"	\"SR\"\n");
		}
		writer.close();
		writer = new BufferedWriter(new FileWriter(journal));
		writer.write("##Section	Jnl\r\n");
		sageData = SheetsQuickstart.getData("journaux ventes et achats");
		for (int i = 0; i < selected.length; i++) {
			cursor = 0;
			if (selected[i][0].equals("true") && !SheetsQuickstart.isInList(compare.get(1), selected[i][1])) {
				while (!selected[i][1].equals(sageData.get(cursor).get(0)) && cursor != sageData.size()) {
					System.out.println(selected[i][1] + "   " + sageData.get(cursor).get(0));
					cursor++;
				}
				writer.write("\"" + sageData.get(cursor).get(0) + "\"	\"" + sageData.get(cursor).get(1) + "\"	\""
						+ sageData.get(cursor).get(2) + "\"\n");
				compare.get(1).add(selected[i][1]);
			}

		}
		sageData = SheetsQuickstart.getData("journaux de banques");
		for (int i = 0; i < selected.length; i++) {
			cursor = 0;
			if (selected[i][0].equals("true") && !SheetsQuickstart.isInList(compare.get(2), selected[i][8])) {
				while (!selected[i][8].equals(sageData.get(cursor).get(0)) && cursor != sageData.size()) {
					cursor++;
				}
				compare.get(2).add(selected[i][8]);
			}
		}
		Collections.sort(compare.get(2));
		for (int i = 0; i < compare.get(2).size(); i++) {
			cursor = 0;
			while (!compare.get(2).get(i).equals(sageData.get(cursor).get(0)) && cursor != sageData.size()) {
				cursor++;
			}
			writer.write("\"" + sageData.get(cursor).get(0) + "\"	\"" + sageData.get(cursor).get(1) + "\"	\""
					+ sageData.get(cursor).get(2) + "\"\n");
		}
		writer.close();
		writer = new BufferedWriter(new FileWriter(modePaiement));
		writer.write("##Section	MdP\n");
		for (int i = 0; i < selected.length; i++) {
			if (selected[i][0].equals("true") && !SheetsQuickstart.isInList(compare.get(3), selected[i][11])) {
				compare.get(3).add(selected[i][11]);
			}

		}
		Collections.sort(compare.get(3));
		for (int i = 0; i < compare.get(3).size(); i++) {
			if (compare.get(3).get(i).equals("VIR")) {
				writer.write("\"VIR\"	\"Virement\"	\"8\"	\"M\"\n");
			}
			if (compare.get(3).get(i).equals("CB")) {
				writer.write("\"CB\"	\"Carte Bleue\"	\"2\"	\n");
			}
			if (compare.get(3).get(i).equals("CHQ")) {
				writer.write("\"CHQ\"	\"Chèque à réception\"	\"3\"	\n");
			}
		}

		writer.close();
		try {

			String fileName;
			boolean create = false;
			JFileChooser chooser = new JFileChooser();
			while (!create) {

				// Dossier Courant
				chooser.setCurrentDirectory(new File("C:\\Users\\Utilisateur\\Desktop"));
				FileFilter filter = new FileNameExtensionFilter("Fichiers texte", new String[] { "txt" });
				chooser.setFileFilter(filter);
				chooser.setSelectedFile(new File("imports Sage.txt"));

				// Affichage et récupération de la réponse de l'utilisateur
				int reponse = chooser.showDialog(chooser, "Enregistrer sous");

				// Si l'utilisateur clique sur OK
				if (reponse == JFileChooser.APPROVE_OPTION) {
					File f = chooser.getSelectedFile();
					if (f.exists()) {
						int result = JOptionPane.showConfirmDialog(chooser,
								"Le fichier existe deja, voulez-vous le remplacer?", "Fichier deja existant",
								JOptionPane.YES_NO_CANCEL_OPTION);
						if (result == JOptionPane.YES_OPTION)
							create = true;
					} else
						create = true;
				}
			}

			fileName = chooser.getSelectedFile().toString();
			File result = new File(fileName);
			result.delete();
			result = new File(fileName);
			SheetsQuickstart.copyFile(mouvement, result);
			SheetsQuickstart.copyFile(tiers, result);
			SheetsQuickstart.copyFile(journal, result);
			SheetsQuickstart.copyFile(modePaiement, result);

			String ObjButtons[] = { "Yes", "No" }; // TODO gerer un mauvais import de Sage
			int PromptResult = JOptionPane.showOptionDialog(null, "L'import des données s'est bien fait sur Sage ?",
					"Import avec succes", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons,
					ObjButtons[1]);
			if (PromptResult == JOptionPane.YES_OPTION) {
				String listTiers = "";
				for (int i = 0; i < compare.get(0).size(); i++) {
					listTiers += compare.get(0).get(i) + "\n";
				}

				JOptionPane PromptResult1 = new JOptionPane();
				PromptResult1.showMessageDialog(null,
						"N'oubliez pas de faire le lettrage des ecritures !\nVoici la liste des comptes a lettrer :\n\n"
								+ listTiers,
						"Lettrage sur Sage", JOptionPane.INFORMATION_MESSAGE);
			}

			mouvement.delete();
			tiers.delete();
			journal.delete();
			modePaiement.delete();

			SheetsQuickstart.init();

		} catch (HeadlessException he) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, he.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (GeneralSecurityException he) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, he.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}

	}

	/**
	 * Code trouve sur internet qui permet de copier coller un fichier dans un autre
	 * fichier
	 * 
	 * @param copy  le fichier a copier
	 * @param paste le fichier dans lequel coller
	 */
	public static void copyFile(File copy, File paste) {
		// Nous déclarons nos objets en dehors du bloc try/catch
		FileInputStream fis = null;
		FileOutputStream fos = null;

		try {
			// On instancie nos objets :
			// fis va lire le fichier
			// fos va écrire dans le nouveau !
			fis = new FileInputStream(copy);
			fos = new FileOutputStream(paste, true);

			// On crée un tableau de byte pour indiquer le nombre de bytes lus à
			// chaque tour de boucle
			byte[] buf = new byte[1];

			// On crée une variable de type int pour y affecter le résultat de
			// la lecture
			// Vaut -1 quand c'est fini
			int n = 0;

			// Tant que l'affectation dans la variable est possible, on boucle
			// Lorsque la lecture du fichier est terminée l'affectation n'est
			// plus possible !
			// On sort donc de la boucle
			while ((n = fis.read(buf)) >= 0) {
				// On écrit dans notre deuxième fichier avec l'objet adéquat
				fos.write(buf);
				// On affiche ce qu'a lu notre boucle au format byte et au
				// format char
				// Nous réinitialisons le buffer à vide
				// au cas où les derniers byte lus ne soient pas un multiple de 8
				// Ceci permet d'avoir un buffer vierge à chaque lecture et ne pas avoir de
				// doublon en fin de fichier
				buf = new byte[1];

			}

		} catch (FileNotFoundException e) {
			// Cette exception est levée si l'objet FileInputStream ne trouve
			// aucun fichier
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			// Celle-ci se produit lors d'une erreur d'écriture ou de lecture
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} finally {
			// On ferme nos flux de données dans un bloc finally pour s'assurer
			// que ces instructions seront exécutées dans tous les cas même si
			// une exception est levée !
			try {
				if (fis != null)
					fis.close();
			} catch (IOException e) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}

			try {
				if (fos != null)
					fos.close();
			} catch (IOException e) {
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	/**
	 * Methode qui permet d'etablir la connexion avec google et qui retourne une
	 * liste de liste d'object correspond a l'excel et la plage desires
	 * 
	 * @param sheetId l'id de la sheet desiree
	 * @param range   la plage desiree
	 * @return une double liste reprenant toutes les valeurs de la plage demandee
	 * @throws IOException
	 * @throws GeneralSecurityException
	 */
	public static List<List<Object>> getSheet(String sheetId, String range)
			throws IOException, GeneralSecurityException {
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		Sheets service = new Sheets.Builder(HTTP_TRANSPORT, SheetsQuickstart.JSON_FACTORY,
				getCredentials(HTTP_TRANSPORT)).setApplicationName(APPLICATION_NAME).build();
		ValueRange response = service.spreadsheets().values().get(sheetId, range).execute();
		return response.getValues();
	}

	public static void main(String... args) throws IOException, GeneralSecurityException {
		SheetsQuickstart.init();
	}
}