import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Color;
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
	public static List<JTextField> jtfs = new ArrayList<JTextField>(); // liste de tous les boutons ou l'on peut ecrire
																		// du texte
	public static List<JComboBox> combos = new ArrayList<JComboBox>(); // liste de tous ceux ou on a une liste
																		// deroulante
	public static List<JLabel> labels = new ArrayList<JLabel>(); // liste de tous les noms des boutons
	public static List<JButton> boutons = new ArrayList<JButton>(); // liste de tous les boutons classiques
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
	 * @param type   "combo", "jtf", "button" sont les seuls champs acceptes
	 * @param x      l'abscisse du coin en haut a gauche du bouton
	 * @param y      l'ordonnee du coin en haut a gauche du bouton
	 * @param width  la largeur du bouton
	 * @param height la hauteur du bouton
	 */
	public static void addBouton(String name, String type, int x, int y, int width, int height) {
		if (type.equals("combo")) { // TODO use this method
			JComboBox temp = new JComboBox();
			temp.setBounds(x, y, width, height);
			combos.add(temp);
			container.add(temp);
		} else if (type.equals("jtf")) {
			JTextField temp = new JTextField();
			temp.setBounds(x, y, width, height);
			JLabel temp2 = new JLabel(name);
			temp.add(temp2);
			labels.add(temp2);
			jtfs.add(temp);
			container.add(temp);
		} else if (type.equals("button")) {
			JButton temp = new JButton(name);
			temp.setBounds(x, y, width, height);
			boutons.add(temp);
			container.add(temp);
		}
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
		if (paiement.equals("VIR"))
			return "Virements!B4:N";
		if (paiement.equals("CHQ"))
			return "Cheques!B4:N";
		if (paiement.equals("CB"))
			return "Carte bancaire!B4:N";
		return null;
	}

	/**
	 * Permet d'obtenir le mode de paiement associe a la plage de l'excel
	 * 
	 * @param range la plage a selectionner dans l'excel
	 * @return le mode de paiement associe, "VIR", "CHQ", "CB"
	 */
	public static String getModePaiement(String range) {
		if (range.equals("Virements!B4:N"))
			return "VIR";
		if (range.equals("Cheques!B4:N"))
			return "CHQ";
		if (range.equals("Carte bancaire!B4:N"))
			return "CB";
		return null;
	}

	/**
	 * Methode qui permet de remettre toutes les listes de boutons a zero
	 */
	public static void initLists() {
		jtfs = new ArrayList<JTextField>(); // remet la liste des bouton ou l'on peut remplir du texte a zero
		combos = new ArrayList<JComboBox>(); // remet la liste des boutons a liste deroulante a zero
		labels = new ArrayList<JLabel>(); // remet la liste des labels (noms des boutons) a zero
		boutons = new ArrayList<JButton>();
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
		Color transparent = new Color(255, 255, 255, 130); // cree une couleur transparente qui permettra par la suite
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
		JMenu donnees = new JMenu("Données"); // cree un menu "donnees"
		JMenu excel = new JMenu("excels"); // cree un menu "excels"
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
		if(size>650) size=650;
		SheetsQuickstart.tabLayout(1400, size); 

		SheetsQuickstart.container.add(tabPanel); // ajoute le tableau au container

		JComboBox club = new JComboBox(); // cree les trois boutons qui permettent de choisir quoi comptabiliser
		JComboBox flux = new JComboBox();
		JComboBox modePaiement = new JComboBox();

		JLabel clubLab = new JLabel("club");
		club.add(clubLab);
		club.setBounds(100, tabPanel.getHeight() + 50, 120, 20); // parametre les tailles et positions des boutons
		flux.setBounds(100, tabPanel.getHeight() + 150, 120, 20);
		modePaiement.setBounds(100, tabPanel.getHeight() + 250, 100, 20);
		List<List<Object>> sheet = SheetsQuickstart.getData("idSheets"); // recupere les clubs correspondants aux excels
		for (int i = 1; i < sheet.size(); i++) { // ajoute tous les noms de clubs au bouton
			club.addItem(sheet.get(i).get(0));
		}
		AutoCompletion.enable(club); // permet de rechercher dans le bouton la valeur qu'on cherche
		flux.addItem("Fiche de caisse - pas encore codé"); // TODO
		flux.addItem("Debit"); // ajoute "debit" au bouton des flux
		flux.addItem("Credit - pas encore codé"); // TODO
		flux.setSelectedItem("Debit");
		modePaiement.addItem("VIR");
		modePaiement.addItem("CHQ");
		modePaiement.addItem("CB");

		JPanel clubPan = new JPanel(); // cree un JPanel qui va nous permettre de customiser le bouton club
		clubPan.setLayout(null);
		clubPan.setBounds(100, tabPanel.getHeight() + 50, 300, 60);
		clubPan.setBackground(transparent); // met la couleur "transparent" en fond de clubPan
		TitledBorder borderTitle = new TitledBorder("Excel choisi"); // cree une bordure avec un titre pour le bouton
		Font f = new Font("Calibri", Font.PLAIN, 16);// cree une police
		borderTitle.setTitleFont(f); // applique la police de toute a l'heure

		clubPan.setBorder(borderTitle); // applique la bordure au bouton customise
		clubPan.add(club); // ajoute le vrai bouton au bouton customise
		combos.add(club); // ajoute les boutons dans les listes de boutons pour qu'on puisse y avoir acces
							// depuis une autre classe
		combos.add(flux);
		combos.get(1).addItem("test");
		combos.add(modePaiement);
		container.add(clubPan); // ajoute les boutons au container
		container.add(flux);
		combos.get(1).addItem("test 2");
		container.add(modePaiement);

		JButton b2 = new JButton("Commencer les ecritures comptables"); // cree un bouton classique pour commencer
		b2.setBounds(400, tabPanel.getHeight() + 200, 400, 50);
		b2.addActionListener(new CommencerEcritures()); // ajoute une action liee au bouton, voir "CommencerEcriture"
		container.add(b2); // ajoute le bouton au container

		frame.setLayout(null);
		frame.setResizable(false); // empeche de redimensionner la fenetre
		frame.setBounds(0, 0, 1400, tabPanel.getHeight()+350);

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
	public static void tabLayout(int widht, int height) {
		Color transparent = new Color(255, 255, 255, 160);
		tabAffiche.setEnabled(false); // empeche que les champs soient modifiable une fois la
		// fenetre
		// ouverte
		tabAffiche.setRowHeight(22); // parametre la hauteur des cellules a 22 pixels
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

		SheetsQuickstart.container.remove(SheetsQuickstart.tabPanel); // enleve le tableau pour le remettre apres, il
																		// s'actualise pas sinon
		String[][] tempTab = { SheetsQuickstart.tab[SheetsQuickstart.numeroLigne] }; // cree le nouveau tableau
		tabAffiche = new JTable(tempTab, SheetsQuickstart.tab[0]);

		SheetsQuickstart.tabLayout(1700, 44);

		container.add(SheetsQuickstart.tabPanel);// remet le tableau dans la fenetre une fois qu'il est bien modifie

		jtfs.get(0).setText(tempTab[0][7]);
		jtfs.get(1).setText(tempTab[0][5]);
		jtfs.get(2).setText(tempTab[0][3]);
		jtfs.get(3).setText(tempTab[0][10]);
		jtfs.get(4).setText(tempTab[0][2]);
		try {
			List<List<Object>> sheetTemp = SheetsQuickstart.getSheet("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
					"Codes analytiques!A2:D");
			String codeAnal = SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][0];
			for (int i = 0; i < sheetTemp.size(); i++) {
				if (sheetTemp.get(i).get(2).equals(codeAnal))
					SheetsQuickstart.combos.get(3).setSelectedItem(sheetTemp.get(i).get(1));
			}
		} catch (IOException exception) {
			System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
		} catch (GeneralSecurityException exception2) {
			System.out.println("merde y'a une erreur mais je sais pas ou - GeneralSecurityException");
		}
	}

	// TODO CREER methode qui actualise les champs preremplis

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
				new FileReader(new File("src/main/resources/Sage Data/" + data + ".txt"))); // va chercher le fichier
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
						+ "\"	D	B	\"" + selected[i][3] + "\"		\"6\"	\"" + selected[i][11] + "\"	\""
						+ selected[i][9] + "\"	\"" + codeAnal + "\"	\"" + selected[i][7] + "\"\n");
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

			mouvement.delete();
			tiers.delete();
			journal.delete();
			modePaiement.delete();

			SheetsQuickstart.init();

		} catch (HeadlessException he) {
			he.printStackTrace();
		} catch (GeneralSecurityException he) {
			System.out.println("merde y'a une erreur mais je sais pas ou - GeneralSecurityException");
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
			e.printStackTrace();
		} catch (IOException e) {
			// Celle-ci se produit lors d'une erreur d'écriture ou de lecture
			e.printStackTrace();
		} finally {
			// On ferme nos flux de données dans un bloc finally pour s'assurer
			// que ces instructions seront exécutées dans tous les cas même si
			// une exception est levée !
			try {
				if (fis != null)
					fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

			try {
				if (fos != null)
					fos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/*
	 * public static void createText(String sheetId, String range, String name,
	 * boolean clubAssocie) // sert a importer // plus facilement // les donnees de
	 * // sage throws IOException, GeneralSecurityException { List<List<Object>>
	 * sheet = SheetsQuickstart.getSheet(sheetId, range); File res = new
	 * File("src/main/resources/Sage Data/" + name + ".txt"); BufferedWriter writer
	 * = new BufferedWriter(new FileWriter(res)); int border = sheet.get(0).size();
	 * if (clubAssocie) border--; for (int i = 0; i < sheet.size(); i++) { for (int
	 * j = 0; j < border; j++) { if (sheet.get(i).get(j) != "") {
	 * writer.write(sheet.get(i).get(j) + ";;"); } else writer.write(" ;;"); }
	 * writer.write("\n"); } writer.close(); }
	 */

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
		// List<List<Object>> test = SheetsQuickstart.getData("codes analytiques");

		/*
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "Clubs!A1:B","idSheets",false); SheetsQuickstart.createText(
		 * "1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ","Codes analytiques!A1:D"
		 * ,"codes analytiques",true);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "JournauxBanque!A1:E","journaux de banques",true);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "CompteBanque!A1:E","comptes bancaires",true);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "JournauxVenteAchat!A1:E","journaux ventes et achats",true);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "CompteCharge!A1:C","comptes de charge",false);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "CompteClient!A1:C","comptes clients",false);
		 * SheetsQuickstart.createText("1fPMx7whVngMUKxdU-RRlYeRodVxsY0ifM9LBocWXhQQ",
		 * "CompteFournisseur!A1:C","comptes fournisseur",false);
		 */

		SheetsQuickstart.init();
	}
}