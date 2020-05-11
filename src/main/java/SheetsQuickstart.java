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
import java.text.SimpleDateFormat;
import java.time.LocalDate;
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
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.TimeUnit;

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
	public static volatile boolean changed;
	public static String spreadsheetId;
	public static String[][] ecritureRestante;
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
		JComponent temp = new JOptionPane();
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
		} else if (type.equals("progress")) {
			temp = new JProgressBar();
			((JProgressBar) temp).setStringPainted(true);
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
					if (i == 4 || i == 5) {
						String ObjButtons[] = { "Yes", "No" };
						int PromptResult = JOptionPane.showOptionDialog(null,
								"Soit dans \"compte facture\" soit dans \"compte fournisseur\", soit les deux l'item selectionne n'est pas dans la liste.\n Voulez-vous vraiment continuer et creer un nouveau compte ?\n",
								"Confirmation de changement d'ecriture", JOptionPane.DEFAULT_OPTION,
								JOptionPane.WARNING_MESSAGE, null, ObjButtons, ObjButtons[1]);
						if (PromptResult != JOptionPane.YES_OPTION) {
							return false;
						} else {
							String nouveauCompte = JOptionPane.showInputDialog(
									"Veuillez saisir le code et le nom du compte que vous voulez creer.\nLa forme que vous devez rentrer etant : \"codeCompte - nomCompte\"");

							String[] codeNomCompte = nouveauCompte.split(" - ");
							boolean testcode = true;
							JOptionPane jop = new JOptionPane();
							if (i == 4 && !nouveauCompte.substring(0, 1).equals("6")) {
								jop.showMessageDialog(null, "Vous devez rentrer un compte 6 !", "Erreur code Compte",
										JOptionPane.ERROR_MESSAGE);
								testcode = false;
								return false;
							}
							if (i == 5 && !nouveauCompte.substring(0, 3).equals("401")) {
								jop.showMessageDialog(null, "Vous devez rentrer un compte 401 !", "Erreur code Compte",
										JOptionPane.ERROR_MESSAGE);
								testcode = false;
								return false;
							}
							for (int j = 0; j < ((DynamicList) components.get(i)).getItemCount(); j++) {
								if (codeNomCompte[0]
										.equals(((DynamicList) components.get(i)).getItemAt(j).split(" - ")[0])) {
									testcode = false;
									jop.showMessageDialog(null,
											"Le code de compte que vous avez rentre est deja attribue",
											"Erreur code compte", JOptionPane.ERROR_MESSAGE);
									return false;
								}
							}
							System.out.println(testcode);
							System.out.println(codeNomCompte[0].length());
							if (codeNomCompte[0].length() > 8) {
								testcode = false;
								jop.showMessageDialog(null,
										"Le code que vous avez rentre est trop long, il doit faire au maximum 8 caracteres",
										"Erreur code compte", JOptionPane.ERROR_MESSAGE);
								return false;
							}
							if (testcode) {

								((DynamicList) components.get(i)).addItem(nouveauCompte);
								((DynamicList) components.get(i)).setSelectedItem(nouveauCompte);

								try {
									File file;
									if (i == 4)
										file = new File("src/main/resources/Data/comptes de charge.txt");
									else if (i == 5)
										file = new File("src/main/resources/Data/comptes fournisseur.txt");
									else
										file = new File("src/main/resources/Data/poubelle.txt");
									BufferedWriter writer = new BufferedWriter(new FileWriter(file, true));
									writer.write(
											codeNomCompte[0] + ";;" + codeNomCompte[1] + ";;" + nouveauCompte + ";;\n");
									writer.close();

								} catch (IOException e) {
									jop.showMessageDialog(null, e.toString(), "Error", JOptionPane.ERROR_MESSAGE);
								}

							} else
								return false;
						}
					} else {
						JOptionPane PromptResult = new JOptionPane();
						PromptResult.showMessageDialog(null,
								"Il y a au moins un des champs de liste deroulante (autre que compte facture et compte fournisseur) ou l'item selectionne n'est pas dans la liste.\n Vous devez choisir un des choix disponibles\n",
								"Confirmation de changement d'ecriture", JOptionPane.ERROR_MESSAGE);
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
		sheetsRange = new String[sheetTemp.size() - 1][sheetTemp.get(0).size()];
		for (int i = 0; i < sheetTemp.size() - 1; i++) {
			for (int j = 0; j < sheetTemp.get(0).size(); j++) {
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
				System.out.println(column + "   " + firstColumn);
			}
		}
	}

	public static String toStringColumns() {
		String res = "       ";
		List<List<Object>> sheetTemp = new ArrayList();
		try {
			sheetTemp = SheetsQuickstart.getData("sheetsRange");
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		System.out.println(sheetTemp.size());
		for (int i = 2; i < sheetTemp.get(0).size(); i++) {
			res += sheetTemp.get(0).get(i) + "  ";
		}
		res += "\n";
		for (int i = 0; i < sheetsColumn.length; i++) {
			if (i == 0)
				res += "VIR :    ";
			if (i == 1)
				res += "CB :     ";
			if (i == 2)
				res += "CHQ :    ";
			for (int j = 0; j < sheetsColumn[0].length; j++) {
				res += sheetsColumn[i][j] + "                         ";
			}
			res += "\n";
		}
		return res;
	}

	public static void hideAllScrollPane() {
		DynamicList temp = new DynamicList();
		for (int i = 0; i < components.size(); i++) {
			System.out.println(components.get(i).getClass());
			if (components.get(i).getClass().equals(temp.getClass())) {
				((DynamicList) components.get(i)).hideScrollPane();
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
	private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
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
		for (int i = 0; i < sheetsRange.length; i++) {
			if (sheetsRange[i][0].equals(paiement))
				return sheetsRange[i][1];
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
		for (int i = 0; i < sheetsRange.length; i++) {
			if (sheetsRange[i][1].equals(range))
				return sheetsRange[i][0];
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
	 * Methode qui permet d'actualiser le tableau du menu repertoriant les ecritures
	 * restantes a comptabiliser
	 * 
	 * @return la hauteur en pixel du tableau
	 */
	public static void updateEcriturerestante() {
		try {
			List<List<Object>> sheetTemp = SheetsQuickstart.getData("idSheets"); // va chercher dans les ressources les
																					// ID
																					// des excels
			List<List<Object>> sheetRange = getData("sheetsRange");
			ecritureRestante = new String[sheetTemp.size()][8]; // cree un tableau de 8 colonnes et dont le nombre de
																// lignes correspond au nombre d'excels differents
			ecritureRestante[0][0] = "Club";
			ecritureRestante[0][1] = "fiches de caisse";
			ecritureRestante[0][2] = "debit - VIR";
			ecritureRestante[0][3] = "debit - CHQ";
			ecritureRestante[0][4] = "debit - CB";
			ecritureRestante[0][5] = "credit - VIR";
			ecritureRestante[0][6] = "credit - CHQ";
			ecritureRestante[0][7] = "credit - CB";

			for (int i = 1; i < sheetTemp.size(); i++) {
				System.out.println(i);
				System.out.println("\n" + sheetTemp.get(i).get(0) + "\n");
				ecritureRestante[i][0] = (String) sheetTemp.get(i).get(0);
				ecritureRestante[i][1] = "0";
				int j = 0;
				TimeUnit.SECONDS.sleep(10);
				List<List<Object>> sT = SheetsQuickstart.getSheet((String) sheetTemp.get(i).get(1),
						(String) sheetRange.get(1).get(1));
				String test1;
				String test2;
				System.out.println(sT.size());
				for (int k = 1; k < sT.size(); k++) {
					test1 = (String) sT.get(k).get(SheetsQuickstart.sheetsColumn[0][4]);
					test2 = (String) sT.get(k).get(sheetsColumn[0][6]);
					if (!test1.equals("") && test2.equals(""))
						j++;
				}
				ecritureRestante[i][2] = Integer.toString(j);
				j = 0;
				TimeUnit.SECONDS.sleep(10);
				sT = SheetsQuickstart.getSheet((String) sheetTemp.get(i).get(1), (String) sheetRange.get(3).get(1));
				System.out.println(sT.size());
				for (int k = 1; k < sT.size(); k++) {
					test1 = (String) sT.get(k).get(sheetsColumn[2][4]);
					test2 = (String) sT.get(k).get(sheetsColumn[2][6]);
					if (!test1.equals("") && test2.equals(""))
						j++;
				}
				ecritureRestante[i][3] = Integer.toString(j);
				j = 0;
				TimeUnit.SECONDS.sleep(10);
				sT = SheetsQuickstart.getSheet((String) sheetTemp.get(i).get(1), (String) sheetRange.get(2).get(1));
				System.out.println(sT.size());
				for (int k = 1; k < sT.size(); k++) {
					test1 = (String) sT.get(k).get(sheetsColumn[1][4]);
					test2 = (String) sT.get(k).get(sheetsColumn[2][6]);
					if (!test1.equals("") && test2.equals(""))
						j++;
				}
				ecritureRestante[i][4] = Integer.toString(j);
				ecritureRestante[i][5] = "0";
				ecritureRestante[i][6] = "0";
				ecritureRestante[i][7] = "0";

			}

			BufferedWriter writer = new BufferedWriter(new FileWriter(new File("src/main/resources/Data/mainTab.txt")));
			for (int i = 0; i < ecritureRestante.length; i++) {
				for (int j = 0; j < ecritureRestante[0].length; j++) {
					writer.write(ecritureRestante[i][j] + ";;");
				}
				writer.write("\n");
			}
			writer.close();
		} catch (IOException |GeneralSecurityException | InterruptedException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.toString(), "Error", JOptionPane.ERROR_MESSAGE);
		} 
	}

	public static void initEcritureRestante() {
		try {
		List<List<Object>> data = getData("mainTab");
		ecritureRestante = new String[data.size()][data.get(0).size()];
		for (int i = 0; i < data.size(); i++) {
			for (int j = 0; j < data.get(0).size(); j++) {
				ecritureRestante[i][j] = (String) data.get(i).get(j);
			}
		}} catch(IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		}
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
		firstThread t1 = new firstThread();
		secondThread t2 = new secondThread();
		t1.start();
		try {
		TimeUnit.SECONDS.sleep(3);
		}catch (InterruptedException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		}
		t2.start();

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
		for (int i = 0; i < container.getComponentCount(); i++) {
			if (container.getComponent(i).equals(tabPanel))
				container.remove(i);
		}
		String[][] tempTab = { SheetsQuickstart.tab[SheetsQuickstart.numeroLigne] }; // cree le nouveau tableau
		tabAffiche = new JTable(tempTab, SheetsQuickstart.tab[0]);

		SheetsQuickstart.tabLayout(1800, 66, 44);

		container.add(SheetsQuickstart.tabPanel);// remet le tableau dans la fenetre une fois qu'il est bien modifie

		int y = 0;
		JTextField temp = new JTextField();
		if (((JComboBox) components.get(10)).getSelectedItem().equals("CB"))
			y = 1;
		if (((JComboBox) components.get(10)).getSelectedItem().equals("CHQ"))
			y = 2;
		((JTextField) components.get(1)).setText(tempTab[0][sheetsColumn[y][0]]); // tempTab[0][0]
		if (((JTextField) components.get(1)).getText().isEmpty()) // TODO enlever cela une fois que les gens
																	// commenceront a rentrer les dates de facture
			((JTextField) components.get(1)).setText(tempTab[0][sheetsColumn[y][4]]);
		((JTextField) components.get(2)).setText(tempTab[0][sheetsColumn[y][1]]);
		((JTextField) components.get(3)).setText(tempTab[0][sheetsColumn[y][2]]);
		((JTextField) components.get(8)).setText(tempTab[0][sheetsColumn[y][4]]);
		if (components.get(11).getClass().equals(temp.getClass()))
			((JTextField) components.get(11)).setText(tempTab[0][sheetsColumn[y][5]]);
		try {
			List<List<Object>> sheetTemp = SheetsQuickstart.getData("codes analytiques");
			String codeAnal = SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[y][3]];
			System.out.println("In try");
			for (int i = 1; i < sheetTemp.size(); i++) {
				System.out.println(codeAnal + "      " + sheetTemp.get(i).get(2));
				if (sheetTemp.get(i).get(2).equals(codeAnal)) {
					System.out.println("in if");
					((DynamicList) SheetsQuickstart.components.get(6))
							.setSelectedItem((String) sheetTemp.get(i).get(1));
				}
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
		File mouvement = new File("src/main/resources/mouvement.txt");
		File tiers = new File("src/main/resources/tiers.txt");
		File journal = new File("src/main/resources/journal.txt");
		File modePaiement = new File("src/main/resources/modePaiement.txt");
		List<List<String>> compare;
		try {
			compare = SheetsQuickstart.writeFile(mouvement, tiers, journal, modePaiement);

			String ObjButtons[] = { "Yes", "No" };
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

				mouvement.delete();
				tiers.delete();
				journal.delete();
				modePaiement.delete();

				String moPaiement = (String) ((JComboBox) components.get(10)).getSelectedItem();
				int moPaie = 0;
				if (moPaiement.equals("CB"))
					moPaie = 1;
				if (moPaiement.equals("CHQ"))
					moPaie = 2;
				SheetsQuickstart.writeSheetDate(spreadsheetId, moPaie);

				SheetsQuickstart.init();
			}
			if (PromptResult == JOptionPane.NO_OPTION) {

				JOptionPane jop1 = new JOptionPane();
				jop1.showMessageDialog(null,
						"Vous allez voir apparaitre un tableau avec toutes les donnees enregistrees pour cet excel.\nVeuillez remplacer \"true\" par \"false\" pour toutes les ecritures qui posent probleme.\n\nLes lignes qui ne contiennent que des points correspondent a des ecritures sautees ou des ecritures deja comptabilisees",
						"Selection des donnees", JOptionPane.INFORMATION_MESSAGE);

				container = new JPanel();
				container.setLayout(null);
				String[] title = new String[selected[0].length];
				title[0] = "ecriture a comptabiliser";
				title[1] = "code journal";
				title[2] = "date facture";
				title[3] = "libelle";
				title[4] = "montant";
				title[5] = "compte facture";
				title[6] = "compte fournisseur";
				title[7] = "code analytique";
				title[8] = "code banque";
				title[9] = "date paiement";
				title[10] = "compte banque";
				title[11] = "mode de paiement";
				title[12] = "numero de cheque";
				title[13] = "numero de ligne";

				toStringSelected();
				tabAffiche = new JTable(selected, title);
				tabAffiche.setRowHeight(40);
				JScrollPane scrollpane = new JScrollPane(tabAffiche);
				frame.setBounds(0, 0, 1500, 700);
				container.setBounds(frame.getBounds());
				tabAffiche.setBounds(0, 0, frame.getWidth(), frame.getHeight() - 200);
				scrollpane.setBounds(tabAffiche.getBounds());
				container.add(scrollpane);

				try {
					JComponent test = components.get(18);
					container.add(components.get(18));
				} catch (IndexOutOfBoundsException e) {
					addBouton("enregistrer", "button", (int) (frame.getWidth() / 2 - 50), frame.getHeight() - 125, 100,
							50);
					((JButton) components.get(18)).addActionListener(new ChangeSelected());
				}

				frame.setContentPane(container);
				frame.setResizable(true);
				frame.setVisible(true);

			}

		} catch (GeneralSecurityException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}

	}

	public static List<List<String>> writeFile(File mouvement, File tiers, File journal, File modePaiement) {
		List<List<String>> compare = new ArrayList(); // des qu'un fournisseur, un journal aura ete selectionne il est
		// ajoute dans cette list de list pour qu'il n'y est pas de
		// doublon
		compare.add(new ArrayList()); // correspond aux tiers deja marques
		compare.add(new ArrayList()); // correspond aux journaux ventes et achats deja marques
		compare.add(new ArrayList()); // correspond aux journaux de banques deja marques
		compare.add(new ArrayList()); // correspond aux mode de paiements deja marques

		try {
			List<List<Object>> sageData = SheetsQuickstart.getData("codes analytiques");

			BufferedWriter writer = new BufferedWriter(new FileWriter(mouvement));
			writer.write("##Transfert\n" + "##Section	Dos\n" + "EUR\n" + "##Section	Mvt\n");
			int compteur = 1;
			for (int i = 0; i < selected.length; i++) {
				if (selected[i][0].equals("true")) {
					String codeAnal = ".";
					int j = 0;
					sageData = SheetsQuickstart.getData("codes analytiques");
					while (codeAnal.equals(".")) {
						if (sageData.get(j).get(1).equals(selected[i][7])) {
							codeAnal = (String) sageData.get(j).get(0);
						}
						j++;
					}
					String[] codeNomCompte = selected[i][5].split(" - ");
					writer.write("\"" + compteur + "\"	\"" + selected[i][1] + "\"	\"" + selected[i][2] + "\"	\""
							+ codeNomCompte[0] + "\"	\"" + codeNomCompte[1] + "\"	\"" + selected[i][4]
							+ "\"	D	B	\"" + selected[i][3] + "\"		\"10\"			\"" + codeAnal + "\"	\""
							+ selected[i][7] + "\"\r\n");
					codeNomCompte = selected[i][6].split(" - ");
					writer.write("\"" + compteur + "\"	\"" + selected[i][1] + "\"	\"" + selected[i][2] + "\"	\""
							+ codeNomCompte[0] + "\"	\"" + codeNomCompte[1] + "\"	\"" + selected[i][4]
							+ "\"	C	B	\"" + selected[i][3] + "\"		\"10\"		\"" + selected[i][2]
							+ "\"\r\n");
					if (Integer.parseInt(selected[i][13]) == i) {
						int numLignePrec = i - 1;
						System.out.println("i :  " + i);
						while (selected[i][13].equals(selected[numLignePrec][13])) {
							selected[i][4] = String
									.valueOf(round(
											Double.parseDouble(selected[i][4].replace(",", "."))
													+ Double.parseDouble(selected[numLignePrec][4].replace(",", ".")),
											2))
									.replace(".", ",");
							numLignePrec--;
						}
						compteur++;
						writer.write("\"" + compteur + "\"	\"" + selected[i][8] + "\"	\"" + selected[i][9] + "\"	\""
								+ codeNomCompte[0] + "\"	\"" + codeNomCompte[1] + "\"	\"" + selected[i][4]
								+ "\"	D	B	\"" + selected[i][3] + "\"	");
						if (selected[i][11].equals("CHQ"))
							writer.write("\"" + selected[i][12] + "\"");
						codeNomCompte = selected[i][10].split(" - ");
						writer.write("	\"6\"	\"" + selected[i][11] + "\"	\"" + selected[i][9] + "\"	\"" + codeAnal
								+ "\"	\"" + selected[i][7] + "\"\n");
						writer.write("\"" + compteur + "\"	\"" + selected[i][8] + "\"	\"" + selected[i][9] + "\"	\""
								+ codeNomCompte[0] + "\"	\"" + codeNomCompte[1] + "\"	\"" + selected[i][4]
								+ "\"	C	B	\"" + selected[i][3] + "\"		\"6\"	\"" + selected[i][11] + "\"	\""
								+ selected[i][9] + "\"\n");
					}
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
				writer.write(
						"\"" + sageData.get(cursor).get(0) + "\"	\"" + sageData.get(cursor).get(1) + "\"	\"SR\"\n");
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
		} catch (HeadlessException he) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, he.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (IOException he) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, he.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		return compare;
	}

	public static double round(double value, int places) {
		if (places < 0)
			throw new IllegalArgumentException();

		BigDecimal bd = BigDecimal.valueOf(value);
		bd = bd.setScale(places, RoundingMode.HALF_UP);
		return bd.doubleValue();
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
			throws GeneralSecurityException, IOException {

		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		Sheets service = new Sheets.Builder(HTTP_TRANSPORT, SheetsQuickstart.JSON_FACTORY,
				getCredentials(HTTP_TRANSPORT)).setApplicationName(APPLICATION_NAME).build();
		ValueRange response = service.spreadsheets().values().get(sheetId, range).execute();
		return response.getValues();
	}

	public static void writeSheet(String spreadsheetId, String range, List<List<Object>> values) {
		try {
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			Sheets service = new Sheets.Builder(HTTP_TRANSPORT, SheetsQuickstart.JSON_FACTORY,
					getCredentials(HTTP_TRANSPORT)).setApplicationName(APPLICATION_NAME).build();
			ValueRange body = new ValueRange().setValues(values);
			UpdateValuesResponse result = service.spreadsheets().values().update(spreadsheetId, range, body)
					.setValueInputOption("USER_ENTERED").execute();
			System.out.printf("%d cells updated.", result.getUpdatedCells());
		} catch (GeneralSecurityException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
	}

	/**
	 * @param spreadsheetId
	 * @param modePaiement  0 pour VIR, 1 pour CB et 2 pour CHQ
	 */
	public static void writeSheetDate(String spreadsheetId, int modePaiement) {
		String range = sheetsRange[modePaiement][1].substring(0, sheetsRange[modePaiement][1].length() - 4);
		range += sheetsRange[modePaiement][8];
		range += sheetsRange[modePaiement][1].substring(sheetsRange[modePaiement][1].length() - 3,
				sheetsRange[modePaiement][1].length() - 2);
		range += ":" + sheetsRange[modePaiement][8];
		System.out.println(range);
		List<List<Object>> values = new ArrayList();
		List<Object> temp = new ArrayList();
		temp.add(new SimpleDateFormat("dd/MM/yyyy").format(new Date()));
		List<Object> temp2 = new ArrayList();
		temp2.add(null);
		for (int i = 0; i < selected.length; i++) {
			if (selected[i][0].equals("true"))
				values.add(temp);
			else
				values.add(temp2);
		}
		writeSheet(spreadsheetId, range, values);

	}

	public static void main(String... args) throws IOException, GeneralSecurityException {
		initEcritureRestante();
		init();
	}
}