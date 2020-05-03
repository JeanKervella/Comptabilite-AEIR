import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.security.GeneralSecurityException;

import javax.swing.JOptionPane;

public class SauterEcriture implements ActionListener {

	@Override
	public void actionPerformed(ActionEvent e) {
			int test=SheetsQuickstart.numeroLigne;
			test++;
			while (test < SheetsQuickstart.tab.length
					&& !SheetsQuickstart.tab[test][11].equals("")) {
				test++;
			}
			if (test >= SheetsQuickstart.tab.length) {
				String ObjButtons[] = { "Yes", "No" };
				int PromptResult = JOptionPane.showOptionDialog(null,
						"Les ecritures sont terminees pour cet excel\nVous ne pourrez plus revenir en arriere apres ce message",
						"Excel termine", JOptionPane.OK_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons,
						ObjButtons[1]);
				if (PromptResult == JOptionPane.OK_OPTION) {
					try {
						SheetsQuickstart.createText();
					} catch (IOException exception) {
						JOptionPane jop = new JOptionPane();
						jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
					} 
				}

			} else {
				SheetsQuickstart.numeroLigne=test;
				SheetsQuickstart.updateEcriture();
				System.out.println(SheetsQuickstart.toStringSelected());
			}
		}

	}

