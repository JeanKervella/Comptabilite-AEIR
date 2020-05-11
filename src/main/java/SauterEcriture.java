import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.security.GeneralSecurityException;

import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;

public class SauterEcriture implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		boolean lastGroup=true;
		lastGroup = SheetsQuickstart.selected[SheetsQuickstart.numeroLigne - 1][13] != ".";
		if (lastGroup) {
			lastGroup=(Integer.parseInt(
					SheetsQuickstart.selected[SheetsQuickstart.numeroLigne - 1][13]) != SheetsQuickstart.numeroLigne
							- 1);
			if(lastGroup){
				System.out.println(SheetsQuickstart.selected[SheetsQuickstart.numeroLigne - 1][13]);
				System.out.println(String.valueOf(SheetsQuickstart.numeroLigne - 1));
				JOptionPane jop = new JOptionPane();
				jop.showMessageDialog(null,
						"Les ecritures precedentes correspondent a un paiement groupe.\n Vous ne pouvez pas passer cette ecriture.\nOu sinon supprimer le ou les ecritures precedentes",
						"Erreur paiement groupe", JOptionPane.ERROR_MESSAGE);
			}
		} if(!lastGroup) {
			int test = SheetsQuickstart.numeroLigne;
			String modePaie = (String) ((JComboBox) SheetsQuickstart.components.get(10)).getSelectedItem();
			int moPaie = 0;
			if (modePaie.equals("VIR"))
				moPaie = 0;
			if (modePaie.equals("CB"))
				moPaie = 1;
			if (modePaie.equals("CHQ"))
				moPaie = 2;
			try {
			test += Integer.parseInt(
					SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][7]]);
			} catch (NumberFormatException e1) {
				test+=1;
			}

			while (test < SheetsQuickstart.tab.length
					&& (!SheetsQuickstart.tab[test][SheetsQuickstart.sheetsColumn[moPaie][6]].equals("")
							|| SheetsQuickstart.tab[test][SheetsQuickstart.sheetsColumn[moPaie][4]].equals(""))) {
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
				try {
				((JProgressBar) SheetsQuickstart.components.get(17))
						.setValue(((JProgressBar) SheetsQuickstart.components.get(17)).getValue() + Integer.parseInt(
								SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][7]]));
				} catch(NumberFormatException e1) {
					((JProgressBar) SheetsQuickstart.components.get(17))
					.setValue(((JProgressBar) SheetsQuickstart.components.get(17)).getValue() +1);
				}
				SheetsQuickstart.numeroLigne = test;
				SheetsQuickstart.updateEcriture();
				System.out.println(SheetsQuickstart.toStringSelected());
			}
		}
	}

}
