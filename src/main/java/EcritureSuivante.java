import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

public class EcritureSuivante implements ActionListener {

	public void actionPerformed(ActionEvent e) {
		boolean selectedItemCorrect = true;
		for (int i = 0; i < 7; i++) {
			boolean test = false;
			for (int j = 0; j < SheetsQuickstart.combos.get(i).getItemCount(); j++) {
				if (SheetsQuickstart.combos.get(i).getItemAt(j)
						.equals(SheetsQuickstart.combos.get(i).getSelectedItem()))
					test = true;
			}
			if (!test)
				selectedItemCorrect = false;
		}
		if (!selectedItemCorrect) {
			String ObjButtons[] = { "Yes", "No" };
			int PromptResult = JOptionPane.showOptionDialog(null,
					"Il y a au moins un des champs de liste deroulante ou l'item selectionne n'est pas dans la liste\n Voulez-vous vraiment continuer ?\nLe logiciel Sage creera un nouveau compte/nouveau journal",
					"Confirmation de changement d'ecriture", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE,
					null, ObjButtons, ObjButtons[1]);
			if (PromptResult == JOptionPane.YES_OPTION) {
				selectedItemCorrect = true;
			}
		}
		if (selectedItemCorrect) {
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][0] = "true";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][1] = (String) SheetsQuickstart.combos.get(0)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][2] = (String) SheetsQuickstart.jtfs.get(0)
					.getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][3] = (String) SheetsQuickstart.jtfs.get(1)
					.getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][4] = (String) SheetsQuickstart.jtfs.get(2)
					.getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][5] = (String) SheetsQuickstart.combos.get(1)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][6] = (String) SheetsQuickstart.combos.get(2)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][7] = (String) SheetsQuickstart.combos.get(3)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][8] = (String) SheetsQuickstart.combos.get(4)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][9] = (String) SheetsQuickstart.jtfs.get(3)
					.getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][10] = (String) SheetsQuickstart.combos.get(5)
					.getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][11] = (String) SheetsQuickstart.combos.get(6)
					.getSelectedItem();
			if (((String) SheetsQuickstart.combos.get(6).getSelectedItem()).equals("CHQ")) {
				SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][12] = (String) SheetsQuickstart.jtfs.get(4)
						.getText();
			} else
				SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][12] = "";

			int test = SheetsQuickstart.numeroLigne;
			test++;
			while (test < SheetsQuickstart.tab.length
					&& !SheetsQuickstart.tab[test][11].equals("")) {
				test++;
			}
			if (test >= SheetsQuickstart.tab.length) {
				String ObjButtons[] = { "Yes", "No" };
				int PromptResult = JOptionPane.showOptionDialog(null,
						"Les ecritures sont terminees pour cet excel\n",
						"Excel termine", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons,
						ObjButtons[1]);
				if (PromptResult == JOptionPane.YES_OPTION) {
					try {
						SheetsQuickstart.createText();
					} catch (IOException exception) {
						System.out.println("merde y'a une erreur mais je sais pas ou - IOException");
					}
				}

			} else {
				SheetsQuickstart.numeroLigne=test;
				SheetsQuickstart.updateEcriture();
				System.out.println(SheetsQuickstart.toStringSelected());
			}
		}

	}

}
