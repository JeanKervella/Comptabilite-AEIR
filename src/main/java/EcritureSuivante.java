import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JTextField;

public class EcritureSuivante implements ActionListener {

	public void actionPerformed(ActionEvent e) {

		if (SheetsQuickstart.selectedCorrect()) {
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][0] = "true";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][1] = (String) ((DynamicList) SheetsQuickstart.components
					.get(0)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][2] = (String) ((JTextField) SheetsQuickstart.components
					.get(1)).getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][3] = (String) ((JTextField) SheetsQuickstart.components
					.get(2)).getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][4] = ((String) ((JTextField) SheetsQuickstart.components
					.get(3)).getText()).replace(".", ",");
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][5] = (String) ((DynamicList) SheetsQuickstart.components
					.get(4)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][6] = (String) ((DynamicList) SheetsQuickstart.components
					.get(5)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][7] = (String) ((DynamicList) SheetsQuickstart.components
					.get(6)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][8] = (String) ((DynamicList) SheetsQuickstart.components
					.get(7)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][9] = (String) ((JTextField) SheetsQuickstart.components
					.get(8)).getText();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][10] = (String) ((DynamicList) SheetsQuickstart.components
					.get(9)).getSelectedItem();
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][11] = (String) ((JComboBox) SheetsQuickstart.components
					.get(10)).getSelectedItem();
			if (((String) ((JComboBox) SheetsQuickstart.components.get(10)).getSelectedItem()).equals("CHQ")) {
				SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][12] = (String) ((JTextField) SheetsQuickstart.components
						.get(11)).getText();
			} else
				SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][12] = ".";
			String modePaie = (String) ((JComboBox) SheetsQuickstart.components.get(10)).getSelectedItem();
			int moPaie = 0;
			if (modePaie.equals("VIR"))
				moPaie = 0;
			if (modePaie.equals("CB"))
				moPaie = 1;
			if (modePaie.equals("CHQ"))
				moPaie = 2;

			if (Integer.parseInt((String) SheetsQuickstart.tabAffiche.getValueAt(0, SheetsQuickstart.sheetsColumn[moPaie][7])) == 1)
				SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][13] = String
						.valueOf(SheetsQuickstart.numeroLigne);
			else {
				if (SheetsQuickstart.selected[SheetsQuickstart.numeroLigne - 1][13].equals(".") || Integer.parseInt(
						SheetsQuickstart.selected[SheetsQuickstart.numeroLigne - 1][13]) == SheetsQuickstart.numeroLigne
								- 1) {
					SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][13] = String
							.valueOf(SheetsQuickstart.numeroLigne - 1 + Integer.parseInt((String) SheetsQuickstart.tabAffiche.getValueAt(0,
									SheetsQuickstart.sheetsColumn[moPaie][7])));
				} else {
					SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][13] =SheetsQuickstart.selected[SheetsQuickstart.numeroLigne-1][13];
				}
			}

			int test = SheetsQuickstart.numeroLigne;
			test++;
			while (test < SheetsQuickstart.tab.length
					&& (!SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][6]]
							.equals("")
							|| SheetsQuickstart.tab[SheetsQuickstart.numeroLigne][SheetsQuickstart.sheetsColumn[moPaie][4]]
									.isEmpty())) {
				test++;
			}
			if (test >= SheetsQuickstart.tab.length) {
				String ObjButtons[] = { "Yes", "No" };
				int PromptResult = JOptionPane.showOptionDialog(null, "Les ecritures sont terminees pour cet excel\n",
						"Excel termine", JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE, null, ObjButtons,
						ObjButtons[1]);
				if (PromptResult == JOptionPane.YES_OPTION) {
					try {
						SheetsQuickstart.createText();
					} catch (IOException exception) {
						JOptionPane jop = new JOptionPane();
						jop.showMessageDialog(null, exception.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
					}
				}

			} else {
				SheetsQuickstart.numeroLigne = test;
				((JProgressBar) SheetsQuickstart.components.get(17))
						.setValue(((JProgressBar) SheetsQuickstart.components.get(17)).getValue() + 1);
				SheetsQuickstart.updateEcriture();
				System.out.println(SheetsQuickstart.toStringSelected());
			}
		}
	}

}
