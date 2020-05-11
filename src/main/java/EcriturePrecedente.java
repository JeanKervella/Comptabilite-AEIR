import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;

public class EcriturePrecedente implements ActionListener {

	public void actionPerformed(ActionEvent e) {

		boolean lastLineModified = false;
		int test =SheetsQuickstart.numeroLigne;
		while (!lastLineModified && test>0) {
			test--;
			if (SheetsQuickstart.selected[test][0] == "true")lastLineModified=true;
		}
		if (lastLineModified) {
			SheetsQuickstart.numeroLigne=test;
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][0] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][1] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][2] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][3] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][4] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][5] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][6] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][7] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][8] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][9] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][10] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][11] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][12] = ".";
			SheetsQuickstart.selected[SheetsQuickstart.numeroLigne][13] = ".";
			SheetsQuickstart.updateEcriture();
			System.out.println(SheetsQuickstart.toStringSelected());
			((JProgressBar) SheetsQuickstart.components.get(17))
			.setValue(((JProgressBar) SheetsQuickstart.components.get(17)).getValue() - 1);
			}else { 
				JOptionPane.showMessageDialog(null,
						"Il n'y a eu aucune ecriture d'ecrite auparavant",
						"Erreur : Aucune Ecriture precedente", JOptionPane.INFORMATION_MESSAGE);
		}

	}

}
