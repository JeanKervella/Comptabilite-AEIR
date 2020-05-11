import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.JOptionPane;

public class ChangeSelected implements ActionListener{

	@Override
	public void actionPerformed(ActionEvent e) {
		for(int i = 0; i<SheetsQuickstart.tabAffiche.getRowCount();i++) {
			for(int j=0;j<SheetsQuickstart.tabAffiche.getColumnCount();j++) {
				SheetsQuickstart.selected[i][j] = (String) SheetsQuickstart.tabAffiche.getValueAt(i, j);
			}
		}
		try {
			SheetsQuickstart.createText();
		} catch (IOException e1) {
			JOptionPane jop = new JOptionPane();
			jop.showMessageDialog(null, e1.getMessage(), "Erreur", JOptionPane.ERROR_MESSAGE);
		}
		
	}

}
