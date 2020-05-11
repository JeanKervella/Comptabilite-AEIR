
public class secondThread extends Thread {

	public secondThread() {
		super();
	}

	public secondThread(String name) {
		super(name);
	}
	
	public void run() {
		SheetsQuickstart.updateEcriturerestante();
	}
}
