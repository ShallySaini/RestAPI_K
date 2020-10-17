package Excel;

import java.io.IOException;
import java.util.ArrayList;

public class getDataFromExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		dataDrivenExcel d = new dataDrivenExcel();
		ArrayList data = d.getData("Logout");
		System.out.println(data);
		System.out.println(d);
	}

}
