package Excel;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDrivenExcel {
	
	public ArrayList getData(String testCaseName) throws IOException {
		
ArrayList ar = new ArrayList<>();
		
		FileInputStream fis = new FileInputStream("C://Shally//SoftwareTesting//ITCareer//Selenium//ExcelDemo.xlsx");
		 
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		
		int sheetCount = workBook.getNumberOfSheets();
		
		for(int i=0; i<sheetCount; i++)
		{
			if(workBook.getSheetName(i).equalsIgnoreCase("Sheet1"))
			{
				XSSFSheet sheet = workBook.getSheetAt(i); //sheet is collection of rows.
				
				//Now scan entire first row to get testcases column
				//Now scan testcase column to get the testcase you want to execute
				//after you grab the testcase row = pull all the data of that row and feed into test.
				Iterator<Row> rows = sheet.iterator();//row is collection of cells.
				Row firstRow = rows.next();
				Iterator<Cell> cells = firstRow.cellIterator();
				int k = 0;
				int column = 0;
				while(cells.hasNext())
				{
				Cell cellValue = cells.next();
				if (cellValue.toString().equalsIgnoreCase("TestCase"))
				{
					column = k;
				}
				k++;
				}
				System.out.println(column);
				//once the testcase name column is identified, scan the full column to get the specified testcase.
				while(rows.hasNext())
				{
					Row r = rows.next();
				if(	r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))//got the row where test case is present.
				{
					//now get all the column values by iterating through the columns.
					Iterator<Cell> cv = r.cellIterator();
					while(cv.hasNext())
					{
						Cell c = cv.next();
						if(c.getCellType()==CellType.STRING)
						{
						ar.add(c.getStringCellValue());
						}
						else
						{
							ar.add(c.getNumericCellValue());
						}
						
					}
				}
				}
				
			}
		}
		return ar;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
		
	}

}
