package dataDriven;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvider {
	
	//multiple sets of data tour tests
	//array
	//Sets of data as 5 arrays from data provider to your tests
	//then your test will run 5 seperate sets of data(arrays)
	
	@Test(dataProvider = "driverTest")
	public void testCaseData(String Greeting,String Communication,String id) {
			
		System.out.println(Greeting+Communication+id);
	}
	DataFormatter formatter= new DataFormatter();
	@DataProvider(name = "driverTest")
	public Object[][] getData() throws IOException {
		
		//every row of excel should be sent to 1 array
		FileInputStream fis = new FileInputStream("C:\\Users\\krish\\OneDrive\\Documents\\excelDriven.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet =wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row =sheet.getRow(0);
		int colcount = row.getLastCellNum();
		
		Object [][]data= new Object[rowCount -1][colcount];
		for (int i = 0; i < rowCount -1; i++) {
			row = sheet.getRow(i+1);
			for (int j = 0; j < colcount; j++) {
				
				XSSFCell cell =row.getCell(j);				
				data[i][j] =formatter.formatCellValue(cell);
			}
		}
		return data;
	}
	

}
