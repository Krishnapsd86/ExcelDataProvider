package dataDriven;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeExcelFile {

	public static void main(String[] args) throws IOException {
		
		//create the workbook 
		XSSFWorkbook wb = new XSSFWorkbook();
		//create the spreadsheet
		XSSFSheet sheet =wb.createSheet("Employee2");
		//create a row object
		XSSFRow row;
		//create map with object
		Map<String,Object[]> createData = new TreeMap<String,Object[]>();
		createData.put("1", new Object[] {"ID","Full Name","Gender"});
		createData.put("2", new Object[] {"1","Krishna Thapaliya","Male"});
		createData.put("3", new Object[] {"2","Bishnu Tamang","Female"});
		createData.put("4", new Object[] {"3","Bikisha Thapaliya","Female"});
		createData.put("5", new Object[] {"4","Arpan Basnet","Male"});
		createData.put("6", new Object[] {"5","Astha Raut","Female"});
		createData.put("7", new Object[] {"6","Tushar Gole","Male"});
		
		//Create a set and set the key
		Set<String> keyId =createData.keySet();
		
		//Iterated the keyid by 
		//row value with increament
		int rowid =0;
		for(String key : keyId) {			
			row =sheet.createRow(rowid++);
			Object[] obj=createData.get(key);
			
			//cell value with increament
			int cellid =0;
			for(Object obj1:obj) {
				Cell cell =row.createCell(cellid++);
				cell.setCellValue((String)obj1);
			}
			try {
				FileOutputStream fos = new FileOutputStream(new File(System.getProperty("user.dir") + "\\Excelfile1.xlsx"));
				wb.write(fos);
				fos.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		

	}

}
