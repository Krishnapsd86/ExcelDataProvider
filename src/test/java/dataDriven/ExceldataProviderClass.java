package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceldataProviderClass {

	public static void main(String[] args) throws IOException {
		        // creating workbook instance to read the file
				// Apache POI --- various to read microsoft
				// Interface classes
				// Work XSSFwork
				// Sheet XSSFSheet
				// Row XSSFRow
				// cell XSSFCell
		
		DataFormatter formatter = new DataFormatter();
		// Create a object of file class to open the file
		File file = new File(System.getProperty("user.dir") + "\\Excelfile.xlsx");
		// Create an object of FileInputStream class to read the file
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rows = sheet.getLastRowNum();
		int col = sheet.getRow(1).getLastCellNum();
		for (int i = 0; i < rows - 1; i++) {
			XSSFRow row = sheet.getRow(i + 1);
			for (int j = 0; j < col; j++) {
				XSSFCell cell = row.getCell(j);
				String data = formatter.formatCellValue(cell);
				System.out.println(data);
			}
			System.out.println(" ");
		}
System.out.println("++++++++++++++++++++++++++++++++++++");
		for (Row row : sheet) {
			for (Cell cell : row) {
				System.out.println(cell);

			}
		}
		System.out.println("");

		

	}

}
