package Practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateSheet {
	
	public static void main(String[] args) throws IOException {
		
		File file =  new File("D:\\WorkSpace_1\\TestFile.xlsx");
		
	FileInputStream input = new FileInputStream(file);
	
	XSSFWorkbook workbook = new XSSFWorkbook(input);

	
	//workbook.createSheet("TextSheet1");
	XSSFSheet sheet = workbook.getSheet("TextSheet");
	sheet.createRow(2);
	XSSFRow row = sheet.getRow(2);
	System.out.println("Row is created");
	row.createCell(2);
	XSSFCell cell = row.getCell(2);
	cell.setCellValue("Name");
	System.out.println("Cell value is set");
	
	FileOutputStream output=new FileOutputStream(file);
	
	workbook.write(output);
	
	System.out.println("Sheet is created in Workbook");
	
	workbook.close();
	
		
	}

}
