package Practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;





public class VtigerOperation {

	public static void main(String[] args) throws IOException {

		createSheet("D:\\Snehal\\TestExcel.xlsx", "TextSheet5", 0, 0);

		vtigerLogin("D:\\Snehal\\TestExcel.xlsx","TextSheet5",0 );



	}

	static void createSheet(String fileName,String sheetName,int rowNum,int coloumnNum ) throws IOException
	{
		File file =  new File(fileName);

		FileInputStream input =new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(input);
		
		workbook.createSheet(sheetName);

		XSSFSheet sheet = workbook.getSheet(sheetName);

		sheet.createRow(rowNum);

		XSSFRow row = sheet.getRow(rowNum);

		row.createCell(coloumnNum);

		XSSFCell cell = row.getCell(coloumnNum);

		cell.setCellValue("UserName");
		String username = "admin";
		String password = "admin";


		for(int i =1;i<4;i++)
		{
			sheet.createRow(i);
			row = sheet.getRow(i);
			row.createCell(coloumnNum);
			cell=row.getCell(coloumnNum);
			cell.setCellValue(username);
			username=username+"1";	
		}


		row = sheet.getRow(rowNum);

		row.createCell(coloumnNum+1);
		cell = row.getCell(coloumnNum+1);

		cell.setCellValue("Password");
		for(int i =1;i<4;i++)
		{
			row = sheet.getRow(i);
			row.createCell(coloumnNum+1);
			cell=row.getCell(coloumnNum+1);
			cell.setCellValue(password);
			password=password+"1";	

		}

		row = sheet.getRow(rowNum);
		row.createCell(coloumnNum+2);
		cell = row.getCell(coloumnNum+2);
		cell.setCellValue("Expected Title");
		
		row = sheet.getRow(rowNum);
		row.createCell(coloumnNum+3);
		cell = row.getCell(coloumnNum+3);
		cell.setCellValue("Status");

		row = sheet.getRow(rowNum+1);
		row.createCell(coloumnNum+2);
		cell = row.getCell(coloumnNum+2);
		cell.setCellValue("Dashboard");

		FileOutputStream output = new FileOutputStream(file);
		workbook.write(output);
		workbook.close();

	}

	static void vtigerLogin(String fileName,String sheetName,int coloumnNum ) throws IOException
	{
		File file =  new File(fileName);

		FileInputStream input =new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(input);

		XSSFSheet sheet = workbook.getSheet(sheetName);
		
		System.setProperty("webdriver.chrome.driver", "D:\\WorkSpace_1\\chromedriver.exe");
		WebDriver driver;

		for(int i = 1; i<4 ; i++) {

			XSSFRow row = sheet.getRow(i);


			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			
			driver.get("https://demo.vtiger.com");
			driver.findElement(By.id("username")).sendKeys(row.getCell(coloumnNum).getStringCellValue());

			driver.findElement(By.id("password")).sendKeys(row.getCell(coloumnNum+1).getStringCellValue());


			driver.findElement(By.xpath("//button[text()='Sign in']")).click();

			row = sheet.getRow(1);
			String result;
	

			if(row.getCell(coloumnNum+2).getStringCellValue().equals(driver.getTitle()))
			{
				 result = "pass";
				System.out.println(result);
				row =sheet.getRow(i);
				row.createCell(coloumnNum+3);
			 XSSFCell cell=	row.getCell(coloumnNum+3);
			 cell.setCellValue(result);

			}
			else {
				result= "fail";

				System.out.println(result);
				row =sheet.getRow(i);
				row.createCell(coloumnNum+3);
				 XSSFCell cell=	row.getCell(coloumnNum+3);
				 cell.setCellValue(result);
			}
			
			FileOutputStream output = new FileOutputStream(file);
			workbook.write(output);
		
			
			
			driver.close();
		


		}
		
		workbook.close();




	}
}



