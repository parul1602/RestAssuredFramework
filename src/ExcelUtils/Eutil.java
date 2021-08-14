package ExcelUtils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Eutil {
	public static void main(String argd[]) {
		//getrowcount();
		getcelldata();
		
	}

	public static void getrowcount()
	{
		XSSFWorkbook workbook;
		try {
			
			workbook = new XSSFWorkbook("C:\\Selenium_workspace\\Excelread\\Data\\half.xlsx");
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			int rowcount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of rows" +rowcount);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		
		
	}
	public static void getcelldata() {
		XSSFWorkbook workbook;
		try {
			
			workbook = new XSSFWorkbook("C:\\Selenium_workspace\\Excelread\\Data\\half.xlsx");
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			String celldata =sheet.getRow(0).getCell(0).getStringCellValue();
			System.out.println(celldata);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		
		
		
	}
}
