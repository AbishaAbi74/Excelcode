package tcs.libertymutual;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Create 2 classes (ExcelCode and TestExcel) under src/main/java in Maven project and paste these codes 

public class ExcelCode {

	public static FileInputStream f;

	public static XSSFWorkbook wb;

	public static XSSFSheet sh;

	

	public static String readStringData(int i, int j)throws IOException{

		f= new FileInputStream("C:\\Users\\abish\\OneDrive\\Documents\\Data.xlsx");

		wb = new XSSFWorkbook(f);

		sh = wb.getSheet("Sheet1");

		XSSFRow r = sh.getRow(i);

		XSSFCell c= r.getCell(j);

		return c.getStringCellValue();	

	}

  public static double readNumericData(int i, int j)throws IOException {

	    f= new FileInputStream("C:\\Users\\abish\\OneDrive\\Documents\\Data.xlsx");

	    wb = new XSSFWorkbook(f);

		sh = wb.getSheet("Sheet1");

		XSSFRow r = sh.getRow(i);

		XSSFCell c= r.getCell(j);

	    return c.getNumericCellValue();

	  

  }

}





