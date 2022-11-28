package exceloperations;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDatabase {
	public static void main(String[] args) throws SQLException, IOException, FileNotFoundException{
		//Database connection
		Connection con=DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","AnushaMadan@1204");
		Statement stmt=con.createStatement(); //Creating a new ststement
		//Through statement object, we can create Sql queries
		//create a new table in the database 'places'
		//In the double quotes, we can write SQL queries
		String sql="create table places (LOCATION_ID decimal(4,0), STREET_ADDRESS varchar(40),POSTAL_CODE varchar(12),CITY varchar(30),STATE_PROVINCE varchar(25),COUNTRY_ID varchar(2))";
		//Have to make sure that these columns are matching with line 46 and 39-44 lines
		stmt.execute(sql); //Execute the statement against the database
		
		//Excel
		FileInputStream fis=new FileInputStream(".\\datafiles\\locations.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet("Locations Data");//sheet name
		
		int rows=sheet.getLastRowNum();
		
		for(int r=1;r<=rows;r++)
		{
			XSSFRow row=sheet.getRow(r); //storing the row data in
			double locId=row.getCell(0).getNumericCellValue(); //gets first cell in first row
			String streatAdd=row.getCell(1).getStringCellValue();//gets second cell in first row
			String postalCode=row.getCell(2).getStringCellValue();
			String city=row.getCell(3).getStringCellValue();
			String stateProv=row.getCell(4).getStringCellValue();
			String countryId=row.getCell(5).getStringCellValue();
			
			sql="insert into places values('"+locId+"', '"+streatAdd+"', '"+postalCode+"', '"+city+"', '"+stateProv+"', '"+countryId+"')";
			stmt.execute(sql);
			stmt.execute("commit");
		}
		
		
		workbook.close();
		fis.close();
		con.close();
		
		System.out.println("Done!!");
		
	}

}
