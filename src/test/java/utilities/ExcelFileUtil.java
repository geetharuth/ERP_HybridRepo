package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {
XSSFWorkbook wb;
//constructor for reading excel path
public ExcelFileUtil(String Excelpath) throws Throwable
{
	FileInputStream fi=new FileInputStream(Excelpath);
	wb= new XSSFWorkbook(fi);
}
//counting no of rows in a sheet
public int rowcount(String sheetname)
{
	 return wb.getSheet(sheetname).getLastRowNum();
}
//method reading cell data
public String getceldata(String sheetname,int row,int column)
{
	String data="";
	if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
	{
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
		data=String.valueOf(celldata);
	}
	else
	{
		data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();	
	}
	return data;
}
	
	//method for write results int new wb
	public void setcellData(String sheetname,int row,int column,String status,String writeExcel) throws Throwable
	{
		//get sheet from wb
		XSSFSheet ws= wb.getSheet(sheetname);
		//get row from sheet
		XSSFRow rownum=ws.getRow(row);
		//create cell in a row
		XSSFCell cell=rownum.createCell(column);
		//write status into cell
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("pass"))
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("fail"))
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo=new FileOutputStream(writeExcel);
		wb.write(fo);
	}
public static void main(String[] args) throws Throwable
{
	ExcelFileUtil xl=new ExcelFileUtil("E:/sample.xlsx");
	int rc=xl.rowcount("emp");
	System.out.println(rc);
	for(int i=1;i<=rc;i++)
	{
		String fname=xl.getceldata("emp", i, 0);
		String mname=xl.getceldata("emp", i, 1);
		String lname=xl.getceldata("emp", i, 2);
		String eid=xl.getceldata("emp", i, 3);
		System.out.println(fname+"  "+mname+"   "+lname+"  "+eid);
	//	xl.setcellData("emp", i, 4, "pass", "E:/sampleresults3.xlsx");
		xl.setcellData("emp", i, 4, "fail", "E:/sampleresults3.xlsx");
		//xl.setcellData("emp", i, 4, "Blocked", "E:/sampleresults3.xlsx");
	}
}
		
	}



