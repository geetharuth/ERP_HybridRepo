package driverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver driver;
String inputpath="./FileInput/Controller.xlsx";
String outputpath="./FileOutput/HybridResults.xlsx";
String TCsheet="MasterTestCases";
ExtentReports reports;
ExtentTest logger;
public void startTest() throws Throwable
{
	String Module_status="";
	String Module_new="";
	//create object for excelfileutil class
	ExcelFileUtil xl=new ExcelFileUtil(inputpath);
	//iterate all rows in TCsheet
	for(int i=1;i<=xl.rowcount(TCsheet);i++)
	{
		if(xl.getceldata(TCsheet, i, 2).equalsIgnoreCase("Y"))
		{
			//read modulename cell and store into one variable
			String TCmodule=xl.getceldata(TCsheet, i, 1);
			//define path of html
			reports =new ExtentReports("./target/Reports/"+TCmodule+"----"+FunctionLibrary.generateDate()+".html");
			logger=reports.startTest(TCmodule);
			
			//iterate all rows in TCmodule
			for(int j=1;j<=xl.rowcount(TCmodule);j++)
			{
				//read cells from TCmodule
				String Description=xl.getceldata(TCmodule, j, 0);
				String Objecttype =xl.getceldata(TCmodule, j, 1);
				String Ltype=xl.getceldata(TCmodule, j, 2);
				String Lvalue=xl.getceldata(TCmodule, j, 3);
				String TestData=xl.getceldata(TCmodule, j, 4);
				try{
					if(Objecttype.equalsIgnoreCase("startBrowser"))
					{
					driver=	FunctionLibrary.startBrowser();
					logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("openUrl"))
					{
					FunctionLibrary.openUrl();
					logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("waitforElement"))
					{
						FunctionLibrary.waitForElement(Ltype, Lvalue, TestData);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("typeAction"))
					{
						FunctionLibrary.typeAction(Ltype, Lvalue, TestData);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("clickAction"))
					{
						FunctionLibrary.clickAction(Ltype, Lvalue);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("Validatetitle"))
					{
						FunctionLibrary.Validatetitle(TestData);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("closebrowser"))
					{
						Thread.sleep(2000);
						FunctionLibrary.closebrowser();
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("dropDownAction"))
					{
						FunctionLibrary.dropDownAction(Ltype, Lvalue, TestData);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("capturestock"))
					{
						FunctionLibrary.capturestock(Ltype, Lvalue);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("stocktable"))
					{
						FunctionLibrary.stocktable();
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("capturesupp"))
					{
						FunctionLibrary.capturesupp(Ltype, Lvalue);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("suppliertable"))
					{
						FunctionLibrary.suppliertable();
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("capturecuss"))
					{
						FunctionLibrary.captureCus(Ltype, Lvalue);
						logger.log(LogStatus.INFO,Description);
					}
					if(Objecttype.equalsIgnoreCase("customertable"))
					{
						FunctionLibrary.customerTable();
						logger.log(LogStatus.INFO,Description);
					}
					
					//write as pass into status cell
					xl.setcellData(TCmodule, j, 5, "pass", outputpath);
					logger.log(LogStatus.PASS,Description);
					Module_status="True";
				}catch (Exception e) 
				{
					System.out.println(e.getMessage());
					//write as fail into status cell
					xl.setcellData(TCmodule, j, 5, "Fail", outputpath);
					logger.log(LogStatus.FAIL,Description);
					Module_status="Flase";
				}
				if(Module_status.equalsIgnoreCase("true"))
				{
					//write as pass into TCsheet
					xl.setcellData(TCsheet, i, 3,"pass",outputpath );
				}
				if(Module_new.equalsIgnoreCase("false"))
				{
					//write as Fail into TCsheet
					xl.setcellData(TCsheet, i, 3,"Fail",outputpath );
				}
				reports.endTest(logger);
				reports.flush();
			}
				
		}
		else
		{
			//write as blocked into status cell for Flag N
			xl.setcellData(TCsheet, i, 3, "Blocked", outputpath);
		}
	}
	
}






}
