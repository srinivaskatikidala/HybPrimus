package DriveFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunLibrary.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver driver;
	
	  WebElement element;

	  @Test

		public void startTest() throws Throwable {
			
			// creating reference object for excel until methods
			ExcelFileUtil excel = new ExcelFileUtil();
			
			// iterating all row in MasterTestCases sheet
			for (int i = 1; i <= excel.rowCount("MasterTestCases"); i++) {
				String ModuleStatus = "";
				if (excel.getData("MasterTestCases", i, 2).equalsIgnoreCase("Y")) {
				
					// store module name into TCModule
					String TCModule = excel.getData("MasterTestCases", i, 1);

					
	ExtentReports report =	new ExtentReports("F:\\Project\\Automation\\Stock_AccountingMVN\\Reports\\"+TCModule+FunctionLibrary.generateDate()+".html");
					
					// iterate all rows in ApplicationLogin sheet
					for (int j = 1; j <= excel.rowCount(TCModule); j++) {

						
						
						// read all columns in ApplicationLogin testcase
		ExtentTest test=report.startTest(TCModule);
						
						String Description = excel.getData(TCModule, j, 0);
						String Object_Type = excel.getData(TCModule, j, 1);
						String Locator_Type = excel.getData(TCModule, j, 2);
						String Locator_Value = excel.getData(TCModule, j, 3);
						String Test_Data = excel.getData(TCModule, j, 4);
						
						try {
							if (Object_Type.equalsIgnoreCase("startBrowser")) {
								driver = FunctionLibrary.startBrowser();
	test.log(LogStatus.INFO, Description);
							} else if (Object_Type.equalsIgnoreCase("openApplication")) {
								FunctionLibrary.openApplication(driver);
			test.log(LogStatus.INFO, Description);

							} else if (Object_Type.equalsIgnoreCase("typeAction")) {
								FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
			test.log(LogStatus.INFO, Description);

							} else if (Object_Type.equalsIgnoreCase("clickAction")) {
								FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
			test.log(LogStatus.INFO, Description);
			
							} else if (Object_Type.equalsIgnoreCase("selectAction")){
								FunctionLibrary.selectAction(element, Locator_Type, Locator_Value, Test_Data);
		test.log(LogStatus.INFO, Description);					
									

							} else if (Object_Type.equalsIgnoreCase("waitForElement")) {
								FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
				test.log(LogStatus.INFO, Description);

							}
							
							else if(Object_Type.equalsIgnoreCase("alert")){
								FunctionLibrary.alert(driver);
				test.log(LogStatus.INFO, Description);				
								
							}

							else if (Object_Type.equalsIgnoreCase("closeBrowser")) {
								FunctionLibrary.closeBrowser(driver);
			test.log(LogStatus.INFO, Description);

							}

							// write as pass into status column
							excel.SetData(TCModule, j, 5, "PASS");
	test.log(LogStatus.PASS, Description);

							ModuleStatus = "true";
							
						} catch (Exception e) {
							System.out.println("exception handled");
							
							
							File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
							FileUtils.copyFile(srcFile, new File("F:\\Project\\Automation\\Stock_AccountingMVN\\Screens\\"
									+ Description + FunctionLibrary.generateDate() + ".png"));
							excel.SetData(TCModule, j, 5, "FAIL");
		test.log(LogStatus.FAIL, Description);
							
							
							ModuleStatus = "false";
							
							System.out.println("module status is" +ModuleStatus);
							
						}
						if (ModuleStatus.equalsIgnoreCase("true")) {
							excel.SetData("MasterTestCases", i, 3, "PASS");
						
						
						} else {
							if (ModuleStatus.equalsIgnoreCase("False")) {
								excel.SetData("MasterTestCases", i, 3, "FAIL");
							} 
							
							break;
						}

						
		report.flush();
		report.endTest(test);
					
					}
				
				} else {

					excel.SetData("MasterTestCases", i, 3, "Not Executed");

				}
			}
		}

	}


