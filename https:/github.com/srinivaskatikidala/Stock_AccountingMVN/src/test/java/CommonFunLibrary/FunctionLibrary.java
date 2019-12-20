package CommonFunLibrary;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import Utilities.PropertyFileUtil;

public class FunctionLibrary {
	
		static WebDriver driver;

		public static WebDriver startBrowser() throws Exception {

			if (PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("Chrome")) {
				System.setProperty("webdriver.chrome.driver",
						"F:\\Project\\Automation\\Stock_AccountingMVN\\CommonJars\\chromedriver.exe");
				driver = new ChromeDriver();
			} else if (PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("firefox")) {
				System.setProperty("webdriver.gecko.driver", "F:\\Project\\Automation\\Stock_AccountingMVN\\CommonJars\\geckodriver.exe");
				driver = new FirefoxDriver();
			} else if (PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("ie")) {
				System.setProperty("webdriver.ie.driver", "./CommonJars/Internetexplorerdriver.exe");
				driver = new InternetExplorerDriver();
			} else {
				System.out.println("Browser value not matching");
			}

			return driver;

		}

		public static void openApplication(WebDriver driver) throws Exception {
			driver.get(PropertyFileUtil.getValueForKey("Url"));
			driver.manage().window().maximize();
		}

		public static void waitForElement(WebDriver driver, String locatortype, String locatorvalue, String waittitme) {

			WebDriverWait mywait = new WebDriverWait(driver, Integer.parseInt(waittitme));

			if (locatortype.equalsIgnoreCase("id")) {
				mywait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorvalue)));
			} else if (locatortype.equalsIgnoreCase("name")) {
				mywait.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorvalue)));
			} else if (locatortype.equalsIgnoreCase("xpath")) {
				mywait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorvalue)));
			} else {
				System.out.println("unable to locate for waitForElement method");
			}

		}

		public static void typeAction(WebDriver driver, String locatortype, String locatorvalue, String testdata) {
			if (locatortype.equalsIgnoreCase("id")) {
				driver.findElement(By.id(locatorvalue)).clear();
				driver.findElement(By.id(locatorvalue)).sendKeys(testdata);
			} else if (locatortype.equalsIgnoreCase("xpath")) {
				driver.findElement(By.xpath(locatorvalue)).clear();
				driver.findElement(By.xpath(locatorvalue)).sendKeys(testdata);
			} else if (locatortype.equalsIgnoreCase("name")) {
				driver.findElement(By.name(locatorvalue)).clear();
				driver.findElement(By.name(locatorvalue)).sendKeys(testdata);
			} else {
				System.out.println("Locator not matching for typeAction method");
			}

		}

		public static void selectAction(WebElement element, String locatortype, String locatorvalue, String testdata ) {

			WebElement dropdownlist = driver.findElement(By.id(locatorvalue));
			
			Select dlist= new Select(dropdownlist);
		
			if (locatortype.equalsIgnoreCase("id")) {
				driver.findElement(By.id(locatorvalue));
				dlist.selectByVisibleText(testdata);
				
			}else if (locatortype.equalsIgnoreCase("name")) {
				driver.findElement(By.name(locatorvalue));
	            dlist.selectByVisibleText(testdata);
			
			} else {

				System.out.println("Locator not matching for selectAction method");

			}
		}

		public static void clickAction(WebDriver driver, String locatortype, String locatorvalue) {
			if (locatortype.equalsIgnoreCase("id")) {
				driver.findElement(By.id(locatorvalue)).click();
			} else if (locatortype.equalsIgnoreCase("xpath")) {
				driver.findElement(By.xpath(locatorvalue)).click();
			} else if (locatortype.equalsIgnoreCase("name")) {
				driver.findElement(By.name(locatorvalue)).click();
			}
		}

		public static void closeBrowser(WebDriver driver) {
			driver.close();
			
		}
		
		public static void alert(WebDriver driver){
			Alert alert = driver.switchTo().alert();
	       alert.accept();
		System.out.println(alert.getText());
		}
		
		
		
		public static String generateDate() {

			Date date = new Date();

			SimpleDateFormat sdf = new SimpleDateFormat("YYYY_MM_dd_hh_mm");
			return sdf.format(date);

		}

	}


