package TestCases;


import static org.testng.Assert.assertTrue;

import java.time.Duration;
import java.util.List;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Header extends All_Functions {

	
	
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");
		getDate();
	}

	@Test(priority = 1, description = "EN-US PV")
	public void verifyMerckManualLanding() {
		try {
			driver = new ChromeDriver();
			driver.get("https://uat102.merckmanuals.com");
			System.out.println("VERSION: EN-US PV");
			verifyHeader();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 2, description = "EN PV")
	public void verifyMSDManualLandings() {
		try {
			driver = new ChromeDriver();
			driver.get("https://uat102.msdmanuals.com");
			System.out.println("VERSION: EN PV");
			verifyHeader();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@AfterClass
	public void CloseBrowser() throws Exception {
		getDate();
		System.out.println("================Browser Closed!================");
	}
}



