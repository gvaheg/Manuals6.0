package TestCases;


import static org.testng.Assert.assertTrue;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Landing extends LandingFunctions {

	
	
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		getDate();
	}

	@Test(priority = 1, description = "EN-US Landing")
	public void verifyMerckManualLanding() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.merckmanuals.com/");
			System.out.println("VERSION: PROD EN-US Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
		
	}
	
	@Test(priority = 2, description = "EN PV Landing")
	public void verifyMSDManualLanding() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/");
			System.out.println("VERSION: PROD EN PV Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 3, description = "DE Landing")
	public void verifyMSDManualLandingDE() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/de");
			System.out.println("VERSION: PROD DE Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 4, description = "ES Landing")
	public void verifyMSDManualLandingES() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/es");
			System.out.println("VERSION: PROD ES Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 5, description = "FR Landing")
	public void verifyMSDManualLandingFR() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/fr");
			System.out.println("VERSION: PROD FR Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 6, description = "IT Landing")
	public void verifyMSDManualLandingIT() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/it");
			System.out.println("VERSION: PROD IT Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 7, description = "JA Landing")
	public void verifyMSDManualLandingJA() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ja-jp");
			System.out.println("VERSION: PROD JA Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 8, description = "KO Landing")
	public void verifyMSDManualLandingKO() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ko");
			System.out.println("VERSION: PROD KO Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 9, description = "PT Landing")
	public void verifyMSDManualLandingPT() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/pt");
			System.out.println("VERSION: PROD PT Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 10, description = "RU Landing")
	public void verifyMSDManualLandingRU() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ru");
			System.out.println("VERSION: PROD RU Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 11, description = "CN Landing")
	public void verifyMSDManualLandingCN() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.cn/");
			System.out.println("VERSION: PROD CN Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 12, description = "VI Landing")
	public void verifyMSDManualLandingVI() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/vi");
			System.out.println("VERSION: PROD VI Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 13, description = "AR Landing")
	public void verifyMSDManualLandingAR() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ar");
			System.out.println("VERSION: PROD AR Landing");
			verifyLandingCases();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 14, description = "UK Landing")
	public void verifyMSDManualLandingUK() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/uk");
			System.out.println("VERSION: PROD UK Landing");
			verifyLandingCases();
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
