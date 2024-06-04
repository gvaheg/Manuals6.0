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

public class PVHomePage extends HomePageFunctions {

	
	
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		getDate();
	}

	@Test(priority = 1, description = "EN-US PV")
	public void verifyMerckManualPV() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.merckmanuals.com/professional");
			System.out.println("VERSION: EN-US PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 2, description = "EN PV")
	public void verifyMSDManualPVs() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/professional");
			System.out.println("VERSION: EN PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 3, description = "DE PV")
	public void verifyMSDManualPVDE() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/de/profi");
			System.out.println("VERSION: PROD DE PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 4, description = "ES PV")
	public void verifyMSDManualPVES() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/es/professional");
			System.out.println("VERSION: PROD ES PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 5, description = "FR PV")
	public void verifyMSDManualPVFR() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/fr/professional");
			System.out.println("VERSION: PROD FR PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 6, description = "IT PV")
	public void verifyMSDManualPVIT() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/it/professionale");
			System.out.println("VERSION: PROD IT PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 7, description = "JA PV")
	public void verifyMSDManualPVJA() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB");
			System.out.println("VERSION: PROD JA PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 8, description = "KO PV")
	public void verifyMSDManualPVKO() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/en-kr/professional");
			System.out.println("VERSION: PROD KO PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 9, description = "PT PV")
	public void verifyMSDManualPVPT() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/pt/profissional");
			System.out.println("VERSION: PROD PT PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 10, description = "RU PV")
	public void verifyMSDManualPVRU() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/ru/%d0%bf%d1%80%d0%be%d1%84%d0%b5%d1%81%d1%81%d0%b8%d0%be%d0%bd%d0%b0%d0%bb%d1%8c%d0%bd%d1%8b%d0%b9");
			System.out.println("VERSION: PROD RU PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 11, description = "CN PV")
	public void verifyMSDManualPVCN() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.cn/professional");
			System.out.println("VERSION: PROD CN PV");
			Thread.sleep(3000);
			// Close Cookies
			try {
				WebElement AcceptCookies = driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(2000);
			try {
				driver.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[3]/a[1]")).click();
			}catch(Exception e) {
				System.out.println("Can't Close Prompt");
				}
			System.out.println("VERSION: PROD CN PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 12, description = "VI PV")
	public void verifyMSDManualPVVI() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/vi/chuy%c3%aan-gia");
			System.out.println("VERSION: PROD VI PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 13, description = "AR PV")
	public void verifyMSDManualPVAR() {
		try {
			driver = new FirefoxDriver();
			driver.get("n/a");
			System.out.println("VERSION: PROD AR PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 14, description = "UK PV")
	public void verifyMSDManualPVUK() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdmanuals.com/uk/professional");
			System.out.println("VERSION: PROD UK PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 15, description = "MM VET PV")
	public void verifyMMVETManualPV() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.merckvetmanual.com/");
			System.out.println("VERSION: PROD MM VET PV");
			verifyPVHome();
			driver.quit();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 16, description = "MSD VET PV")
	public void verifyMSDVETManualPV() {
		try {
			driver = new FirefoxDriver();
			driver.get("https://uat102.msdvetmanual.com/");
			System.out.println("VERSION: PROD MSD VET PV");
			verifyPVHome();
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
