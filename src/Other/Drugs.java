package Other;

import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Drugs extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

	}

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd.get("https://www.merckmanuals.com/professional");
			System.out.println("VERSION: PROD EN-US PV");
			verifyDrugs();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "EN-US CV")
	public void PT_PV() {
		try {
			wd.get("https://www.merckmanuals.com/home");
			System.out.println("VERSION: EN-US CV");
			verifyDrugs();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@AfterClass
	public void CloseBrowser() throws Exception {
		wd.close();
		System.out.println("================Browser Closed!================");
	}
}
