package Resources;

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

public class Podcasts extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		getDate();
		//wd.manage().window().setSize(new Dimension(400, 860));
		//wd.manage().window().setPosition(new Point(1200, 0));
		wd.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		
	}

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd.get("https://www.merckmanuals.com/professional/resourcespages/podcasts");
			System.out.println("VERSION: PROD EN-US PV");
			verifyPodcasts();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "EN-US CV")
	public void EN_US_CV() {
		try {
			wd.get("https://www.merckmanuals.com/home/resourcespages/podcasts");
			System.out.println("VERSION: PROD EN-US CV");
			verifyPodcasts();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 22, description = "EN MSD PV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://east.msdmanuals.com/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN MSD PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 23, description = "EN MSD CV")
	public void EN_MSD_CV() {
		try {
			wd.get("https://east.msdmanuals.com/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN MSD CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	
	
	@AfterClass
	public void CloseBrowser() throws Exception {
		wd.close();
		getDate();
		System.out.println("================Browser Closed!================");
	}
}
