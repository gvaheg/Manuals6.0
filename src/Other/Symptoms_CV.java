package Other;

import java.time.Duration;
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

public class Symptoms_CV extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

	}

	@Test(priority = 1, description = "EN-US CV")
	public void EN_US_PV() {
		try {
			wd.get("https://east.merckmanuals.com/home/symptoms");
			System.out.println("VERSION: PROD EN-US CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "PT CV")
	public void PT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/pt/casa/symptoms");
			System.out.println("VERSION: PROD PT CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 3, description = "JA CV")
	public void JA_PV() {
		try {
			wd.get("https://east.msdmanuals.com/ja-jp/%E3%83%9B%E3%83%BC%E3%83%A0/symptoms");
			System.out.println("VERSION: PROD JA CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 4, description = "FR CV")
	public void FR_PV() {
		try {
			wd.get("https://east.msdmanuals.com/fr/accueil/symptoms");
			System.out.println("VERSION: PROD FR CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 5, description = "ES CV")
	public void ES_PV() {
		try {
			wd.get("https://east.msdmanuals.com/es/hogar/symptoms");
			System.out.println("VERSION: PROD ES CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 6, description = "DE CV")
	public void DE_PV() {
		try {
			wd.get("https://east.msdmanuals.com/de/heim/symptoms");
			System.out.println("VERSION: PROD DE CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 7, description = "IT CV")
	public void IT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/it/casa/symptoms");
			System.out.println("VERSION: PROD IT CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 8, description = "RU CV")
	public void RU_PV() {
		try {
			wd.get("https://east.msdmanuals.com/ru/%d0%b4%d0%be%d0%bc%d0%b0/symptoms");
			try {
				wd.findElement(By.xpath(".//div[@class='confirmationpopup__answer']/a[1]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD RU CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 9, description = "ZH CV")
	public void ZH_PV() {
		try {
			wd.get("https://www.msdmanuals.cn/home/symptoms");
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[2]/a[1]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD ZH CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 10, description = "AR CV")
	public void AR_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ar/home/symptoms");
			System.out.println("VERSION: PROD AR CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}


	@Test(priority = 11, description = "EN MSD CV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://east.msdmanuals.com/home/symptoms");
			System.out.println("VERSION: PROD EN MSD CV");
			verifySymptoms_CV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 12, description = "KO CV")
	public void KO_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ko/%ed%99%88/symptoms");
			System.out.println("VERSION: PROD KO CV");
			verifySymptoms_CV();
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
