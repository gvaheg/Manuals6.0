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

public class Emergencies extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		wd.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

	}

	@Test(priority = 1, description = "EN-US CV")
	public void EN_US_PV() {
		try {
			wd.get("https://east.merckmanuals.com/home/first-aid");
			System.out.println("VERSION: PROD EN-US CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "PT CV")
	public void PT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/pt/casa/emerg%c3%aancias-e-les%c3%b5es");
			System.out.println("VERSION: PROD PT CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 3, description = "JA CV")
	public void JA_PV() {
		try {
			wd.get("https://east.msdmanuals.com/ja-jp/%E3%83%9B%E3%83%BC%E3%83%A0/%E6%95%91%E5%91%BD%E3%83%BB%E5%BF%9C%E6%80%A5%E6%89%8B%E5%BD%93");
			System.out.println("VERSION: PROD JA CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 4, description = "FR CV")
	public void FR_PV() {
		try {
			wd.get("https://east.msdmanuals.com/fr/accueil/urgences-et-blessures");
			System.out.println("VERSION: PROD FR CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 5, description = "ES CV")
	public void ES_PV() {
		try {
			wd.get("https://east.msdmanuals.com/es/hogar/primeros-auxilios");
			System.out.println("VERSION: PROD ES CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 6, description = "DE CV")
	public void DE_PV() {
		try {
			wd.get("https://east.msdmanuals.com/de/heim/notf%C3%A4lle-und-verletzungen");
			System.out.println("VERSION: PROD DE CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 7, description = "IT CV")
	public void IT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/it/casa/emergenze-e-lesioni");
			System.out.println("VERSION: PROD IT CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 8, description = "RU CV")
	public void RU_PV() {
		try {
			wd.get("https://east.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/first-aid");
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[2]/a[1]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD RU CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 9, description = "CN CV")
	public void ZH_PV() {
		try {
			wd.get("https://www.msdmanuals.cn/home/%E6%80%A5%E7%97%87%E5%92%8C%E6%84%8F%E5%A4%96%E4%BA%8B%E6%95%85");
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[2]/a[1]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD CN CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 10, description = "AR CV")
	public void MM_VET() {
		try {
			wd.get("https://east.msdmanuals.com/ar/home/first-aid");
			System.out.println("VERSION: PROD AR CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}


	@Test(priority = 11, description = "EN MSD CV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://east.msdmanuals.com/home/first-aid");
			System.out.println("VERSION: PROD EN MSD CV");
			verifyEmergencies();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 12, description = "KO CV")
	public void KO_CV() {
		try {
			wd.get("https://www.msdmanuals.com/ko/%ed%99%88/%ec%9d%91%ea%b8%89%ec%83%81%ed%99%a9-%eb%b0%8f-%ec%86%90%ec%83%81");
			System.out.println("VERSION: PROD KO CV");
			verifyEmergencies();
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
