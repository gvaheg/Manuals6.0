package Resources;

import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class LetterSpine extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");
		wd = new ChromeDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		wd.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

	}

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd.get("https://www.merckmanuals.com/professional");
			System.out.println("VERSION: PROD EN-US PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "PT PV")
	public void PT_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/pt/profissional/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD PT PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 3, description = "JA PV")
	public void JA_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/pages-with-widgets/%E5%8C%BB%E5%AD%A6%E8%A8%88%E7%AE%97%E3%83%84%E3%83%BC%E3%83%AB%EF%BC%88%E5%AD%A6%E7%BF%92%E7%94%A8%EF%BC%89?mode=list");
			System.out.println("VERSION: PROD JA PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 4, description = "FR PV")
	public void FR_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/fr/professional/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD FR PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 5, description = "ES PV")
	public void ES_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/es/professional/pages-with-widgets/calculadoras-cl%C3%ADnicas?mode=list");
			System.out.println("VERSION: PROD ES PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 6, description = "DE PV")
	public void DE_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/de/profi/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD DE PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 7, description = "IT PV")
	public void IT_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/it/professionale/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD IT PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 8, description = "RU PV")
	public void RU_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/pages-with-widgets/clinical-calculators?mode=list");
			try {
				wd.findElement(By.xpath(".//div/div[2]/div/a[2]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD RU PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 9, description = "ZH PV")
	public void ZH_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/pages-with-widgets/clinical-calculators?mode=list");
			try {
				wd.findElement(By.xpath(".//div/div[2]/div/a[2]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD ZH PV");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 10, description = "MM VET")
	public void MM_VET() {
		try {
			wd.get("https://uat93.merckvetmanual.com/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD MM VET");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 11, description = "MSD VET")
	public void MSD_VET() {
		try {
			wd.get("https://uat93.msdvetmanual.com/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD MSD VET");
			verifyLetterSpine();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 12, description = "EN MSD PV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://uat93.msdmanuals.com/professional/pages-with-widgets/clinical-calculators?mode=list");
			System.out.println("VERSION: PROD EN MSD PV");
			verifyLetterSpine();
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
