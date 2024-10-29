package Resources;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class ClinicalCalculators extends All_Functions  {
	

	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		getDate();
	}
	
	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.merckmanuals.com/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection&order=bysection");
			System.out.println("VERSION: PROD EN-US PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "PT PV")
	public void PT_PV() {
		
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/pt/profissional/pages-with-widgets/calculadoras-cl%C3%ADnicas?mode=list&order=bysection");
			System.out.println("VERSION: PROD PT PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 3, description = "JA PV")
	public void JA_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/%E5%8C%BB%E5%AD%A6%E8%A8%88%E7%AE%97%E3%83%84%E3%83%BC%E3%83%AB%EF%BC%88%E5%AD%A6%E7%BF%92%E7%94%A8%EF%BC%89?mode=list&order=bysection");
			System.out.println("VERSION: PROD JA PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 4, description = "FR PV")
	public void FR_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/fr/professional/pages-with-widgets/calculateurs-cliniques?mode=list&order=bysection");
			System.out.println("VERSION: PROD FR PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 5, description = "ES PV")
	public void ES_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/es/professional/pages-with-widgets/calculadoras-cl%C3%ADnicas?mode=list&order=bysection");
			System.out.println("VERSION: PROD ES PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 6, description = "DE PV")
	public void DE_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/de/profi/pages-with-widgets/klinische-rechner?mode=list&order=bysection");
			System.out.println("VERSION: PROD DE PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 7, description = "IT PV")
	public void IT_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/it/professionale/pages-with-widgets/calcolatori-clinici?mode=list&order=bysection");
			System.out.println("VERSION: PROD IT PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 8, description = "RU PV")
	public void RU_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/ru/professional/pages-with-widgets/%D0%BA%D0%BB%D0%B8%D0%BD%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D0%B5-%D0%BA%D0%B0%D0%BB%D1%8C%D0%BA%D1%83%D0%BB%D1%8F%D1%82%D0%BE%D1%80%D1%8B?mode=list&order=bysection");
			// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
							AcceptCookies.click();
						} catch (Exception e) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(3000);
			System.out.println("VERSION: PROD RU PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 9, description = "CN PV")
	public void ZH_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.cn/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			Thread.sleep(3000);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(2000);
			try {wd.findElement(By.xpath("//button[@class='ChineseModalPopup_languageSelectorPopupVersionButton__j7M_0']")).click();
			}catch(Exception e) {System.out.println("Can't Close Prompt");}
			System.out.println("VERSION: PROD CN PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 10, description = "MM VET")
	public void MM_VET() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.merckvetmanual.com/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			System.out.println("VERSION: PROD MM VET");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 11, description = "MSD VET")
	public void MSD_VET() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdvetmanual.com/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			System.out.println("VERSION: PROD MSD VET");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 12, description = "EN MSD PV")
	public void EN_MSD_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			System.out.println("VERSION: PROD EN MSD PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 13, description = "VI PV")
	public void VI_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/vi/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			System.out.println("VERSION: PROD VI PV");
			verifyClinicalCalculators();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
		@Test(priority = 14, description = "UK PV")
	public void UK_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/uk/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection");
			System.out.println("VERSION: PROD UK PV");
			verifyClinicalCalculators();
			wd.close();
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

