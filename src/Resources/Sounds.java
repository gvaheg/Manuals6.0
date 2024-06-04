package Resources;

import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Sounds extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		getDate();
		wd.manage().window().setSize(new Dimension(400, 860));
		wd.manage().window().setPosition(new Point(1200, 0));
		//wd.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		
	}

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd.get("https://www.merckmanuals.com/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN-US PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 2, description = "EN-US CV")
	public void EN_US_CV() {
		try {
			wd.get("https://www.merckmanuals.com/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN-US CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 3, description = "PT PV")
	public void PT_PV() {
		try {
			wd.get("https://www.msdmanuals.com/pt/profissional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD PT PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 4, description = "PT CV")
	public void PT_CV() {
		try {
			wd.get("https://www.msdmanuals.com/pt/casa/pages-with-widgets/%C3%A1udio?mode=list");
			System.out.println("VERSION: PROD PT CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 5, description = "JA PV")
	public void JA_PV() {
		try {
			wd.get("https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/%E3%82%AA%E3%83%BC%E3%83%87%E3%82%A3%E3%82%AA?mode=list");
			System.out.println("VERSION: PROD JA PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 6, description = "JA CV")
	public void JA_CV() {
		try {
			wd.get("https://www.msdmanuals.com/ja-jp/home/pages-with-widgets/%E3%82%AA%E3%83%BC%E3%83%87%E3%82%A3%E3%82%AA?mode=list");
			System.out.println("VERSION: PROD JA CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 7, description = "FR PV")
	public void FR_PV() {
		try {
			wd.get("https://www.msdmanuals.com/fr/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD FR PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 8, description = "FR CV")
	public void FR_CV() {
		try {
			wd.get("https://www.msdmanuals.com/fr/accueil/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD FR CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 9, description = "ES PV")
	public void ES_PV() {
		try {
			wd.get("https://www.msdmanuals.com/es/professional/pages-with-widgets/sonar?region=jp&mode=list");
			System.out.println("VERSION: PROD ES PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 10, description = "ES CV")
	public void ES_CV() {
		try {
			wd.get("https://www.msdmanuals.com/es/hogar/pages-with-widgets/sonar?region=&mode=list");
			System.out.println("VERSION: PROD ES CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 11, description = "DE PV")
	public void DE_PV() {
		try {
			wd.get("https://www.msdmanuals.com/de/profi/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD DE PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 12, description = "DE CV")
	public void DE_CV() {
		try {
			wd.get("https://www.msdmanuals.com/de/heim/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD DE CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 13, description = "IT PV")
	public void IT_PV() {
		try {
			wd.get("https://www.msdmanuals.com/it/professionale/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD IT PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 14, description = "IT CV")
	public void IT_CV() {
		try {
			wd.get("https://www.msdmanuals.com/it/casa/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD IT CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 15, description = "RU PV")
	public void RU_PV() {
		try {
			wd.get("https://www.msdmanuals.com/ru/professional/pages-with-widgets/%D0%B0%D1%83%D0%B4%D0%B8%D0%BE?mode=list");
			System.out.println("VERSION: PROD RU PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 16, description = "RU CV")
	public void RU_CV() {
		try {
			wd.get("https://www.msdmanuals.com/ru/home/pages-with-widgets/%D0%B0%D1%83%D0%B4%D0%B8%D0%BE?mode=list");
			Thread.sleep(2000);
			Thread.sleep(2000);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(2000);
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[3]/a[1]")).click();
			}catch(Exception e) {
				System.out.println("Can't Close Prompt");
				}
			System.out.println("VERSION: PROD RU CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 17, description = "CN PV")
	public void CN_PV() {
		try {
			wd.get("https://www.msdmanuals.cn/professional/pages-with-widgets/audio?mode=list");
			Thread.sleep(2000);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(2000);
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[3]/a[1]")).click();
			}catch(Exception e) {
				System.out.println("Can't Close Prompt");
				}
			System.out.println("VERSION: PROD CN PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 18, description = "CN CV")
	public void CN_CV() {
		try {
			wd.get("https://www.msdmanuals.cn/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD CN CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 19, description = "KO CV")
	public void KO_CV() {
		try {
			wd.get("https://www.msdmanuals.com/ko/home/pages-with-widgets/%EC%98%A4%EB%94%94%EC%98%A4?mode=list");
			System.out.println("VERSION: PROD KO CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 20, description = "MM VET")
	public void MM_VET() {
		try {
			wd.get("https://www.merckvetmanual.com/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD MM VET");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 21, description = "MSD VET")
	public void MSD_VET() {
		try {
			wd.get("https://www.msdvetmanual.com/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD MSD VET");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 22, description = "EN MSD PV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://www.msdmanuals.com/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN MSD PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 23, description = "EN MSD CV")
	public void EN_MSD_CV() {
		try {
			wd.get("https://www.msdmanuals.com/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN MSD CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 24, description = "AR MSD CV")
	public void AR_MSD_CV() {
		try {
			wd.get("https://www.msdmanuals.com/ar/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD EN MSD CV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 25, description = "VI PV")
	public void VI_MSD_PV() {
		try {
			wd.get("https://www.msdmanuals.com/vi/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD VI PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	
	@Test(priority = 26, description = "UK PV")
	public void UK_PV() {
		try {
			wd.get("https://www.msdmanuals.com/uk/professional/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD UK PV");
			verifySounds();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 27, description = "HI CV")
	public void HI_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/hi/home/pages-with-widgets/audio?mode=list");
			System.out.println("VERSION: PROD HI CV");
			verifySounds();
			wd.close();
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
