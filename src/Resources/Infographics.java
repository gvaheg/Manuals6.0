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

public class Infographics extends All_Functions{
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		wd = new ChromeDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		getDate();
		wd.manage().window().setSize(new Dimension(200, 860));
		wd.manage().window().setPosition(new Point(1000, 0));
		wd.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		
	}

	@Test(priority = 1, description = "EN-US CV")
	public void EN_US_PV() {
		try {
			wd.get("https://east.merckmanuals.com/home/pages-with-widgets/infographics?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD EN-US CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 2, description = "EN MSD CV")
	public void EN_US_CV() {
		try {
			wd.get("https://east.msdmanuals.com/home/pages-with-widgets/infographics?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD EN MSD CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 3, description = "MM VET")
	public void MM_VET() {
		try {
			wd.get("https://east.merckvetmanual.com/pages-with-widgets/infographics?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MM VET");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 4, description = "MSD VET")
	public void MSD_VET() {
		try {
			wd.get("https://east.msdvetmanual.com/pages-with-widgets/infographics?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD VET");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 5, description = "ES CV")
	public void ES_CV() {
		try {
			wd.get("https://east.msdmanuals.com/es/hogar/pages-with-widgets/infograf%c3%ada?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD ES CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 6, description = "DE CV")
	public void DE_CV() {
		try {
			wd.get("https://east.msdmanuals.com/de/heim/pages-with-widgets/infografik?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD DE CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 7, description = "FR CV")
	public void FR_CV() {
		try {
			wd.get("https://east.msdmanuals.com/fr/accueil/pages-with-widgets/infographie?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD FR CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 8, description = "IT CV")
	public void IT_CV() {
		try {
			wd.get("https://east.msdmanuals.com/it/casa/pages-with-widgets/infografica?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD IT CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 9, description = "JA CV")
	public void JA_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ja-jp/%e3%83%9b%e3%83%bc%e3%83%a0/pages-with-widgets/%e5%bd%b9%e7%ab%8b%e3%81%a4%e5%81%a5%e5%ba%b7%e6%83%85%e5%a0%b1?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD JA CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 10, description = "KO CV")
	public void KO_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ko/%ed%99%88/pages-with-widgets/%e6%83%85%e5%a0%b1%e7%94%bb%e5%83%8f?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD KO CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 11, description = "PT CV")
	public void PT_CV() {
		try {
			wd.get("https://east.msdmanuals.com/pt/casa/pages-with-widgets/infogr%c3%a1ficos?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD PT CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 12, description = "RU CV")
	public void RU_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ru/%d0%b4%d0%be%d0%bc%d0%b0/pages-with-widgets/%d0%b8%d0%bd%d1%84%d0%be%d0%b3%d1%80%d0%b0%d1%84%d0%b8%d0%ba%d0%b0?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD RU CV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(3000);
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[2]/a[1]")).click();
			} catch (Exception e) {
			}
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 13, description = "CN CV")
	public void ZH_CV() {
		try {
			wd.get("https://www.msdmanuals.cn/home/pages-with-widgets/%e4%bf%a1%e6%81%af%e5%9b%be?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD CN CV");
			verifyInfographics();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 14, description = "AR CV")
	public void AR_CV() {
		try {
			wd.get("https://east.msdmanuals.com/ar/home/pages-with-widgets/infographics?mode=list&order=alphabetical");
			System.out.println("VERSION: PROD MSD AR CV");
			verifyInfographics();
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
