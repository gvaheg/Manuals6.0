package Resources;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class LabTests extends All_Functions {

	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		getDate();
	}
	

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.merckmanuals.com/professional/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD EN-US PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "EN-US CV")
	public void EN_US_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.merckmanuals.com/home/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD EN-US CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 3, description = "PT PV")
	public void PT_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/pt/profissional/pages-with-widgets/an%c3%a1lises-laboratoriais?mode=list");
			System.out.println("VERSION: PROD PT PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 4, description = "PT CV")
	public void PT_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/pt/casa/pages-with-widgets/an%c3%a1lises-laboratoriais?mode=list");
			System.out.println("VERSION: PROD PT CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 5, description = "JA PV")
	public void JA_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/ja-jp/%e3%83%97%e3%83%ad%e3%83%95%e3%82%a7%e3%83%83%e3%82%b7%e3%83%a7%e3%83%8a%e3%83%ab/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD JA PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 6, description = "JA CV")
	public void JA_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/ja-jp/%e3%83%9b%e3%83%bc%e3%83%a0/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD JA CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 7, description = "FR PV")
	public void FR_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/fr/professional/pages-with-widgets/tests-de-labo?mode=list");
			System.out.println("VERSION: PROD FR PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 8, description = "FR CV")
	public void FR_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/fr/accueil/pages-with-widgets/tests-de-labo?mode=list");
			System.out.println("VERSION: PROD FR CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 9, description = "ES PV")
	public void ES_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/es/professional/pages-with-widgets/pruebas-anal%c3%adticas?mode=list");
			System.out.println("VERSION: PROD ES PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 10, description = "ES CV")
	public void ES_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/es/hogar/pages-with-widgets/pruebas-anal%c3%adticas?mode=list");
			System.out.println("VERSION: PROD ES CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 11, description = "DE PV")
	public void DE_PV() {
		try {
			wd = new FirefoxDriver();
			//wd.get("https://east.msdmanuals.com/de/profi/pages-with-widgets/images?mode=list");
			System.out.println("VERSION: PROD DE PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 12, description = "DE CV")
	public void DE_CV() {
		try {
			wd = new FirefoxDriver();
			//wd.get("https://east.msdmanuals.com/de/heim/pages-with-widgets/images?mode=list");
			System.out.println("VERSION: PROD DE CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 13, description = "IT PV")
	public void IT_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/it/professionale/pages-with-widgets/esami-di-laboratorio?mode=list");
			System.out.println("VERSION: PROD IT PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 14, description = "IT CV")
	public void IT_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/it/casa/pages-with-widgets/esami-di-laboratorio?mode=list");
			System.out.println("VERSION: PROD IT CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 15, description = "RU PV")
	public void RU_PV() {
		try {
			wd = new FirefoxDriver();
			//wd.get("https://east.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/pages-with-widgets/images?mode=list");
			try {wd.findElement(By.xpath(".//div/div[2]/div/a[2]")).click();}catch(Exception e) {}
			System.out.println("VERSION: PROD RU PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 16, description = "RU CV")
	public void RU_CV() {
		try {
			wd = new FirefoxDriver();
			//wd.get("https://east.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/pages-with-widgets/images?mode=list");
			// Close Cookies
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
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 17, description = "CN PV")
	public void ZH_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.cn/professional/pages-with-widgets/%e5%ae%9e%e9%aa%8c%e5%ae%a4%e6%a3%80%e6%b5%8b?mode=list");
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
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 18, description = "CN CV")
	public void ZH_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.cn/home/pages-with-widgets/%e5%ae%9e%e9%aa%8c%e5%ae%a4%e6%a3%80%e6%b5%8b?mode=list");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			try {wd.findElement(By.xpath(".//div/div[2]/div/a[2]")).click();}catch(Exception e) {}
			System.out.println("VERSION: PROD CN CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 19, description = "KO CV")
	public void KO_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://www.msdmanuals.com/ko/%ed%99%88/pages-with-widgets/%ec%8b%a4%ed%97%98%ec%8b%a4-%ea%b2%80%ec%82%ac?mode=list");
			System.out.println("VERSION: PROD KO CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	@Test(priority = 20, description = "MM VET")
	public void MM_VET() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.merckvetmanual.com/pages-with-widgets/images?mode=list");
			System.out.println("VERSION: PROD MM VET");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	@Test(priority = 21, description = "MSD VET")
	public void MSD_VET() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.msdvetmanual.com/pages-with-widgets/images?mode=list");
			System.out.println("VERSION: PROD MSD VET");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
	
	@Test(priority = 22, description = "EN PV")
	public void EN_PV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.msdmanuals.com/professional/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD EN PV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 23, description = "EN CV")
	public void EN_CV() {
		try {
			wd = new FirefoxDriver();
			wd.get("https://east.msdmanuals.com/home/pages-with-widgets/lab-tests?mode=list");
			System.out.println("VERSION: PROD EN CV");
			verifyLabTests();
			wd.close();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
	
		@Test(priority = 24, description = "AR MSD CV")
		public void AR_MSD_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ar/home/pages-with-widgets/%d8%a7%d8%ae%d8%aa%d8%a8%d8%a7%d8%b1%d8%a7%d8%aa-%d9%85%d8%ae%d8%aa%d8%a8%d8%b1%d9%8a%d8%a9?mode=list");
				System.out.println("VERSION: PROD EN MSD CV");
				verifyLabTests();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
	}
		
		@Test(priority = 25, description = "VI MSD PV")
		public void VI_MSD_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/vi/chuy%c3%aan-gia/pages-with-widgets/lab-tests?mode=list");
				System.out.println("VERSION: PROD VI MSD PV");
				verifyLabTests();
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
