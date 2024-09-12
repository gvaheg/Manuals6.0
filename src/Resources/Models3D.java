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

public class Models3D extends All_Functions {
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
				wd.get("https://www.merckmanuals.com/professional/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD EN-US PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 2, description = "EN-US CV")
		public void EN_US_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.merckmanuals.com/home/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD EN-US CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 3, description = "PT PV")
		public void PT_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/pt/profissional/pages-with-widgets/modelos-3d?mode=list");
				System.out.println("VERSION: PROD PT PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 4, description = "PT CV")
		public void PT_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/pt/casa/pages-with-widgets/modelos-3d?mode=list");
				System.out.println("VERSION: PROD PT CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 5, description = "JA PV")
		public void JA_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD JA PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 6, description = "JA CV")
		public void JA_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ja-jp/home/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD JA CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		
		
		@Test(priority = 7, description = "FR PV")
		public void FR_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/fr/professional/pages-with-widgets/mod%C3%A8les-3d?mode=list");
				System.out.println("VERSION: PROD FR PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 8, description = "FR CV")
		public void FR_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/fr/accueil/pages-with-widgets/mod%C3%A8les-3d?mode=list");
				System.out.println("VERSION: PROD FR CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 9, description = "ES PV")
		public void ES_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/es/professional/pages-with-widgets/modelos-3d?mode=list");
				System.out.println("VERSION: PROD ES PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 10, description = "ES CV")
		public void ES_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/es/hogar/pages-with-widgets/modelos-3d?mode=list");
				System.out.println("VERSION: PROD ES CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 11, description = "DE PV")
		public void DE_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/de/profi/pages-with-widgets/3d-modelle?mode=list");
				System.out.println("VERSION: PROD DE PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 12, description = "DE CV")
		public void DE_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/de/heim/pages-with-widgets/3d-modelle?mode=list");
				System.out.println("VERSION: PROD DE CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 13, description = "IT PV")
		public void IT_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/it/professionale/pages-with-widgets/modelli-3d?mode=list");
				System.out.println("VERSION: PROD IT PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 14, description = "IT CV")
		public void IT_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/it/casa/pages-with-widgets/modelli-3d?mode=list");
				System.out.println("VERSION: PROD IT CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		

		@Test(priority = 15, description = "RU PV")
		public void RU_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ru/professional/pages-with-widgets/3d-%EB%AA%A8%EB%8D%B8?mode=list");
				System.out.println("VERSION: PROD RU PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 16, description = "RU CV")
		public void RU_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ru/home/pages-with-widgets/3d-%D0%BC%D0%BE%D0%B4%D0%B5%D0%BB%D0%B8?mode=list");
				Thread.sleep(2000);
				CloseCookies();
				Thread.sleep(2000);
				try {wd.findElement(By.xpath("//button[@class='ChineseModalPopup_languageSelectorPopupVersionButton__j7M_0']")).click();
				}catch(Exception e) {System.out.println("Can't Close Prompt");}
				System.out.println("VERSION: PROD RU CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
/*
		@Test(priority = 17, description = "CN PV")
		public void ZH_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.cn/professional/pages-with-widgets/3d-models?mode=list");
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
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 18, description = "CN CV")
		public void ZH_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.cn/home/pages-with-widgets/biodigital?mode=list");
				try {
					wd.findElement(By.xpath(".//div/div[2]/div/a[2]")).click();
				} catch (Exception e) {
				}
				System.out.println("VERSION: PROD CN CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
*/
		
		@Test(priority = 19, description = "KO CV")
		public void KO_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ko/home/pages-with-widgets/3d-%EB%AA%A8%EB%8D%B8?mode=list");
				System.out.println("VERSION: PROD KO CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		
		@Test(priority = 20, description = "EN MSD PV")
		public void EN_MSD_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/professional/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD EN MSD PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 21, description = "EN MSD CV")
		public void EN_MSD_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/home/pages-with-widgets/3d-models?mode=list");
				System.out.println("VERSION: PROD EN MSD CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
			}
			
			
		
		@Test(priority = 22, description = "AR CV")
		public void AR_MSD_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/ar/home/pages-with-widgets/%D9%86%D9%85%D8%A7%D8%B0%D8%AC-%D8%AB%D9%84%D8%A7%D8%AB%D9%8A%D8%A9-%D8%A7%D9%84%D8%A3%D8%A8%D8%B9%D8%A7%D8%AF?mode=list");
				System.out.println("VERSION: PROD AR CV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}

		@Test(priority = 23, description = "VI PV")
		public void VI_MSD_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/vi/professional/pages-with-widgets/c%C3%A1c-m%C3%B4-h%C3%ACnh-3d?mode=list");
				System.out.println("VERSION: PROD VI PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		
		@Test(priority = 24, description = "UK PV")
		public void UK_MSD_PV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/uk/professional/pages-with-widgets/%D1%82%D1%80%D0%B8%D0%B2%D0%B8%D0%BC%D1%96%D1%80%D0%BD%D1%96-%D0%BC%D0%BE%D0%B4%D0%B5%D0%BB%D1%96?mode=list");
				System.out.println("VERSION: PROD UK PV");
				verify3DModels();
				wd.close();
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		@Test(priority = 25, description = "HI CV")
		public void HI_CV() {
			try {
				wd = new FirefoxDriver();
				wd.get("https://www.msdmanuals.com/hi/home/pages-with-widgets/biodigital?mode=list");
				System.out.println("VERSION: PROD HI CV");
				verify3DModels();
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
