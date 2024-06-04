package Other;

import java.time.Duration;
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

public class Symptoms_PV extends All_Functions {
	// Before Class
	@BeforeClass
	public void startBrowser() throws Exception {
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wb = new XSSFWorkbook();
		System.out.println("Browser Started!");
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
	}

	@Test(priority = 1, description = "EN-US PV")
	public void EN_US_PV() {
		try {
			wd.get("https://east.merckmanuals.com/professional/resourcespages/symptoms");
			System.out.println("VERSION: PROD EN-US PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 2, description = "PT PV")
	public void PT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/pt/profissional/resourcespages/symptoms");
			System.out.println("VERSION: PROD PT PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
/*
	@Test(priority = 3, description = "JA PV")
	public void JA_PV() {
		try {
			wd.get("");
			System.out.println("VERSION: PROD JA PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}
*/
	@Test(priority = 4, description = "FR PV")
	public void FR_PV() {
		try {
			wd.get("https://east.msdmanuals.com/fr/professional/resourcespages/symptoms");
			System.out.println("VERSION: PROD FR PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 5, description = "ES PV")
	public void ES_PV() {
		try {
			wd.get("https://east.msdmanuals.com/es/professional/resourcespages/symptoms");
			System.out.println("VERSION: PROD ES PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 6, description = "DE PV")
	public void DE_PV() {
		try {
			wd.get("https://east.msdmanuals.com/de/profi/resourcespages/symptoms");
			System.out.println("VERSION: PROD DE PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 7, description = "IT PV")
	public void IT_PV() {
		try {
			wd.get("https://east.msdmanuals.com/it/professionale/resourcespages/symptoms");
			System.out.println("VERSION: PROD IT PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 8, description = "RU PV")
	public void RU_PV() {
		try {
			wd.get("https://east.msdmanuals.com/ru/%d0%bf%d1%80%d0%be%d1%84%d0%b5%d1%81%d1%81%d0%b8%d0%be%d0%bd%d0%b0%d0%bb%d1%8c%d0%bd%d1%8b%d0%b9/resourcespages/symptoms");
			System.out.println("VERSION: PROD RU PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			try {
				wd.findElement(By.xpath(".//div[@class='confirmationpopup__answer']/a[1]")).click();
			} catch (Exception e) {
			}

			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 9, description = "CN PV")
	public void ZH_PV() {
		try {
			wd.get("https://www.msdmanuals.cn/professional/resourcespages/symptoms");
			try {
				wd.findElement(By.xpath("//*[@id=\"access-confirmation-popup\"]/div/div/div/div/div[2]/a[1]")).click();
			} catch (Exception e) {
			}
			System.out.println("VERSION: PROD CN PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 10, description = "VI PV")
	public void AR_CV() {
		try {
			wd.get("https://east.msdmanuals.com/vi/chuy%c3%aan-gia/resourcespages/symptoms");
			System.out.println("VERSION: PROD VI PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
		} catch (Exception e) {
			System.out.println("Page Error!");
		}
	}

	@Test(priority = 11, description = "EN MSD PV")
	public void EN_MSD_PV() {
		try {
			wd.get("https://east.msdmanuals.com/professional/resourcespages/symptoms");
			System.out.println("VERSION: PROD EN MSD PV");
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			verifySymptoms_PV();
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
