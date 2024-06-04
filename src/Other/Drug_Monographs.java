package Other;

import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Drug_Monographs {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[2];
		String[] Versions = new String[2];

		// English (EN) PV
		mainURLs[0] = "https://west.merckmanuals.com/professional/";
		Versions[0] = "VERSION: English (EN) PV - Drugs";
		mainURLs[1] = "https://west.merckmanuals.com/home/";
		Versions[1] = "VERSION: English (EN) CV - Drugs";

		Thread.sleep(3000);
		// Main Loop
		for (int j = 0; j < 2; j++) {

			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				// Close Cookies
				try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
				// End Close Cookies
				String CurrentVersion = Versions[j];
				Thread.sleep(3000);
				try {
					WebElement linkByText = wd.findElement(By.linkText("DRUG INFO"));
					linkByText.click();
					Thread.sleep(3000);
				} catch (Exception e) {
					System.out.println("Drug Info CANNOT CLICK");
				}
				try {
					WebElement linksByText = wd.findElement(By.linkText("Drugs by Name, Generic and Brand"));
					linksByText.click();
				} catch (Exception e) {
					System.out.println("Drugs by name CANNOT CLICK");
				}
				
				Thread.sleep(3000);
				// Close Cookies
				try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
				// End Close Cookies
				
				int letterCount = wd.findElements(By.xpath("//*[@id=\"maincontent\"]/section/div/div[3]/div/a")).size();
				System.out.println("We have Letters: " + letterCount);
				
				for (int s = 1; s < letterCount + 1; s++) {
					try {
						WebElement Letter = wd.findElement(By.xpath("//*[@id=\"maincontent\"]/section/div/div[3]/div/a["+s+"]"));
						Letter.click();
						System.out.println("Letter: " + Letter.getText());
					} catch (Exception e) {
						System.out.println("Can't Click");
					}
					
					try {
						int rowCount = wd.findElements(By.xpath(".//div[1]/section["+s+"]/div/table/tbody/tr")).size();
						System.out.println("Rows in this section: " + rowCount);
						for (int i = 1; i < rowCount+1; i++) {
							
							WebElement elementToClick = wd.findElement(By.xpath(".//div[1]/section["+s+"]/div/table/tbody/tr["+i+"]/td[1]/div"));
							System.out.println(elementToClick.getText());
							elementToClick.click();
							try {
								wd.findElement(By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
								try {
									WebElement List_Title = wd.findElement(By.xpath("//*[@id=\"section_0\"]/b"));
									System.out.println("PASS");

								} catch (Exception e) {
									try {
										wd.findElement(By.xpath(".//div[@class='monographLinks']/a")).click();
										
										WebElement List_Title = wd.findElement(By.xpath("//*[@id=\"section_0\"]/b"));
										System.out.println("PASS");
										wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']")).click();

									} catch (Exception ez) {
										System.out.println("FAIL");
										wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']")).click();
									}
								}
								try {
									wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']")).click();
									

								} catch (Exception e) {
									System.out.println("Can't close popup");
								}
								
							} catch (Exception e) {
								System.out.println("CANNOT OPEN DRUG, NO URL");
							}
						}
						Thread.sleep(5000);

					} catch (Exception e) {
						System.out.println("ERROR@!");
						Thread.sleep(5000);
					}
					
					
				}
				

				
				
				
				
				
/*
				try {
					
				} catch (Exception e) {
					
				}
*/
				
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		// Close excel and browser
		wd.close();
		System.out.println("Browser Closed!");
	}

}
