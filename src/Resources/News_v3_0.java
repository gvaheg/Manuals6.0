package Resources;

import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.io.FileOutputStream;

public class News_v3_0 {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		System.out.println("Browser Started!");
		wd.manage().window().setSize(new Dimension(400, 860));
		wd.manage().window().setPosition(new Point(1200, 0));
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[23];
		String[] Versions = new String[23];

		mainURLs[0] = "http://www.merckmanuals.com/professional/pages-with-widgets/news-list";
		Versions[0] = "ENGLISH PV";
		mainURLs[1] = "http://www.msdmanuals.com/professional/pages-with-widgets/news-list";
		Versions[1] = "MSD ENGLISH PV";
		mainURLs[2] = "http://www.msdmanuals.com/es/professional/pages-with-widgets/noticias-y-comentarios";
		Versions[2] = "Spanish PV";
		mainURLs[3] = "http://www.msdmanuals.com/fr/professional/pages-with-widgets/news-list";
		Versions[3] = "French PV";
		
		mainURLs[4] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/pages-with-widgets/%E3%83%8B%E3%83%A5%E3%83%BC%E3%82%B9";
		Versions[4] = "Japanese PV";
		mainURLs[5] = "http://www.msdmanuals.cn/%E4%B8%93%E4%B8%9A";
		Versions[5] = "Chinese PV";
		mainURLs[6] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/pages-with-widgets/news-list";
		Versions[6] = "Russian PV";
		
		mainURLs[7] = "http://www.msdmanuals.com/pt/profissional/pages-with-widgets/news-list";
		Versions[7] = "Portuguese Professional";
		mainURLs[8] = "http://www.msdmanuals.com/it/professionale/pages-with-widgets/news-list";
		Versions[8] = "Italian PV";
		mainURLs[9] = "http://www.msdmanuals.com/de/profi/pages-with-widgets/news-list";
		Versions[9] = "German PV";
		mainURLs[10] = "http://www.merckmanuals.com/home/pages-with-widgets/news-list";
		Versions[10] = "ENGLISH CV";
		mainURLs[11] = "http://www.msdmanuals.com/home/pages-with-widgets/news-list";
		Versions[11] = "MSD ENGLISH CV";
		mainURLs[12] = "http://www.msdmanuals.com/es/hogar/pages-with-widgets/noticias-y-comentarios";
		Versions[12] = "Spanish CV";
		mainURLs[13] = "http://www.msdmanuals.com/fr/accueil/pages-with-widgets/news-list";
		Versions[13] = "French CV";
		
		mainURLs[14] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/pages-with-widgets/%E3%83%8B%E3%83%A5%E3%83%BC%E3%82%B9";
		Versions[14] = "Japanese CV";
		mainURLs[15] = "http://www.msdmanuals.cn/%E9%A6%96%E9%A1%B5";
		Versions[15] = "Chinese CV";
		mainURLs[16] = "http://www.msdmanuals.com/ko/%ED%99%88/pages-with-widgets/%EC%B5%9C%EC%8B%A0-%EB%89%B4%EC%8A%A4";
		Versions[16] = "Korean CV";
		mainURLs[17] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/pages-with-widgets/news-list";
		Versions[17] = "Russian CV";
		
		mainURLs[18] = "http://www.msdmanuals.com/pt/casa/pages-with-widgets/news-list";
		Versions[18] = "Portuguese Consumer";
		mainURLs[19] = "http://www.msdmanuals.com/it/casa/pages-with-widgets/news-list";
		Versions[19] = "Italian CV";
		mainURLs[20] = "http://www.msdmanuals.com/de/heim/pages-with-widgets/news-list";
		Versions[20] = "German CV";
		mainURLs[21] = "http://www.merckvetmanual.com/pages-with-widgets/news-list";
		Versions[21] = "MM Vet Manual";
		mainURLs[22] = "http://www.msdvetmanual.com/pages-with-widgets/news-list";
		Versions[22] = "MSD Vet Manual";
		
		
		XSSFSheet sheet = wb.createSheet("News");
		// Main Loop
		for (int j = 0; j < 23; j++) {
			Row rowN = sheet.createRow(j);

			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				//Thread.sleep(7000);

				System.out.println(Versions[j]);
				String CurrentVersion = Versions[j];
				rowN.createCell(0).setCellValue(Versions[j]);
				// Handle modal windows
				if (CurrentVersion.contains("Russian CV")) {
					wd.findElement(By.xpath(".//div[2]/div/a[2]")).click();

				} else if (CurrentVersion.contains("Chinese PV")) {
					wd.findElement(By.xpath(".//div[2]/div/a[2]")).click();
				}

				wd.findElement(By.xpath(".//main/section[1]/div[2]/h3/a")).click();
				try {
					WebElement News_Date = wd.findElement(By.xpath(".//div[@class='news__detail--date']"));
					System.out.println(News_Date.getAttribute("innerHTML"));
					rowN.createCell(1).setCellValue(News_Date.getAttribute("innerHTML"));
				} catch (Exception e) {
					rowN.createCell(1).setCellValue("N/A");
				}

			} catch (Exception e) {
				System.out.println("N/A");
				rowN.createCell(1).setCellValue("N/A");
			}
			FileOutputStream fout = new FileOutputStream("C:\\Users\\gvahe\\Desktop\\News.xlsx");
			wb.write(fout);
			fout.close();
			System.out.println("Written Successfully!");

		}

		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
