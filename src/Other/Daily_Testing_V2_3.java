package Other;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.concurrent.ThreadLocalRandom;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;

public class Daily_Testing_V2_3 {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(15));
		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[25];
		String[] Versions = new String[25];
		

		mainURLs[0] = "https://east.merckmanuals.com/professional";
		Versions[0] = "EN US PV";
		mainURLs[1] = "https://east.msdmanuals.com/professional";
		Versions[1] = "EN PV";
		mainURLs[2] = "https://east.msdmanuals.com/es/professional";
		Versions[2] = "ES PV";
		mainURLs[3] = "https://east.msdmanuals.com/fr/professional";
		Versions[3] = "FR PV";
		mainURLs[4] = "https://east.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB";
		Versions[4] = "JA PV";
		mainURLs[5] = "https://east.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A";
		Versions[5] = "ZH PV";
		mainURLs[6] = "https://east.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9";
		Versions[6] = "RU PV";
		mainURLs[7] = "https://east.msdmanuals.com/pt/profissional";
		Versions[7] = "PT PV";
		mainURLs[8] = "https://east.msdmanuals.com/it/professionale";
		Versions[8] = "IT PV";
		mainURLs[9] = "https://east.msdmanuals.com/de/profi";
		Versions[9] = "DE PV";
		mainURLs[10] = "https://east.merckmanuals.com/home";
		Versions[10] = "EN US CV";
		mainURLs[11] = "https://east.msdmanuals.com/home";
		Versions[11] = "EN CV";
		mainURLs[12] = "https://east.msdmanuals.com/es/hogar";
		Versions[12] = "ES CV";
		mainURLs[13] = "https://east.msdmanuals.com/fr/accueil";
		Versions[13] = "FR CV";
		mainURLs[14] = "https://east.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0";
		Versions[14] = "JA CV";
		mainURLs[15] = "https://east.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5";
		Versions[15] = "ZH CV";
		mainURLs[16] = "https://east.msdmanuals.com/ko/%ED%99%88";
		Versions[16] = "KO CV";
		mainURLs[17] = "https://east.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0";
		Versions[17] = "RU CV";
		mainURLs[18] = "https://east.msdmanuals.com/pt/casa";
		Versions[18] = "PT CV";
		mainURLs[19] = "https://east.msdmanuals.com/it/casa";
		Versions[19] = "IT CV";
		mainURLs[20] = "https://east.msdmanuals.com/de/heim";
		Versions[20] = "DE CV";
		mainURLs[21] = "https://east.merckvetmanual.com";
		Versions[21] = "MM Vet";
		mainURLs[22] = "https://east.msdvetmanual.com";
		Versions[22] = "MSD Vet";
		mainURLs[23] = "https://east.msdmanuals.com/ar/home";
		Versions[23] = "AR CV";
		mainURLs[24] = "https://east.msdmanuals.com/vi/chuy%C3%AAn-gia";
		Versions[24] = "VI PV";

		// Main Loop
		for (int j = 0; j < 25; j++) {
			Thread.sleep(3000);
			XSSFSheet sheet = wb.createSheet(Versions[j]);
			Row rowHeading0 = sheet.createRow(0);
			rowHeading0.createCell(0).setCellValue("Test Page");
			rowHeading0.createCell(1).setCellValue("Page Title");
			rowHeading0.createCell(2).setCellValue("Element On Page");
			rowHeading0.createCell(3).setCellValue("Page URL");
			rowHeading0.createCell(4).setCellValue("Status");
			rowHeading0.createCell(5).setCellValue("Comments 1");
			rowHeading0.createCell(6).setCellValue("Comments 2");
			rowHeading0.createCell(7).setCellValue("Comments 3");
			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				String CurrentVersion = Versions[j];
				System.out.println(Versions[j]);
				wd.navigate().refresh();
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies

				// Handle modal windows
				if (CurrentVersion.contains("RU CV")) {
					Thread.sleep(3000);
					try {wd.findElement(By.xpath("//*[text()[contains(.,'Да')]]")).click();}catch(Exception e) {}
					Thread.sleep(1000);

				} else if (CurrentVersion.contains("ZH PV")) {
					Thread.sleep(3000);
					try {wd.findElement(By.xpath("//*[text()[contains(.,'我是专业医学人士')]]")).click();}catch(Exception e) {}
					Thread.sleep(1000);
				}

			} catch (Exception e) {
				System.out.println("N/A");
			}
			

			// START ROW 1 FEATURED CONTENT
			Row row1 = sheet.createRow(1);
			row1.createCell(0).setCellValue("Home Page");
			try {
				try {
					String title = wd.getTitle();
					String currentURL = wd.getCurrentUrl();
					System.out.println("Title is: " + title);
					row1.createCell(1).setCellValue(title);
					System.out.println("URL is: " + currentURL);
					row1.createCell(3).setCellValue(currentURL);
				} catch (Exception e) {
					System.out.println("No Title");
				}
				row1.createCell(2).setCellValue("Featured Article");
				try {
					WebElement FeaturedArticle = wd.findElement(By.xpath(".//div[@class='box__header']/h2"));
					System.out.println(FeaturedArticle.getText());
					row1.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail");
					row1.createCell(4).setCellValue("Fail");
				}
			} catch (Exception e) {
				System.out.println("Home Page Failed");
				row1.createCell(4).setCellValue("Home Page Failed");
			}
			// END ROW 1
			
			//Thread.sleep(3000);

			// START ROW 3 QUIZ ON HOME PAGE
			Row row3 = sheet.createRow(3);
			row3.createCell(0).setCellValue("Home Page");
			try {
				String title = wd.getTitle();
				String currentURL = wd.getCurrentUrl();
				System.out.println("Title is: " + title);
				row3.createCell(1).setCellValue(title);
				System.out.println("URL is: " + currentURL);
				row3.createCell(3).setCellValue(currentURL);
			} catch (Exception e) {
				System.out.println("No Title");
			}
			row3.createCell(2).setCellValue("QUIZ on home page");
			try {
				wd.findElement(By.xpath(".//div/form/ol/li/input")).click();
				wd.findElement(By.xpath("//div/form/input[@class='quiz__submit']")).click();
				try {
					// Close Popup
					Thread.sleep(1000);
					wd.switchTo().defaultContent();
					wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']"))
							.click();
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Popup");
					ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
					wd.switchTo().window(tabs.get(0));
					Thread.sleep(2000);
					wd.switchTo().defaultContent();
					wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']"))
							.click();
				}

				row3.createCell(4).setCellValue("Pass");
			} catch (Exception e) {
				wd.get(mainURLs[j]);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception ez) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				System.out.println("Fail QUIZ");
				row3.createCell(4).setCellValue("Fail");
			}

			// END ROW 3
			
			
			//Thread.sleep(4000);

			// START ROW 4 NEWS AND COMMENTARY
			Row row4 = sheet.createRow(4);
			row4.createCell(0).setCellValue("Home Page");
			try {
				String title = wd.getTitle();
				String currentURL = wd.getCurrentUrl();
				System.out.println("Title is: " + title);
				row4.createCell(1).setCellValue(title);
				System.out.println("URL is: " + currentURL);
				row4.createCell(3).setCellValue(currentURL);
			} catch (Exception e) {
				System.out.println("No Title");
			}
			row4.createCell(2).setCellValue("News on home page");
			try {
				String currentURL = wd.getCurrentUrl();
				WebElement NewsCommentary = wd.findElement(By.xpath(".//div[@class='latestnews__entries']/h3/a[1]"));
				row4.createCell(4).setCellValue("Pass");
				NewsCommentary.click();
				Thread.sleep(2000);
				try {
					WebElement NewsTitle = wd.findElement(By.xpath(".//div[@class='news__detail--header']/h3"));
					row4.createCell(5).setCellValue("Latest News Title: " + NewsTitle.getText());
					WebElement News_Date = wd.findElement(By.xpath(".//div[@class='news__detail--date']"));
					System.out.println(News_Date.getAttribute("innerHTML"));
					row4.createCell(6).setCellValue("Latest News Date: " + News_Date.getAttribute("innerHTML"));

				} catch (Exception e) {
					row4.createCell(6).setCellValue("No Date");
				}
				wd.get(currentURL);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
			} catch (Exception e) {
				System.out.println("Fail");
				row4.createCell(4).setCellValue("Fail");
			}
			System.out.println("GOOD TO GO 1 ............");
			// END ROW 4
			
			
			Thread.sleep(4000);

			// START ROW 6 Calculators on home page
			Row row6 = sheet.createRow(6);
			row6.createCell(0).setCellValue("Home Page");
			try {
				String title = wd.getTitle();
				String currentURL = wd.getCurrentUrl();
				System.out.println("Title is: " + title);
				row6.createCell(1).setCellValue(title);
				System.out.println("URL is: " + currentURL);
				row6.createCell(3).setCellValue(currentURL);
			} catch (Exception e) {
				System.out.println("No Title");
			}
			row6.createCell(2).setCellValue("Calculators on home page");
			try {
				String currentURL = wd.getCurrentUrl();

				try {
					WebElement Calc1 = wd.findElement(By.xpath("//*[@id=\"ui-id-3\"]"));
					Calc1.click();
					System.out.println("Calculator Title1: " + Calc1.getText());
					row6.createCell(5).setCellValue("Calc1: " + Calc1.getText());
					row6.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail Calculator 1");
					row6.createCell(5).setCellValue("FAIL - NO CALCULATORS");
				}
				try {
					WebElement Calc2 = wd.findElement(By.xpath("//*[@id=\"ui-id-4\"]"));
					Calc2.click();
					System.out.println("Calculator Title2: " + Calc2.getText());
					row6.createCell(6).setCellValue("Calc2: " + Calc2.getText());
					row6.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail Calculator 2");
					row6.createCell(6).setCellValue("FAIL - 1 calculator only");
				}
				try {
					WebElement Calc3 = wd.findElement(By.xpath("//*[@id=\"ui-id-5\"]"));
					Calc3.click();
					System.out.println("Calculator Title3: " + Calc3.getText());
					row6.createCell(7).setCellValue("Calc3: " + Calc3.getText());
					row6.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail Calculator 3");
					row6.createCell(7).setCellValue("FAIL - 2 calculators only");
				}
				try {
					WebElement Calc4 = wd.findElement(By.xpath("//*[@id=\"ui-id-6\"]"));
					Calc4.click();
					System.out.println("Calculator Title4: " + Calc4.getText());
					row6.createCell(8).setCellValue("Calc4: " + Calc4.getText());
					row6.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail Calculator 4");
					row6.createCell(8).setCellValue("FAIL - 3 calculators only");
				}
				Thread.sleep(2000);
				wd.get(currentURL);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
			} catch (Exception e) {
				System.out.println("Fail Calculators");
				row6.createCell(4).setCellValue("Fail");
			}
			
			// END ROW 6
			Thread.sleep(4000);

			// START ROW 2 TOPIC PAGE
			Row row2 = sheet.createRow(2);
			row2.createCell(0).setCellValue("Topic Page");
			

			try {
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				
				// Select Random Topic
				try {
					//Thread.sleep(2000);
					try {
						Actions actions = new Actions(wd);
						WebElement target = wd.findElement(By.xpath("//*[@id=\"l-group--header-letterpine\"]/div[1]/div[3]/nav/div[2]"));
						actions.moveToElement(target).perform();
						}catch(Exception ea) {
						System.out.println("Cannot Click on Medical Topics");
					}
					
					try {
						System.out.println("Selecting random topic...");
						List<WebElement> Chapters = wd.findElements(By.xpath("//*[@class='mainmenu__item item1 topics col4 mainmenu__healthtopics']//div[@class='maintab__list--2col']/a"));
						int sizeChapters = Chapters.size();
						System.out.println("Size of Chapters: "+ sizeChapters);
						//Choose Random Chapter
						Random r=new Random();
						int linkNo=r.nextInt(sizeChapters);
						WebElement link=Chapters.get(linkNo);
						System.out.println(link.getText());
						Thread.sleep(2000);
						link.click();
					} catch (Exception e) {
						System.out.println("Can't Click on Chapter");
					}
					Thread.sleep(1000);
					// Close Cookies
					try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
					// End Close Cookies
					try {
						System.out.println("Selecting random topic...");
						List<WebElement> Topics = wd.findElements(By.xpath("//li[@class='medicalsection__section']/h3/a"));
						int sizeTopics = Topics.size();
						System.out.println("Size of Topics: "+ sizeTopics);
						//Choose Random Topic
						Random r=new Random();
						int linkNo=r.nextInt(sizeTopics);
						WebElement link=Topics.get(linkNo);
						System.out.println(link.getText());
						Thread.sleep(2000);
						link.click();
					} catch (Exception e) {
						System.out.println("Can't Click on Topic");
					}
					// Close Cookies
					try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
					// End Close Cookies

				} catch (Exception e) {
					System.out.println("Fail");
					row2.createCell(4).setCellValue("Can't Choose a Topic");
				}

				try {
					String title = wd.getTitle();
					String currentURL = wd.getCurrentUrl();
					System.out.println("Title is: " + title);
					row2.createCell(1).setCellValue(title);
					System.out.println("URL is: " + currentURL);
					row2.createCell(3).setCellValue(currentURL);
				} catch (Exception e) {
					System.out.println("No Title");
				}
				try {
					WebElement TopicTitle = wd
							.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
					System.out.println(TopicTitle.getText() + ": Pass");
					row2.createCell(2).setCellValue(TopicTitle.getText());
					row2.createCell(4).setCellValue("Pass");
				} catch (Exception e) {
					System.out.println("Fail");
					row2.createCell(4).setCellValue("Fail");
				}
			} catch (Exception e) {
				System.out.println("Topic Page Failed");
				row2.createCell(4).setCellValue("Topic Page Failed");
			}

			// START ROW 5 ALSO READ
			Row row5 = sheet.createRow(5);
			row5.createCell(0).setCellValue("Topic Page");
			try {
				String title = wd.getTitle();
				String currentURL = wd.getCurrentUrl();
				System.out.println("Title is: " + title);
				row5.createCell(1).setCellValue(title);
				System.out.println("URL is: " + currentURL);
				row5.createCell(3).setCellValue(currentURL);
			} catch (Exception e) {
				System.out.println("No Title");
			}
			row5.createCell(2).setCellValue("Also Read");
			try {
				WebElement AlsoRead = wd.findElement(By.xpath("//*[@id=\"slick-slide00\"]"));
				System.out.println("Also Read: Pass");
				row5.createCell(4).setCellValue("Pass");
				try {
					WebElement AlsoReadImage1 = wd.findElement(By.xpath("//*[@id=\"slick-slide00\"]/a[1]/img[@src]"));
					System.out.println("Image: Pass");
					row5.createCell(5).setCellValue("Image: Pass");

				} catch (Exception e) {
					System.out.println("Problem with Images");
					row5.createCell(5).setCellValue("Image: Fail");
				}

			} catch (Exception e) {
				System.out.println("Fail");
				row5.createCell(4).setCellValue("Fail");
			}

			// END ROW 5

			// START ROW 7 ALSO OF INTEREST
			Row row7 = sheet.createRow(7);
			row7.createCell(0).setCellValue("Topic Page");
			try {
				String title = wd.getTitle();
				String currentURL = wd.getCurrentUrl();
				System.out.println("Title is: " + title);
				row7.createCell(1).setCellValue(title);
				System.out.println("URL is: " + currentURL);
				row7.createCell(3).setCellValue(currentURL);
			} catch (Exception e) {
				System.out.println("No Title");
			}
			row7.createCell(2).setCellValue("Also Of Interest");
			try {
				WebElement AlsoOfInterest = wd.findElement(By.xpath(".//header/h2[@class='also-of-interest__title']"));
				System.out.println("Also Of Interest: Pass");
				row7.createCell(4).setCellValue("Pass");
			} catch (Exception e) {
				System.out.println("Fail");
				row7.createCell(4).setCellValue("Fail");
			}

			// END ROW 5
			
			
			 DateFormat dateFormat = new SimpleDateFormat("MM.dd.yyyy");
			 
			 //get current date time with Date()
			 Date date = new Date();
			 
			 // Now format the date
			 String date1= dateFormat.format(date);
			 
			 // Print the Date
			System.out.println("Current date and time is " +date1);
			FileOutputStream fout = new FileOutputStream("C:\\users\\gvahe\\Desktop\\Daily_Testing_"+date1+".xlsx");
			wb.write(fout);
			fout.close();
			System.out.println("Written Successfully!");
			Thread.sleep(2000);
		}

		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
