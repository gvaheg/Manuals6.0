package TESTCHROME;


import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

public class All_Functions2COPY {

	private static final boolean Figures = false;
	// START DRIVER
	static WebDriver wd;
	static XSSFWorkbook wb;
	
	// VERIFY Quizzes
	public void verifyQuizzes() throws Exception {
		Thread.sleep(5000);
		getDate();
		wd.manage().window().setSize(new Dimension(450, 840));
		wd.manage().window().setPosition(new Point(900, 0));
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
		Actions action = new Actions(wd);
		wd.get(currentURL);
		try {
			wait50.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		} catch (Exception e) {
			System.out.println("Refreshing the page!");
			wd.navigate().refresh();
			wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		}
		// Close Cookies
		try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
		Thread.sleep(1000);
		// End Close Cookies
	
		try {
			//Thread.sleep(8000);
			int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
			System.out.println("We have Sections: " + sectionCount);
			int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
			System.out.println("Total Quizzes: " + TotalCount);
			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			Row rowHeading4 = sheet.createRow(4);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			rowHeading4.createCell(0).setCellValue("SECTION");
			rowHeading4.createCell(1).setCellValue("ID");
			rowHeading4.createCell(2).setCellValue("TITLE");
			rowHeading4.createCell(3).setCellValue("LOCATION TITLE");
			rowHeading4.createCell(4).setCellValue("LOCATION (PASS/FAIL)");
			rowHeading4.createCell(5).setCellValue("LOCATION URL");
			rowHeading4.createCell(6).setCellValue("LOCATION URL RESPONSE");
			rowHeading4.createCell(7).setCellValue("CHOICE 1");
			rowHeading4.createCell(8).setCellValue("CHOICE 2");
			rowHeading4.createCell(9).setCellValue("CHOICE 3");
			rowHeading4.createCell(10).setCellValue("CHOICE 4");
			rowHeading4.createCell(11).setCellValue("DE HREF");
			rowHeading4.createCell(12).setCellValue("ES HREF");
			rowHeading4.createCell(13).setCellValue("FR HREF");
			rowHeading4.createCell(14).setCellValue("IT HREF");
			rowHeading4.createCell(15).setCellValue("JA HREF");
			rowHeading4.createCell(16).setCellValue("PT HREF");
			rowHeading4.createCell(17).setCellValue("RU HREF");
			rowHeading4.createCell(18).setCellValue("VI HREF");
			rowHeading4.createCell(19).setCellValue("CN HREF");
			rowHeading4.createCell(20).setCellValue("AR HREF");


			int rowNum = 5;
			for (int s = 2; s < sectionCount + 1; s++) {
				// Close Cookies
				try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
				// End Close Cookies
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					// Main For Loop
					try {
						// Thread.sleep(400);
						Row rowN = sheet.createRow(rowNum++);
						try {
							WebElement Section = wd.findElement(By.xpath(
									".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}

						WebElement List_Title = wd
								.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
						System.out.println("List Title: " + List_Title.getText());
						rowN.createCell(2).setCellValue(List_Title.getAttribute("innerText"));
						WebElement elementToClick = wd
								.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
						try {
							JavascriptExecutor js = (JavascriptExecutor) wd;
							WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
							js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
							js.executeScript("javascript:window.scrollBy(0,-100)");
						}catch (Exception e) {
							System.out.println("Can't SCROLLL!!!");
						}
						elementToClick.click();
						
						// SWITCH TO IFRAME
						
						
						try {
							WebElement iFrameCalc = wd.findElement(By.cssSelector("iframe[class='IFrames_Iframe__WVuGl']"));
							wd.switchTo().frame(iFrameCalc);
						} catch (Exception e) {
							System.out.println("Can't Switch");
						}
						
						// Test Choices
						try {
							String choice1 = wd.findElement(By.xpath(".//*[@id='choices']/li[1]/div[2]")).getText();
							System.out.println("Choice 1: " + choice1);
							rowN.createCell(7).setCellValue(choice1);
						} catch (Exception e) {
							System.out.println("No options");
						}
						try {
							String choice2 = wd.findElement(By.xpath(".//*[@id='choices']/li[2]/div[2]")).getText();
							System.out.println("Choice 2: " + choice2);
							rowN.createCell(8).setCellValue(choice2);
						} catch (Exception e) {
							System.out.println("No options");
						}
						try {
							String choice3 = wd.findElement(By.xpath(".//*[@id='choices']/li[3]/div[2]")).getText();
							System.out.println("Choice 3: " + choice3);
							rowN.createCell(9).setCellValue(choice3);
						} catch (Exception e) {
							System.out.println("No options");
						}
						try {
							String choice4 = wd.findElement(By.xpath(".//*[@id='choices']/li[4]/div[2]")).getText();
							System.out.println("Choice 4: " + choice4);
							rowN.createCell(10).setCellValue(choice4);
						} catch (Exception e) {
							System.out.println("No options");
						}
						// Choose choice
						try {
							wd.findElement(By.xpath(".//*[@id='choices']/li[1]/div[1]/input")).click();
							try {
								wd.findElement(By.xpath(".//*[@id='question_inner']/a")).click();
								
							} catch (Exception et) {
								wd.findElement(By.xpath(".//*[@id='question_inner']/div[4]/a")).click();
							}
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Can't Choose Option");
						}

						// Get location, verify links
						try {
							wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//*[@id='reference_text']//a")));
							WebElement moreinfo = wd.findElement(By.xpath(".//*[@id='reference_text']//a"));
							System.out.println("Location " + i + ": " + moreinfo.getText());
							rowN.createCell(3).setCellValue(moreinfo.getText());
							rowN.createCell(4).setCellValue("Pass");
							String linkUrl = moreinfo.getAttribute("href");
							System.out.println("Location URL: " + linkUrl);
							rowN.createCell(5).setCellValue(linkUrl);

							// VERIFY LINKS ARE ACTIVE
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								try {
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								} catch (Exception e) {
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								}
								// Close Cookies
								try {
										WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception e) {
										System.out.println("Can't Close Cookies");
									}
								// End Close Cookies
								try {
									try {
										wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/h1")));
									} catch (Exception e) {
										System.out.println("Refreshing the page!");
										wd.navigate().refresh();
										wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/h1")));
									}
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[1]/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(6).setCellValue("OK");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(6).setCellValue("Not Found");
								}
								//START HREFs
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
									System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(11).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
									System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(12).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
									System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(13).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
									System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(14).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
									System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(15).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
									System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(16).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
									System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(17).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
									System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(18).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
									System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(19).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
									System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(20).setCellValue("N/A");
								}
								//END HREFs
								
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(6).setCellValue("CANNOT VERIFY");
							}

						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(4).setCellValue("Fail");
						} // END get location, verify links

						try {
							// Close Popup
							Thread.sleep(1000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Cannot Close Popup");
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.switchTo().window(tabs.get(0));
							Thread.sleep(2000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
						}

					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}

					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Quizzes.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}

	// VERIFY SOUNDS
	public void verifySounds() throws Exception {
		ManageWindows();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
		WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
		String currentURL = wd.getCurrentUrl();
		wd.get(currentURL);
		confirmPage(wait20, wait50);
		CloseCookies();
		try {
			int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
			System.out.println("We have Sections: " + sectionCount);
			int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
			System.out.println("Total Sounds: " + TotalCount);
			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			Row rowHeading4 = sheet.createRow(4);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			rowHeading4.createCell(0).setCellValue("SECTION");
			rowHeading4.createCell(1).setCellValue("ID");
			rowHeading4.createCell(2).setCellValue("TITLE");
			rowHeading4.createCell(3).setCellValue("CREDITS");
			rowHeading4.createCell(4).setCellValue("DESCRIPTION");
			rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
			rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
			rowHeading4.createCell(7).setCellValue("LOCATION URL");
			rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");
			rowHeading4.createCell(9).setCellValue("DE HREF");
			rowHeading4.createCell(10).setCellValue("ES HREF");
			rowHeading4.createCell(11).setCellValue("FR HREF");
			rowHeading4.createCell(12).setCellValue("IT HREF");
			rowHeading4.createCell(13).setCellValue("JA HREF");
			rowHeading4.createCell(14).setCellValue("PT HREF");
			rowHeading4.createCell(15).setCellValue("RU HREF");
			rowHeading4.createCell(16).setCellValue("VI HREF");
			rowHeading4.createCell(17).setCellValue("CN HREF");
			rowHeading4.createCell(18).setCellValue("AR HREF");

			int rowNum = 5;

			for (int s = 2; s < sectionCount + 1; s++) {
				// Close Cookies
				try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
				// End Close Cookies
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					// Main For Loop

					try {
						Row rowN = sheet.createRow(rowNum++);
						try {
							WebElement Section = wd.findElement(
									By.xpath(".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}

						try {
							WebElement elementToClick = wd
									.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
							try {
								JavascriptExecutor js = (JavascriptExecutor) wd;
								WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
								js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
								js.executeScript("javascript:window.scrollBy(0,-100)");
							}catch (Exception e) {
								System.out.println("Can't SCROLLL!!!");
							}
							
							elementToClick.click();
						} catch (Exception e) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON ELEMENT");
							try {
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								System.out.println("Can't CLICK ON ELEMENT 2nd time");
							}
						}

						try {
							WebElement Title = wd
									.findElement(By.xpath(".//div[@class='active carousel-item']/div/h3"));
							System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
							// Write ID and Title
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Can't get the title. Reloading...");
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(2000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.close();
								Thread.sleep(3000);
								wd.switchTo().window(tabs.get(1));
								wd.get(currentURL);
								// Close Cookies
								try {
									WebElement AcceptCookies = wd
											.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
									AcceptCookies.click();
								} catch (Exception ez) {
									System.out.println("Can't Close Cookies");
								}
								// End Close Cookies
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								WebElement Title = wd
										.findElement(By.xpath(".//div[@class='active carousel-item']/div/h3"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception ex) {
								System.out.println("Can't get title second time...");
							}

						}
					
						
						
				
						// Show Details
						try {
							WebElement showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
							showDetails.click();
						} catch (Exception ex) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON SHOW DETAILS");
						}
						
						
						// Get & write Description
						try {
							WebElement description = wd.findElement(By.xpath(
									"//div[1]/div[1]/div/ul/div/p"));
							System.out.println("Description: " + description.getAttribute("innerText"));
							rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Description: CANT FIND");
							rowN.createCell(4).setCellValue("");
						}
						// Get & write Credits
						try {
							WebElement credits = wd
									.findElement(By.xpath("//div[1]/div[2]/div/div/p"));
							System.out.println("Credits: " + credits.getAttribute("innerText"));
							rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Credits: CANT FIND");
							rowN.createCell(3).setCellValue("");
						}
						

						// Get location, verify links
						try {
							wait10.until(ExpectedConditions.visibilityOfElementLocated(
									By.xpath("//a[@class='SliderMobile_topic__VrGLg']")));
							WebElement location = wd
									.findElement(By.xpath("//a[@class='SliderMobile_topic__VrGLg']"));
							System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
							// Excel Values
							rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
							rowN.createCell(6).setCellValue("Pass");
							String linkUrl = location.getAttribute("href");
							rowN.createCell(7).setCellValue(linkUrl);

							System.out.println(linkUrl);

							// VERIFY LINKS ARE ACTIVE
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								try {
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								} catch (Exception e) {
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								}
								try {
									// Close Cookies
									try {
										WebElement AcceptCookies = wd
												.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception e) {
										System.out.println("Can't Close Cookies");
									}
									// End Close Cookies
									try {
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									} catch (Exception e) {
										System.out.println("Refreshing the page!");
										wd.navigate().refresh();
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									}
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[1]/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(8).setCellValue("OK");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(8).setCellValue("Not Found");
								}
								// START HREFs
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
									System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(11).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
									System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(12).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
									System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(13).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
									System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(14).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
									System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(15).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
									System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(16).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
									System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(17).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
									System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(18).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
									System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(19).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
									System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(20).setCellValue("N/A");
								}
								// END HREFs
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(8).setCellValue("CANNOT VERIFY");
							}

						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(6).setCellValue("Fail");
						} // END get location, verify links

						
						
						
						try {
							// Close Popup
							//Thread.sleep(1000);
							wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Cannot Close Popup");
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.switchTo().window(tabs.get(0));
							Thread.sleep(2000);
							wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
						}

					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Sounds.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}

	
	// VERIFY CALCULATORS
	public void verifyClinicalCalculators() throws Exception {
		ManageWindows();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
		String currentURL = wd.getCurrentUrl();
		wd.get(currentURL);
		confirmPage(wait20, wait50);
		CloseCookies();
		
			try {
				CloseCookies();
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				// Excel Create Headings
				//String currentLanguage = wd.findElement(By.xpath("//*[@id='currentLanguage']")).getAttribute("value");
				//String currentEdition = wd.findElement(By.xpath("//*[@id='currentEdition']")).getAttribute("value");
				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				Row rowHeading4 = sheet.createRow(4);
				rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
				rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
				rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading4.createCell(0).setCellValue("SECTION");
				rowHeading4.createCell(1).setCellValue("ID");
				rowHeading4.createCell(2).setCellValue("LIST TITLE");
				rowHeading4.createCell(3).setCellValue("CALCULATOR URL");
				rowHeading4.createCell(4).setCellValue("Calculator Error");
				rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(7).setCellValue("LOCATION URL");
				rowHeading4.createCell(8).setCellValue("LOCATION URL (OK/Not Found)");
				rowHeading4.createCell(9).setCellValue("POPUP TITLE");
				rowHeading4.createCell(11).setCellValue("DE HREF");
				rowHeading4.createCell(12).setCellValue("ES HREF");
				rowHeading4.createCell(13).setCellValue("FR HREF");
				rowHeading4.createCell(14).setCellValue("IT HREF");
				rowHeading4.createCell(15).setCellValue("JA HREF");
				rowHeading4.createCell(16).setCellValue("PT HREF");
				rowHeading4.createCell(17).setCellValue("RU HREF");
				rowHeading4.createCell(18).setCellValue("VI HREF");
				rowHeading4.createCell(19).setCellValue("CN HREF");
				rowHeading4.createCell(20).setCellValue("AR HREF");

				int rowNum = 5;
				for (int s = 2; s < sectionCount + 1; s++) {
					int rowCount = wd.findElements(By.xpath(".//main//div[" + s + "]//div/div[2]/ul/li")).size();
					System.out.println("Rows in this section: " + rowCount);

					// Main For Loop
					for (int i = 1; i < rowCount + 1; i++) {
						try {
							Row rowN = sheet.createRow(rowNum++);
							//openResource(rowN, i, s);
							try {
								WebElement Section = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]/div/div/h2"));
								rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
								System.out.println("Section: " + Section.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
							} catch (Exception e) {
								System.out.println("Can't Find Section");
							}
							
							try {	
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a"));
								try {
									JavascriptExecutor js = (JavascriptExecutor) wd;
									WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]"));
									js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
									js.executeScript("javascript:window.scrollBy(0,-100)");
								}catch (Exception e) {
									System.out.println("Can't SCROLLL!!!");
								}
								elementToClick.click();
								WebElement Title = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a/span"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
								
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("Can't CLICK ON ELEMENT");
								try {
									WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
								} catch (Exception ex) {
									Thread.sleep(2000);
									WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
									System.out.println("Can't CLICK ON ELEMENT 2nd time");
								}
							}
						
							
							
							
							/*
							// SWITCH TO IFRAME
							try {
								Thread.sleep(4000);
								WebElement iFrameCalc = wd.findElement(By.xpath("/html/body/iframe[1]"));
								wd.switchTo().frame(iFrameCalc);
							} catch (Exception e) {
								System.out.println("Can't Switch");
							}

							try {
								wait10.until(ExpectedConditions
										.visibilityOfElementLocated(By.xpath("//h3[@class='PopupTable_title__mqaZ0']")));
								WebElement CalcTitle = wd.findElement(By.xpath("//h3[@class='PopupTable_title__mqaZ0']"));
								System.out.println(
										"Popup Title: " + CalcTitle.getAttribute("innerText") + ", Calculator: PASS");
								rowN.createCell(9).setCellValue(CalcTitle.getAttribute("innerText"));
								rowN.createCell(4).setCellValue("OK");

							} catch (Exception e) {
								System.out.println("Calculator: FAIL (ERROR PAGE)");
								rowN.createCell(9).setCellValue("");
								rowN.createCell(4).setCellValue("ERROR PAGE");
							}
							wd.switchTo().defaultContent();

							// GET CALCULATOR URL
							try {
								WebElement CalcLink = wd.findElement(By.xpath("/html/body/iframe[1]"));
								System.out.println("Calculator URL: " + CalcLink.getAttribute("src"));
								rowN.createCell(3).setCellValue(CalcLink.getAttribute("src"));

							} catch (Exception e) {
								System.out.println("Calculator URL: NOT FOUND");
								rowN.createCell(3).setCellValue("NOT FOUND");
							}
							*/
							
							
							
							
							// Show Details
							try {
								WebElement showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
								showDetails.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								System.out.println("Can't CLICK ON SHOW DETAILS");
							}
							try {
								WebElement location = wd.findElement(By.xpath(".//div[2]/a[@class='SliderMobile_topic__VrGLg']"));
								System.out.println("Location " + i + ": " + location.getAttribute("innerText"));

								// Excel Values
								rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
								rowN.createCell(6).setCellValue("Pass");
								String linkUrl = location.getAttribute("href");
								rowN.createCell(7).setCellValue(linkUrl);
								System.out.println(linkUrl);

								// VERIFY LINKS ARE ACTIVE
								try {
									((JavascriptExecutor) wd).executeScript("window.open()");
									Thread.sleep(1000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									try {
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									} catch (Exception e) {
										Thread.sleep(3000);
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									}

									try {
										CloseCookies();
									
										WebElement TopicTitle = wd.findElement(By.xpath(".//div[1]/h1"));
										System.out.println("Location Topic Title: " + TopicTitle.getText());
										rowN.createCell(8).setCellValue("OK");
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Location Topic Title: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									
									// START HREFs
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
										System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(11).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
										System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(12).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
										System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(13).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
										System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(14).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
										System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(15).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
										System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(16).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
										System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(17).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
										System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(18).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
										System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(19).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
										System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(20).setCellValue("N/A");
									}
									// END HREFs
									wd.close();
									wd.switchTo().window(tabs.get(0));

								} catch (Exception e) {
									Thread.sleep(2000);
									System.out.println("URL: CANNOT VERIFY");
									wd.close();
									rowN.createCell(8).setCellValue("CANNOT VERIFY");
								}
								
								

							} catch (Exception e) {
								System.out.println("Location " + i + ": Fail");
								Thread.sleep(3000);
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(6).setCellValue("Fail");
							} // END get location, verify links
							
							
							
							ClosePopup();
							
							
						} catch (Exception e) {
							System.out.println("ERROR! Resource page is not responding. Relopening the page... ");
							ErrorMain(currentURL, e);
						}
						
						
						try {
							FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Calculators.xlsx");
							wb.write(fout);
							fout.close();
							System.out.println("Written Successfully!");
						} catch (Exception e) {
							System.out.println("Cannot SAVE File!");
						}
					}
					
					
					
				}
			} catch (Exception e) {
				System.out.println("ERROR! Web page is not responding. Relopening the page... ");
				ErrorMain(currentURL,e);
			}
		
	}

	
	// VERIFY Infographics
	public void verifyInfographics() throws Exception {
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(40));
		Actions action = new Actions(wd);
		Thread.sleep(10000);
		wd.get(currentURL);
		Thread.sleep(2000);
		// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				
		
		try {
			Thread.sleep(10000);

			int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
			System.out.println("We have Sections: " + sectionCount);
			int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
			System.out.println("Total Infographics: " + TotalCount);

			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			Row rowHeading4 = sheet.createRow(4);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			rowHeading4.createCell(0).setCellValue("SECTION");
			rowHeading4.createCell(1).setCellValue("ID");
			rowHeading4.createCell(2).setCellValue("TITLE");
			rowHeading4.createCell(3).setCellValue("IMAGE");
			rowHeading4.createCell(4).setCellValue("URL");
			rowHeading4.createCell(5).setCellValue("URL RESPONSE");

			int rowNum = 5;
			for (int s = 2; s < sectionCount + 1; s++) {
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {

					// Main For Loop

					try {
						// Thread.sleep(400);
						Row rowN = sheet.createRow(rowNum++);

						try {
							WebElement Section = wd.findElement(By.xpath(
									".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}
						rowN.createCell(1).setCellValue(i);

						WebElement List_Title = wd
								.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
						System.out.println("List Title: " + List_Title.getText());
						rowN.createCell(2).setCellValue(List_Title.getAttribute("innerText"));
						WebElement elementToClick = wd
								.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
						try {
							JavascriptExecutor js = (JavascriptExecutor) wd;
							WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
							js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
							js.executeScript("javascript:window.scrollBy(0,-100)");
						}catch (Exception e) {
							System.out.println("Can't SCROLLL!!!");
						}
						elementToClick.click();

						// Test Image Present
						try {
							Boolean ImagePresent = wd.findElements(By.xpath(".//div/div/div/img")).size() > 0;
							System.out.println("Image " + i + ": " + ImagePresent);
							rowN.createCell(3).setCellValue(ImagePresent);

						} catch (Exception e) {
							rowN.createCell(3).setCellValue("NO IMAGE");
							System.out.println("Image: NO IMAGE");
						}

						// Get URL & Test
						try {
							WebElement LocationLink = wd.findElement(
									By.xpath(".//div[2]/div/a[@class='multimedia__description--citation']"));
							// VERIFY LINKS ARE ACTIVE
							String linkUrl = LocationLink.getAttribute("href");
							System.out.println("URL " + i + ": " + linkUrl);
							try {
								URL url = new URL(linkUrl);
								rowN.createCell(4).setCellValue(linkUrl);
								HttpURLConnection httpURLConnect = (HttpURLConnection) url.openConnection();
								httpURLConnect.setConnectTimeout(2000);
								httpURLConnect.connect();
								if (httpURLConnect.getResponseCode() == 200) {
									System.out.println(
											"URL RESPONSE: " + linkUrl + " - " + httpURLConnect.getResponseMessage());
									// Excel
									rowN.createCell(5).setCellValue(httpURLConnect.getResponseMessage());
								}
								if (httpURLConnect.getResponseCode() == HttpURLConnection.HTTP_NOT_FOUND) {
									System.out.println(
											"URL RESPONSE: " + linkUrl + " - " + httpURLConnect.getResponseMessage()
													+ " - " + HttpURLConnection.HTTP_NOT_FOUND);
									// Excel
									rowN.createCell(5).setCellValue(httpURLConnect.getResponseMessage());
								}
							} catch (Exception e) {
							} // END VERIFY LINKS ACTIVE
						} catch (Exception e) {
							System.out.println("URL " + i + ": NO URL");
							rowN.createCell(4).setCellValue("NO URL");
							rowN.createCell(5).setCellValue("N/A");
						}

						try {
							// Close Popup
							Thread.sleep(1000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Cannot Close Popup");
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.switchTo().window(tabs.get(0));
							Thread.sleep(2000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
						}

					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}

					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Infographics.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}

	
	// VERIFY 3D Models
	public void verify3DModels() throws Exception {
		String currentURL = wd.getCurrentUrl();
		wd.manage().window().setSize(new Dimension(450, 840));
		wd.manage().window().setPosition(new Point(900, 0));
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(20));
		Actions action = new Actions(wd);
		wd.get(currentURL);
		try {
			wait50.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		} catch (Exception e) {
			System.out.println("Refreshing the page!");
			wd.navigate().refresh();
			wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		}
		// Close Cookies
		try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
		Thread.sleep(1000);
		// End Close Cookies
		try {
			int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
			System.out.println("We have Sections: " + sectionCount);
			int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
			System.out.println("Total Infographics: " + TotalCount);

			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			Row rowHeading4 = sheet.createRow(4);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			rowHeading4.createCell(0).setCellValue("SECTION");
			rowHeading4.createCell(1).setCellValue("ID");
			rowHeading4.createCell(2).setCellValue("TITLE");
			rowHeading4.createCell(3).setCellValue("3D MODEL EXIST");
			rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
			rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
			rowHeading4.createCell(7).setCellValue("LOCATION URL");
			rowHeading4.createCell(8).setCellValue("LOCATION URL RESPONSE");
			// rowHeading4.createCell(4).setCellValue("MODEL TITLE");
			// rowHeading4.createCell(5).setCellValue("MODEL DESCRIPTION");
			rowHeading4.createCell(11).setCellValue("DE HREF");
			rowHeading4.createCell(12).setCellValue("ES HREF");
			rowHeading4.createCell(13).setCellValue("FR HREF");
			rowHeading4.createCell(14).setCellValue("IT HREF");
			rowHeading4.createCell(15).setCellValue("JA HREF");
			rowHeading4.createCell(16).setCellValue("PT HREF");
			rowHeading4.createCell(17).setCellValue("RU HREF");
			rowHeading4.createCell(18).setCellValue("VI HREF");
			rowHeading4.createCell(19).setCellValue("CN HREF");
			rowHeading4.createCell(20).setCellValue("AR HREF");
			int rowNum = 5;
			for (int s = 2; s < sectionCount + 1; s++) {
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					// Close Cookies
					try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception e) {
							System.out.println("Can't Close Cookies");
						}
					// End Close Cookies
					// Main For Loop
					try {
						// Thread.sleep(400);
						Row rowN = sheet.createRow(rowNum++);

						try {
							WebElement Section = wd.findElement(
									By.xpath(".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}
						rowN.createCell(1).setCellValue(i);

						WebElement List_Title = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
						System.out.println("List Title: " + List_Title.getText());
						rowN.createCell(2).setCellValue(List_Title.getAttribute("innerText"));
						try {
							WebElement elementToClick = wd
									.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
							try {
								JavascriptExecutor js = (JavascriptExecutor) wd;
								WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
								js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
								js.executeScript("javascript:window.scrollBy(0,-100)");
							}catch (Exception e) {
								System.out.println("Can't SCROLLL!!!");
							}
							
							elementToClick.click();
						} catch (Exception e) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON ELEMENT");
							try {
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								System.out.println("Can't CLICK ON ELEMENT 2nd time");
							}
						}
						
						
						// test 3d model exist
						try {
							WebElement ModelLink = wd
									.findElement(By.xpath(".//div[@class='IFramePopupContent_iframe__gHA1p']/iframe"));
							System.out.println(i + ": 3D Model EXIST");
							rowN.createCell(3).setCellValue("YES");
						} catch (Exception e) {
							System.out.println(i + ": CANNOT FIND 3D MODEL");
							rowN.createCell(3).setCellValue("NO");
						}

						
						// Show Details
						try {
							WebElement showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
							showDetails.click();
						} catch (Exception ex) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON SHOW DETAILS");
						}
						// Get location, verify links
						try {
							wait10.until(ExpectedConditions.visibilityOfElementLocated(
									By.xpath("//a[@class='SliderMobile_topic__VrGLg']")));
							WebElement location = wd
									.findElement(By.xpath("//a[@class='SliderMobile_topic__VrGLg']"));
							System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
							// Excel Values
							rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
							rowN.createCell(6).setCellValue("Pass");
							String linkUrl = location.getAttribute("href");
							rowN.createCell(7).setCellValue(linkUrl);

							System.out.println(linkUrl);

							// VERIFY LINKS ARE ACTIVE
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								try {
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								} catch (Exception e) {
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								}
								try {
									// Close Cookies
									try {
										WebElement AcceptCookies = wd
												.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception e) {
										System.out.println("Can't Close Cookies");
									}
									// End Close Cookies
									try {
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									} catch (Exception e) {
										System.out.println("Refreshing the page!");
										wd.navigate().refresh();
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									}
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[1]/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(8).setCellValue("OK");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(8).setCellValue("Not Found");
								}
								// START HREFs
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
									System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(11).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
									System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(12).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
									System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(13).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
									System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(14).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
									System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(15).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
									System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(16).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
									System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(17).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
									System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(18).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
									System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(19).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
									System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(20).setCellValue("N/A");
								}
								// END HREFs
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(8).setCellValue("CANNOT VERIFY");
							}

						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(6).setCellValue("Fail");
						} // END get location, verify links
						
						
						
						
						
						
						

						try {
							// Close Popup
							Thread.sleep(1000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Cannot Close Popup");
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.switchTo().window(tabs.get(0));
							Thread.sleep(2000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
						}

					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}

					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\3DModels.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}

	// VERIFY FIGURES
	public void verifyFigures() throws Exception {
		ManageWindows();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
		WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
		String currentURL = wd.getCurrentUrl();
		wd.get(currentURL);
		confirmPage(wait20, wait50);
		CloseCookies();
			try {
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading4 = sheet.createRow(4);
				int HeadN = 0;
				rowHeading4.createCell(HeadN++).setCellValue("SECTION");
				rowHeading4.createCell(HeadN++).setCellValue("ID");
				rowHeading4.createCell(HeadN++).setCellValue("TITLE");
				rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
				rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
				rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
				rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
				rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
				rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
				excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
				CloseCookies();
				int rowNum = 5;
				for (int s = 1; s < 26; s++) {
						int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div/div/div[" + s + "]//div/div[2]/ul/li")).size();
						System.out.println("Rows in this section: " + rowCount);
					// Main For Loop
					for (int i = 1; i < rowCount + 1; i++) {
						try {
							Row rowN = sheet.createRow(rowNum++);
							openResource(rowN, i, s);
							getCredits(rowN, i);
							getFileName(rowN, wait10);
							getLocation(rowN, wait10, wait20, wait50, i, rowN, HeadN, i);
							ClosePopup();
						} catch (Exception e) {
							System.out.println("ERROR! Resource page is not responding. Relopening the page... ");
							Thread.sleep(3000);
							((JavascriptExecutor) wd).executeScript("window.open()");
							Thread.sleep(2000);
							ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
							Thread.sleep(2000);
							wd.close();
							wd.switchTo().window(tabs1.get(1));
							Thread.sleep(2000);
							wd.get(currentURL);
							Thread.sleep(5000);
							CloseCookies();
							return;
						}
						try {
							FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Figures.xlsx");
							wb.write(fout);
							fout.close();
							System.out.println("Written Successfully!");
						} catch (Exception e) {
							System.out.println("Cannot SAVE File!");
						}
					}
					
					
				}
				
				
				
				
			} catch (Exception e) {
				System.out.println("ERROR! Web page is not responding. Relopening the page... ");
				ErrorMain(currentURL, e);
			}
		
	}

	
	// VERIFY IMAGES 1
	public void verifyImages1() throws Exception {
		ManageWindows();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
		WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
		String currentURL = wd.getCurrentUrl();
		wd.get(currentURL);
		confirmPage(wait20, wait50);
		CloseCookies();
		int retryCount = 0;
		while (retryCount < 6) {
			try {
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				// Excel Create Headings
				//String currentLanguage = wd.findElement(By.xpath("//*[@id='currentLanguage']")).getAttribute("value");
				//String currentEdition = wd.findElement(By.xpath("//*[@id='currentEdition']")).getAttribute("value");
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading4 = sheet.createRow(4);
				int HeadN = 0;
				rowHeading4.createCell(HeadN++).setCellValue("SECTION");
				rowHeading4.createCell(HeadN++).setCellValue("ID");
				rowHeading4.createCell(HeadN++).setCellValue("TITLE");
				rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
				rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
				rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
				rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
				rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
				rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
				excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);

				int rowNum = 5;
				for (int s = 2; s < sectionCount + 1; s++) {
					// Close Cookies
					try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception e) {
							System.out.println("Can't Close Cookies");
						}
					Thread.sleep(1000);
					// End Close Cookies
					int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div/div/div[" + s + "]//div/div[2]/ul/li")).size();
					System.out.println("Rows in this section: " + rowCount);

					// Main For Loop
					for (int i = 1; i < rowCount + 1; i++) {
						try {
							Row rowN = sheet.createRow(rowNum++);
							openResource(rowN, i, s);
							getCredits(rowN,i);
							//getFileName(rowN, wait10);
							getLocation(rowN, wait10, wait20, wait50, i, rowN, HeadN, i);
							ClosePopup();
						} catch (Exception e) {
							System.out.println("ERROR! Resource page is not responding. Relopening the page... ");
							ErrorMain(currentURL, e);
						}
						try {
							FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Images1.xlsx");
							wb.write(fout);
							fout.close();
							System.out.println("Written Successfully!");
						} catch (Exception e) {
							System.out.println("Cannot SAVE File!");
						}
					}
				}
			} catch (Exception e) {
				System.out.println("ERROR! Web page is not responding. Relopening the page... ");
				wd.quit();
		        wd = new FirefoxDriver();
		        wd.navigate().to(currentURL); // Navigate back to the page
		        retryCount++;
			}
	}
	}
	
	// VERIFY IMAGES 2
	public void verifyImages2() throws Exception {
			ManageWindows();
			WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
			WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
			WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
			WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
			String currentURL = wd.getCurrentUrl();
			wd.get(currentURL);
			confirmPage(wait20, wait50);
			CloseCookies();
			int retryCount = 0;
			while (retryCount < 6) {
				try {
					int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
					System.out.println("We have Sections: " + sectionCount);
					int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
					System.out.println("Total Number: " + TotalCount);
					// Excel Create Headings
					//String currentLanguage = wd.findElement(By.xpath("//*[@id='currentLanguage']")).getAttribute("value");
					//String currentEdition = wd.findElement(By.xpath("//*[@id='currentEdition']")).getAttribute("value");
					XSSFSheet sheet = wb.createSheet();
					Row rowHeading4 = sheet.createRow(4);
					int HeadN = 0;
					rowHeading4.createCell(HeadN++).setCellValue("SECTION");
					rowHeading4.createCell(HeadN++).setCellValue("ID");
					rowHeading4.createCell(HeadN++).setCellValue("TITLE");
					rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
					rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
					rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
					rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
					rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
					rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
					excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);

					int rowNum = 5;
					for (int s = 2; s < sectionCount + 1; s++) {
						// Close Cookies
						try {
								WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
								AcceptCookies.click();
							} catch (Exception e) {
								System.out.println("Can't Close Cookies");
							}
						Thread.sleep(1000);
						// End Close Cookies
						int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div/div/div[" + s + "]//div/div[2]/ul/li")).size();
						System.out.println("Rows in this section: " + rowCount);

						// Main For Loop
						for (int i = 1; i < rowCount + 1; i++) {
							try {
								Row rowN = sheet.createRow(rowNum++);
								openResource(rowN, i, s);
								getCredits(rowN,i);
								//getFileName(rowN, wait10);
								getLocation(rowN, wait10, wait20, wait50, i, rowN, HeadN, i);
								ClosePopup();
							} catch (Exception e) {
								System.out.println("ERROR! Resource page is not responding. Relopening the page... ");
								ErrorMain(currentURL, e);
							}
							try {
								FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Images2.xlsx");
								wb.write(fout);
								fout.close();
								System.out.println("Written Successfully!");
							} catch (Exception e) {
								System.out.println("Cannot SAVE File!");
							}
						}
					}
				} catch (Exception e) {
					System.out.println("ERROR! Web page is not responding. Relopening the page... ");
					wd.quit();
			        wd = new FirefoxDriver();
			        wd.navigate().to(currentURL); // Navigate back to the page
			        retryCount++;
				}
		}
		}
	
	// VERIFY IMAGES 3
	public void verifyImages3() throws Exception {
			ManageWindows();
			WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
			WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
			WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
			WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
			String currentURL = wd.getCurrentUrl();
			wd.get(currentURL);
			confirmPage(wait20, wait50);
			CloseCookies();
			int retryCount = 0;
			while (retryCount < 6) {
				try {
					int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
					System.out.println("We have Sections: " + sectionCount);
					int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
					System.out.println("Total Number: " + TotalCount);
					// Excel Create Headings
					//String currentLanguage = wd.findElement(By.xpath("//*[@id='currentLanguage']")).getAttribute("value");
					//String currentEdition = wd.findElement(By.xpath("//*[@id='currentEdition']")).getAttribute("value");
					XSSFSheet sheet = wb.createSheet();
					Row rowHeading4 = sheet.createRow(4);
					int HeadN = 0;
					rowHeading4.createCell(HeadN++).setCellValue("SECTION");
					rowHeading4.createCell(HeadN++).setCellValue("ID");
					rowHeading4.createCell(HeadN++).setCellValue("TITLE");
					rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
					rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
					rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
					rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
					rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
					rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
					rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
					excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);

					int rowNum = 5;
					for (int s = 2; s < sectionCount + 1; s++) {
						// Close Cookies
						try {
								WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
								AcceptCookies.click();
							} catch (Exception e) {
								System.out.println("Can't Close Cookies");
							}
						Thread.sleep(1000);
						// End Close Cookies
						int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div/div/div[" + s + "]//div/div[2]/ul/li")).size();
						System.out.println("Rows in this section: " + rowCount);

						// Main For Loop
						for (int i = 1; i < rowCount + 1; i++) {
							try {
								Row rowN = sheet.createRow(rowNum++);
								openResource(rowN, i, s);
								getCredits(rowN,i);
								//getFileName(rowN, wait10);
								getLocation(rowN, wait10, wait20, wait50, i, rowN, HeadN, i);
								ClosePopup();
							} catch (Exception e) {
								System.out.println("ERROR! Resource page is not responding. Relopening the page... ");
								ErrorMain(currentURL, e);
							}
							try {
								FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Images3.xlsx");
								wb.write(fout);
								fout.close();
								System.out.println("Written Successfully!");
							} catch (Exception e) {
								System.out.println("Cannot SAVE File!");
							}
						}
					}
				} catch (Exception e) {
					System.out.println("ERROR! Web page is not responding. Relopening the page... ");
					wd.quit();
			        wd = new FirefoxDriver();
			        wd.navigate().to(currentURL); // Navigate back to the page
			        retryCount++;
				}
		}
		}

	// VERIFY VIDEOS
	public void verifyVideos() throws Exception {
		Thread.sleep(2000);
		getDate();
		wd.manage().window().setSize(new Dimension(450, 840));
		wd.manage().window().setPosition(new Point(900, 0));
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
		WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
		WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(20));
		Actions action = new Actions(wd);
		wd.get(currentURL);
		try {
			wait50.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		} catch (Exception e) {
			System.out.println("Refreshing the page!");
			wd.navigate().refresh();
			wait20.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
		}
		// Close Cookies
		try {
			WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
			AcceptCookies.click();
		} catch (Exception e) {
			System.out.println("Can't Close Cookies");
		}
		Thread.sleep(1000);
		// End Close Cookies
		try {
			int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
			System.out.println("We have Sections: " + sectionCount);
			int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
			System.out.println("Total Number: " + TotalCount);
			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			Row rowHeading4 = sheet.createRow(4);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			rowHeading4.createCell(0).setCellValue("SECTION");
			rowHeading4.createCell(1).setCellValue("ID");
			rowHeading4.createCell(2).setCellValue("TITLE");
			rowHeading4.createCell(3).setCellValue("CREDITS");
			rowHeading4.createCell(4).setCellValue("DESCRIPTION");
			rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
			rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
			rowHeading4.createCell(7).setCellValue("LOCATION URL");
			rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");
			rowHeading4.createCell(11).setCellValue("DE HREF");
			rowHeading4.createCell(12).setCellValue("ES HREF");
			rowHeading4.createCell(13).setCellValue("FR HREF");
			rowHeading4.createCell(14).setCellValue("IT HREF");
			rowHeading4.createCell(15).setCellValue("JA HREF");
			rowHeading4.createCell(16).setCellValue("PT HREF");
			rowHeading4.createCell(17).setCellValue("RU HREF");
			rowHeading4.createCell(18).setCellValue("VI HREF");
			rowHeading4.createCell(19).setCellValue("CN HREF");
			rowHeading4.createCell(20).setCellValue("AR HREF");
			int rowNum = 5;

			for (int s = 2; s < sectionCount + 1; s++) {
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					
					// Main For Loop
					try {
						Row rowN = sheet.createRow(rowNum++);

						try {
							WebElement Section = wd.findElement(
									By.xpath(".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}

						try {
							WebElement elementToClick = wd
									.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
							try {
								JavascriptExecutor js = (JavascriptExecutor) wd;
								WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
								js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
								js.executeScript("javascript:window.scrollBy(0,-100)");
							}catch (Exception e) {
								System.out.println("Can't SCROLLL!!!");
							}
							
							elementToClick.click();
						} catch (Exception e) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON ELEMENT");
							try {
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								System.out.println("Can't CLICK ON ELEMENT 2nd time");
							}
						}

						try {
							WebElement Title = wd
									.findElement(By.xpath(".//div[@class='active carousel-item']/div/h3"));
							System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
							// Write ID and Title
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Can't get the title. Reloading...");
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(2000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.close();
								Thread.sleep(3000);
								wd.switchTo().window(tabs.get(1));
								wd.get(currentURL);
								// Close Cookies
								try {
									WebElement AcceptCookies = wd
											.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
									AcceptCookies.click();
								} catch (Exception ez) {
									System.out.println("Can't Close Cookies");
								}
								// End Close Cookies
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								WebElement Title = wd
										.findElement(By.xpath(".//div[@class='active carousel-item']/div/h3"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception ex) {
								System.out.println("Can't get title second time...");
							}

						}
						// Show Details
						try {
							WebElement showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
							showDetails.click();
						} catch (Exception ex) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON SHOW DETAILS");
						}
						// Get & write Description
						try {
							WebElement description = wd.findElement(By.xpath(
									".//div[1]/div/div//p"));
							System.out.println("Description: " + description.getAttribute("innerText"));
							rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Description: CANT FIND");
							rowN.createCell(4).setCellValue("");
						}
						// Get & write Credits
						try {
							WebElement credits = wd
									.findElement(By.xpath("//div[1]/div[2]/div//p"));
							System.out.println("Credits: " + credits.getAttribute("innerText"));
							rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Credits: CANT FIND");
							rowN.createCell(3).setCellValue("");
						}

						// Get location, verify links
						try {
							wait10.until(ExpectedConditions.visibilityOfElementLocated(
									By.xpath("//a[@class='SliderMobile_topic__VrGLg']")));
							WebElement location = wd
									.findElement(By.xpath("//a[@class='SliderMobile_topic__VrGLg']"));
							System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
							// Excel Values
							rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
							rowN.createCell(6).setCellValue("Pass");
							String linkUrl = location.getAttribute("href");
							rowN.createCell(7).setCellValue(linkUrl);

							System.out.println(linkUrl);

							// VERIFY LINKS ARE ACTIVE
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								try {
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								} catch (Exception e) {
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								}
								try {
									// Close Cookies
									try {
										WebElement AcceptCookies = wd
												.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception e) {
										System.out.println("Can't Close Cookies");
									}
									// End Close Cookies
									try {
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									} catch (Exception e) {
										System.out.println("Refreshing the page!");
										wd.navigate().refresh();
										wait20.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//div[1]/h1")));
									}
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[1]/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(8).setCellValue("OK");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(8).setCellValue("Not Found");
								}
								// START HREFs
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
									System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(11).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
									System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(12).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
									System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(13).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
									System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(14).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
									System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(15).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
									System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(16).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
									System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(17).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
									System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(18).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
									System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(19).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
									System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(20).setCellValue("N/A");
								}
								// END HREFs
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(8).setCellValue("CANNOT VERIFY");
							}

						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(6).setCellValue("Fail");
						} // END get location, verify links

						
						try {
							// Close Popup
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
						} catch (Exception e) {
							System.out.println("Cannot Close Popup!");
							Thread.sleep(2000);
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.switchTo().window(tabs.get(0));
							Thread.sleep(2000);
							wd.switchTo().defaultContent();
							wd.findElement(By
									.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
									.click();
						}

					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd
									.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
						
					}
					
					
					
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Videos.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");

				}
				
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}
	
	
	// VERIFY VIDEO IDs
		public void verifyVideosIDs() throws Exception {
			Thread.sleep(2000);
			getDate();
			wd.manage().window().setSize(new Dimension(450, 840));
			wd.manage().window().setPosition(new Point(900, 0));
			wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
			String currentURL = wd.getCurrentUrl();
			WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
			WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
			WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(20));
			Actions action = new Actions(wd);
			wd.get(currentURL);
			try {
				wait50.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			} catch (Exception e) {
				System.out.println("Refreshing the page!");
				wd.navigate().refresh();
				wait20.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			}
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			Thread.sleep(1000);
			// End Close Cookies
			try {
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				Row rowHeading4 = sheet.createRow(4);
				rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
				rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
				rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading4.createCell(0).setCellValue("SECTION");
				rowHeading4.createCell(1).setCellValue("ID");
				rowHeading4.createCell(2).setCellValue("TITLE");
				rowHeading4.createCell(3).setCellValue("CREDITS");
				rowHeading4.createCell(4).setCellValue("DESCRIPTION");
				rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(7).setCellValue("LOCATION URL");
				rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");
				rowHeading4.createCell(11).setCellValue("DE HREF");
				rowHeading4.createCell(12).setCellValue("ES HREF");
				rowHeading4.createCell(13).setCellValue("FR HREF");
				rowHeading4.createCell(14).setCellValue("IT HREF");
				rowHeading4.createCell(15).setCellValue("JA HREF");
				rowHeading4.createCell(16).setCellValue("PT HREF");
				rowHeading4.createCell(17).setCellValue("RU HREF");
				rowHeading4.createCell(18).setCellValue("VI HREF");
				rowHeading4.createCell(19).setCellValue("CN HREF");
				rowHeading4.createCell(20).setCellValue("AR HREF");
				int rowNum = 5;

				for (int s = 2; s < sectionCount + 1; s++) {
					// Close Cookies
					try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
					// End Close Cookies
					int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
					for (int i = 1; i < rowCount + 1; i++) {
						// Main For Loop
						try {
							Row rowN = sheet.createRow(rowNum++);

							try {
								WebElement Section = wd.findElement(
										By.xpath(".//main/div[1]/div/div[" + s + "]/div/div/h2"));
								rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
								System.out.println("Section: " + Section.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
							} catch (Exception e) {
								System.out.println("Can't Find Section");
							}

							try {
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								try {
									JavascriptExecutor js = (JavascriptExecutor) wd;
									WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
									js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
									js.executeScript("javascript:window.scrollBy(0,-100)");
								}catch (Exception e) {
									System.out.println("Can't SCROLLL!!!");
								}
								
								elementToClick.click();
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("Can't CLICK ON ELEMENT");
								try {
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
								} catch (Exception ex) {
									Thread.sleep(2000);
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
									System.out.println("Can't CLICK ON ELEMENT 2nd time");
								}
							}

							try {
								WebElement Title = wd
										.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								// Write ID and Title
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception e) {
								System.out.println("Can't get the title. Reloading...");
								try {
									((JavascriptExecutor) wd).executeScript("window.open()");
									Thread.sleep(2000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									wd.close();
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(currentURL);
									// Close Cookies
									try {
										WebElement AcceptCookies = wd
												.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception ez) {
										System.out.println("Can't Close Cookies");
									}
									// End Close Cookies
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
									WebElement Title = wd
											.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
									System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
									rowN.createCell(1).setCellValue(i);
									rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
								} catch (Exception ex) {
									System.out.println("Can't get title second time...");
								}

							}
							
							
							try {
								System.out.println("WRITE HERE!");
							} catch (Exception e) {
								
								
								System.out.println("BIG ERRORRRR!!!");
								
							}
							
							
							
							

					
							try {
								// Close Popup
								wd.switchTo().defaultContent();
								wd.findElement(By
										.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
										.click();
							} catch (Exception e) {
								System.out.println("Cannot Close Popup!");
								Thread.sleep(2000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.switchTo().window(tabs.get(0));
								Thread.sleep(2000);
								wd.switchTo().defaultContent();
								wd.findElement(By
										.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
										.click();
							}

						} catch (Exception e) {
							System.out.println("ERROR! Web page is not responding. Relopening the page... ");
							((JavascriptExecutor) wd).executeScript("window.open()");
							Thread.sleep(1000);
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.close();
							Thread.sleep(2000);
							wd.switchTo().window(tabs.get(1));
							wd.get(currentURL);
							// Close Cookies
							try {
								WebElement AcceptCookies = wd
										.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
								AcceptCookies.click();
							} catch (Exception ez) {
								System.out.println("Can't Close Cookies");
							}
							// End Close Cookies
							Thread.sleep(10000);
						}
						FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Videos.xlsx");
						wb.write(fout);
						fout.close();
						System.out.println("Written Successfully!");

					}
				}

			} catch (Exception e) {
				System.out.println("Page Error? What is going on?");
				((JavascriptExecutor) wd).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
				wd.close();
				Thread.sleep(2000);
				wd.switchTo().window(tabs.get(1));
				wd.get(currentURL);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception ez) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				Thread.sleep(10000);
			}
		}
	
	
	// VERIFY LAB TESTS
	public void verifyLabTests() throws Exception {
		Thread.sleep(10000);
		getDate();
		wd.manage().window().setSize(new Dimension(450, 840));
		wd.manage().window().setPosition(new Point(900, 0));
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(40));
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(40));
		Actions action = new Actions(wd);
		wd.get(currentURL);
		// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
		try {
			Thread.sleep(7000);
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				Row rowHeading4 = sheet.createRow(4);
				rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
				rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
				rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading4.createCell(0).setCellValue("SECTION");
				rowHeading4.createCell(1).setCellValue("ID");
				rowHeading4.createCell(2).setCellValue("TITLE");
				rowHeading4.createCell(3).setCellValue("CREDITS");
				rowHeading4.createCell(4).setCellValue("DESCRIPTION");
				rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(7).setCellValue("LOCATION URL");
				rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");
				rowHeading4.createCell(11).setCellValue("DE HREF");
				rowHeading4.createCell(12).setCellValue("ES HREF");
				rowHeading4.createCell(13).setCellValue("FR HREF");
				rowHeading4.createCell(14).setCellValue("IT HREF");
				rowHeading4.createCell(15).setCellValue("JA HREF");
				rowHeading4.createCell(16).setCellValue("PT HREF");
				rowHeading4.createCell(17).setCellValue("RU HREF");
				rowHeading4.createCell(18).setCellValue("VI HREF");
				rowHeading4.createCell(19).setCellValue("CN HREF");
				rowHeading4.createCell(20).setCellValue("AR HREF");
				int rowNum = 5;
				

			for (int s = 2; s < sectionCount + 1; s++) {
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					// Main For Loop
					try {
						Row rowN = sheet.createRow(rowNum++);

						try {
							WebElement Section = wd.findElement(By.xpath(
									".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}
						
						try {
							WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
							elementToClick.click();
						} catch (Exception e) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON ELEMENT");
							try {
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								System.out.println("Can't CLICK ON ELEMENT 2nd time");
							}
						}
						
						int size = wd.findElements(By.tagName("iframe")).size();
						System.out.println("Total Number of iFrames: "+size);
						Thread.sleep(2000);
						if (size>5) {
							wd.switchTo().frame(5);
						}else {wd.switchTo().frame(4);}
						
						Thread.sleep(1000);
						try {
							WebElement Title = wd.findElement(By.xpath("//*[@id=\"block-pagetitle\"]/div/h1/div/div/div"));
							System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
							// Write ID and Title
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Can't get the title. Reloading...");
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.close();
								Thread.sleep(2000);
								wd.switchTo().window(tabs.get(1));
								wd.get(currentURL);
								// Close Cookies
								try {
									WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
									AcceptCookies.click();
								} catch (Exception ez) {
									System.out.println("Can't Close Cookies");
								}
								// End Close Cookies
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								WebElement Title = wd.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception ex) {
								System.out.println("Can't get title second time...");
							}
							
						}
						wd.switchTo().defaultContent();
						
						Thread.sleep(2000);


						// Get location, verify links
						try {
							wait.until(ExpectedConditions.visibilityOfElementLocated(
									By.xpath(".//div[3]/div[2]/a")));
							WebElement location = wd
									.findElement(By.xpath(".//div[3]/div[2]/a"));
							System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
							// Excel Values
							rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
							rowN.createCell(6).setCellValue("Pass");
							String linkUrl = location.getAttribute("href");
							rowN.createCell(7).setCellValue(linkUrl);

							System.out.println(linkUrl);

							// VERIFY LINKS ARE ACTIVE
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								try {
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								} catch (Exception e) {
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(linkUrl);
								}
								try {
									wait.until(ExpectedConditions.visibilityOfElementLocated(
											By.xpath(".//div[1]/h1")));
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[1]/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(8).setCellValue("OK");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(8).setCellValue("Not Found");
								}
								//START HREFs
								Thread.sleep(1000);
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
									System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(11).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
									System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(12).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
									System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(13).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
									System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(14).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
									System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(15).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
									System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(16).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
									System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(17).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
									System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(18).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
									System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(19).setCellValue("N/A");
								}
								try {
									WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
									System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
									rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
								} catch (Exception e) {
									System.out.println("Couldn't get HREF");
									rowN.createCell(20).setCellValue("N/A");
								}
								
								//END HREFs
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(8).setCellValue("CANNOT VERIFY");
							}

						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(6).setCellValue("Fail");
						} // END get location, verify links
						
						
							try {
								// Close Popup
								wd.switchTo().defaultContent();
								wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
							} catch (Exception e) {
									System.out.println("Cannot Close Popup!");
									Thread.sleep(2000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									wd.switchTo().window(tabs.get(0));
									Thread.sleep(2000);
									wd.switchTo().defaultContent();
									wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
							}

						
							
					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\LabTests.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
					
					
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}
	
	// VERIFY TABLES
		public void verifyTables() throws Exception {
			Thread.sleep(3000);
			getDate();
			wd.manage().window().setSize(new Dimension(450, 840));
			wd.manage().window().setPosition(new Point(900, 0));
			wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
			String currentURL = wd.getCurrentUrl();
			WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
			WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
			WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(20));
			Actions action = new Actions(wd);
			wd.get(currentURL);
			try {
				wait50.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			} catch (Exception e) {
				System.out.println("Refreshing the page!");
				wd.navigate().refresh();
				wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			}
			// Close Cookies
			try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
			Thread.sleep(1000);
			// End Close Cookies
			try {
					int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
					System.out.println("We have Sections: " + sectionCount);
					int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
					System.out.println("Total Number: " + TotalCount);
					// Get Date
					DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
					Date date = new Date();
					String date1 = dateFormat.format(date);
					// Excel Create Headings
					XSSFSheet sheet = wb.createSheet();
					Row rowHeading0 = sheet.createRow(0);
					Row rowHeading1 = sheet.createRow(1);
					Row rowHeading2 = sheet.createRow(2);
					Row rowHeading3 = sheet.createRow(3);
					Row rowHeading4 = sheet.createRow(4);
					rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
					rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
					rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
					rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
					rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
					rowHeading4.createCell(0).setCellValue("SECTION");
					rowHeading4.createCell(1).setCellValue("ID");
					rowHeading4.createCell(2).setCellValue("TITLE");
					rowHeading4.createCell(3).setCellValue("EMPTY");
					rowHeading4.createCell(4).setCellValue("EMPTY");
					rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
					rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
					rowHeading4.createCell(7).setCellValue("LOCATION URL");
					rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");
					rowHeading4.createCell(11).setCellValue("DE HREF");
					rowHeading4.createCell(12).setCellValue("ES HREF");
					rowHeading4.createCell(13).setCellValue("FR HREF");
					rowHeading4.createCell(14).setCellValue("IT HREF");
					rowHeading4.createCell(15).setCellValue("JA HREF");
					rowHeading4.createCell(16).setCellValue("PT HREF");
					rowHeading4.createCell(17).setCellValue("RU HREF");
					rowHeading4.createCell(18).setCellValue("VI HREF");
					rowHeading4.createCell(19).setCellValue("CN HREF");
					rowHeading4.createCell(20).setCellValue("AR HREF");
					int rowNum = 5;
					

				for (int s = 2; s < sectionCount + 1; s++) {
					// Close Cookies
					try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception e) {
							System.out.println("Can't Close Cookies");
						}
					// End Close Cookies
					int rowCount = wd.findElements(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li")).size();
					System.out.println("Rows in this section: " + rowCount);
					
					// Main For Loop
					for (int i = 1; i < rowCount + 1; i++) {
						
						
						try {
							Row rowN = sheet.createRow(rowNum++);
							
							
							
							try {
								WebElement Section = wd.findElement(
										By.xpath(".//main/div[1]/div//div[" + s + "]/div/div/h2"));
								rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
								System.out.println("Section: " + Section.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
							} catch (Exception e) {
								System.out.println("Can't Find Section");
							}

							try {
								WebElement elementToClick = wd
										.findElement(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								try {
									JavascriptExecutor js = (JavascriptExecutor) wd;
									WebElement ScrollTo = wd.findElement(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li[" + i +"]"));
									js.executeScript("arguments[0].scrollIntoView(true);",ScrollTo);
									js.executeScript("javascript:window.scrollBy(0,-100)");
								}catch (Exception e) {
									System.out.println("Can't SCROLLL!!!");
								}
								
								elementToClick.click();
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("Can't CLICK ON ELEMENT");
								try {
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
								} catch (Exception ex) {
									Thread.sleep(2000);
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
									System.out.println("Can't CLICK ON ELEMENT 2nd time");
								}
							}

							try {
								WebElement Title = wd
										.findElement(By.xpath(".//h3[@class='PopupTable_title__mqaZ0']"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								// Write ID and Title
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception e) {
								System.out.println("Can't get the title. Reloading...");
								try {
									((JavascriptExecutor) wd).executeScript("window.open()");
									Thread.sleep(2000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									wd.close();
									Thread.sleep(3000);
									wd.switchTo().window(tabs.get(1));
									wd.get(currentURL);
									// Close Cookies
									try {
										WebElement AcceptCookies = wd
												.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
										AcceptCookies.click();
									} catch (Exception ez) {
										System.out.println("Can't Close Cookies");
									}
									// End Close Cookies
									WebElement elementToClick = wd
											.findElement(By.xpath(".//main/div[1]/div//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
									elementToClick.click();
									WebElement Title = wd
											.findElement(By.xpath(".//div[@class='active carousel-item']/div/h3"));
									System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
									rowN.createCell(1).setCellValue(i);
									rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
								} catch (Exception ex) {
									System.out.println("Can't get title second time...");
								}

							}
							// Show Details
							try {
								WebElement showDetails = wd.findElement(By.xpath("//span[@class='PopupTable_showDetails__WDJpC']"));
								showDetails.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								System.out.println("Can't CLICK ON SHOW DETAILS");
							}
							

							// Get location, verify links
							try {
								wait10.until(ExpectedConditions.visibilityOfElementLocated(
										By.xpath("//a[@class='InThisTopic_location__VJhFE']")));
								WebElement location = wd
										.findElement(By.xpath("//a[@class='InThisTopic_location__VJhFE']"));
								System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
								// Excel Values
								rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
								rowN.createCell(6).setCellValue("Pass");
								String linkUrl = location.getAttribute("href");
								rowN.createCell(7).setCellValue(linkUrl);

								System.out.println(linkUrl);

								// VERIFY LINKS ARE ACTIVE
								try {
									((JavascriptExecutor) wd).executeScript("window.open()");
									Thread.sleep(1000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									try {
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									} catch (Exception e) {
										Thread.sleep(3000);
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									}
									try {
										// Close Cookies
										try {
											WebElement AcceptCookies = wd
													.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
											AcceptCookies.click();
										} catch (Exception e) {
											System.out.println("Can't Close Cookies");
										}
										// End Close Cookies
										try {
											wait20.until(ExpectedConditions.visibilityOfElementLocated(
													By.xpath(".//div[1]/h1")));
										} catch (Exception e) {
											System.out.println("Refreshing the page!");
											wd.navigate().refresh();
											wait20.until(ExpectedConditions.visibilityOfElementLocated(
													By.xpath(".//div[1]/h1")));
										}
										WebElement TopicTitle = wd
												.findElement(By.xpath(".//div[1]/h1"));
										System.out.println("Location Topic Title: " + TopicTitle.getText());
										rowN.createCell(8).setCellValue("OK");
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Location Topic Title: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									// START HREFs
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
										System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(11).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(11).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
										System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(12).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(12).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
										System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(13).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(13).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
										System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(14).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(14).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
										System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(15).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(15).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
										System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(16).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(16).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
										System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(17).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(17).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
										System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(18).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(18).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
										System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(19).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(19).setCellValue("N/A");
									}
									try {
										WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
										System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
										rowN.createCell(20).setCellValue(HrefLink.getAttribute("href"));
									} catch (Exception e) {
										System.out.println("Couldn't get HREF");
										rowN.createCell(20).setCellValue("N/A");
									}
									// END HREFs
									wd.close();
									wd.switchTo().window(tabs.get(0));
								} catch (Exception e) {
									Thread.sleep(2000);
									System.out.println("URL: CANNOT VERIFY");
									rowN.createCell(8).setCellValue("CANNOT VERIFY");
								}

							} catch (Exception e) {
								System.out.println("Location " + i + ": Fail");
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(6).setCellValue("Fail");
							} // END get location, verify links

							
							try {
								// Close Popup
								wd.switchTo().defaultContent();
								wd.findElement(By
										.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
										.click();
							} catch (Exception e) {
								System.out.println("Cannot Close Popup!");
								Thread.sleep(2000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.switchTo().window(tabs.get(0));
								Thread.sleep(2000);
								wd.switchTo().defaultContent();
								wd.findElement(By
										.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
										.click();
							}

						} catch (Exception e) {
							System.out.println("ERROR! Web page is not responding. Relopening the page... ");
							((JavascriptExecutor) wd).executeScript("window.open()");
							Thread.sleep(1000);
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.close();
							Thread.sleep(2000);
							wd.switchTo().window(tabs.get(1));
							wd.get(currentURL);
							// Close Cookies
							try {
								WebElement AcceptCookies = wd
										.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
								AcceptCookies.click();
							} catch (Exception ez) {
								System.out.println("Can't Close Cookies");
							}
							// End Close Cookies
							Thread.sleep(10000);
							
						}
						
						
						
						
						FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Tables.xlsx");
						wb.write(fout);
						fout.close();
						System.out.println("Written Successfully!");
						
						
					}
				}

			} catch (Exception e) {
				System.out.println("Page Error? What is going on?");
				((JavascriptExecutor) wd).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
				wd.close();
				Thread.sleep(2000);
				wd.switchTo().window(tabs.get(1));
				wd.get(currentURL);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception ez) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				Thread.sleep(10000);
			}
		}
		
	// Get Date
	public void getDate() throws Exception {
		// Create object of SimpleDateFormat class and decide the format
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		// get current date time with Date()
		Date date = new Date();
		// Now format the date
		String date1 = dateFormat.format(date);
		// Print the Date
		System.out.println("Current date/time: " + date1);
	}
	
	// VERIFY LetterSpine
	public void verifyLetterSpine() throws Exception {
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(5));
		Actions action = new Actions(wd);
		try {
			
			// Excel Create Headings
			XSSFSheet sheet = wb.createSheet();
			Row rowHeading0 = sheet.createRow(0);
			rowHeading0.createCell(0).setCellValue("LETTER");
			rowHeading0.createCell(1).setCellValue("TOTAL NUMBER");
			rowHeading0.createCell(2).setCellValue("TITLE");
			int rowNum = 1;
			
			int letterCount = wd.findElements(By.xpath("//*[@class=\"letterspine__letter\"]")).size();
			System.out.println("We have Letters: " + letterCount);
			
			for (int s = 2; s < letterCount + 1; s++) {
				try {
					WebElement Letter = wd.findElement(By.xpath("//*[@class=\"letterspine__letter\"]["+s+"]"));
					Letter.click();
					System.out.println("Letter: " + Letter.getText());
				} catch (Exception e) {
					System.out.println("Can't Click");
				}
				Thread.sleep(1000);
				int titleCount = wd.findElements(By.xpath("//*[@class=\"modalspine__toplevel\"]")).size();
				System.out.println("Titles under this letter: " + titleCount);
				
				
				for (int i = 1; i < titleCount + 1; i++) {
					
					Row rowN = sheet.createRow(rowNum++);
					rowN.createCell(1).setCellValue(titleCount);
					// Get Title
					WebElement Title = wd.findElement(By.xpath("//*[@class=\"modalspine__toplevel\"]["+i+"]"));
					System.out.println("Title: "+Title.getText());
					rowN.createCell(2).setCellValue(Title.getText());
					
					
					String TwoLetter = Title.getText();
					System.out.println(TwoLetter.substring(0,2));
					rowN.createCell(0).setCellValue(TwoLetter.substring(0,2));
					
				}
				
				/*
				
				int zCount = wd.findElements(By.xpath("//*[@class=\"modalspine__letterswrapper two-letters\"]")).size();
				System.out.println("ZCount: "+zCount);
				for (int z = 1; z < zCount + 1; z++) {
					//Row rowN = sheet.createRow(rowNum++);
					// 
					WebElement GetTwoLetters = wd.findElement(By.xpath("//*[@class=\"modalspine__letterswrapper two-letters\"]/div["+z+"]"));
					System.out.println("Title: "+GetTwoLetters.getText());
					//rowN.createCell(3).setCellValue(GetTwoLetters.getText());
					
					
					
				}
				
				
				*/
				FileOutputStream fout = new FileOutputStream("C:\\TestResults\\LetterSpine.xlsx");
				wb.write(fout);
				fout.close();
				System.out.println("Written Successfully!");
				
				
			}

			
			
			
			
			
		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}
	
	// VERIFY PODCASTS
		public void verifyPodcasts() throws Exception {
			//Thread.sleep(10000);
			String currentURL = wd.getCurrentUrl();
			WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(15));
			Actions action = new Actions(wd);
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			try {
				Thread.sleep(2000);
				int sectionCount = wd.findElements(By.xpath(".//div[@class='row']")).size();
				System.out.println("We have Sections: " + sectionCount);

				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				Row rowHeading4 = sheet.createRow(4);
				rowHeading0.createCell(0).setCellValue("PODCASTS");
				rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading4.createCell(0).setCellValue("ID");
				rowHeading4.createCell(1).setCellValue("SEASON/EPISODE");
				rowHeading4.createCell(2).setCellValue("ALT TITLE");
				rowHeading4.createCell(3).setCellValue("URL");
				rowHeading4.createCell(4).setCellValue("FEATURED IMAGE (PASS/FAIL)");
				rowHeading4.createCell(5).setCellValue("PODCAST (PASS/FAIL)");
				rowHeading4.createCell(6).setCellValue("ANCHOR TITLE");
				rowHeading4.createCell(7).setCellValue("ANCHOR URL");
				rowHeading4.createCell(8).setCellValue("APPLE TITLE");
				rowHeading4.createCell(9).setCellValue("APPLE URL");
				rowHeading4.createCell(10).setCellValue("GOOGLE TITLE");
				rowHeading4.createCell(11).setCellValue("GOOGLE URL");
				rowHeading4.createCell(12).setCellValue("SPOTIFY TITLE");
				rowHeading4.createCell(13).setCellValue("SPOTIFY URL");

				int rowNum = 5;

				for (int s = 2; s < sectionCount + 1; s++) {
					int rowCount = wd.findElements(By.xpath(".//div["+s+"]/div[@class='column2']/a")).size();
					System.out.println("Rows in this section: " + rowCount);
					
					for (int i = 1; i < rowCount + 1; i++) {
						// Main For Loop

						try {
							Thread.sleep(400);
							Row rowN = sheet.createRow(rowNum++);
							rowN.createCell(0).setCellValue(i);
							try {
								try {
									WebElement image = wd.findElement(By.xpath(".//div["+s+"]/div[@class='column2']/a["+i+"]/img"));
									String altAttribute = image.getAttribute("alt");
									System.out.println("ALT TEXT " + i + ": " + altAttribute);
									rowN.createCell(2).setCellValue(altAttribute);
									rowN.createCell(4).setCellValue("Pass");
								}catch (Exception e) {
									System.out.println("ALT TEXT " + i + ": MISSING");
									rowN.createCell(2).setCellValue("MISSING");
									rowN.createCell(4).setCellValue("Fail");
								}
								
								WebElement location = wd.findElement(By.xpath(".//div["+s+"]/div[@class='column2']/a["+i+"]"));
								System.out.println("Location " + i + ": " + location);
								String linkUrl = location.getAttribute("href");
								System.out.println(linkUrl);
								rowN.createCell(3).setCellValue(linkUrl);
								

								// VERIFY LOCATION
								try {
									((JavascriptExecutor) wd).executeScript("window.open()");
									Thread.sleep(1000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									try {
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									} catch (Exception e) {
										Thread.sleep(3000);
										wd.switchTo().window(tabs.get(1));
										wd.get(linkUrl);
									}
									
									try {
										wait.until(ExpectedConditions.visibilityOfElementLocated(
												By.xpath(".//h1[@class='news__detail--title']")));
										WebElement TopicTitle = wd
												.findElement(By.xpath(".//h1[@class='news__detail--title']"));
										System.out.println("Podcast Title: " + TopicTitle.getText());
										rowN.createCell(5).setCellValue("Pass");
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Podcast Title: NOT FOUND");
										rowN.createCell(5).setCellValue("Fail");
									}
									
									// Season and Episode
									try {
									
										WebElement Season = wd
												.findElement(By.xpath(".//div[@class='news__detail--main']/p[1]"));
										System.out.println("Podcast Episode #: " + Season.getText());
										rowN.createCell(1).setCellValue(Season.getText());
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Podcast Episode #: NOT FOUND");
										rowN.createCell(1).setCellValue("Not Found");
									}
									// ANCHOR LINK
									try {
										WebElement Link1 = wd.findElement(By.xpath(".//div[@class='news__detail--main']/a[1]"));
										String Link1URL = Link1.getAttribute("href");
										System.out.println("Anchor: " + Link1URL);
										rowN.createCell(7).setCellValue(Link1URL);
										((JavascriptExecutor) wd).executeScript("window.open()");
										Thread.sleep(1000);
										ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
										try {
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										} catch (Exception e) {
											Thread.sleep(3000);
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										}
										try {
											wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='styles__segmentInfo___4yt5o']/h2")));
											WebElement AnchorTitle = wd.findElement(By.xpath(".//div[@class='styles__segmentInfo___4yt5o']/h2"));
											System.out.println("Anchor Title: " + AnchorTitle.getText());
											rowN.createCell(6).setCellValue(AnchorTitle.getText());
										} catch (Exception e) {
											System.out.println("Anchor Title: " + "NOT FOUND");
										}
										
										wd.close();
										wd.switchTo().window(tabs1.get(1));
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Anchor: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									// APPLE LINK
									try {
										WebElement Link1 = wd.findElement(By.xpath(".//div[@class='news__detail--main']/a[2]"));
										String Link1URL = Link1.getAttribute("href");
										System.out.println("Apple: " + Link1URL);
										rowN.createCell(9).setCellValue(Link1URL);
										((JavascriptExecutor) wd).executeScript("window.open()");
										Thread.sleep(1000);
										ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
										try {
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										} catch (Exception e) {
											Thread.sleep(3000);
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										}
										try {
											wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//header/h1/span[1]")));
											WebElement AppleTitle = wd.findElement(By.xpath(".//header/h1/span[1]"));
											System.out.println("Apple Title: " + AppleTitle.getText());
											rowN.createCell(8).setCellValue(AppleTitle.getText());
										} catch (Exception e) {
											System.out.println("Apple Title: " + "NOT FOUND");
										}
										
										wd.close();
										wd.switchTo().window(tabs1.get(1));
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Apple: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									// Google LINK
									try {
										WebElement Link1 = wd.findElement(By.xpath(".//div[@class='news__detail--main']/a[3]"));
										String Link1URL = Link1.getAttribute("href");
										System.out.println("Google: " + Link1URL);
										rowN.createCell(11).setCellValue(Link1URL);
										((JavascriptExecutor) wd).executeScript("window.open()");
										Thread.sleep(1000);
										ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
										try {
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										} catch (Exception e) {
											Thread.sleep(3000);
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										}
										try {
											wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='wv3SK']")));
											WebElement GoogleTitle = wd.findElement(By.xpath(".//div[@class='wv3SK']"));
											System.out.println("Google Title: " + GoogleTitle.getText());
											rowN.createCell(10).setCellValue(GoogleTitle.getText());
										} catch (Exception e) {
											System.out.println("Google Title: " + "NOT FOUND");
										}
										
										wd.close();
										wd.switchTo().window(tabs1.get(1));
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Google: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									
									
									// Spotify LINK
									try {
										WebElement Link1 = wd.findElement(By.xpath(".//div[@class='news__detail--main']/a[4]"));
										String Link1URL = Link1.getAttribute("href");
										System.out.println("Google: " + Link1URL);
										rowN.createCell(13).setCellValue(Link1URL);
										((JavascriptExecutor) wd).executeScript("window.open()");
										Thread.sleep(1000);
										ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
										try {
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										} catch (Exception e) {
											Thread.sleep(3000);
											wd.switchTo().window(tabs1.get(2));
											wd.get(Link1URL);
										}
										try {
											wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//section/div[1]/div[3]/h1")));
											WebElement SpotifyLink = wd.findElement(By.xpath(".//section/div[1]/div[3]/h1"));
											System.out.println("Spotify Title: " + SpotifyLink.getText());
											rowN.createCell(12).setCellValue(SpotifyLink.getText());
										} catch (Exception e) {
											System.out.println("Spotify Title: " + "NOT FOUND");
										}
										
										wd.close();
										wd.switchTo().window(tabs1.get(1));
										Thread.sleep(1000);
									} catch (Exception e) {
										System.out.println("Spotify: NOT FOUND");
										rowN.createCell(8).setCellValue("Not Found");
									}
									
									wd.close();
									wd.switchTo().window(tabs.get(0));
								} catch (Exception e) {
									Thread.sleep(2000);
									System.out.println("URL: CANNOT VERIFY");
									rowN.createCell(8).setCellValue("CANNOT VERIFY");
								}

							} catch (Exception e) {
								System.out.println("Location " + i + ": Fail");
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(6).setCellValue("Fail");
							} // END get location, verify links


						} catch (Exception e) {
							System.out.println("ERROR! Web page is not responding. Relopening the page... ");
							((JavascriptExecutor) wd).executeScript("window.open()");
							Thread.sleep(1000);
							ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
							wd.close();
							Thread.sleep(2000);
							wd.switchTo().window(tabs.get(1));
							wd.get(currentURL);
							// Close Cookies
							try {
								WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
								AcceptCookies.click();
							} catch (Exception ez) {
								System.out.println("Can't Close Cookies");
							}
							// End Close Cookies
							Thread.sleep(10000);
						}
						FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Podcasts.xlsx");
						wb.write(fout);
						fout.close();
						System.out.println("Written Successfully!");
					}
					
					
				}

			} catch (Exception e) {
				System.out.println("Page Error? What is going on?");
				((JavascriptExecutor) wd).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
				wd.close();
				Thread.sleep(2000);
				wd.switchTo().window(tabs.get(1));
				wd.get(currentURL);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
					AcceptCookies.click();
				} catch (Exception ez) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				Thread.sleep(10000);
			}
		}
	
	// TEMPLATE
	public void verifyTemplate() throws Exception {
		getDate();
		wd.manage().window().setSize(new Dimension(300, 840));
		wd.manage().window().setPosition(new Point(900, 0));
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(5));
		Actions action = new Actions(wd);
		
		try {
			//Thread.sleep(8000);
				int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
				System.out.println("We have Sections: " + sectionCount);
				int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
				System.out.println("Total Number: " + TotalCount);
				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				Row rowHeading4 = sheet.createRow(4);
				rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
				rowHeading0.createCell(0).setCellValue("TOTAL SECTIONS: " + sectionCount);
				rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading4.createCell(0).setCellValue("SECTION");
				rowHeading4.createCell(1).setCellValue("ID");
				rowHeading4.createCell(2).setCellValue("TITLE");
				rowHeading4.createCell(3).setCellValue("CREDITS");
				rowHeading4.createCell(4).setCellValue("DESCRIPTION");
				rowHeading4.createCell(5).setCellValue("LOCATION TITLE");
				rowHeading4.createCell(6).setCellValue("LOCATION (PASS/FAIL)");
				rowHeading4.createCell(7).setCellValue("LOCATION URL");
				rowHeading4.createCell(8).setCellValue("LOC. URL RESPONSE");

				int rowNum = 5;
				

			for (int s = 2; s < sectionCount + 1; s++) {
				int rowCount = wd.findElements(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li")).size();
				System.out.println("Rows in this section: " + rowCount);
				for (int i = 1; i < rowCount + 1; i++) {
					// Main For Loop
					try {
						Row rowN = sheet.createRow(rowNum++);

						try {
							WebElement Section = wd.findElement(By.xpath(
									".//main/div[1]/div/div[" + s + "]/div/div/h2"));
							rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
							System.out.println("Section: " + Section.getAttribute("innerText"));
							rowN.createCell(1).setCellValue(i);
						} catch (Exception e) {
							System.out.println("Can't Find Section");
						}
						
						try {
							WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
							elementToClick.click();
						} catch (Exception e) {
							Thread.sleep(2000);
							System.out.println("Can't CLICK ON ELEMENT");
							try {
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
							} catch (Exception ex) {
								Thread.sleep(2000);
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								System.out.println("Can't CLICK ON ELEMENT 2nd time");
							}
						}
						
						try {
							WebElement Title = wd.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
							System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
							// Write ID and Title
							rowN.createCell(1).setCellValue(i);
							rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
						} catch (Exception e) {
							System.out.println("Can't get the title. Reloading...");
							try {
								((JavascriptExecutor) wd).executeScript("window.open()");
								Thread.sleep(1000);
								ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
								wd.close();
								Thread.sleep(2000);
								wd.switchTo().window(tabs.get(1));
								wd.get(currentURL);
								// Close Cookies
								try {
									WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
									AcceptCookies.click();
								} catch (Exception ez) {
									System.out.println("Can't Close Cookies");
								}
								// End Close Cookies
								WebElement elementToClick = wd.findElement(By.xpath(".//main/div[1]/div/div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"));
								elementToClick.click();
								WebElement Title = wd.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
								System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
								rowN.createCell(1).setCellValue(i);
								rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
							} catch (Exception ex) {
								System.out.println("Can't get title second time...");
							}
							
						}

							try {
								// Close Popup
								wd.switchTo().defaultContent();
								wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
							} catch (Exception e) {
									System.out.println("Cannot Close Popup!");
									Thread.sleep(2000);
									ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
									wd.switchTo().window(tabs.get(0));
									Thread.sleep(2000);
									wd.switchTo().defaultContent();
									wd.findElement(By.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']")).click();
							}

						
							
					} catch (Exception e) {
						System.out.println("ERROR! Web page is not responding. Relopening the page... ");
						((JavascriptExecutor) wd).executeScript("window.open()");
						Thread.sleep(1000);
						ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
						wd.close();
						Thread.sleep(2000);
						wd.switchTo().window(tabs.get(1));
						wd.get(currentURL);
						// Close Cookies
						try {
							WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
							AcceptCookies.click();
						} catch (Exception ez) {
							System.out.println("Can't Close Cookies");
						}
						// End Close Cookies
						Thread.sleep(10000);
					}
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Template.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
					
					
				}
			}

		} catch (Exception e) {
			System.out.println("Page Error? What is going on?");
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.close();
			Thread.sleep(2000);
			wd.switchTo().window(tabs.get(1));
			wd.get(currentURL);
			// Close Cookies
			try {
				WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
				AcceptCookies.click();
			} catch (Exception ez) {
				System.out.println("Can't Close Cookies");
			}
			// End Close Cookies
			Thread.sleep(10000);
		}
	}

	//////////////////////////REUSABLE METHODS
	
	
	// Manage Windows
	public void ManageWindows(){
		//wd.manage().window().maximize();
		wd.manage().window().setSize(new Dimension(400, 820));
		wd.manage().window().setPosition(new Point(780, 0));
		//wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
		Actions action = new Actions(wd);
	}
	
	// Close Cookies
	public void CloseCookies() throws InterruptedException {
		try {
			WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
			AcceptCookies.click();
		} catch (Exception e) {
			System.out.println("Can't Close Cookies");
		}
		Thread.sleep(1000);
	}
	
	// Confirm page opens
		public void confirmPage(WebDriverWait wait50,WebDriverWait wait20) {
			try {
				wait20.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			} catch (Exception e) {
				System.out.println("Refreshing the page!");
				wd.navigate().refresh();
				wait20.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			}
		}
	
	// Excel Rows
	public void excelRows(XSSFSheet sheet,int sectionCount,int TotalCount,int HeadN,Row rowHeading4) {
		try {
			// Get Date
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
			Date date = new Date();
			String date1 = dateFormat.format(date);
			//Main Excel Rows for all
			Row rowHeading0 = sheet.createRow(0);
			Row rowHeading1 = sheet.createRow(1);
			Row rowHeading2 = sheet.createRow(2);
			Row rowHeading3 = sheet.createRow(3);
			rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
			rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
			rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
			rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
			rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
			int HrefN = HeadN++;
			rowHeading4.createCell(HrefN++).setCellValue("DE HREF");
			rowHeading4.createCell(HrefN++).setCellValue("ES HREF");
			rowHeading4.createCell(HrefN++).setCellValue("FR HREF");
			rowHeading4.createCell(HrefN++).setCellValue("IT HREF");
			rowHeading4.createCell(HrefN++).setCellValue("JA HREF");
			rowHeading4.createCell(HrefN++).setCellValue("PT HREF");
			rowHeading4.createCell(HrefN++).setCellValue("RU HREF");
			rowHeading4.createCell(HrefN++).setCellValue("VI HREF");
			rowHeading4.createCell(HrefN++).setCellValue("CN HREF");
			rowHeading4.createCell(HrefN++).setCellValue("AR HREF");
			rowHeading4.createCell(HrefN++).setCellValue("UK HREF");
			
		} catch (Exception e) {
			System.out.println("Cannot Create Excel Rows!!!");
		}
		
	}
	
	// Open Each Resource Item
	public void openResource(Row rowN,int i, int s) throws InterruptedException{
		
		//Main Title
		try {
			WebElement Title = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a"));
			System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
			// Write ID and Title
			rowN.createCell(1).setCellValue(i);
			rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
		} catch (Exception e) {
		System.out.println("CANNOT GET TITLE");
		rowN.createCell(1).setCellValue(i);
		rowN.createCell(2).setCellValue("CANNOT GET TITLE");
		}
		//Open Resource
		JavascriptExecutor js = (JavascriptExecutor) wd;
        String liXPath = ".//main/div[1]//div["+s+"]//div/div[2]/ul/li[" + i + "]/a";
        WebElement elementToClick = wd.findElement(By.xpath(liXPath));
        js.executeScript("arguments[0].scrollIntoView({block: 'center'});", elementToClick);
        Thread.sleep(500);
		elementToClick.click();
		try {
			WebElement Section = wd.findElement(By.xpath(".//main/div[1]//div["+s+"]/div/div/h2"));
			rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
			System.out.println("Section: " + Section.getAttribute("innerText"));
			rowN.createCell(1).setCellValue(i);
		} catch (Exception e) {
			System.out.println("Can't Find Section");
		}
	}
	
	// Get Credits and Descriptions
	public void getCredits(Row rowN, int i) {

		// Get & write Description
		try {
			WebElement description = wd.findElement(By.xpath(
					".//div["+i+"]/div/div[3]/div[2]/div[3]/div[1]/div/p"));
			System.out.println("Description: " + description.getAttribute("innerText"));
			rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("Description: CANT FIND");
			rowN.createCell(4).setCellValue("");
		}
		// Get & write Credits
		try {
			WebElement credits = wd
					.findElement(By.xpath(".//div[\"+i+\"]/div/div[3]/div[2]/div[3]/div[2]/div/p"));
			System.out.println("Credits: " + credits.getAttribute("innerText"));
			rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("");
			rowN.createCell(3).setCellValue("");
		}

	}

	//GET FILE NAME AND VERIFY IMAGE
	public void getFileName(Row rowN, WebDriverWait wait10) {
		try {
			WebElement FileName = wd.findElement(By.xpath(".//div[@class='multimedia__image--wrapper']/img"));
			URL FileURL = new URL(FileName.getAttribute("src"));
			System.out.println("Src attribute is: " + FileURL);
			System.out.println("path = " + FileURL.getPath());
			rowN.createCell(9).setCellValue(FileName.getAttribute("src"));
			rowN.createCell(10).setCellValue(FileURL.getPath());
			String ImageURL = FileName.getAttribute("src");
			// VERIFY LINKS ARE ACTIVE
			try {
				((JavascriptExecutor) wd).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
				try {
					wd.switchTo().window(tabs.get(1));
					wd.get(ImageURL);
				} catch (Exception e) {
					Thread.sleep(3000);
					wd.switchTo().window(tabs.get(1));
					wd.get(ImageURL);
				}
				try {
					try {
						wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/img")));
					} catch (Exception e) {
						System.out.println("Refreshing the page!");
						wd.navigate().refresh();
						wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/img")));
					}
					WebElement ImagePresent = wd.findElement(By.xpath("/html/body/img"));
					System.out.println("Image: FOUND");
					rowN.createCell(11).setCellValue("FOUND");
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Image: NOT FOUND");
					rowN.createCell(11).setCellValue("NOT FOUND");
				}
				wd.close();
				wd.switchTo().window(tabs.get(0));
			} catch (Exception e) {
				Thread.sleep(2000);
				System.out.println("CANNOT VERIFY IMAGE");
				rowN.createCell(11).setCellValue("CANNOT VERIFY");
			}
		} catch (Exception e) {
			System.out.println("CANNOT GET FILE NAME");
			rowN.createCell(9).setCellValue("No File PATH");
			rowN.createCell(10).setCellValue("No File URL");
			rowN.createCell(11).setCellValue("No Image");
		}
	}
			
	
	// Get Location
	public void getLocation(Row rowN,WebDriverWait wait10,WebDriverWait wait20, WebDriverWait wait30, int i, Row rowHeading4, int HeadN, int s) throws IOException, InterruptedException {
	try {
		WebElement location = wd.findElement(By.xpath(".//div["+i+"]/div/div[3]/div[3]/div/div/ul/li/a"));
		System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
		
		// Excel Values
		rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
		rowN.createCell(6).setCellValue("Pass");
		String linkUrl = location.getAttribute("href");
		rowN.createCell(7).setCellValue(linkUrl);
		System.out.println(linkUrl);
		
		// VERIFY LINKS ARE ACTIVE
		try {
			((JavascriptExecutor) wd).executeScript("window.open()");
			Thread.sleep(1000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			try {
				wd.switchTo().window(tabs.get(1));
				wd.get(linkUrl);
			} catch (Exception e) {
				Thread.sleep(3000);
				wd.switchTo().window(tabs.get(1));
				wd.get(linkUrl);
			}
			try {
				CloseCookies();
				try {
					wait30.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/h1")));
				} catch (Exception e) {
					System.out.println("Refreshing the page!");
					wd.navigate().refresh();
					wait30.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/h1")));
				}
				WebElement TopicTitle = wd
						.findElement(By.xpath(".//div[1]/h1"));
				System.out.println("Location Topic Title: " + TopicTitle.getText());
				rowN.createCell(8).setCellValue("OK");
				Thread.sleep(1000);
			} catch (Exception e) {
				System.out.println("Location Topic Title: NOT FOUND");
				rowN.createCell(8).setCellValue("Not Found");
			}
			//START HREFS
			int HrefNum = HeadN++;
			// Find the meta tag with content attribute
			WebElement metaTag = wd.findElement(By.cssSelector("meta[name='edition']"));
	        // Get the content attribute value
	        String content = metaTag.getAttribute("content");
			// Check if the content attribute equals "Veterinary"
	        if (content != null && content.contains("Veterinary")) {
	            System.out.println("The meta content is 'Veterinary'. No HREFS");
	        } else {
	            System.out.println("The meta content is not 'Veterinary'. Testing HREFs...");
	            try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='de']"));
					System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='es']"));
					System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
					System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='it']"));
					System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
					System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
					System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
					System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='zh']"));
					System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
					System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='ar']"));
					System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = wd.findElement(By.xpath(".//link[@hreflang='uk']"));
					System.out.println("UK HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
	        }
			//END HREFs
			wd.close();
			wd.switchTo().window(tabs.get(0));
			
		} catch (Exception e) {
			Thread.sleep(2000);
			System.out.println("URL: CANNOT VERIFY");
			wd.close();
			rowN.createCell(8).setCellValue("CANNOT VERIFY");
		}
		
	} catch (Exception e) {
		System.out.println("Location " + i + ": Fail");
		Thread.sleep(3000);
		rowN.createCell(1).setCellValue(i);
		rowN.createCell(6).setCellValue("Fail");
	} // END get location, verify links
	
}
	
	// Close Popup
	public void ClosePopup() throws InterruptedException {
		
		try {
			// Close Popup
			wd.switchTo().defaultContent();
			wd.findElement(By
					.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
					.click();
		} catch (Exception e) {
			System.out.println("Cannot Close Popup!");
			Thread.sleep(2000);
			ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
			wd.switchTo().window(tabs.get(0));
			Thread.sleep(2000);
			wd.switchTo().defaultContent();
			wd.findElement(By
					.xpath(".//p[@class='modal_btnClose__eatUo modal_headerElement__9HT5x']"))
					.click();
		}
	}
	
	
	//HANDLING ERRORS
	public void ErrorMain(String currentURL, Exception e) throws Exception{
		Thread.sleep(3000);
		((JavascriptExecutor) wd).executeScript("window.open()");
		Thread.sleep(2000);
		ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
		Thread.sleep(2000);
		wd.close();
		wd.switchTo().window(tabs1.get(1));
		wd.get(currentURL);
		Thread.sleep(5000);
		Thread.sleep(2000);
		CloseCookies();
	}
		

	
	
	
	
}
