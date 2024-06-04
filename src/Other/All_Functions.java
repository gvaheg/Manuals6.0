package Other;

import java.io.FileOutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;
import java.util.concurrent.TimeUnit;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class All_Functions {

	// START DRIVER
	static WebDriver wd;
	static XSSFWorkbook wb;
	
	// VERIFY Daily
		public void verifyDaily() throws Exception {
			wd.manage().window().maximize();
			wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
			WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(5));
			Actions action = new Actions(wd);
				try {
					// Excel Create Headings
					XSSFSheet sheet = wb.createSheet();
					Row rowHeading0 = sheet.createRow(0);
					rowHeading0.createCell(0).setCellValue("Test Page");
					rowHeading0.createCell(1).setCellValue("Page Title");
					rowHeading0.createCell(2).setCellValue("Element On Page");
					rowHeading0.createCell(3).setCellValue("URL");
					rowHeading0.createCell(4).setCellValue("Status");
					
					
					try {
						//Home Page
						Row row1 = sheet.createRow(1);
						row1.createCell(0).setCellValue("Home Page");
						try {
							try {
								String title = wd.getTitle();
								String currentURL = wd.getCurrentUrl();
								System.out.println("Title is: "+title);
								row1.createCell(1).setCellValue(title);
								System.out.println("URL is: "+currentURL);
								row1.createCell(3).setCellValue(currentURL);
							}catch(Exception e) {
								System.out.println("No Title");
							}
							try {
								WebElement FeaturedContent = wd.findElement(By.xpath("//*[@class=\"box__header\"]"));
								System.out.println(FeaturedContent.getText()+": Pass");
								row1.createCell(2).setCellValue(FeaturedContent.getText());
								row1.createCell(4).setCellValue("Pass");
							} catch (Exception e) {
								System.out.println("Fail");
								row1.createCell(4).setCellValue("Fail");
							}
						}catch(Exception e) {
							System.out.println("Home Page Failed");
							row1.createCell(4).setCellValue("Home Page Failed");
						}
						
						
						//Topic Page
						Row row2 = sheet.createRow(2);
						row2.createCell(0).setCellValue("Topic Page");
						try {
							// Select Random Chapter
							try {
								WebElement dropdown = wd.findElement(By.xpath("//div[@class='mainmenu__title mainmenu__title--nolink']"));
								// Putting in a loop to select different values every time
								// Click on drop down
								dropdown.click();
								// Waiting till options in drop down are visible
								wait.until(ExpectedConditions.visibilityOf(wd.findElement(By.xpath("//div[@class='maintab__header']"))));
								// Getting list of options
								List<WebElement> itemsInDropdown = wd.findElements(By.xpath("//div[@class='maintab maintab--health ']//div[@class='maintab__left']//a[@class='maintab__item']"));
								// Getting size of options available
								int size = itemsInDropdown.size();
								System.out.println(size/2);
								// Generate a random number with in range
								int randnMumber = ThreadLocalRandom.current().nextInt(0, size/2);
								// Selecting random value
								itemsInDropdown.get(randnMumber).click();
								Thread.sleep(2000);
							} catch (Exception e) {
								System.out.println("Fail");
								row2.createCell(4).setCellValue("Can't Select Dropdown Item");
							}
							// Select Random Topic
							try {
								wait.until(ExpectedConditions.visibilityOf(wd.findElement(By.xpath("//h3[@class='medicalsection__caption']"))));
								List<WebElement> topicTitles = wd.findElements(By.xpath("//h3[@class='medicalsection__caption']"));
								int size = topicTitles.size();
								System.out.println(size/2);
								int randnMumber = ThreadLocalRandom.current().nextInt(0, size/2);
								// Selecting random value
								topicTitles.get(randnMumber).click();
								Thread.sleep(2000);
							} catch (Exception e) {
								System.out.println("Fail");
								row2.createCell(4).setCellValue("Can't Click On Topic");
							}
								
							try {
								String title = wd.getTitle();
								String currentURL = wd.getCurrentUrl();
								System.out.println("Title is: "+title);
								row2.createCell(1).setCellValue(title);
								System.out.println("URL is: "+currentURL);
								row2.createCell(3).setCellValue(currentURL);
							}catch(Exception e) {
								System.out.println("No Title");
							}
							try {
								WebElement TopicTitle = wd.findElement(By.xpath("//h1[@class='topic__header--title topic__header--title--animate']"));
								System.out.println(TopicTitle.getText()+": Pass");
								row2.createCell(2).setCellValue(TopicTitle.getText());
								row2.createCell(4).setCellValue("Pass");
							} catch (Exception e) {
								System.out.println("Fail");
								row2.createCell(4).setCellValue("Fail");
							}
						}catch(Exception e) {
							System.out.println("Topic Page Failed");
							row2.createCell(4).setCellValue("Topic Page Failed");
						}
						
						FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Daily_Testing.xlsx");
						wb.write(fout);
						fout.close();
						System.out.println("Written Successfully!");
						Thread.sleep(5000);

					} catch (Exception e) {
						System.out.println("ERROR@!");
						Thread.sleep(5000);
					}
						
						
				} catch (Exception e) {
					System.out.println("Page Error!");
				}
			
			
			
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
			
			for (int s = 1; s < letterCount + 1; s++) {
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
			Thread.sleep(10000);
		}
	}
	
	// VERIFY Drugs
	public void verifyDrugs() throws Exception {
		String currentURL = wd.getCurrentUrl();
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(5));
		Actions action = new Actions(wd);
			try {
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet();
				Row rowHeading0 = sheet.createRow(0);
				rowHeading0.createCell(0).setCellValue("LETTER");
				rowHeading0.createCell(1).setCellValue("DRUG TITLE");
				rowHeading0.createCell(2).setCellValue("STATUS");
				int rowNum = 1;
				
				
				Thread.sleep(3000);
				WebElement linkByText = wd.findElement(By.linkText("DRUG INFO"));
				linkByText.click();
				Thread.sleep(3000);
				WebElement linksByText = wd.findElement(By.linkText("Drugs by Name, Generic and Brand"));
				linksByText.click();
				Thread.sleep(2000);
				
				int letterCount = wd.findElements(By.xpath("/html/body/div[2]/div[6]/section/div/section/main/div[3]/div/a")).size();
				System.out.println("We have Letters: " + letterCount);
				
				for (int s = 1; s < 27; s++) {
					
					try {
						WebElement Letter = wd.findElement(By.xpath("/html/body/div[2]/div[6]/section/div/section/main/div[3]/div/a["+s+"]"));
						Letter.click();
						System.out.println("Letter: " + Letter.getText());
						
					} catch (Exception e) {
						System.out.println("Can't Click On Letter");
					}
					
					
					Thread.sleep(5000);
					try {
						int rowCount = wd.findElements(By.xpath(".//section["+s+"]/div/table/tbody/tr")).size();
						System.out.println("Rows in this section: " + rowCount);
						for (int i = 1; i < rowCount+1; i++) {
							Row rowN = sheet.createRow(rowNum++);
							
							
							WebElement Title = wd.findElement(By.xpath(".//section["+s+"]/div/table/tbody/tr["+i+"]/td[1]/div"));
							System.out.println(Title.getText());
							rowN.createCell(1).setCellValue(Title.getText());
							Title.click();
							/*
							String OneLetter = Title.getText();
							System.out.println(OneLetter.substring(0,1));
							rowN.createCell(0).setCellValue(OneLetter.substring(0,1));
							*/
							try {
								wd.findElement(By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
								try {
									WebElement AnyText = wd.findElement(By.xpath("//*[@id=\"section_0\"]/b"));
									System.out.println("PASS");
									rowN.createCell(2).setCellValue("Pass");

								} catch (Exception e) {
									try {
										wd.findElement(By.xpath(".//div[@class='monographLinks']/a")).click();
										
										WebElement AnyText = wd.findElement(By.xpath("//*[@id=\"section_0\"]/b"));
										System.out.println("PASS");
										rowN.createCell(2).setCellValue("Pass");
										wd.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']")).click();

									} catch (Exception ez) {
										System.out.println("FAIL");
										rowN.createCell(2).setCellValue("Fail");
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
								rowN.createCell(2).setCellValue("CANNOT OPEN DRUG, NO URL");
							}
							
							FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Drugs.xlsx");
							wb.write(fout);
							fout.close();
							System.out.println("Written Successfully!");
						}
						Thread.sleep(5000);

					} catch (Exception e) {
						System.out.println("ERROR@!");
						Thread.sleep(5000);
					}
					
					
				}
				

				
				
				
				
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		
		
		
	}
	
	// VERIFY Emergencies
		public void verifyEmergencies() throws Exception {
			String currentURL = wd.getCurrentUrl();
			WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(5));
			Actions action = new Actions(wd);
			wd.manage().window().maximize();
			String Version = wd.findElement(By.xpath(".//meta[3]")).getAttribute("lang");
			
			try {
				// Get Date
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
				Date date = new Date();
				String date1 = dateFormat.format(date);
				// Excel Create Headings
				XSSFSheet sheet = wb.createSheet(Version);
				Row rowHeading0 = sheet.createRow(0);
				Row rowHeading1 = sheet.createRow(1);
				Row rowHeading2 = sheet.createRow(2);
				Row rowHeading3 = sheet.createRow(3);
				rowHeading0.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
				rowHeading1.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
				rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
				rowHeading3.createCell(0).setCellValue("Title on First Aid page");
				rowHeading3.createCell(1).setCellValue("Title on Topic Page");
				rowHeading3.createCell(2).setCellValue("URL");
				rowHeading3.createCell(3).setCellValue("Titles Match (PASS/FAIL)");
				rowHeading3.createCell(4).setCellValue("URL ERROR (PASS/FAIL)");
				
				int rowNum = 4;
				//Thread.sleep(3000);
				
				// TESTING EMERGENCIES
				int EmergenciesCount = wd.findElements(By.xpath(".//ul[@class='emergencies__images--list']/li")).size();
				System.out.println("We have Emergencies: " + EmergenciesCount);
				for (int s = 1; s < EmergenciesCount + 1; s++) {
					try {
						Row rowN = sheet.createRow(rowNum++);
						WebElement EmergencyTitle = wd.findElement(By.xpath(".//ul[@class='emergencies__images--list']/li["+s+"]/a"));
						String actualTitle = EmergencyTitle.getText();
						System.out.println("Emergency Title: " + EmergencyTitle.getText());
						rowN.createCell(0).setCellValue(EmergencyTitle.getText());
						
						// Get location, verify links
						try {
							String linkUrl = EmergencyTitle.getAttribute("href");
							System.out.println("Location URL: " + linkUrl);
							rowN.createCell(2).setCellValue(linkUrl);
							if (linkUrl == null || linkUrl.isEmpty()) {
								System.out.println("URL is either not configured for anchor tag or it is empty");
								rowN.createCell(4).setCellValue("Fail");
								continue;
							}
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
											By.xpath(".//div[@class='topic__headings']/div/div/h1")));
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(1).setCellValue(TopicTitle.getText());
									// START TITLE MATCH CHECK
									String expectedTitle = TopicTitle.getText();
									if(actualTitle.equalsIgnoreCase(expectedTitle)) 
									{
										System.out.println("Titles Match");
										rowN.createCell(3).setCellValue("Pass");
									}else {
										System.out.println("Titles DO NOT Match");
										rowN.createCell(3).setCellValue("Fail");
									}
									// END TITLE MATCH CHECK
									rowN.createCell(4).setCellValue("Pass");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
								}
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								
							}
							
						} catch (Exception e) {
							System.out.println("Location Topic Title: NOT FOUND");
							rowN.createCell(4).setCellValue("Fail");
						} // END get location, verify links
						
					} catch (Exception e) {
						System.out.println("Can't Find Emergencies");
					}
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Emergencies.xlsx");
					wb.write(fout);
					fout.close();
					System.out.println("Written Successfully!");
				}
				Thread.sleep(3000);
				
				
				// TESTING INJURES
				int InjuresCount = wd.findElements(By.xpath(".//ul[@class='emergencies__other--list']/li")).size();
				System.out.println("We have Injures: " + InjuresCount);
				for (int i = 1; i < InjuresCount + 1; i++) {
					try {
						Row rowN = sheet.createRow(rowNum++);
						WebElement InjuresTitle = wd.findElement(By.xpath(".//ul[@class='emergencies__other--list']/li["+i+"]/a"));
						System.out.println("Injures Title: " + InjuresTitle.getText());
						String actualTitle = InjuresTitle.getText();
						rowN.createCell(0).setCellValue(InjuresTitle.getText());
						
						// Get location, verify links
						try {
							String linkUrl = InjuresTitle.getAttribute("href");
							System.out.println("Location URL: " + linkUrl);
							rowN.createCell(2).setCellValue(linkUrl);
							
							if (linkUrl == null || linkUrl.isEmpty()) {
								System.out.println("URL is either not configured for anchor tag or it is empty");
								rowN.createCell(4).setCellValue("Fail");
								continue;
							}
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
											By.xpath(".//div[@class='topic__headings']/div/div/h1")));
									WebElement TopicTitle = wd
											.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(1).setCellValue(TopicTitle.getText());
									// START TITLE MATCH CHECK
									String expectedTitle = TopicTitle.getText();
									if(actualTitle.equalsIgnoreCase(expectedTitle)) 
									{
										System.out.println("Titles Match");
										rowN.createCell(3).setCellValue("Pass");
									}else {
										System.out.println("Titles DO NOT Match");
										rowN.createCell(3).setCellValue("Fail");
									}
									//END MATCH CHECK
									rowN.createCell(4).setCellValue("Pass");
									Thread.sleep(1000);
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(4).setCellValue("Fail");
									
								}
								
								try {
									wd.close();
									wd.switchTo().window(tabs.get(0));
								} catch (Exception e) {
									Thread.sleep(3000);
									System.out.println("Cannot Close Window");
								}
								
								
								
								
								
								
								
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								
							}
						} catch (Exception e) {
							System.out.println("Location Topic Title: NOT FOUND");
							rowN.createCell(4).setCellValue("Fail");
						} // END get location, verify links
					} catch (Exception e) {
						System.out.println("Can't Find Injures");
					}
					FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Emergencies.xlsx");
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
				Thread.sleep(10000);
			}
		}
		
		
		// VERIFY Symptoms CV
				public void verifySymptoms_CV() throws Exception {
					wd.manage().window().setSize(new Dimension(450, 840));
					wd.manage().window().setPosition(new Point(900, 0));
					wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
					String currentURL = wd.getCurrentUrl();
					WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(3));
					Actions action = new Actions(wd);
					// Close Cookies
					try {
						WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
						AcceptCookies.click();
					} catch (Exception e) {
						System.out.println("Can't Close Cookies");
					}
					// End Close Cookies
					String Version = wd.findElement(By.xpath(".//meta[3]")).getAttribute("lang");
					
					try {
						// Get Date
						DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
						Date date = new Date();
						String date1 = dateFormat.format(date);
						// Excel Create Headings
						XSSFSheet sheet = wb.createSheet(Version);
						Row rowHeading0 = sheet.createRow(0);
						Row rowHeading1 = sheet.createRow(1);
						Row rowHeading2 = sheet.createRow(2);
						Row rowHeading3 = sheet.createRow(3);
						rowHeading0.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
						rowHeading1.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
						rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
						rowHeading3.createCell(0).setCellValue("Title on First Aid page");
						rowHeading3.createCell(1).setCellValue("Title on Topic Page");
						rowHeading3.createCell(2).setCellValue("URL");
						rowHeading3.createCell(3).setCellValue("URL ERROR (PASS/FAIL)");
						
						
						int sectionCount = wd.findElements(By.xpath(".//*[@id='maincontent']/section//section" )).size();
						System.out.println("We have Sections: " + sectionCount);
						
						int rowNum = 4;
						//Thread.sleep(3000);
						for (int s = 1; s < sectionCount + 1; s++) {
							int rowCount = wd.findElements(By.xpath(".//*[@id='maincontent']/section//section[" + s + "]/div[2]/div/ul/li")).size();
							System.out.println("Rows in this section: " + rowCount);
							for (int i = 1; i < rowCount + 1; i++) {
								// Main For Loop
								
								try {
									Row rowN = sheet.createRow(rowNum++);
									WebElement SymptomTitle = wd.findElement(By.xpath(".//*[@id='maincontent']/section//section["+s+"]/div[2]/div[2]/ul[@class='symptoms__container--row']/li["+i+"]/a"));
									//String actualTitle = SymptomTitle.getText();
									System.out.println("Symptom Title: " + SymptomTitle.getText());
									rowN.createCell(0).setCellValue(SymptomTitle.getText());
									
									// Get location, verify links
									try {
										String linkUrl = SymptomTitle.getAttribute("href");
										System.out.println("Location URL: " + linkUrl);
										rowN.createCell(2).setCellValue(linkUrl);
										if (linkUrl == null || linkUrl.isEmpty()) {
											System.out.println("URL is either not configured for anchor tag or it is empty");
											rowN.createCell(3).setCellValue("Fail");
											continue;
										}
										// VERIFY LINKS ARE ACTIVE
										try {
											((JavascriptExecutor) wd).executeScript("window.open()");
											Thread.sleep(1000);
											ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
											try {
												wd.switchTo().window(tabs.get(1));
												wd.get(linkUrl);
												System.out.println("Got title");
											} catch (Exception e) {
												System.out.println("Couldn't get title");
												Thread.sleep(3000);
												wd.switchTo().window(tabs.get(1));
												wd.get(linkUrl);
												System.out.println("Triet to get title again");
											}
											try {
												wait.until(ExpectedConditions.visibilityOfElementLocated(
														By.xpath(".//div[@class='topic__headings']/div/div/h1")));
												WebElement TopicTitle = wd
														.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
												System.out.println("Location Topic Title: " + TopicTitle.getText());
												rowN.createCell(1).setCellValue(TopicTitle.getText());
												/*
												// START TITLE MATCH CHECK
												String expectedTitle = TopicTitle.getText();
												if(actualTitle.equalsIgnoreCase(expectedTitle)) 
												{
													System.out.println("Titles Match");
													rowN.createCell(3).setCellValue("Pass");
												}else {
													System.out.println("Titles DO NOT Match");
													rowN.createCell(3).setCellValue("Fail");
												}
												// END TITLE MATCH CHECK
												*/
												rowN.createCell(3).setCellValue("Pass");
												Thread.sleep(1000);
											} catch (Exception e) {
												System.out.println("Location Topic Title: NOT FOUND");
											}
											wd.close();
											wd.switchTo().window(tabs.get(0));
										} catch (Exception e) {
											Thread.sleep(2000);
											System.out.println("URL: CANNOT VERIFY");
											
										}
										
									} catch (Exception e) {
										System.out.println("Location Topic Title: NOT FOUND");
										rowN.createCell(3).setCellValue("Fail");
									} // END get location, verify links
									
								} catch (Exception e) {
									System.out.println("Can't Find Symptoms");
								}
								FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Symptoms_CV.xlsx");
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
						Thread.sleep(10000);
					}
				}
		
				// VERIFY Symptoms_PV
				public void verifySymptoms_PV() throws Exception {
					wd.manage().window().setSize(new Dimension(450, 840));
					wd.manage().window().setPosition(new Point(900, 0));
					wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
					String currentURL = wd.getCurrentUrl();
					WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(3));
					Actions action = new Actions(wd);
					
					String Version = wd.findElement(By.xpath(".//meta[3]")).getAttribute("lang");
					
					try {
						// Get Date
						DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
						Date date = new Date();
						String date1 = dateFormat.format(date);
						// Excel Create Headings
						XSSFSheet sheet = wb.createSheet(Version);
						Row rowHeading0 = sheet.createRow(0);
						Row rowHeading1 = sheet.createRow(1);
						Row rowHeading2 = sheet.createRow(2);
						Row rowHeading3 = sheet.createRow(3);
						rowHeading0.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
						rowHeading1.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
						rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
						rowHeading3.createCell(0).setCellValue("Title on First Aid page");
						rowHeading3.createCell(1).setCellValue("Title on Topic Page");
						rowHeading3.createCell(2).setCellValue("URL");
						rowHeading3.createCell(3).setCellValue("URL ERROR (PASS/FAIL)");
						/*
						rowHeading3.createCell(5).setCellValue("DE HREF");
						rowHeading3.createCell(6).setCellValue("ES HREF");
						rowHeading3.createCell(7).setCellValue("FR HREF");
						rowHeading3.createCell(8).setCellValue("IT HREF");
						rowHeading3.createCell(9).setCellValue("JA HREF");
						rowHeading3.createCell(10).setCellValue("PT HREF");
						rowHeading3.createCell(11).setCellValue("RU HREF");
						rowHeading3.createCell(12).setCellValue("VI HREF");
						rowHeading3.createCell(13).setCellValue("CN HREF");
						 */
						Thread.sleep(3000);
						int sectionCount = wd.findElements(By.xpath(".//section/main/section")).size();
						System.out.println("We have Sections: " + sectionCount);
						
						int rowNum = 4;
						//Thread.sleep(3000);
						for (int s = 1; s < sectionCount + 1; s++) {
							int rowCount = wd.findElements(By.xpath(".//section/main/section["+ s +"]/div[2]/div[2]/ul/li")).size();
							System.out.println("Rows in this section: " + rowCount);
							for (int i = 1; i < rowCount + 1; i++) {
								// Main For Loop
								
								try {
									Row rowN = sheet.createRow(rowNum++);
									WebElement SymptomTitle = wd.findElement(By.xpath(".//main/section["+s+"]/div[2]/div[2]/ul[@class='symptoms__container--row']/li["+i+"]/a"));
									String actualTitle = SymptomTitle.getText();
									System.out.println("Symptom Title: " + SymptomTitle.getText());
									rowN.createCell(0).setCellValue(SymptomTitle.getText());
										
									// Get location, verify links
									try {
										String linkUrl = SymptomTitle.getAttribute("href");
										System.out.println("Location URL: " + linkUrl);
										rowN.createCell(2).setCellValue(linkUrl);
										if (linkUrl == null || linkUrl.isEmpty()) {
											System.out.println("URL is either not configured for anchor tag or it is empty");
											rowN.createCell(3).setCellValue("Fail");
											continue;
										}
										
										// VERIFY LINKS ARE ACTIVE
										try {
											((JavascriptExecutor) wd).executeScript("window.open()");
											Thread.sleep(1000);
											ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
											try {
												wd.switchTo().window(tabs.get(1));
												wd.get(linkUrl);
												System.out.println("Got title");
											} catch (Exception e) {
												System.out.println("Couldn't get title");
												Thread.sleep(3000);
												wd.switchTo().window(tabs.get(1));
												wd.get(linkUrl);
												System.out.println("Triet to get title again");
											}
											try {
												wait.until(ExpectedConditions.visibilityOfElementLocated(
														By.xpath(".//div[@class='topic__headings']/div/div/h1")));
												WebElement TopicTitle = wd
														.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
												System.out.println("Location Topic Title: " + TopicTitle.getText());
												rowN.createCell(1).setCellValue(TopicTitle.getText());
												/*
												// START TITLE MATCH CHECK
												String expectedTitle = TopicTitle.getText();
												if(actualTitle.equalsIgnoreCase(expectedTitle)) 
												{
													System.out.println("Titles Match");
													rowN.createCell(3).setCellValue("Pass");
												}else {
													System.out.println("Titles DO NOT Match");
													rowN.createCell(3).setCellValue("Fail");
												}
												// END TITLE MATCH CHECK
												 */
												rowN.createCell(3).setCellValue("Pass");
												Thread.sleep(1000);
												/*
												//START HREFs
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='de']"));
													System.out.println("DE HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(5).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(5).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='es']"));
													System.out.println("ES HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(6).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(6).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='fr']"));
													System.out.println("FR HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(7).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(7).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='it']"));
													System.out.println("IT HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(8).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(8).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
													System.out.println("JA HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(9).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(9).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='pt']"));
													System.out.println("PT HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(10).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(10).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='ru']"));
													System.out.println("RU HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(11).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(11).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='vi']"));
													System.out.println("VI HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(12).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(12).setCellValue("N/A");
												}
												try {
													WebElement HrefLinkAR = wd.findElement(By.xpath(".//link[@hreflang='zh-CN']"));
													System.out.println("CN HREF Link: " + HrefLinkAR.getAttribute("href"));
													rowN.createCell(13).setCellValue(HrefLinkAR.getAttribute("href"));
												} catch (Exception e) {
													System.out.println("Couldn't get HREF");
													rowN.createCell(13).setCellValue("N/A");
												}
												//END HREFs
												 */
											} catch (Exception e) {
												System.out.println("Location Topic Title: NOT FOUND");
											}
											wd.close();
											wd.switchTo().window(tabs.get(0));
										} catch (Exception e) {
											Thread.sleep(2000);
											System.out.println("URL: CANNOT VERIFY");
											
										}
										
									} catch (Exception e) {
										System.out.println("Location Topic Title: NOT FOUND");
										rowN.createCell(3).setCellValue("Fail");
									} // END get location, verify links
									

								} catch (Exception e) {
									System.out.println("Can't Find Symptoms");
								}
								
								FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Symptoms_PV.xlsx");
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
						Thread.sleep(10000);
					}
				}
		
		

		
}