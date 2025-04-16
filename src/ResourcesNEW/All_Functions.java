package ResourcesNEW;


import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.platform.commons.function.Try;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.UnreachableBrowserException;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.netty.handler.timeout.TimeoutException;

public class All_Functions {

	private static final boolean Figures = false;
    // START DRIVER
    static WebDriver wd;
    static XSSFWorkbook wb;

    // Setter for WebDriver
    public static void setWebDriver(WebDriver driver) {
        wd = driver;
        
    }
    
   	// VERIFY TABLES
       public void verifyTables() throws Exception {
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
               rowHeading4.createCell(HeadN++).setCellValue("NONE");
               rowHeading4.createCell(HeadN++).setCellValue("NONE");
               rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
               rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
               rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
               rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
               excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
               CloseCookies();
               int rowNum = 5;

               for (int s = 2; s < 36; s++) {
                   // SECTIONS>
                   String sectionText = null;
                   while (true) {
                       try {
                           if (s > 36) {
                               break; // Exit the loop if s exceeds 34
                           }
                           WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                           sectionText = section.getAttribute("innerText"); // Store section text
                           break; // Exit the loop if the section is found
                       } catch (Exception e) {
                           System.out.println("Can't Find Section for s = " + s);
                           s++; // Increment s if the section is not found
                       }
                   }
                   // <END SECTIONS
                   int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                   System.out.println("Rows in this section: " + rowCount);
                   // Main For Loop
                   for (int i = 1; i < rowCount + 1; i++) {
                       try {
                           Row rowN = sheet.createRow(rowNum++);
                           // Section Text to Excel
                           rowN.createCell(0).setCellValue(sectionText);
                           System.out.println("Section: " + sectionText);
                           rowN.createCell(1).setCellValue(i);
                           // Run Main Codes
                           openResource(rowN, i, s);
                        // ShowDetails
                   			try {
                   				WebElement showDetails = null;
                   				try {
                   					showDetails = wd.findElement(By.xpath("//span[@class='PopupTable_showDetails__WDJpC']"));
                   				} catch (Exception e) {
                   					Thread.sleep(1000);
                   					showDetails = wd.findElement(By.xpath("//*[@id='iframepopupundefined']/div[2]/div[2]/div[1]/div[1]/span[2]"));
                   				}
                   				if (showDetails != null) {
                   				showDetails.click();
                   				}
                   			} catch (Exception ex) {
                   				Thread.sleep(2000);
                   				System.out.println("Can't CLICK ON SHOW DETAILS");
                   			}
                   			
                   			try {
                   			WebElement location = wd.findElement(By.xpath(".//a[@class='InThisTopic_location__VJhFE ']"));
                   			System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
                   		 
                   		        if (location != null) {
                   		            // Element found, proceed with operations
                   		            // Save location details in Excel
                   		            rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
                   		            rowN.createCell(6).setCellValue("Pass");
                   		            String linkUrl = location.getAttribute("href");
                   		            rowN.createCell(7).setCellValue(linkUrl);
                   		            System.out.println(linkUrl);

                   		            // Verify the location URL with an HTTP request
                   		            boolean isURLValid = verifyURL(linkUrl);

                   		            if (isURLValid) {
                   		                System.out.println("URL is valid.");
                   		                rowN.createCell(8).setCellValue("OK");
                   		            } else {
                   		                System.out.println("URL is not valid.");
                   		                rowN.createCell(8).setCellValue("Not Found");

                   		                // Highlight cell in red for invalid URL
                   		                highlightCellRed(sheet, rowN.getCell(8));
                   		            }
                   		        } else {
                   		            // Element not found with either xpath
                   		            System.out.println("Location not found.");
                   		            rowN.createCell(6).setCellValue("Fail");

                   		            // Highlight cell in red for location failure
                   		            highlightCellRed(sheet, rowN.getCell(6));
                   		        }
                   		        
                   		    } catch (Exception e) {
                   		        System.out.println("Location " + i + ": Fail");
                   		        rowN.createCell(6).setCellValue("Fail");

                   		        // Highlight cell in red for location failure
                   		        highlightCellRed(sheet, rowN.getCell(6));
                   		    }
                           closePopup();
                       } catch (Exception e) {
                           System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                           FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Tables.xlsx");
                           wb.write(fout);
                           fout.close();
                           System.out.println("Written Successfully!");
                       } catch (Exception e) {
                           System.out.println("Cannot SAVE File!");
                       }
                   }
               }

           } catch (Exception e) {
               System.out.println("ERROR! Web page is not responding. Reopening the page... ");
               ErrorMain(currentURL, e);
           }
       }
    
 // VERIFY QUIZZES
    public void verifyQuizzes() throws Exception {
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
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            rowHeading4.createCell(HeadN++).setCellValue("CHOICE 1");
            rowHeading4.createCell(HeadN++).setCellValue("CHOICE 2");
            rowHeading4.createCell(HeadN++).setCellValue("CHOICE 3");
            rowHeading4.createCell(HeadN++).setCellValue("CHOICE 4");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                	    // OPEN RESOURCE
                	    try {
                	        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
                	        WebElement titleElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                	            ".//main/div[1]//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"
                	        )));
                	        String titleText = titleElement.getAttribute("innerText");
                	        System.out.println("Title " + i + ": " + titleText);

                	        // Write ID and Title
                	        rowN.createCell(1).setCellValue(i);
                	        rowN.createCell(2).setCellValue(titleText);

                	        // Scroll to the resource item and click it
                	        JavascriptExecutor js = (JavascriptExecutor) wd;
                	        js.executeScript("arguments[0].scrollIntoView({block: 'center'});", titleElement);
                	        Thread.sleep(500); // Small pause to ensure scrolling completes

                	        int retries = 3; // Retry mechanism
                	        while (retries > 0) {
                	            try {
                	                titleElement.click();
                	                System.out.println("Clicked on resource item " + i);
                	                break; // Exit loop if successful
                	            } catch (Exception e) {
                	                System.out.println("Failed to click on resource item " + i + ", retrying...");
                	                retries--;
                	                Thread.sleep(1000);
                	            }
                	        }
                	        if (retries == 0) {
                	            throw new Exception("Failed to click on resource item " + i + " after retries");
                	        }

                	    } catch (Exception e) {
                	        System.out.println("Failed to click on resource item " + i);
                	        rowN.createCell(1).setCellValue(i);
                	        rowN.createCell(2).setCellValue("CANNOT GET TITLE");
                	        wd.navigate().refresh();
                	    }
                        
                        
                        
                        
                        
                        
                        

                     // SWITCH TO IFRAME
						try {
							//wait50.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("iframe[class='IFrames_Iframe__WVuGl']")));
							Thread.sleep(1000);
							WebElement iFrameCalc = wd.findElement(By.cssSelector("iframe[class='IFrames_Iframe__WVuGl']"));
							wd.switchTo().frame(iFrameCalc);
						} catch (Exception e) {
							System.out.println("Can't Switch");
						}
						
						// Test Choices
						try {
							String choice1 = wd.findElement(By.xpath("//*[@id='choices']/li[1]/div[2]")).getText();
							System.out.println("Choice 1: " + choice1);
							rowN.createCell(7).setCellValue(choice1);
						} catch (Exception e) {
							System.out.println("No options");
						}
						try {
							String choice2 = wd.findElement(By.xpath("//*[@id='choices']/li[2]/div[2]")).getText();
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
							WebElement moreinfo = wd.findElement(By.xpath(".//*[@id='reference_text']//a"));
							System.out.println("Location " + i + ": " + moreinfo.getText());
							rowN.createCell(3).setCellValue(moreinfo.getText());
							rowN.createCell(4).setCellValue("Pass");
							String linkUrl = moreinfo.getAttribute("href");
							System.out.println("Location URL: " + linkUrl);
							rowN.createCell(5).setCellValue(linkUrl);

							// VERIFY QUIZ LINKS ARE ACTIVE
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
								CloseCookies();
								try {
									try {
										wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1[@class='readable TopicHead_topicHeaderTittle__miyQz ']")));
									} catch (Exception e) {
										System.out.println("Refreshing the page!");
										wd.navigate().refresh();
										wait20.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1[@class='readable TopicHead_topicHeaderTittle__miyQz ']")));
									}
									WebElement TopicTitle = wd
											.findElement(By.xpath("//h1[@class='readable TopicHead_topicHeaderTittle__miyQz ']"));
									System.out.println("Location Topic Title: " + TopicTitle.getText());
									rowN.createCell(6).setCellValue("OK");
								} catch (Exception e) {
									System.out.println("Location Topic Title: NOT FOUND");
									rowN.createCell(6).setCellValue("Not Found");
								}
								wd.close();
								wd.switchTo().window(tabs.get(0));
							} catch (Exception e) {
								Thread.sleep(2000);
								System.out.println("URL: CANNOT VERIFY");
								rowN.createCell(6).setCellValue("CANNOT VERIFY");
							}
							
						} catch (Exception e) {
							System.out.println("Location " + i + ": Fail");
							rowN.createCell(6).setCellValue("Fail");
						} // END get location, verify links

                        
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Quizzes.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
    
    
 // VERIFY MODELS 3D
    public void verifyModels3D() throws Exception {
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
            rowHeading4.createCell(HeadN++).setCellValue("3D MODEL EXIST");
            rowHeading4.createCell(HeadN++).setCellValue("NONE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            rowHeading4.createCell(HeadN++).setCellValue("CALCULATOR TITLE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                     // test 3d model exist
						try {
							WebElement ModelLink = wd
									.findElement(By.xpath(".//div[@class='IFramePopupContent_iframebiodigital__szAqO']/iframe"));
							System.out.println(i + ": 3D Model EXIST");
							rowN.createCell(3).setCellValue("YES");
						} catch (Exception e) {
							System.out.println(i + ": CANNOT FIND 3D MODEL");
							rowN.createCell(3).setCellValue("NO");
						}
                        ShowDetails();
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Models3D.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
    
    
 // VERIFY CALCULATORS
    public void verifyCalculators() throws Exception {
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
            rowHeading4.createCell(HeadN++).setCellValue("CALCULATOR URL");
            rowHeading4.createCell(HeadN++).setCellValue("CALCULATOR ERROR");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            rowHeading4.createCell(HeadN++).setCellValue("CALCULATOR TITLE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        

						// SWITCH TO IFRAME
						int z = i-1;
						try {
						    WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
						    // Wait for the iframe to be present
						    WebElement iFrameCalc = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#iframepopup" + z + " > div.IFramePopupContent_spaceForCarousalArrow__Wp4db > div.IFramePopupContent_content__WSH_v > div > iframe")));
						    wd.switchTo().frame(iFrameCalc);
						    // Optional small wait to ensure contents inside iframe are loaded
						    new WebDriverWait(wd, Duration.ofSeconds(5)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[@class='medCalcFontTitleBox']")));

						} catch (Exception e) {
						    System.out.println("Can't Switch");
						}

						try {
						    WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
						    // Wait for the calculator title to be visible
						    WebElement CalcTitle = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@class='medCalcFontTitleBox']")));
						    System.out.println("Popup Title: " + CalcTitle.getAttribute("innerText") + ", Calculator: PASS");
						    rowN.createCell(9).setCellValue(CalcTitle.getAttribute("innerText"));
						    rowN.createCell(4).setCellValue("OK");

						} catch (Exception e) {
						    System.out.println("Calculator: FAIL (ERROR PAGE)");
						    rowN.createCell(9).setCellValue("");
						    rowN.createCell(4).setCellValue("ERROR PAGE");
						}
						
						try {
							wd.switchTo().defaultContent();
						} catch (Exception e) {
							System.out.println("Can't Switch TO DEFAULT");
						}

						// GET CALCULATOR URL
						try {
							WebElement CalcLink = wd.findElement(By.cssSelector("#iframepopup"+z+ "> div.IFramePopupContent_spaceForCarousalArrow__Wp4db > div.IFramePopupContent_content__WSH_v > div > iframe"));
							System.out.println("Calculator URL: " + CalcLink.getAttribute("src"));
							rowN.createCell(3).setCellValue(CalcLink.getAttribute("src"));

						} catch (Exception e) {
							System.out.println("Calculator URL: NOT FOUND");
							rowN.createCell(3).setCellValue("NOT FOUND");
						}
                        
                        
                        ShowDetails();
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
    
    
    
 // VERIFY VIDEOS
    public void verifyVideos() throws Exception {
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
            rowHeading4.createCell(HeadN++).setCellValue("VIDEO NAME");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        
                        // GET VIDEO NAME
                		// Assume titleNumber is set for the current title (e.g., 1 for Title 1, 2 for Title 2, etc.)
						int titleNumber = i;  
						int frameIndex = titleNumber - 1;  // Adjust because list indices start at 0
						// Get all iframes with the common locator
						List<WebElement> frames = wd.findElements(By.cssSelector("iframe.IFrames_Iframe__WVuGl"));
						System.out.println("Found " + frames.size() + " iFrame(s).");
						String fileName = "";
						if (frames.size() > frameIndex) {
						    try {
						        // Switch back to the main document first
						        wd.switchTo().defaultContent();
						        // Use an explicit wait (10 seconds; adjust if needed) and switch into the iframe corresponding to the title number
						        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
						        wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(frames.get(frameIndex)));
						        // Since the <title> element is inside the <head> (and not visible), wait for its presence
						        WebElement titleElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html/head/title")));
						        // Wait until the title element's text is non-empty (this handles timing issues)
						        wait.until(driver -> {
						            String text = titleElement.getAttribute("textContent");
						            return text != null && !text.trim().isEmpty();
						        });
						        // Try to get the file name using textContent; if empty, try innerHTML as a fallback
						        fileName = titleElement.getAttribute("textContent").trim();
						        if (fileName.isEmpty()) {
						            fileName = titleElement.getAttribute("innerHTML").trim();
						        }
						        System.out.println("Title " + titleNumber + " File Name: " + fileName);
						    } catch (Exception e) {
						        System.out.println("Error retrieving file name for Title " + titleNumber + ": " + e.getMessage());
						    } finally {
						        // Always switch back to default content before moving on
						        wd.switchTo().defaultContent();
						    }
						} else {
						    System.out.println("Not enough iframes for Title " + titleNumber + ". Using fallback logic.");
						    // Optionally, add fallback logic here (for example, iterating over frames or using a different locator)
						}
						
						
						// Now write the file name into the Excel row, column index 9 (10th column)
						rowN.createCell(9).setCellValue(fileName);
                		getCredits(rowN, i);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Videos.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
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

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
 

	// VERIFY IMAGES
    public void verifyImages() throws Exception {
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

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Images.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
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
            //rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
            //rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
            //rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        //getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
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
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Sounds.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
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
	
		public void excelRows(XSSFSheet sheet, int sectionCount, int TotalCount, int HeadN, Row rowHeading4) {
		    try {
		        // Get Date
		        DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		        Date date = new Date();
		        String date1 = dateFormat.format(date);

		        // Main Excel Rows for all
		        Row rowHeading0 = sheet.createRow(0);
		        Row rowHeading1 = sheet.createRow(1);
		        Row rowHeading2 = sheet.createRow(2);
		        Row rowHeading3 = sheet.createRow(3);
		        rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
		        rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
		        rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
		        rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
		        rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());
		    } catch (Exception e) {
		        System.out.println("Cannot Create Excel Rows: " + e.getMessage());
		    }
		}


		
	
	// Open Each Resource Item
	public void openResource(Row rowN, int i, int s) throws InterruptedException {
	    int maxTestsPerSession = 30;
	    if (i % maxTestsPerSession == 0) {
	        String currentURL = wd.getCurrentUrl();
	     // Quit the current driver
	        wd.quit();

	        // Initialize a new Firefox driver in headless mode
	        FirefoxOptions options = new FirefoxOptions();
	        if (i >= maxTestsPerSession) {
	            System.out.println("RELOADING TO CONTINUE...");
	           options.setHeadless(true);
	        } else {
	            System.out.println("CANNOT RELOAD TO CONTINUE...");
	           options.setHeadless(false);
	        }

	        // Restart the driver
	        wd = new FirefoxDriver(options);
	        Thread.sleep(1000);

	        // Navigate to the last URL
	        wd.get(currentURL);
	        Thread.sleep(2000);

	        // Perform any necessary actions to reset or prepare the session
	        ManageWindows();
	        handlePopups();
	    }

	    // Main Title
	    try {
	        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
	        WebElement titleElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
	            ".//main/div[1]//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"
	        )));
	        String titleText = titleElement.getAttribute("innerText");
	        System.out.println("Title " + i + ": " + titleText);

	        // Write ID and Title
	        rowN.createCell(1).setCellValue(i);
	        rowN.createCell(2).setCellValue(titleText);

	        // Scroll to the resource item and click it
	        JavascriptExecutor js = (JavascriptExecutor) wd;
	        js.executeScript("arguments[0].scrollIntoView({block: 'center'});", titleElement);
	        Thread.sleep(500); // Small pause to ensure scrolling completes

	        int retries = 3; // Retry mechanism
	        while (retries > 0) {
	            try {
	                titleElement.click();
	                System.out.println("Clicked on resource item " + i);
	                break; // Exit loop if successful
	            } catch (Exception e) {
	                System.out.println("Failed to click on resource item " + i + ", retrying...");
	                retries--;
	                Thread.sleep(1000);
	            }
	        }
	        if (retries == 0) {
	            throw new Exception("Failed to click on resource item " + i + " after retries");
	        }

	    } catch (Exception e) {
	        System.out.println("Failed to click on resource item " + i);
	        rowN.createCell(1).setCellValue(i);
	        rowN.createCell(2).setCellValue("CANNOT GET TITLE");
	        wd.navigate().refresh();
	    }
	}

	
	
	// Get Credits and Descriptions
	public void getCredits(Row rowN, int i) {

		// Get & write Description
		try {
			WebElement description = wd.findElement(By.xpath(
					"//div[@class='SliderMobile_description__7xhlL']/div/div/p"));
			System.out.println("Description: " + description.getAttribute("innerText"));
			rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("Description: CANT FIND");
			rowN.createCell(4).setCellValue("");
		}
		// Get & write Credits
		try {
			WebElement credits = wd
					.findElement(By.xpath("//div[@class='SliderMobile_credits___BXrj']/div/div/p"));
			System.out.println("Credits: " + credits.getAttribute("innerText"));
			rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("");
			rowN.createCell(3).setCellValue("");
		}

	}
	

	public void getFileName(Row rowN, XSSFSheet sheet) {
	    try {
	        // Increase timeout to handle slower image loading
	        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(20));

	        // Locate the image element
	        WebElement FileName = wait.until(ExpectedConditions.visibilityOfElementLocated(
	            By.cssSelector(".active.carousel-item .PopupContent_multiMediaImageContainer__wunjw > div.PopupContent_imageCarousel__uy4Mq > img")
	        ));

	        // Extract the URL from the src attribute
	        String srcAttribute = FileName.getAttribute("src");
	        URL FileURL = new URL(srcAttribute);

	        System.out.println("Src attribute is: " + FileURL);
	        System.out.println("Path = " + FileURL.getPath());

	        // Write image details to the Excel sheet
	        rowN.createCell(9).setCellValue(srcAttribute);
	        rowN.createCell(10).setCellValue(FileURL.getPath());

	        // VERIFY IMAGE EXISTS
	        if (verifyImageURLWithRetries(FileURL)) {
	            System.out.println("Image: FOUND");
	            rowN.createCell(11).setCellValue("FOUND");
	        } else {
	            System.out.println("Image: NOT FOUND");
	            rowN.createCell(11).setCellValue("NOT FOUND");
	            highlightCellRed(sheet, rowN.getCell(11));
	        }
	    } catch (Exception e) {
	        System.out.println("CANNOT GET FILE NAME: " + e.getMessage());
	        rowN.createCell(9).setCellValue("No File PATH");
	        rowN.createCell(10).setCellValue("No File URL");
	        rowN.createCell(11).setCellValue("NOT FOUND");
	        highlightCellRed(sheet, rowN.getCell(11));
	    }
	}

	private boolean verifyImageURLWithRetries(URL fileURL) {
	    int retries = 3; // Number of retry attempts
	    while (retries > 0) {
	        try {
	            // Introduce a small delay between retries
	            Thread.sleep(500);

	            // Open a connection to the URL
	            HttpURLConnection connection = (HttpURLConnection) fileURL.openConnection();

	            // Set headers to mimic a browser
	            connection.setRequestMethod("GET");
	            connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
	            connection.setRequestProperty("Accept", "image/webp,*/*");
	            connection.setConnectTimeout(5000);
	            connection.setReadTimeout(5000);
	            connection.setInstanceFollowRedirects(true); // Follow redirects

	            // Get the response code
	            int responseCode = connection.getResponseCode();

	            if (responseCode >= 200 && responseCode < 300) {
	                System.out.println("Image exists with response code: " + responseCode);
	                return true; // Image exists
	            } else {
	                System.out.println("Image check failed with response code: " + responseCode);
	            }
	        } catch (IOException e) {
	            System.out.println("Error verifying URL: " + e.getMessage());
	        } catch (InterruptedException ie) {
	            Thread.currentThread().interrupt();
	            System.out.println("Thread interrupted during retry.");
	        }
	        retries--;
	    }
	    return false; // Image verification failed after retries
	}



	// Highlight a cell in red
	private void highlightCellRed(XSSFSheet sheet, Cell cell) {
	    CellStyle style = sheet.getWorkbook().createCellStyle();
	    style.setFillForegroundColor(IndexedColors.RED.getIndex());
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    cell.setCellStyle(style);
	}

	// Highlight a cell in yellow
	private void highlightCellYellow(XSSFSheet sheet, Cell cell) {
	    CellStyle style = sheet.getWorkbook().createCellStyle();
	    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    cell.setCellStyle(style);
	}


	// Helper method to verify the image URL
	private boolean verifyImageURL(URL fileURL) {
	    try {
	        HttpURLConnection connection = (HttpURLConnection) fileURL.openConnection();
	        connection.setRequestMethod("HEAD");
	        connection.setConnectTimeout(5000); // Timeout for faster verification
	        connection.setReadTimeout(5000);
	        int responseCode = connection.getResponseCode();

	        // Check if the response code indicates the image exists
	        if (responseCode >= 200 && responseCode < 300) {
	            return true;
	        }
	    } catch (Exception e) {
	        System.out.println("Error verifying image URL: " + e.getMessage());
	    }
	    return false;
	}

	
	// ShowDetails
		public void ShowDetails() throws InterruptedException {
			try {
				WebElement showDetails = null;
				try {
					showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
				} catch (Exception e) {
					Thread.sleep(1000);
					showDetails = wd.findElement(By.xpath("//*[@id='iframepopupundefined']/div[2]/div[2]/div[1]/div[1]/span[2]"));
				}
				if (showDetails != null) {
				showDetails.click();
				}
			} catch (Exception ex) {
				Thread.sleep(2000);
				System.out.println("Can't CLICK ON SHOW DETAILS");
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
		Thread.sleep(2000);
		wd.get(currentURL);
		Thread.sleep(5000);
		CloseCookies();
	}
		

	// Get Location
	public void getLocation(Row rowN, XSSFSheet sheet, WebDriverWait wait10, WebDriverWait wait20, WebDriverWait wait30, int i, Row rowHeading4, int HeadN, int s) throws IOException, InterruptedException {
	    try {
	        WebElement location = locateElement(
	           
	            "//div[@class='accordionSection_carouselContainer__Y5UzM']//div/div[1]//div[2]/a",
	            "//div[2]/div[1]/div/div/ul/li/a",
	            "//a[@class='InThisTopic_location__VJhFE ']"
	            
	        );

	        if (location != null) {
	            // Element found, proceed with operations
	            System.out.println("Location " + i + ": " + location.getAttribute("innerText"));

	            // Save location details in Excel
	            rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
	            rowN.createCell(6).setCellValue("Pass");
	            String linkUrl = location.getAttribute("href");
	            rowN.createCell(7).setCellValue(linkUrl);
	            System.out.println(linkUrl);

	            // Verify the location URL with an HTTP request
	            boolean isURLValid = verifyURL(linkUrl);

	            if (isURLValid) {
	                System.out.println("URL is valid.");
	                rowN.createCell(8).setCellValue("OK");
	            } else {
	                System.out.println("URL is not valid.");
	                rowN.createCell(8).setCellValue("Not Found");

	                // Highlight cell in red for invalid URL
	                highlightCellRed(sheet, rowN.getCell(8));
	            }
	        } else {
	            // Element not found with either xpath
	            System.out.println("Location not found.");
	            rowN.createCell(6).setCellValue("Fail");

	            // Highlight cell in red for location failure
	            highlightCellRed(sheet, rowN.getCell(6));
	        }
	    } catch (Exception e) {
	        System.out.println("Location " + i + ": Fail");
	        rowN.createCell(6).setCellValue("Fail");

	        // Highlight cell in red for location failure
	        highlightCellRed(sheet, rowN.getCell(6));
	    }
	}

	
	
	
	private int waitTimeoutSeconds = 10; // Example timeout in seconds

	private WebElement locateElement(String primaryXPath, String fallbackXPath, String fallbackXPath2) {
	    WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(waitTimeoutSeconds)); // Use Duration
	    WebElement element = null;

	    try {
	        System.out.println("Attempting to find element with primary XPath: " + primaryXPath);
	        // Wait until the element located by the primary XPath is present in the DOM
	        element = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(primaryXPath)));
	        //System.out.println("Element found using primary XPath.");
	        return element;
	    } catch (Exception e) { // Catch TimeoutException or other potential exceptions from wait.until
	        System.out.println("Primary XPath failed: " + e.getMessage());
	        // Optionally log the exception e if more detail is needed
	    }

	    try {
	        System.out.println("Attempting to find element with fallback XPath: " + fallbackXPath);
	        // Wait for the fallback XPath
	        element = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(fallbackXPath)));
	        System.out.println("Element found using fallback XPath.");
	        return element;
	    } catch (Exception e) {
	        System.out.println("Fallback XPath failed: " + e.getMessage());
	    }

	    try {
	        System.out.println("Attempting to find element with second fallback XPath: " + fallbackXPath2);
	        // Wait for the second fallback XPath
	        element = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(fallbackXPath2)));
	        System.out.println("Element found using second fallback XPath.");
	        return element;
	    } catch (Exception e) {
	        System.out.println("Second fallback XPath failed: " + e.getMessage());
	    }

	    // If we reach here, none of the XPaths worked within the timeout
	    System.out.println("Element not found using any of the provided XPaths within the timeout period.");
	    return null; // Return null explicitly if nothing was found
	}

	/*
	// UPD 04/10/25 Helper method to verify the URL with an HTTP request
	private boolean verifyURL(String linkUrl) {
	    if (linkUrl == null || linkUrl.isEmpty()) return false;

	    try {
	        HttpURLConnection connection = (HttpURLConnection) new URL(linkUrl).openConnection();
	        connection.setRequestMethod("GET");  // use GET for better compatibility
	        connection.setInstanceFollowRedirects(true);
	        connection.setConnectTimeout(5000);
	        connection.setReadTimeout(5000);
	        connection.setRequestProperty("User-Agent", "Mozilla/5.0"); // mimic a real browser
	        connection.connect();

	        int responseCode = connection.getResponseCode();
	        System.out.println("GET response code: " + responseCode);

	        // Accept 2xx and 3xx (redirects)
	        return responseCode >= 200 && responseCode < 400;

	    } catch (Exception e) {
	        System.out.println("Error verifying URL: " + e.getMessage());
	        return false;
	    }
	}
*/

	
	
	

	// Add a helper to close popups
	private void closePopup() throws InterruptedException {
	    try {
	        wd.switchTo().defaultContent();
	        WebElement closeButton = wd.findElement(By.xpath(".//p[contains(@class,'modal_btnClose__eatUo') and contains(@class,'modal_headerElement__9HT5x')]"));
	        if (closeButton.isDisplayed()) {
	            closeButton.click();
	            System.out.println("Popup closed.");
	        } else {
	            System.out.println("Popup close button not visible.");
	        }
	    } catch (NoSuchElementException e) {
	        System.out.println("Popup close button not found.");
	    } catch (Exception e) {
	        System.out.println("Cannot Close Popup: " + e.getMessage());
	        Thread.sleep(2000);
	        ArrayList<String> tabs = new ArrayList<>(wd.getWindowHandles());
	        wd.switchTo().window(tabs.get(0));
	        Thread.sleep(2000);
	        wd.switchTo().defaultContent();
	        try {
	            wd.findElement(By.xpath(".//p[contains(@class,'modal_btnClose__eatUo') and contains(@class,'modal_headerElement__9HT5x')]")).click();
	        } catch (Exception innerEx) {
	            System.out.println("Failed to close popup after retry: " + innerEx.getMessage());
	        }
	    }
	}

    
    public void handlePopups() {
        try {
            System.out.println("Handling popups...");
            WebElement acceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
            acceptCookies.click();
            System.out.println("Cookies popup handled.");
        } catch (Exception e) {
            System.out.println("Cookies popup not found.");
        }

        try {
            WebElement languageSelector = wd.findElement(By.xpath("//*[@class='ChineseModalPopup_languageSelectorPopupVersionButton__j7M_0']"));
            languageSelector.click();
            System.out.println("Language selector popup handled.");
        } catch (Exception e) {
            System.out.println("Language selector popup not found.");
        }
    }
	
    public void verifyLanguageVersion(String url, String version, boolean hasPopup) {
        try {
            System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
            All_Functions.wd = wd; // Set the WebDriver instance in All_Functions
            System.out.println("Navigating to URL: " + url);
            wd.get(url);
            WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")));
            System.out.println("Page loaded successfully.");
            if (hasPopup) {
                handlePopups();
            }
           
            System.out.println("VERSION: " + version);
            verifyFigures();
        } catch (Exception e) {
            System.out.println("Page Error for version: " + version);
            e.printStackTrace();
        } finally {
            if (wd != null) {
                wd.close();
                System.out.println("Browser closed.");
            }
        }
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    // Method to verify the URL with retries and random user-agents to avoid 403s and 429s
    public boolean verifyURL(String linkUrl) {
        if (linkUrl == null || linkUrl.isEmpty()) {
            return false;
        }

        // A list of common user agents to mimic various browsers
        String[] userAgents = {
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
            "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:40.0) Gecko/20100101 Firefox/40.0",
            "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; AS; rv:11.0) like Gecko",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:49.0) Gecko/20100101 Firefox/49.0",
            "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
        };

        int maxRetries = 3;
        int retryCount = 0;
        boolean success = false;

        while (retryCount < maxRetries) {
            try {
                HttpURLConnection connection = (HttpURLConnection) new URL(linkUrl).openConnection();
                connection.setRequestMethod("GET");
                connection.setInstanceFollowRedirects(true);
                connection.setConnectTimeout(5000);
                connection.setReadTimeout(5000);
                connection.setRequestProperty("User-Agent", getRandomUserAgent(userAgents)); // Randomized user-agent

                connection.connect();

                int responseCode = connection.getResponseCode();
                System.out.println("GET response code: " + responseCode);

                if (responseCode >= 200 && responseCode < 400) {
                    success = true;
                    break; // Exit loop if successful
                }

                // Handle specific status codes like 403 or 429 (too many requests)
                if (responseCode == 403 || responseCode == 429) {
                    System.out.println("Received " + responseCode + ". Retrying...");
                    Thread.sleep(getRandomBackoffTime()); // Wait before retrying
                } else {
                    break; // Other errors, no retry needed
                }

            } catch (Exception e) {
                System.out.println("Error verifying URL: " + e.getMessage());
                break; // If exception occurs, break out of the loop
            }
            retryCount++;
        }

        return success;
    }

    // Utility method to get a random User-Agent
    private String getRandomUserAgent(String[] userAgents) {
        Random rand = new Random();
        return userAgents[rand.nextInt(userAgents.length)];
    }

    // Utility method to get a random backoff time in milliseconds
    private int getRandomBackoffTime() {
        Random rand = new Random();
        return 2000 + rand.nextInt(3000); // Random time between 2000ms and 5000ms
    }

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
	
}
