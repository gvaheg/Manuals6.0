package Other;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class CombinedScript2 {

    public static void main(String[] args) {
        System.out.println("Starting combined script...");

        // Set up Firefox WebDriver in headless mode
        System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
        FirefoxOptions options = new FirefoxOptions();
        options.addArguments("--headless");

        WebDriver driver = new FirefoxDriver(options);
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
        System.out.println("WebDriver initialized.");

        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("URL Status & Merck Instances");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("URL");
        headerRow.createCell(1).setCellValue("Status Code");
        headerRow.createCell(2).setCellValue("Status Message");
        headerRow.createCell(3).setCellValue("Path 1");

        // Specify output file path and ensure the directory exists
        String outputFilePath = "C:\\TestResults\\URL_Status_and_Merck.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.getParentFile().mkdirs();

        int rowNum = 1;
        int maxPaths = 1;

        try {
            // Retrieve URLs from sitemap
            List<String> urls = getUrlsFromSitemap("https://www.msdmanuals.com/sitemap.ashx");
            System.out.println("URLs retrieved from sitemap: " + urls.size());

            for (String url : urls) {
                System.out.println("Processing URL: " + url);

                int statusCode = -1;
                String statusMessage = "";
                List<String> xpaths = new ArrayList<>();

                try {
                    statusCode = getStatusCodeAndCheckPage(driver, url, xpaths);
                    statusMessage = (statusCode == 200) ? "OK" : "Error or Forbidden";
                } catch (TimeoutException e) {
                    statusMessage = "Timeout";
                    System.err.println("Timeout occurred for URL: " + url);
                } catch (WebDriverException e) {
                    statusMessage = "WebDriver Error";
                    System.err.println("WebDriver exception for URL: " + url + " - " + e.getMessage());
                } catch (Exception e) {
                    statusMessage = "Other Error";
                    System.err.println("Error processing URL: " + url + " - " + e.getMessage());
                }

                maxPaths = Math.max(maxPaths, xpaths.size());

                // Write results to Excel
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(url);
                row.createCell(1).setCellValue(statusCode);
                row.createCell(2).setCellValue(statusMessage);

                int colNum = 3;
                for (String xpath : xpaths) {
                    row.createCell(colNum++).setCellValue(xpath);
                }
                System.out.println("Results written for URL: " + url);

                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
                }

                Thread.sleep(3000); // Delay to avoid rate limiting
            }

            // Adjust headers for additional paths
            for (int i = 1; i < maxPaths; i++) {
                headerRow.createCell(3 + i).setCellValue("Path " + (i + 1));
            }

            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
            }

            System.out.println("Excel file created successfully at " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            driver.quit();
            System.out.println("WebDriver closed.");
        }
    }

    private static int getStatusCodeAndCheckPage(WebDriver driver, String url, List<String> xpaths) {
        int statusCode = 200; // Default to 200 for Selenium-based requests
        try {
            driver.get(url);

            JavascriptExecutor js = (JavascriptExecutor) driver;

            // Hide footer to avoid irrelevant matches
            try {
                WebElement footer = driver.findElement(By.cssSelector("[data-testid='footer']"));
                js.executeScript("arguments[0].style.display='none';", footer);
            } catch (Exception e) {
                System.out.println("Footer not found for: " + url);
            }

            // Extract XPaths for "Merck" instances
            xpaths.addAll((List<String>) js.executeScript(
                "function getXPath(element) {" +
                "    if (element.id) {" +
                "        return 'id(\"' + element.id + '\")';" +
                "    }" +
                "    if (element === document.body) {" +
                "        return '/html/body';" +
                "    }" +
                "    if (!element.parentNode) {" +
                "        return ''; " + // Skip if parentNode is null
                "    }" +
                "    var ix = 0;" +
                "    var siblings = element.parentNode.childNodes;" +
                "    for (var i = 0; i < siblings.length; i++) {" +
                "        var sibling = siblings[i];" +
                "        if (sibling === element) {" +
                "            return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';" +
                "        }" +
                "        if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {" +
                "            ix++;" +
                "        }" +
                "    }" +
                "}" +
                "return Array.from(document.querySelectorAll('*'))" +
                "    .filter(e => e.textContent.toLowerCase().includes('merck') &&" +
                "        (!e.children || Array.from(e.children).every(child => !child.textContent.toLowerCase().includes('merck'))))" +
                "    .map(getXPath);"
            ));

        } catch (Exception e) {
            System.err.println("Error fetching page content for: " + url + " - " + e.getMessage());
            statusCode = -1;
        }
        return statusCode;
    }

    private static List<String> getUrlsFromSitemap(String sitemapUrl) {
        List<String> urls = new ArrayList<>();
        try {
            HttpURLConnection connection = (HttpURLConnection) new URL(sitemapUrl).openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("User-Agent", "Mozilla/5.0");
            connection.setConnectTimeout(10000);
            connection.setReadTimeout(10000);

            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(new String(connection.getInputStream().readAllBytes()))));
            NodeList urlNodes = doc.getElementsByTagName("loc");

            for (int i = 0; i < urlNodes.getLength(); i++) {
                urls.add(urlNodes.item(i).getTextContent());
            }
            connection.disconnect();
        } catch (Exception e) {
            System.err.println("Error fetching sitemap: " + e.getMessage());
        }
        return urls;
    }
}
