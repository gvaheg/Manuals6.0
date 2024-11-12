package Other;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.JavascriptExecutor;
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

public class Branding {

    public static void main(String[] args) {
        System.out.println("Starting test...");

        // Set up Firefox WebDriver in headless mode
        System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
        FirefoxOptions options = new FirefoxOptions();
        options.addArguments("--headless");
        WebDriver driver = new FirefoxDriver(options);
        System.out.println("WebDriver initialized.");

        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Merck Instances");

        // Create initial header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("URL");
        headerRow.createCell(2).setCellValue("Path 1");

        // Specify output file path and ensure the directory exists
        String outputFilePath = "C:\\TestResults\\Merck_Instances.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.getParentFile().mkdirs(); // Create directory if it doesn't exist

        int rowNum = 1; // Start from the second row since the first is the header
        int idCounter = 1; // Unique ID for each URL
        int maxPaths = 1; // Track the maximum number of "Path" columns needed

        try {
            // Set implicit wait using Duration
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

            // Retrieve URLs from sitemap
            List<String> urls = getUrlsFromSitemap("https://www.msdmanuals.com/sitemap.ashx");
            System.out.println("URLs retrieved from sitemap: " + urls.size());

            // Loop through each URL and check for "Merck"
            for (String url : urls) {
                System.out.println("Processing URL: " + url);
                driver.get(url);

                // Hide the footer from the search
                try {
                    WebElement footer = driver.findElement(By.cssSelector("[data-testid='footer']"));
                    ((JavascriptExecutor) driver).executeScript("arguments[0].style.display='none';", footer);
                    System.out.println("Footer hidden for URL: " + url);
                } catch (Exception e) {
                    System.out.println("Footer not found on page: " + url);
                }

                // Use JavaScript to find elements containing "Merck" and return their specific XPaths
                List<String> xpaths = (List<String>) ((JavascriptExecutor) driver).executeScript(
                    "function getXPath(element) {" +
                    "    if (element.id !== '') {" +
                    "        return 'id(\"' + element.id + '\")';" +
                    "    }" +
                    "    if (element === document.body) {" +
                    "        return '/html/body';" +
                    "    }" +
                    "    var index = 0;" +
                    "    var siblings = element.parentNode.childNodes;" +
                    "    for (var i = 0; i < siblings.length; i++) {" +
                    "        var sibling = siblings[i];" +
                    "        if (sibling === element) {" +
                    "            return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (index + 1) + ']';" +
                    "        }" +
                    "        if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {" +
                    "            index++;" +
                    "        }" +
                    "    }" +
                    "}" +
                    "var footer = document.querySelector('#__next > div:nth-child(1) > div:nth-child(2) > footer');" +
                    "var elements = Array.from(document.querySelectorAll('*'))" +
                    "    .filter(e => e.textContent.toLowerCase().includes('merck') &&" +
                    "        !Array.from(e.children).some(child => child.textContent.toLowerCase().includes('merck')) &&" +
                    "        (!footer || !footer.contains(e))" +
                    "    );" +
                    "return elements.map(e => getXPath(e));");

                // Create a new row for each URL
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(idCounter++);
                row.createCell(1).setCellValue(url);

                // Track the maximum number of paths for header adjustment
                maxPaths = Math.max(maxPaths, xpaths.size());

                // Write paths to consecutive columns in the row
                int colNum = 2; // Starting column for "Path"
                if (xpaths.isEmpty()) {
                    row.createCell(2).setCellValue("No 'Merck' found");
                } else {
                    for (String xpath : xpaths) {
                        row.createCell(colNum++).setCellValue(xpath);
                    }
                }
                System.out.println("Data for URL written to Excel.");

                // Save Excel file after each URL is processed
                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
                }
            }

            // Adjust header to add "Path 2", "Path 3", ..., "Path N" based on maxPaths
            for (int i = 1; i < maxPaths; i++) {
                headerRow.createCell(2 + i).setCellValue("Path " + (i + 1));
            }

            // Save final adjustments to Excel
            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
            }

            System.out.println("Excel file 'Merck_Instances.xlsx' created successfully at " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            // Close the driver after crawling
            driver.quit();
            System.out.println("WebDriver closed.");
        }
    }

    // Method to fetch and parse URLs from the sitemap
    private static List<String> getUrlsFromSitemap(String sitemapUrl) {
        List<String> urls = new ArrayList<>();
        HttpURLConnection connection = null;
        try {
            connection = (HttpURLConnection) new URL(sitemapUrl).openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("User-Agent", "Mozilla/5.0");

            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(new String(connection.getInputStream().readAllBytes()))));
            NodeList urlNodes = doc.getElementsByTagName("loc");

            for (int i = 0; i < urlNodes.getLength(); i++) {
                urls.add(urlNodes.item(i).getTextContent());
            }
            System.out.println("Sitemap URLs parsed successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
        return urls;
    }
}
