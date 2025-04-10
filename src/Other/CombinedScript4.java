// CombinedScript.java
package Other;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Color;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.Duration;
import java.util.*;

public class CombinedScript4 {

    public static void main(String[] args) {
        System.out.println("Starting script...");

        // WebDriver setup
        System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
        FirefoxOptions options = new FirefoxOptions();
        options.addArguments("--headless");

        WebDriver driver = new FirefoxDriver(options);
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
        System.out.println("WebDriver initialized.");

        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("URL Status & Merck Paths");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("URL");
        headerRow.createCell(1).setCellValue("Status Code");
        headerRow.createCell(2).setCellValue("Status Message");
        headerRow.createCell(3).setCellValue("Path 1");

        // Output file path
        String outputFilePath = "C:\\TestResults\\Final_XPaths.xlsx";

        // Cell styles
        XSSFCellStyle greenStyle = createCellStyle(workbook, new Color(144, 238, 144)); // Green
        XSSFCellStyle redStyle = createCellStyle(workbook, new Color(255, 99, 71)); // Red
        

        int rowNum = 1;

        try {
            List<String> urls = getUrlsFromSitemap("https://www.msdmanuals.com/sitemap.ashx");
            if (urls.isEmpty()) {
                System.err.println("No URLs retrieved from sitemap.");
                return;
            }
            System.out.println("Total URLs retrieved: " + urls.size());

            int urlIndex = 1;
            for (String url : urls) {
                System.out.printf("Processing URL (%d/%d): %s%n", urlIndex++, urls.size(), url);

                int statusCode = -1;
                String statusMessage = "";
                List<Map<String, Object>> results = new ArrayList<>();

                try {
                    statusCode = getStatusCodeWithRetry(url, 3);
                    if (statusCode != 200) {
                        statusMessage = (statusCode == 403) ? "403 Forbidden" : "Non-200 Status";
                        System.out.println("Skipping URL due to status: " + statusCode);
                    } else {
                        statusCode = extractFinalPaths(driver, url, results);
                        statusMessage = "OK";
                    }
                } catch (Exception e) {
                    statusMessage = "Error";
                    System.err.println("Error processing URL: " + url + " - " + e.getMessage());
                }

                // Write to Excel
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(url);
                row.createCell(1).setCellValue(statusCode);
                row.createCell(2).setCellValue(statusMessage);

                int colNum = 3;
                for (Map<String, Object> result : results) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue((String) result.get("xpath"));

                    // Handle potential null values for "hasRahway"
                    Object hasRahwayObj = result.get("hasRahway");
                    boolean hasRahway = hasRahwayObj != null && (Boolean) hasRahwayObj;
                    cell.setCellStyle(hasRahway ? greenStyle : redStyle);
                }

                // Update headers
                for (int i = 1; i < results.size(); i++) {
                    headerRow.createCell(3 + i).setCellValue("Path " + (i + 1));
                }

                // Save Excel file
                try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                    workbook.write(fileOut);
                }

                System.out.println("Results saved for URL: " + url);
            }

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

    private static int extractFinalPaths(WebDriver driver, String url, List<Map<String, Object>> results) {
        int statusCode = 200;
        try {
            driver.get(url);

            JavascriptExecutor js = (JavascriptExecutor) driver;

            // Extract only final paths
            results.addAll((List<Map<String, Object>>) js.executeScript(
                "function getXPath(element) {" +
                "    if (element.id) return 'id(\"' + element.id + '\")';" +
                "    if (element === document.body) return '/html/body';" +
                "    if (!element.parentNode) return '';" +
                "    let ix = 0;" +
                "    const siblings = element.parentNode.childNodes;" +
                "    for (let i = 0; i < siblings.length; i++) {" +
                "        if (siblings[i] === element) return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';" +
                "        if (siblings[i].nodeType === 1 && siblings[i].tagName === element.tagName) ix++;" +
                "    }" +
                "}" +
                "return Array.from(document.querySelectorAll('*'))" +
                "    .filter(e => e.textContent.toLowerCase().includes('merck'))" +
                "    .filter(e => Array.from(e.children).every(child => !child.textContent.toLowerCase().includes('merck')))" + // Exclude intermediate paths
                "    .map(e => {" +
                "        const parent = e.closest('p');" +
                "        const containsRahway = parent && parent.textContent.toLowerCase().includes('rahway');" +
                "        return { xpath: getXPath(e), hasRahway: containsRahway };" +
                "    });"
            ));
        } catch (Exception e) {
            System.err.println("Error extracting paths for: " + url);
            statusCode = -1;
        }
        return statusCode;
    }

    private static XSSFCellStyle createCellStyle(Workbook workbook, Color color) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(color, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static List<String> getUrlsFromSitemap(String sitemapUrl) {
        List<String> urls = new ArrayList<>();
        try {
            HttpURLConnection connection = (HttpURLConnection) new URL(sitemapUrl).openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("User-Agent", "Mozilla/5.0");

            BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            StringBuilder sitemapContent = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sitemapContent.append(line);
            }
            reader.close();

            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(sitemapContent.toString())));
            NodeList urlNodes = doc.getElementsByTagName("loc");

            for (int i = 0; i < urlNodes.getLength(); i++) {
                urls.add(urlNodes.item(i).getTextContent());
            }

            connection.disconnect();
        } catch (Exception e) {
            System.err.println("Error retrieving sitemap: " + e.getMessage());
        }
        return urls;
    }

    private static int getStatusCodeWithRetry(String url, int maxRetries) {
        int retries = 0;
        int statusCode = -1;
        while (retries < maxRetries) {
            try {
                HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("User-Agent", "Mozilla/5.0");
                statusCode = connection.getResponseCode();
                connection.disconnect();
                if (statusCode != 403) break;
                System.out.println("Retrying for 403 status...");
                Thread.sleep(30000);
            } catch (Exception e) {
                retries++;
            }
        }
        return statusCode;
    }
}
