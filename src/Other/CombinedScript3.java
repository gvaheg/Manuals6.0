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
import java.util.stream.Collectors;

public class CombinedScript3 {

    public static void main(String[] args) {
        System.out.println("Starting combined script...");

        // Set up WebDriver in headless mode
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
        headerRow.createCell(3).setCellValue("Final XPath");

        // Output file path
        String outputFilePath = "C:\\TestResults\\URL_Status_and_MerckDE.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.getParentFile().mkdirs();

        // Create green and red cell styles
        XSSFCellStyle greenCellStyle = (XSSFCellStyle) workbook.createCellStyle();
        greenCellStyle.setFillForegroundColor(new XSSFColor(new Color(144, 238, 144), null));
        greenCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle redCellStyle = (XSSFCellStyle) workbook.createCellStyle();
        redCellStyle.setFillForegroundColor(new XSSFColor(new Color(255, 99, 71), null));
        redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        int rowNum = 1;

        try {
            List<String> urls = getUrlsFromSitemap("https://www.msdmanuals.com/de/sitemap.ashx");
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
                        System.out.println("Loading URL in browser...");
                        statusCode = checkPageAndExtractMerckInstances(driver, url, results);
                        statusMessage = "OK";
                    }
                } catch (Exception e) {
                    statusMessage = "Error";
                    System.err.println("Error processing URL: " + url + " - " + e.getMessage());
                }

                // Deduplicate and filter results for only final XPaths
                List<String> finalXPaths = results.stream()
                        .map(r -> (String) r.get("xpath"))
                        .distinct()
                        .collect(Collectors.toList());

                // Write results to Excel
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(url);
                row.createCell(1).setCellValue(statusCode);
                row.createCell(2).setCellValue(statusMessage);

                int colNum = 3;
                for (String xpath : finalXPaths) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(xpath);

                    boolean hasRahway = results.stream()
                            .filter(r -> xpath.equals(r.get("xpath")))
                            .anyMatch(r -> (boolean) r.get("hasRahway"));

                    cell.setCellStyle(hasRahway ? greenCellStyle : redCellStyle);
                }

                System.out.println("Results written for URL: " + url);

                // Save workbook after each URL
                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
                }
            }

            System.out.println("Excel file created successfully at: " + outputFilePath);

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

    private static int getStatusCodeWithRetry(String url, int maxRetries) {
        int retries = 0;
        int statusCode = -1;

        while (retries < maxRetries) {
            try {
                HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("User-Agent", "Mozilla/5.0");
                connection.setConnectTimeout(10000);
                connection.setReadTimeout(10000);

                statusCode = connection.getResponseCode();
                connection.disconnect();

                if (statusCode == 403) {
                    retries++;
                    System.out.println("403 Forbidden for URL: " + url + ". Retrying in 30 seconds...");
                    Thread.sleep(30000);
                } else {
                    break;
                }
            } catch (Exception e) {
                retries++;
                System.err.println("Error fetching status for URL: " + url + " - " + e.getMessage());
            }
        }
        return statusCode;
    }

    private static int checkPageAndExtractMerckInstances(WebDriver driver, String url, List<Map<String, Object>> results) {
        int statusCode = 200;
        try {
            driver.get(url);

            JavascriptExecutor js = (JavascriptExecutor) driver;

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
                "    .filter(e => e.textContent.trim().toLowerCase() === 'merck' &&" + // Match exact text
                "        e.offsetParent !== null &&" + // Must be visible
                "        !e.closest('footer') &&" + // Exclude footer
                "        (!e.children || e.children.length === 0))" + // Final element check
                "    .map(e => ({ xpath: getXPath(e), hasRahway: false }));"
            ));

        } catch (Exception e) {
            System.err.println("Error processing page content for: " + url);
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
}
