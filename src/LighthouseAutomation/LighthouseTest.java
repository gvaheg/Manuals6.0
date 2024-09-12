package LighthouseAutomation;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class LighthouseTest {

    // Run Lighthouse for a given URL
    public static void runLighthouse(String url, String outputPath) {
        try {
            // Update this with the correct path to your lighthouse.cmd
            String lighthousePath = "C:/Users/vahep/AppData/Roaming/npm/lighthouse.cmd";

            // Build the command for mobile emulation and multiple categories
            List<String> command = new ArrayList<>();
            command.add(lighthousePath);
            command.add(url);
            command.add("--output");
            command.add("json");
            command.add("--output-path");
            command.add(outputPath);
            command.add("--chrome-flags=--headless --user-agent='Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1' --window-size=375,812");
            command.add("--only-categories=performance,accessibility,best-practices,seo");

            ProcessBuilder processBuilder = new ProcessBuilder(command);
            processBuilder.redirectErrorStream(true); // Merge stdout and stderr
            Process process = processBuilder.start();

            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println(line);
            }

            int exitCode = process.waitFor();
            System.out.println("Lighthouse process for " + url + " exited with code: " + exitCode);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Sanitize URLs to generate valid file names
    public static String sanitizeFileName(String url) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] hash = digest.digest(url.getBytes(StandardCharsets.UTF_8));
            StringBuilder hexString = new StringBuilder();
            for (byte b : hash) {
                hexString.append(Integer.toHexString(0xFF & b));
            }
            return hexString.toString(); // Return the hashed version of the URL as the filename
        } catch (NoSuchAlgorithmException e) {
            e.printStackTrace();
            return url.replaceAll("[^a-zA-Z0-9.-]", "_"); // Fallback to simple sanitization
        }
    }

    // Parse Lighthouse JSON report
    public static JsonObject parseLighthouseReport(String filePath) {
        JsonObject results = new JsonObject();
        try (FileReader reader = new FileReader(filePath)) {
            JsonObject jsonObject = JsonParser.parseReader(reader).getAsJsonObject();
            results.addProperty("performance", getValue(jsonObject, "categories", "performance"));
            results.addProperty("accessibility", getValue(jsonObject, "categories", "accessibility"));
            results.addProperty("best-practices", getValue(jsonObject, "categories", "best-practices"));
            results.addProperty("seo", getValue(jsonObject, "categories", "seo"));
            results.addProperty("first-contentful-paint", getAuditValue(jsonObject, "first-contentful-paint"));
            results.addProperty("largest-contentful-paint", getAuditValue(jsonObject, "largest-contentful-paint"));
            results.addProperty("total-blocking-time", getAuditValue(jsonObject, "total-blocking-time"));
            results.addProperty("cumulative-layout-shift", getAuditValue(jsonObject, "cumulative-layout-shift"));
            results.addProperty("speed-index", getAuditValue(jsonObject, "speed-index"));
        } catch (Exception e) {
            e.printStackTrace();
        }
        return results;
    }

    // Safely get audit metrics from JSON report
    public static String getAuditValue(JsonObject jsonObject, String key) {
        try {
            return jsonObject.getAsJsonObject("audits").getAsJsonObject(key).get("displayValue").getAsString();
        } catch (Exception e) {
            return "N/A";
        }
    }

    // Safely get category scores from JSON report
    public static double getValue(JsonObject jsonObject, String category, String key) {
        try {
            return jsonObject.getAsJsonObject(category).getAsJsonObject(key).get("score").getAsDouble();
        } catch (Exception e) {
            return 0.0;
        }
    }

    // Write results to Excel after each audit
    public static void writeResultsToExcel(String fileName, String url, String versionName, JsonObject result) throws IOException {
        Workbook workbook;
        Sheet sheet;
        
        File file = new File(fileName);
        
        if (file.exists()) {
            // Open the existing workbook
            FileInputStream fileInputStream = new FileInputStream(file);
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheet("Lighthouse Results");
        } else {
            // Create a new workbook and sheet if the file doesn't exist
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Lighthouse Results");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Date");
            headerRow.createCell(1).setCellValue("Version Name");
            headerRow.createCell(2).setCellValue("URL");
            headerRow.createCell(3).setCellValue("Performance");
            headerRow.createCell(4).setCellValue("Accessibility");
            headerRow.createCell(5).setCellValue("Best Practices");
            headerRow.createCell(6).setCellValue("SEO");
            headerRow.createCell(7).setCellValue("First Contentful Paint");
            headerRow.createCell(8).setCellValue("Largest Contentful Paint");
            headerRow.createCell(9).setCellValue("Total Blocking Time");
            headerRow.createCell(10).setCellValue("Cumulative Layout Shift");
            headerRow.createCell(11).setCellValue("Speed Index");
        }

        // Get the current date
        LocalDate today = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        String formattedDate = today.format(formatter);

        // Get the last row number
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowNum + 1);
        
        // Add data to the new row
        row.createCell(0).setCellValue(formattedDate); // Add today's date
        row.createCell(1).setCellValue(versionName);   // Add version name
        row.createCell(2).setCellValue(url);           // Add URL

        row.createCell(3).setCellValue(result.has("performance") ? result.get("performance").getAsDouble() : 0);
        row.createCell(4).setCellValue(result.has("accessibility") ? result.get("accessibility").getAsDouble() : 0);
        row.createCell(5).setCellValue(result.has("best-practices") ? result.get("best-practices").getAsDouble() : 0);
        row.createCell(6).setCellValue(result.has("seo") ? result.get("seo").getAsDouble() : 0);
        row.createCell(7).setCellValue(result.has("first-contentful-paint") ? result.get("first-contentful-paint").getAsString() : "N/A");
        row.createCell(8).setCellValue(result.has("largest-contentful-paint") ? result.get("largest-contentful-paint").getAsString() : "N/A");
        row.createCell(9).setCellValue(result.has("total-blocking-time") ? result.get("total-blocking-time").getAsString() : "N/A");
        row.createCell(10).setCellValue(result.has("cumulative-layout-shift") ? result.get("cumulative-layout-shift").getAsString() : "N/A");
        row.createCell(11).setCellValue(result.has("speed-index") ? result.get("speed-index").getAsString() : "N/A");

        // Write to the Excel file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }

        // Close the workbook
        workbook.close();
    }


    public static void main(String[] args) {
    	// Define the file name for the Excel report
        String fileName = "LighthouseResults.xlsx";
        
        // Check if the file exists and delete it if it does (this clears previous results)
        File file = new File(fileName);
        if (file.exists()) {
            boolean deleted = file.delete(); // Delete the file
            if (deleted) {
                System.out.println("Previous Lighthouse results file deleted.");
            } else {
                System.out.println("Failed to delete previous Lighthouse results file.");
            }
        }
    	
        // List of URLs to audit
        List<String> urls = new ArrayList<>();
        urls.add("https://www.merckmanuals.com/professional"); // 0
        urls.add("https://www.msdmanuals.com/professional"); // 1
        urls.add("https://www.msdmanuals.com/de/profi"); // 2
        urls.add("https://www.msdmanuals.com/es/professional"); // 3
        urls.add("https://www.msdmanuals.com/fr/professional"); // 4
        urls.add("https://www.msdmanuals.com/it/professionale"); // 5
        urls.add("https://www.msdmanuals.com/ja-jp/professional"); // 6
        urls.add("https://www.msdmanuals.com/pt/profissional"); // 7
        urls.add("https://www.msdmanuals.com/ru/professional"); // 8
        urls.add("https://www.msdmanuals.cn/professional"); // 9
        urls.add("https://www.msdmanuals.com/vi/professional"); // 10
        urls.add("https://www.merckmanuals.com/home"); // 11
        urls.add("https://www.msdmanuals.com/home"); // 12
        urls.add("https://www.msdmanuals.com/de/heim"); // 13
        urls.add("https://www.msdmanuals.com/es/hogar"); // 14
        urls.add("https://www.msdmanuals.com/fr/accueil"); // 15
        urls.add("https://www.msdmanuals.com/it/casa"); // 16
        urls.add("https://www.msdmanuals.com/ja-jp/home"); // 17
        urls.add("https://www.msdmanuals.com/ko/home"); // 18
        urls.add("https://www.msdmanuals.com/pt/casa"); // 19
        urls.add("https://www.msdmanuals.com/ru/home"); // 20
        urls.add("https://www.msdmanuals.cn/home"); // 21
        urls.add("https://www.msdmanuals.com/ar/home"); // 22
        urls.add("https://www.merckmanuals.com/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris?query=acne"); // 23
        urls.add("https://www.msdmanuals.com/en-jp/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris"); // 24
        urls.add("https://www.msdmanuals.com/de/profi/erkrankungen-der-haut/akne-und-verwandte-erkrankungen/acne-vulgaris"); // 25
        urls.add("https://www.msdmanuals.com/es/professional/trastornos-dermatol%C3%B3gicos/acn%C3%A9-y-trastornos-relacionados/acn%C3%A9-vulgar"); // 26
        urls.add("https://www.msdmanuals.com/fr/professional/troubles-dermatologiques/acn%C3%A9-et-pathologies-apparent%C3%A9es/acn%C3%A9-vulgaire"); // 27
        urls.add("https://www.msdmanuals.com/it/professionale/disturbi-dermatologici/acne-e-disturbi-correlati/acne-vulgaris"); // 28
        urls.add("https://www.msdmanuals.com/ja-jp/professional/14-%E7%9A%AE%E8%86%9A%E7%96%BE%E6%82%A3/%E3%81%96%E7%98%A1%E3%81%8A%E3%82%88%E3%81%B3%E9%96%A2%E9%80%A3%E7%96%BE%E6%82%A3/%E5%B0%8B%E5%B8%B8%E6%80%A7%E3%81%96%E7%98%A1"); // 29
        urls.add("https://www.msdmanuals.com/pt/profissional/dist%C3%BArbios-dermatol%C3%B3gicos/acne-e-doen%C3%A7as-relacionadas/acne-vulgar"); // 30
        urls.add("https://www.msdmanuals.com/ru/professional/%D0%B4%D0%B5%D1%80%D0%BC%D0%B0%D1%82%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B0%D1%8F-%D0%BF%D0%B0%D1%82%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%8F/%D0%B0%D0%BA%D0%BD%D0%B5-%D0%B8-%D1%81%D0%B2%D1%8F%D0%B7%D0%B0%D0%BD%D0%BD%D1%8B%D0%B5-%D1%81-%D0%BD%D0%B8%D0%BC%D0%B8-%D1%80%D0%B0%D1%81%D1%81%D1%82%D1%80%D0%BE%D0%B9%D1%81%D1%82%D0%B2%D0%B0/%D0%B2%D1%83%D0%BB%D1%8C%D0%B3%D0%B0%D1%80%D0%BD%D1%8B%D0%B5-%D1%83%D0%B3%D1%80%D0%B8"); // 31
        urls.add("https://www.msdmanuals.cn/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris"); // 32
        urls.add("https://www.msdmanuals.com/vi/professional/r%E1%BB%91i-lo%E1%BA%A1n-da-li%E1%BB%85u/tr%E1%BB%A9ng-c%C3%A1-v%C3%A0-c%C3%A1c-b%E1%BB%87nh-l%C3%BD-li%C3%AAn-quan/tr%E1%BB%A9ng-c%C3%A1"); // 33
        urls.add("https://www.merckmanuals.com/home/skin-disorders/acne-and-related-disorders/acne?query=acne"); // 34
        urls.add("https://www.msdmanuals.com/home/skin-disorders/acne-and-related-disorders/acne"); // 35
        urls.add("https://www.msdmanuals.com/de/heim/hauterkrankungen/akne-und-verwandte-erkrankungen/akne"); // 36
        urls.add("https://www.msdmanuals.com/es/hogar/trastornos-de-la-piel/acn%C3%A9-y-trastornos-relacionados/acn%C3%A9"); // 37
        urls.add("https://www.msdmanuals.com/fr/accueil/troubles-cutane%C3%A9s/acn%C3%A9-et-troubles-associ%C3%A9s/acn%C3%A9"); // 38
        urls.add("https://www.msdmanuals.com/it/casa/patologie-della-cute/acne-e-disturbi-correlati/acne"); // 39
        urls.add("https://www.msdmanuals.com/ja-jp/home/17-%E7%9A%AE%E8%86%9A%E3%81%AE%E7%97%85%E6%B0%97/%E3%81%AB%E3%81%8D%E3%81%B3%E3%81%A8%E9%96%A2%E9%80%A3%E7%96%BE%E6%82%A3/%E3%81%AB%E3%81%8D%E3%81%B3%EF%BC%88%E3%81%96%E7%98%A1%EF%BC%89"); // 40
        urls.add("https://www.msdmanuals.com/ko/home/%ED%94%BC%EB%B6%80-%EC%A7%88%ED%99%98/%EC%97%AC%EB%93%9C%EB%A6%84-%EB%B0%8F-%EA%B4%80%EB%A0%A8-%EC%A7%88%ED%99%98/%EC%97%AC%EB%93%9C%EB%A6%84"); // 41
        urls.add("https://www.msdmanuals.com/pt/casa/dist%C3%BArbios-da-pele/acne-e-dist%C3%BArbios-relacionados/acne"); // 42
        urls.add("https://www.msdmanuals.com/ru/home/%D0%BA%D0%BE%D0%B6%D0%BD%D1%8B%D0%B5-%D0%B7%D0%B0%D0%B1%D0%BE%D0%BB%D0%B5%D0%B2%D0%B0%D0%BD%D0%B8%D1%8F/%D1%83%D0%B3%D1%80%D0%B5%D0%B2%D0%B0%D1%8F-%D1%81%D1%8B%D0%BF%D1%8C-%D0%B8-%D1%80%D0%BE%D0%B4%D1%81%D1%82%D0%B2%D0%B5%D0%BD%D0%BD%D1%8B%D0%B5-%D0%B7%D0%B0%D0%B1%D0%BE%D0%BB%D0%B5%D0%B2%D0%B0%D0%BD%D0%B8%D1%8F/%D1%83%D0%B3%D1%80%D0%B5%D0%B2%D0%B0%D1%8F-%D1%81%D1%8B%D0%BF%D1%8C"); // 43
        urls.add("https://www.msdmanuals.cn/home/skin-disorders/acne-and-related-disorders/acne"); // 44
        urls.add("https://www.msdmanuals.com/ar/home/%D8%A7%D9%84%D8%A7%D8%B6%D8%B7%D8%B1%D8%A7%D8%A8%D8%A7%D8%AA-%D8%A7%D9%84%D8%AC%D9%84%D8%AF%D9%8A%D9%91%D9%8E%D8%A9/%D8%AD%D8%A8%D9%91-%D8%A7%D9%84%D8%B4%D9%91%D9%8E%D8%A8%D8%A7%D8%A8-%D9%88%D8%A7%D9%84%D8%A7%D8%B6%D8%B7%D8%B1%D8%A7%D8%A8%D8%A7%D8%AA-%D8%B0%D9%8E%D8%A7%D8%AA-%D8%A7%D9%84%D8%B5%D9%91%D9%90%D9%84%D8%A9/%D8%AD%D8%A8-%D8%A7%D9%84%D8%B4%D8%A8%D8%A7%D8%A8-%D8%A7%D9%84%D8%B9%D8%AF%D9%91"); // 45
        urls.add("https://www.merckmanuals.com/professional/injuries-poisoning/fractures/ankle-fractures?query=ankle%20fractures"); // 46
        urls.add("https://www.msdmanuals.com/professional/injuries-poisoning/fractures/ankle-fractures"); // 47
        urls.add("https://www.msdmanuals.com/de/profi/verletzungen-vergiftungen/frakturen/sprunggelenkfrakturen"); // 48
        urls.add("https://www.msdmanuals.com/es/professional/lesiones-y-envenenamientos/fracturas/fracturas-del-tobillo"); // 49
        urls.add("https://www.msdmanuals.com/fr/professional/blessures-empoisonnement/fractures/fractures-de-cheville"); // 50
        urls.add("https://www.msdmanuals.com/it/professionale/traumi-avvelenamento/fratture/fratture-alla-caviglia"); // 51
        urls.add("https://www.msdmanuals.com/ja-jp/professional/22-%E5%A4%96%E5%82%B7%E3%81%A8%E4%B8%AD%E6%AF%92/%E9%AA%A8%E6%8A%98/%E8%B6%B3%E9%96%A2%E7%AF%80%E9%AA%A8%E6%8A%98"); // 52
        urls.add("https://www.msdmanuals.com/pt/profissional/les%C3%B5es-intoxica%C3%A7%C3%A3o/fraturas/fraturas-do-tornozelo"); // 53
        urls.add("https://www.msdmanuals.com/ru/professional/%D1%82%D1%80%D0%B0%D0%B2%D0%BC%D1%8B-%D0%BE%D1%82%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%B8%D1%8F/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B-%D0%B3%D0%BE%D0%BB%D0%B5%D0%BD%D0%BE%D1%81%D1%82%D0%BE%D0%BF%D0%BD%D0%BE%D0%B3%D0%BE-%D1%81%D1%83%D1%81%D1%82%D0%B0%D0%B2%D0%B0"); // 54
        urls.add("https://www.msdmanuals.cn/professional/injuries-poisoning/fractures/ankle-fractures"); // 55
        urls.add("https://www.msdmanuals.com/vi/professional/ch%E1%BA%A5n-th%C6%B0%C6%A1ng-ng%E1%BB%99-%C4%91%E1%BB%99c/g%C3%A3y-x%C6%B0%C6%A1ng/g%C3%A3y-x%C6%B0%C6%A1ng-c%E1%BB%95-ch%C3%A2n"); // 56
        urls.add("https://www.merckmanuals.com/home/injuries-and-poisoning/fractures/ankle-fractures?query=ankle%20fractures"); // 57
        urls.add("https://www.msdmanuals.com/home/injuries-and-poisoning/fractures/ankle-fractures?query=ankle%20fractures"); // 58
        urls.add("https://www.msdmanuals.com/de/heim/verletzungen-und-vergiftung/frakturen/kn%C3%B6chelfrakturen"); // 59
        urls.add("https://www.msdmanuals.com/es/hogar/traumatismos-y-envenenamientos/fracturas/fracturas-de-tobillo"); // 60
        urls.add("https://www.msdmanuals.com/fr/accueil/l%C3%A9sions-et-intoxications/fractures/fractures-de-la-cheville"); // 61
        urls.add("https://www.msdmanuals.com/it/casa/lesioni-e-avvelenamento/fratture/fratture-della-caviglia"); // 62
        urls.add("https://www.msdmanuals.com/ja-jp/home/25-%E5%A4%96%E5%82%B7%E3%81%A8%E4%B8%AD%E6%AF%92/%E9%AA%A8%E6%8A%98/%E8%B6%B3%E9%A6%96%E3%81%AE%E9%AA%A8%E6%8A%98"); // 63
        urls.add("https://www.msdmanuals.com/ko/home/%EB%B6%80%EC%83%81-%EB%B0%8F-%EC%A4%91%EB%8F%85/%EA%B3%A8%EC%A0%88/%EB%B0%9C%EB%AA%A9-%EA%B3%A8%EC%A0%88"); // 64
        urls.add("https://www.msdmanuals.com/pt/casa/les%C3%B5es-e-envenenamentos/fraturas/fraturas-do-tornozelo"); // 65
        urls.add("https://www.msdmanuals.com/ru/home/%D1%82%D1%80%D0%B0%D0%B2%D0%BC%D1%8B-%D0%B8-%D0%BE%D1%82%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%B8%D1%8F/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B-%D0%BB%D0%BE%D0%B4%D1%8B%D0%B6%D0%BA%D0%B8"); // 66
        urls.add("https://www.msdmanuals.cn/home/injuries-and-poisoning/fractures/ankle-fractures"); // 67
        urls.add("https://www.msdmanuals.com/ar/home/%D8%A7%D9%84%D8%A5%D8%B5%D8%A7%D8%A8%D8%A7%D8%AA-%D9%88%D8%A7%D9%84%D8%AA%D9%91%D9%8E%D8%B3%D9%85%D9%91%D9%8F%D9%85/%D8%A7%D9%84%D9%83%D8%B3%D9%88%D8%B1/%D9%83%D9%8F%D8%B3%D9%88%D8%B1%D9%8F-%D8%A7%D9%84%D9%83%D8%A7%D8%AD%D9%84"); // 68


        // List of version names for each URL
        List<String> versionNames = new ArrayList<>();
        versionNames.add("EN PV Home Merck"); // 0
        versionNames.add("EN PV Home"); // 1
        versionNames.add("DE PV Home"); // 2
        versionNames.add("ES PV Home"); // 3
        versionNames.add("FR PV Home"); // 4
        versionNames.add("IT PV Home"); // 5
        versionNames.add("JA PV Home"); // 6
        versionNames.add("PT PV Home"); // 7
        versionNames.add("RU PV Home"); // 8
        versionNames.add("CN PV Home"); // 9
        versionNames.add("VI PV Home"); // 10
        versionNames.add("EN CV Home Merck"); // 11
        versionNames.add("EN CV Home"); // 12
        versionNames.add("DE CV Home"); // 13
        versionNames.add("ES CV Home"); // 14
        versionNames.add("FR CV Home"); // 15
        versionNames.add("IT CV Home"); // 16
        versionNames.add("JA CV Home"); // 17
        versionNames.add("KO CV Home"); // 18
        versionNames.add("PT CV Home"); // 19
        versionNames.add("RU CV Home"); // 20
        versionNames.add("CN CV Home"); // 21
        versionNames.add("AR CV Home"); // 22
        versionNames.add("EN PV Acne Merck"); // 23
        versionNames.add("EN PV Acne"); // 24
        versionNames.add("DE PV Acne"); // 25
        versionNames.add("ES PV Acne"); // 26
        versionNames.add("FR PV Acne"); // 27
        versionNames.add("IT PV Acne"); // 28
        versionNames.add("JA PV Acne"); // 29
        versionNames.add("PT PV Acne"); // 30
        versionNames.add("RU PV Acne"); // 31
        versionNames.add("CN PV Acne"); // 32
        versionNames.add("VI PV Acne"); // 33
        versionNames.add("EN CV Acne Merck"); // 34
        versionNames.add("EN CV Acne"); // 35
        versionNames.add("DE CV Acne"); // 36
        versionNames.add("ES CV Acne"); // 37
        versionNames.add("FR CV Acne"); // 38
        versionNames.add("IT CV Acne"); // 39
        versionNames.add("JA CV Acne"); // 40
        versionNames.add("KO CV Acne"); // 41
        versionNames.add("PT CV Acne"); // 42
        versionNames.add("RU CV Acne"); // 43
        versionNames.add("CN CV Acne"); // 44
        versionNames.add("AR CV Acne"); // 45
        versionNames.add("EN PV Ankle Fractures Merck"); // 46
        versionNames.add("EN PV Ankle Fractures"); // 47
        versionNames.add("DE PV Ankle Fractures"); // 48
        versionNames.add("ES PV Ankle Fractures"); // 49
        versionNames.add("FR PV Ankle Fractures"); // 50
        versionNames.add("IT PV Ankle Fractures"); // 51
        versionNames.add("JA PV Ankle Fractures"); // 52
        versionNames.add("PT PV Ankle Fractures"); // 53
        versionNames.add("RU PV Ankle Fractures"); // 54
        versionNames.add("CN PV Ankle Fractures"); // 55
        versionNames.add("VI PV Ankle Fractures"); // 56
        versionNames.add("EN CV Ankle Fractures Merck"); // 57
        versionNames.add("EN CV Ankle Fractures"); // 58
        versionNames.add("DE CV Ankle Fractures"); // 59
        versionNames.add("ES CV Ankle Fractures"); // 60
        versionNames.add("FR CV Ankle Fractures"); // 61
        versionNames.add("IT CV Ankle Fractures"); // 62
        versionNames.add("JA CV Ankle Fractures"); // 63
        versionNames.add("KO CV Ankle Fractures"); // 64
        versionNames.add("PT CV Ankle Fractures"); // 65
        versionNames.add("RU CV Ankle Fractures"); // 66
        versionNames.add("CN CV Ankle Fractures"); // 67
        versionNames.add("AR CV Ankle Fractures"); // 68

        // Loop through each URL, run Lighthouse, parse the result, and write to Excel
        for (int i = 0; i < urls.size(); i++) {
            String url = urls.get(i);
            String versionName = versionNames.get(i);
            String outputPath = "lighthouse-report-" + sanitizeFileName(url) + ".json";

            // Run Lighthouse for the current URL
            runLighthouse(url, outputPath);

            // Parse the audit results from the report
            JsonObject result = parseLighthouseReport(outputPath);

            // Write the result to the Excel file after each audit
            try {
                writeResultsToExcel(fileName, url, versionName, result);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Lighthouse audits completed and results written to " + fileName);
    }
}
