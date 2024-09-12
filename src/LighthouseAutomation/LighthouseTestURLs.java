package LighthouseAutomation;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import com.google.gson.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LighthouseTestURLs {

    // This will hold all URLs and their respective version names
    private static final String[] urls = {
        "https://www.merckmanuals.com/professional",  // 0
        "https://www.msdmanuals.com/professional",  // 1
        "https://www.msdmanuals.com/de/profi",  // 2
        "https://www.msdmanuals.com/es/professional",  // 3
        "https://www.msdmanuals.com/fr/professional",  // 4
        "https://www.msdmanuals.com/it/professionale",  // 5
        "https://www.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB",  // 6
        "https://www.msdmanuals.com/pt/profissional",  // 7
        "https://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9",  // 8
        "https://www.msdmanuals.cn/professional",  // 9
        "https://www.msdmanuals.com/vi/chuy%C3%AAn-gia",  // 10
        "https://www.merckmanuals.com/home",  // 11
        "https://www.msdmanuals.com/home",  // 12
        "https://www.msdmanuals.com/de/heim",  // 13
        "https://www.msdmanuals.com/es/hogar",  // 14
        "https://www.msdmanuals.com/fr/accueil",  // 15
        "https://www.msdmanuals.com/it/casa",  // 16
        "https://www.msdmanuals.com/ja-jp/%E3%83%9B%E3%83%BC%E3%83%A0",  // 17
        "https://www.msdmanuals.com/ko/%ED%99%88",  // 18
        "https://www.msdmanuals.com/pt/casa",  // 19
        "https://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0",  // 20
        "https://www.msdmanuals.cn/home",  // 21
        "https://www.msdmanuals.com/ar/home",  // 22
        "https://www.merckmanuals.com/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris?query=acne",  // 23
        "https://www.msdmanuals.com/en-jp/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris",  // 24
        "https://www.msdmanuals.com/de/profi/erkrankungen-der-haut/akne-und-verwandte-erkrankungen/acne-vulgaris",  // 25
        "https://www.msdmanuals.com/es/professional/trastornos-dermatol%C3%B3gicos/acn%C3%A9-y-trastornos-relacionados/acn%C3%A9-vulgar",  // 26
        "https://www.msdmanuals.com/fr/professional/troubles-dermatologiques/acn%C3%A9-et-pathologies-apparent%C3%A9es/acn%C3%A9-vulgaire",  // 27
        "https://www.msdmanuals.com/it/professionale/disturbi-dermatologici/acne-e-disturbi-correlati/acne-vulgaris",  // 28
        "https://www.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/14-%E7%9A%AE%E8%86%9A%E7%96%BE%E6%82%A3/%E3%81%96%E7%98%A1%E3%81%8A%E3%82%88%E3%81%B3%E9%96%A2%E9%80%A3%E7%96%BE%E6%82%A3/%E5%B0%81%E5%B8%B8%E6%80%A7%E3%81%96%E7%98%A1",  // 29
        "https://www.msdmanuals.com/pt/profissional/dist%C3%BArbios-dermatol%C3%B3gicos/acne-e-doen%C3%A7as-relacionadas/acne-vulgar",  // 30
        "https://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/%D0%B4%D0%B5%D1%80%D0%BC%D0%B0%D1%82%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B0%D1%8F-%D0%BF%D0%B0%D1%82%D0%BE%D0%BB%D0%BE%D0%B3%D0%B8%D1%8F/%D0%B0%D0%BA%D0%BD%D0%B5-%D0%B8-%D1%81%D0%B2%D1%8F%D0%B7%D0%B0%D0%BD%D0%BD%D1%8B%D0%B5-%D1%81-%D0%BD%D0%B8%D0%BC%D0%B8-%D1%80%D0%B0%D1%81%D1%81%D1%82%D1%80%D0%BE%D0%B9%D1%81%D1%82%D0%B2%D0%B0/%D0%B2%D1%83%D0%BB%D1%8C%D0%B3%D0%B0%D1%80%D0%BD%D1%8B%D0%B5-%D1%83%D0%B3%D1%80%D0%B8",  // 31
        "https://www.msdmanuals.cn/professional/dermatologic-disorders/acne-and-related-disorders/acne-vulgaris",  // 32
        "https://www.msdmanuals.com/vi/chuy%C3%AAn-gia/r%E1%BB%91i-lo%E1%BA%A1n-da-li%E1%BB%85u/tr%E1%BB%A9ng-c%C3%A1-v%C3%A0-c%C3%A1c-b%E1%BB%87nh-l%C3%BD-li%C3%AAn-quan/tr%E1%BB%A9ng-c%C3%A1",  // 33
        "https://www.merckmanuals.com/home/skin-disorders/acne-and-related-disorders/acne?query=acne",  // 34
        "https://www.msdmanuals.com/home/skin-disorders/acne-and-related-disorders/acne",  // 35
        "https://www.msdmanuals.com/de/heim/hauterkrankungen/akne-und-verwandte-erkrankungen/akne",  // 36
        "https://www.msdmanuals.com/es/hogar/trastornos-de-la-piel/acn%C3%A9-y-trastornos-relacionados/acn%C3%A9",  // 37
        "https://www.msdmanuals.com/fr/accueil/troubles-cutane%C3%A9s/acn%C3%A9-et-troubles-associ%C3%A9s/acn%C3%A9",  // 38
        "https://www.msdmanuals.com/it/casa/patologie-della-cute/acne-e-disturbi-correlati/acne",  // 39
        "https://www.msdmanuals.com/ja-jp/%E3%83%9B%E3%83%BC%E3%83%A0/17-%E7%9A%AE%E8%86%9A%E3%81%AE%E7%96%BE%E6%82%A3/%E3%81%AB%E3%81%8D%E3%81%B3%E3%81%A8%E9%96%A2%E9%80%A3%E7%96%BE%E6%82%A3/%E3%81%AB%E3%81%8D%E3%81%B3-%E3%81%96%E7%98%A1",  // 40
        "https://www.msdmanuals.com/ko/%ED%99%88/%ED%94%BC%EB%B6%80-%EC%A7%88%ED%99%98/%EC%97%AC%EB%93%9C%EB%A6%84-%EB%B0%8F-%EA%B4%80%EB%A0%A8-%EC%A7%88%ED%99%98/%EC%97%AC%EB%93%9C%EB%A6%84",  // 41
        "https://www.msdmanuals.com/pt/casa/dist%C3%BArbios-da-pele/acne-e-dist%C3%BArbios-relacionados/acne",  // 42
        "https://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/%D0%BA%D0%BE%D0%B6%D0%BD%D1%8B%D0%B5-%D0%B7%D0%B0%D0%B1%D0%BE%D0%BB%D0%B5%D0%B2%D0%B0%D0%BD%D0%B8%D1%8F/%D1%83%D0%B3%D1%80%D0%B5%D0%B2%D0%B0%D1%8F-%D1%81%D1%8B%D0%BF%D1%8C-%D0%B8-%D1%80%D0%BE%D0%B4%D1%81%D1%82%D0%B2%D0%B5%D0%BD%D0%BD%D1%8B%D0%B5-%D0%B7%D0%B0%D0%B1%D0%BE%D0%BB%D0%B5%D0%B2%D0%B0%D0%BD%D0%B8%D1%8F/%D1%83%D0%B3%D1%80%D0%B5%D0%B2%D0%B0%D1%8F-%D1%81%D1%8B%D0%BF",  // 43
        "https://www.msdmanuals.cn/home/skin-disorders/acne-and-related-disorders/acne",  // 44
        "https://www.msdmanuals.com/ar/home/%D8%A7%D9%84%D8%A7%D8%B6%D8%B7%D8%B1%D8%A7%D8%A8%D8%A7%D8%AA-%D8%A7%D9%84%D8%AC%D9%84%D8%AF%D9%8A%D9%91%D9%8E%D8%A9/%D8%AD%D8%A8%D9%91-%D8%A7%D9%84%D8%B4%D9%91%D8%A8%D8%A7%D8%A8-%D9%88%D8%A7%D9%84%D8%A7%D8%B6%D8%B7%D8%B1%D8%A7%D8%A8%D8%A7%D8%AA-%D8%B0%D8%A7%D8%AA-%D8%A7%D9%84%D8%B5%D9%91%D9%8E%D9%84%D8%A9/%D8%AD%D8%A8-%D8%A7%D9%84%D8%B4%D8%A8%D8%A7%D8%A8-%D8%A7%D9%84%D8%B9%D8%AF%D9%91",  // 45
        "https://www.merckmanuals.com/professional/injuries-poisoning/fractures/ankle-fractures?query=ankle%20fractures",  // 46
        "https://www.msdmanuals.com/professional/injuries-poisoning/fractures/ankle-fractures",  // 47
        "https://www.msdmanuals.com/de/profi/verletzungen-vergiftungen/frakturen/sprunggelenkfrakturen",  // 48
        "https://www.msdmanuals.com/es/professional/lesiones-y-envenenamientos/fracturas/fracturas-del-tobillo",  // 49
        "https://www.msdmanuals.com/fr/professional/blessures-empoisonnement/fractures/fractures-de-cheville",  // 50
        "https://www.msdmanuals.com/it/professionale/traumi-avvelenamento/fratture-alla-caviglia",  // 51
        "https://www.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/22-%E5%A4%96%E5%82%B7%E3%81%A8%E4%B8%AD%E6%AF%92/%E9%AA%A8%E6%8A%98-%E8%84%B1%E8%87%BC-%E3%81%8A%E3%82%88%E3%81%B3%E6%8D%BB%E6%8C%AB/%E8%B6%B3%E9%96%A2%E7%AF%80%E9%AA%A8%E6%8A%98",  // 52
        "https://www.msdmanuals.com/pt/profissional/les%C3%B5es-intoxica%C3%A7%C3%A3o/fraturas/fraturas-do-tornozelo",  // 53
        "https://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/%D1%82%D1%80%D0%B0%D0%B2%D0%BC%D1%8B-%D0%BE%D1%82%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%B8%D1%8F/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B-%D0%B3%D0%BE%D0%BB%D0%B5%D0%BD%D0%BE%D1%81%D1%82%D0%BE%D0%BF%D0%BD%D0%BE%D0%B3%D0%BE-%D1%81%D1%83%D1%81%D1%82%D0%B0%D0%B2%D0%B0",  // 54
        "https://www.msdmanuals.cn/professional/injuries-poisoning/fractures/ankle-fractures",  // 55
        "https://www.msdmanuals.com/vi/chuy%C3%AAn-gia/ch%E1%BA%A5n-th%C6%B0%C6%A1ng-ng%E1%BB%99-%C4%91%E1%BB%99c/g%C3%A3y-x%C6%B0%C6%A1ng/g%C3%A3y-x%C6%B0%C6%A1ng-c%E1%BB%95-ch%C3%A2n",  // 56
        "https://www.merckmanuals.com/home/injuries-and-poisoning/fractures/ankle-fractures?query=ankle%20fractures",  // 57
        "https://www.msdmanuals.com/home/injuries-and-poisoning/fractures/ankle-fractures?query=ankle%20fractures",  // 58
        "https://www.msdmanuals.com/de/heim/verletzungen-und-vergiftung/frakturen/kn%C3%B6chelfrakturen",  // 59
        "https://www.msdmanuals.com/es/hogar/traumatismos-y-envenenamientos/fracturas/fracturas-de-tobillo",  // 60
        "https://www.msdmanuals.com/fr/accueil/l%C3%A9sions-et-intoxications/fractures/fractures-de-la-cheville",  // 61
        "https://www.msdmanuals.com/it/casa/lesioni-e-avvelenamento/fratture/fratture-della-caviglia",  // 62
        "https://www.msdmanuals.com/ja-jp/%E3%83%9B%E3%83%BC%E3%83%A0/25-%E5%A4%96%E5%82%B7%E3%81%A8%E4%B8%AD%E6%AF%92/%E9%AA%A8%E6%8A%98/%E8%B6%B3%E9%A6%96%E3%81%AE%E9%AA%A8%E6%8A%98",  // 63
        "https://www.msdmanuals.com/ko/%ED%99%88/%EB%B6%80%EC%83%81-%EB%B0%8F-%EC%A4%91%EB%8F%85/%EA%B3%A8%EC%A0%88/%EB%B0%9C%EB%AA%A9-%EA%B3%A8%EC%A0%88",  // 64
        "https://www.msdmanuals.com/pt/casa/les%C3%B5es-e-envenenamentos/fraturas/fraturas-do-tornozelo",  // 65
        "https://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/%D1%82%D1%80%D0%B0%D0%B2%D0%BC%D1%8B-%D0%B8-%D0%BE%D1%82%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%B8%D1%8F/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B/%D0%BF%D0%B5%D1%80%D0%B5%D0%BB%D0%BE%D0%BC%D1%8B-%D0%BB%D0%BE%D0%B4%D1%8B%D0%B6%D0%BA%D0%B8",  // 66
        "https://www.msdmanuals.cn/home/injuries-and-poisoning/fractures/ankle-fractures",  // 67
        "https://www.msdmanuals.com/ar/home/%D8%A7%D9%84%D8%A5%D8%B5%D8%A7%D8%A8%D8%A7%D8%AA-%D9%88%D8%A7%D9%84%D8%AA%D9%91%D9%8E%D8%B3%D9%85%D9%91%D9%8F%D9%85/%D8%AD%D8%A7%D9%84%D8%A7%D8%AA%D9%8F-%D8%A7%D9%84%D9%83%D8%B3%D9%88%D8%B1-%D9%88%D8%A7%D9%84%D8%AE%D9%84%D9%88%D8%B9-%D9%88%D8%A7%D9%84%D8%A7%D9%84%D8%AA%D9%88%D8%A7%D8%A1%D8%A7%D8%AA/%D9%83%D9%8F%D8%B3%D9%88%D8%B1%D9%8F-%D8%A7%D9%84%D9%83%D8%A7%D8%AD%D9%84",  // 68

    };

    private static final String[] versions = {
        "EN PV Home Merck",  // 0
        "EN PV Home",  // 1
        "DE PV Home",  // 2
        "ES PV Home",  // 3
        "FR PV Home",  // 4
        "IT PV Home",  // 5
        "JA PV Home",  // 6
        "PT PV Home",  // 7
        "RU PV Home",  // 8
        "CN PV Home",  // 9
        "VI PV Home",  // 10
        "EN CV Home Merck",  // 11
        "EN CV Home",  // 12
        "DE CV Home",  // 13
        "ES CV Home",  // 14
        "FR CV Home",  // 15
        "IT CV Home",  // 16
        "JA CV Home",  // 17
        "KO CV Home",  // 18
        "PT CV Home",  // 19
        "RU CV Home",  // 20
        "CN CV Home",  // 21
        "AR CV Home",  // 22
        "EN PV Acne Merck",  // 23
        "EN PV Acne",  // 24
        "DE PV Acne",  // 25
        "ES PV Acne",  // 26
        "FR PV Acne",  // 27
        "IT PV Acne",  // 28
        "JA PV Acne",  // 29
        "PT PV Acne",  // 30
        "RU PV Acne",  // 31
        "CN PV Acne",  // 32
        "VI PV Acne",  // 33
        "EN CV Acne Merck",  // 34
        "EN CV Acne",  // 35
        "DE CV Acne",  // 36
        "ES CV Acne",  // 37
        "FR CV Acne",  // 38
        "IT CV Acne",  // 39
        "JA CV Acne",  // 40
        "KO CV Acne",  // 41
        "PT CV Acne",  // 42
        "RU CV Acne",  // 43
        "CN CV Acne",  // 44
        "AR CV Acne",  // 45
        "EN PV Ankle Fractures Merck",  // 46
        "EN PV Ankle Fractures",  // 47
        "DE PV Ankle Fractures",  // 48
        "ES PV Ankle Fractures",  // 49
        "FR PV Ankle Fractures",  // 50
        "IT PV Ankle Fractures",  // 51
        "JA PV Ankle Fractures",  // 52
        "PT PV Ankle Fractures",  // 53
        "RU PV Ankle Fractures",  // 54
        "CN PV Ankle Fractures",  // 55
        "VI PV Ankle Fractures",  // 56
        "EN CV Ankle Fractures Merck",  // 57
        "EN CV Ankle Fractures",  // 58
        "DE CV Ankle Fractures",  // 59
        "ES CV Ankle Fractures",  // 60
        "FR CV Ankle Fractures",  // 61
        "IT CV Ankle Fractures",  // 62
        "JA CV Ankle Fractures",  // 63
        "KO CV Ankle Fractures",  // 64
        "PT CV Ankle Fractures",  // 65
        "RU CV Ankle Fractures",  // 66
        "CN CV Ankle Fractures",  // 67
        "AR CV Ankle Fractures"  // 68

    };

    public static void main(String[] args) throws IOException, InterruptedException {
        // Date format to add today's date in the report
        String date = new SimpleDateFormat("yyyy-MM-dd").format(new Date());

        // Create an Excel workbook to store results
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Lighthouse Results");

        // Set up Excel headers
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

        // Iterate over URLs and versions, run Lighthouse, and capture results
        for (int i = 0; i < urls.length; i++) {
            String url = urls[i];
            String version = versions[i];

            // Run Lighthouse and collect results
            String jsonFile = runLighthouse(url, version);
            if (jsonFile != null) {
                parseAndWriteToExcel(jsonFile, workbook, sheet, i + 1, date, version, url);
            } else {
                System.out.println("Failed to run Lighthouse for URL: " + url);
            }
        }

        // Save the Excel file
        try (FileOutputStream fileOut = new FileOutputStream("LighthouseResults.xlsx")) {
            workbook.write(fileOut);
        }

        workbook.close();
        System.out.println("Lighthouse audit complete and results written to LighthouseResults.xlsx");
    }

    // Method to run Lighthouse for a given URL and return the path to the JSON report
    public static String runLighthouse(String url, String version) throws IOException, InterruptedException {
        String fileName = "lighthouse-report-" + version.replaceAll("\\s+", "-") + ".json";
        String[] command = {"lighthouse", url, "--output", "json", "--output-path", fileName, "--preset=mobile", "--quiet"};
        Process process = new ProcessBuilder(command).start();
        process.waitFor();

        File jsonFile = new File(fileName);
        if (jsonFile.exists()) {
            return fileName;
        } else {
            return null;
        }
    }

    // Method to parse Lighthouse JSON report and write data to Excel
    public static void parseAndWriteToExcel(String jsonFile, Workbook workbook, Sheet sheet, int rowIndex, String date, String version, String url) {
        try (FileReader reader = new FileReader(jsonFile)) {
            JsonObject json = JsonParser.parseReader(reader).getAsJsonObject();
            Row row = sheet.createRow(rowIndex);

            // Fill in the Date, Version Name, and URL
            row.createCell(0).setCellValue(date);
            row.createCell(1).setCellValue(version);
            row.createCell(2).setCellValue(url);

            // Get Lighthouse metrics and write them to the row
            JsonObject categories = json.getAsJsonObject("categories");
            row.createCell(3).setCellValue(categories.getAsJsonObject("performance").get("score").getAsDouble());
            row.createCell(4).setCellValue(categories.getAsJsonObject("accessibility").get("score").getAsDouble());
            row.createCell(5).setCellValue(categories.getAsJsonObject("best-practices").get("score").getAsDouble());
            row.createCell(6).setCellValue(categories.getAsJsonObject("seo").get("score").getAsDouble());

            JsonObject audits = json.getAsJsonObject("audits");
            row.createCell(7).setCellValue(audits.getAsJsonObject("first-contentful-paint").get("numericValue").getAsDouble());
            row.createCell(8).setCellValue(audits.getAsJsonObject("largest-contentful-paint").get("numericValue").getAsDouble());
            row.createCell(9).setCellValue(audits.getAsJsonObject("total-blocking-time").get("numericValue").getAsDouble());
            row.createCell(10).setCellValue(audits.getAsJsonObject("cumulative-layout-shift").get("numericValue").getAsDouble());
            row.createCell(11).setCellValue(audits.getAsJsonObject("speed-index").get("numericValue").getAsDouble());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
