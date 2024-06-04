package Source_Code;

import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Get_Source_ALL_EN {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");
		wd = new ChromeDriver();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[262];
		String[] Versions = new String[262];

		// English (EN-AU) PV
		mainURLs[0] = "https://east.msdmanuals.com/en-au/";
		Versions[0] = "VERSION: English (EN-AU) PV - Landing Page";
		mainURLs[1] = "https://east.msdmanuals.com/en-au/professional";
		Versions[1] = "VERSION: English (EN-AU) PV - Home Page";
		mainURLs[2] = "https://east.msdmanuals.com/en-au/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[2] = "VERSION: English (EN-AU) PV - Topic Page";
		mainURLs[3] = "https://east.msdmanuals.com/en-au/professional/pages-with-widgets/figures?mode=list";
		Versions[3] = "VERSION: English (EN-AU) PV - Figures Page";
		mainURLs[4] = "https://east.msdmanuals.com/en-au/professional/resourcespages/about-the-msd-manuals";
		Versions[4] = "VERSION: English (EN-AU) PV - About Page";
		mainURLs[5] = "https://east.msdmanuals.com/en-au/professional/resourcespages/disclaimer";
		Versions[5] = "VERSION: English (EN-AU) PV - Disclaimer Page";
		mainURLs[6] = "https://east.msdmanuals.com/en-au/professional/content/permissions";
		Versions[6] = "VERSION: English (EN-AU) PV - Permissions Page";
		mainURLs[7] = "https://east.msdmanuals.com/en-au/professional/resourcespages/privacy";
		Versions[7] = "VERSION: English (EN-AU) PV - Privacy Page";
		mainURLs[8] = "https://east.msdmanuals.com/en-au/professional/authors";
		Versions[8] = "VERSION: English (EN-AU) PV - Contributors Page";
		mainURLs[9] = "https://east.msdmanuals.com/en-au/professional/content/termsofuse";
		Versions[9] = "VERSION: English (EN-AU) PV - Terms Of Use Page";
		mainURLs[10] = "https://east.msdmanuals.com/en-au/professional/content/licensing";
		Versions[10] = "VERSION: English (EN-AU) PV - Licensing Page";
		mainURLs[11] = "https://east.msdmanuals.com/en-au/professional/content/contact-us";
		Versions[11] = "VERSION: English (EN-AU) PV - Contact Us Page";
		mainURLs[12] = "https://east.msdmanuals.com/en-au/professional/resourcespages/global-medical-knowledge-2020";
		Versions[12] = "VERSION: English (EN-AU) PV - Global Medical Knowledge Page";
		mainURLs[13] = "https://east.msdmanuals.com/en-au/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[13] = "VERSION: English (EN-AU) PV - Normal Laboratory Values Page";
		mainURLs[14] = "https://east.msdmanuals.com/en-au/professional/pages-with-widgets/case-studies?mode=list";
		Versions[14] = "VERSION: English (EN-AU) PV - Case Studies Page";

		// English (EN-AU) CV
		mainURLs[15] = "https://east.msdmanuals.com/en-au/home";
		Versions[15] = "VERSION: English (EN-AU) CV - Home Page";
		mainURLs[16] = "https://east.msdmanuals.com/en-au/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[16] = "VERSION: English (EN-AU) CV - Topic Page";
		mainURLs[17] = "https://east.msdmanuals.com/en-au/professional/pages-with-widgets/figures?mode=list";
		Versions[17] = "VERSION: English (EN-AU) CV - Figures Page";
		mainURLs[18] = "https://east.msdmanuals.com/en-au/home/resourcespages/about-the-manuals";
		Versions[18] = "VERSION: English (EN-AU) CV - About Page";
		mainURLs[19] = "https://east.msdmanuals.com/en-au/home/resourcespages/disclaimer";
		Versions[19] = "VERSION: English (EN-AU) CV - Disclaimer Page";
		mainURLs[20] = "https://east.msdmanuals.com/en-au/home/content/permissions";
		Versions[20] = "VERSION: English (EN-AU) CV - Permissions Page";
		mainURLs[21] = "https://east.msdmanuals.com/en-au/home/resourcespages/privacy";
		Versions[21] = "VERSION: English (EN-AU) CV - Privacy Page";
		mainURLs[22] = "https://east.msdmanuals.com/en-au/home/authors";
		Versions[22] = "VERSION: English (EN-AU) CV - Contributors Page";
		mainURLs[23] = "https://east.msdmanuals.com/en-au/home/content/termsofuse";
		Versions[23] = "VERSION: English (EN-AU) CV - Terms Of Use Page";
		mainURLs[24] = "https://east.msdmanuals.com/en-au/home/content/licensing";
		Versions[24] = "VERSION: English (EN-AU) CV - Licensing Page";
		mainURLs[25] = "https://east.msdmanuals.com/en-au/home/content/contact-us";
		Versions[25] = "VERSION: English (EN-AU) CV - Contact Us Page";
		mainURLs[26] = "https://east.msdmanuals.com/en-au/home/resourcespages/global-medical-knowledge-2020";
		Versions[26] = "VERSION: English (EN-AU) CV - Global Medical Knowledge Page";
		mainURLs[27] = "https://east.msdmanuals.com/en-au/home/pages-with-widgets/infographics?mode=list";
		Versions[27] = "VERSION: English (EN-AU) CV - Infographics Page";
		mainURLs[28] = "https://east.msdmanuals.com/en-au/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[28] = "VERSION: English (EN-AU) CV - The One-Page Manual of Health Page";
/*
		// English (EN-CA) PV
		mainURLs[29] = "https://east.merckmanuals.com/en-ca/";
		Versions[29] = "VERSION: English (EN-CA) PV - Landing Page";
		mainURLs[30] = "https://east.merckmanuals.com/en-ca/professional";
		Versions[30] = "VERSION: English (EN-CA) PV - Home Page";
		mainURLs[31] = "https://east.merckmanuals.com/en-ca/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[31] = "VERSION: English (EN-CA) PV - Topic Page";
		mainURLs[32] = "https://east.merckmanuals.com/en-ca/professional/pages-with-widgets/figures?mode=list";
		Versions[32] = "VERSION: English (EN-CA) PV - Figures Page";
		mainURLs[33] = "https://east.merckmanuals.com/en-ca/professional/resourcespages/about-the-msd-manuals";
		Versions[33] = "VERSION: English (EN-CA) PV - About Page";
		mainURLs[34] = "https://east.merckmanuals.com/en-ca/professional/resourcespages/disclaimer";
		Versions[34] = "VERSION: English (EN-CA) PV - Disclaimer Page";
		mainURLs[35] = "https://east.merckmanuals.com/en-ca/professional/content/permissions";
		Versions[35] = "VERSION: English (EN-CA) PV - Permissions Page";
		mainURLs[36] = "https://east.merckmanuals.com/en-ca/professional/resourcespages/privacy";
		Versions[36] = "VERSION: English (EN-CA) PV - Privacy Page";
		mainURLs[37] = "https://east.merckmanuals.com/en-ca/professional/authors";
		Versions[37] = "VERSION: English (EN-CA) PV - Contributors Page";
		mainURLs[38] = "https://east.merckmanuals.com/en-ca/professional/content/termsofuse";
		Versions[38] = "VERSION: English (EN-CA) PV - Terms Of Use Page";
		mainURLs[39] = "https://east.merckmanuals.com/en-ca/professional/content/licensing";
		Versions[39] = "VERSION: English (EN-CA) PV - Licensing Page";
		mainURLs[40] = "https://east.merckmanuals.com/en-ca/professional/content/contact-us";
		Versions[40] = "VERSION: English (EN-CA) PV - Contact Us Page";
		mainURLs[41] = "https://east.merckmanuals.com/en-ca/professional/resourcespages/global-medical-knowledge-2020";
		Versions[41] = "VERSION: English (EN-CA) PV - Global Medical Knowledge Page";
		mainURLs[42] = "https://east.merckmanuals.com/en-ca/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[42] = "VERSION: English (EN-CA) PV - Normal Laboratory Values Page";
		mainURLs[43] = "https://east.merckmanuals.com/en-ca/professional/pages-with-widgets/case-studies?mode=list";
		Versions[43] = "VERSION: English (EN-CA) PV - Case Studies Page";

		// English (EN-CA) CV
		mainURLs[44] = "https://east.merckmanuals.com/en-ca/home";
		Versions[44] = "VERSION: English (EN-CA) CV - Home Page";
		mainURLs[45] = "https://east.merckmanuals.com/en-ca/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[45] = "VERSION: English (EN-CA) CV - Topic Page";
		mainURLs[46] = "https://east.merckmanuals.com/en-ca/professional/pages-with-widgets/figures?mode=list";
		Versions[46] = "VERSION: English (EN-CA) CV - Figures Page";
		mainURLs[47] = "https://east.merckmanuals.com/en-ca/home/resourcespages/about-the-manuals";
		Versions[47] = "VERSION: English (EN-CA) CV - About Page";
		mainURLs[48] = "https://east.merckmanuals.com/en-ca/home/resourcespages/disclaimer";
		Versions[48] = "VERSION: English (EN-CA) CV - Disclaimer Page";
		mainURLs[49] = "https://east.merckmanuals.com/en-ca/home/content/permissions";
		Versions[49] = "VERSION: English (EN-CA) CV - Permissions Page";
		mainURLs[50] = "https://east.merckmanuals.com/en-ca/home/resourcespages/privacy";
		Versions[50] = "VERSION: English (EN-CA) CV - Privacy Page";
		mainURLs[51] = "https://east.merckmanuals.com/en-ca/home/authors";
		Versions[51] = "VERSION: English (EN-CA) CV - Contributors Page";
		mainURLs[52] = "https://east.merckmanuals.com/en-ca/home/content/termsofuse";
		Versions[52] = "VERSION: English (EN-CA) CV - Terms Of Use Page";
		mainURLs[53] = "https://east.merckmanuals.com/en-ca/home/content/licensing";
		Versions[53] = "VERSION: English (EN-CA) CV - Licensing Page";
		mainURLs[54] = "https://east.merckmanuals.com/en-ca/home/content/contact-us";
		Versions[54] = "VERSION: English (EN-CA) CV - Contact Us Page";
		mainURLs[55] = "https://east.merckmanuals.com/en-ca/home/resourcespages/global-medical-knowledge-2020";
		Versions[55] = "VERSION: English (EN-CA) CV - Global Medical Knowledge Page";
		mainURLs[56] = "https://east.merckmanuals.com/en-ca/home/pages-with-widgets/infographics?mode=list";
		Versions[56] = "VERSION: English (EN-CA) CV - Infographics Page";
		mainURLs[57] = "https://east.merckmanuals.com/en-ca/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[57] = "VERSION: English (EN-CA) CV - The One-Page Manual of Health Page";
*/
		// English (EN-GB) PV
		mainURLs[58] = "https://east.msdmanuals.com/en-gb/";
		Versions[58] = "VERSION: English (EN-GB) PV - Landing Page";
		mainURLs[59] = "https://east.msdmanuals.com/en-gb/professional";
		Versions[59] = "VERSION: English (EN-GB) PV - Home Page";
		mainURLs[60] = "https://east.msdmanuals.com/en-gb/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[60] = "VERSION: English (EN-GB) PV - Topic Page";
		mainURLs[61] = "https://east.msdmanuals.com/en-gb/professional/pages-with-widgets/figures?mode=list";
		Versions[61] = "VERSION: English (EN-GB) PV - Figures Page";
		mainURLs[62] = "https://east.msdmanuals.com/en-gb/professional/resourcespages/about-the-msd-manuals";
		Versions[62] = "VERSION: English (EN-GB) PV - About Page";
		mainURLs[63] = "https://east.msdmanuals.com/en-gb/professional/resourcespages/disclaimer";
		Versions[63] = "VERSION: English (EN-GB) PV - Disclaimer Page";
		mainURLs[64] = "https://east.msdmanuals.com/en-gb/professional/content/permissions";
		Versions[64] = "VERSION: English (EN-GB) PV - Permissions Page";
		mainURLs[65] = "https://east.msdmanuals.com/en-gb/professional/resourcespages/privacy";
		Versions[65] = "VERSION: English (EN-GB) PV - Privacy Page";
		mainURLs[66] = "https://east.msdmanuals.com/en-gb/professional/authors";
		Versions[66] = "VERSION: English (EN-GB) PV - Contributors Page";
		mainURLs[67] = "https://east.msdmanuals.com/en-gb/professional/content/termsofuse";
		Versions[67] = "VERSION: English (EN-GB) PV - Terms Of Use Page";
		mainURLs[68] = "https://east.msdmanuals.com/en-gb/professional/content/licensing";
		Versions[68] = "VERSION: English (EN-GB) PV - Licensing Page";
		mainURLs[69] = "https://east.msdmanuals.com/en-gb/professional/content/contact-us";
		Versions[69] = "VERSION: English (EN-GB) PV - Contact Us Page";
		mainURLs[70] = "https://east.msdmanuals.com/en-gb/professional/resourcespages/global-medical-knowledge-2020";
		Versions[70] = "VERSION: English (EN-GB) PV - Global Medical Knowledge Page";
		mainURLs[71] = "https://east.msdmanuals.com/en-gb/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[71] = "VERSION: English (EN-GB) PV - Normal Laboratory Values Page";
		mainURLs[72] = "https://east.msdmanuals.com/en-gb/professional/pages-with-widgets/case-studies?mode=list";
		Versions[72] = "VERSION: English (EN-GB) PV - Case Studies Page";

		// English (EN-GB) CV
		mainURLs[73] = "https://east.msdmanuals.com/en-gb/home";
		Versions[73] = "VERSION: English (EN-GB) CV - Home Page";
		mainURLs[74] = "https://east.msdmanuals.com/en-gb/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[74] = "VERSION: English (EN-GB) CV - Topic Page";
		mainURLs[75] = "https://east.msdmanuals.com/en-gb/professional/pages-with-widgets/figures?mode=list";
		Versions[75] = "VERSION: English (EN-GB) CV - Figures Page";
		mainURLs[76] = "https://east.msdmanuals.com/en-gb/home/resourcespages/about-the-manuals";
		Versions[76] = "VERSION: English (EN-GB) CV - About Page";
		mainURLs[77] = "https://east.msdmanuals.com/en-gb/home/resourcespages/disclaimer";
		Versions[77] = "VERSION: English (EN-GB) CV - Disclaimer Page";
		mainURLs[78] = "https://east.msdmanuals.com/en-gb/home/content/permissions";
		Versions[78] = "VERSION: English (EN-GB) CV - Permissions Page";
		mainURLs[79] = "https://east.msdmanuals.com/en-gb/home/resourcespages/privacy";
		Versions[79] = "VERSION: English (EN-GB) CV - Privacy Page";
		mainURLs[80] = "https://east.msdmanuals.com/en-gb/home/authors";
		Versions[80] = "VERSION: English (EN-GB) CV - Contributors Page";
		mainURLs[81] = "https://east.msdmanuals.com/en-gb/home/content/termsofuse";
		Versions[81] = "VERSION: English (EN-GB) CV - Terms Of Use Page";
		mainURLs[82] = "https://east.msdmanuals.com/en-gb/home/content/licensing";
		Versions[82] = "VERSION: English (EN-GB) CV - Licensing Page";
		mainURLs[83] = "https://east.msdmanuals.com/en-gb/home/content/contact-us";
		Versions[83] = "VERSION: English (EN-GB) CV - Contact Us Page";
		mainURLs[84] = "https://east.msdmanuals.com/en-gb/home/resourcespages/global-medical-knowledge-2020";
		Versions[84] = "VERSION: English (EN-GB) CV - Global Medical Knowledge Page";
		mainURLs[85] = "https://east.msdmanuals.com/en-gb/home/pages-with-widgets/infographics?mode=list";
		Versions[85] = "VERSION: English (EN-GB) CV - Infographics Page";
		mainURLs[86] = "https://east.msdmanuals.com/en-gb/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[86] = "VERSION: English (EN-GB) CV - The One-Page Manual of Health Page";

		// English (EN-KR) PV
		mainURLs[87] = "https://east.msdmanuals.com/en-kr/";
		Versions[87] = "VERSION: English (EN-KR) PV - Landing Page";
		mainURLs[88] = "https://east.msdmanuals.com/en-kr/professional";
		Versions[88] = "VERSION: English (EN-KR) PV - Home Page";
		mainURLs[89] = "https://east.msdmanuals.com/en-kr/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[89] = "VERSION: English (EN-KR) PV - Topic Page";
		mainURLs[90] = "https://east.msdmanuals.com/en-kr/professional/pages-with-widgets/figures?mode=list";
		Versions[90] = "VERSION: English (EN-KR) PV - Figures Page";
		mainURLs[91] = "https://east.msdmanuals.com/en-kr/professional/resourcespages/about-the-msd-manuals";
		Versions[91] = "VERSION: English (EN-KR) PV - About Page";
		mainURLs[92] = "https://east.msdmanuals.com/en-kr/professional/resourcespages/disclaimer";
		Versions[92] = "VERSION: English (EN-KR) PV - Disclaimer Page";
		mainURLs[93] = "https://east.msdmanuals.com/en-kr/professional/content/permissions";
		Versions[93] = "VERSION: English (EN-KR) PV - Permissions Page";
		mainURLs[94] = "https://east.msdmanuals.com/en-kr/professional/resourcespages/privacy";
		Versions[94] = "VERSION: English (EN-KR) PV - Privacy Page";
		mainURLs[95] = "https://east.msdmanuals.com/en-kr/professional/authors";
		Versions[95] = "VERSION: English (EN-KR) PV - Contributors Page";
		mainURLs[96] = "https://east.msdmanuals.com/en-kr/professional/content/termsofuse";
		Versions[96] = "VERSION: English (EN-KR) PV - Terms Of Use Page";
		mainURLs[97] = "https://east.msdmanuals.com/en-kr/professional/content/licensing";
		Versions[97] = "VERSION: English (EN-KR) PV - Licensing Page";
		mainURLs[98] = "https://east.msdmanuals.com/en-kr/professional/content/contact-us";
		Versions[98] = "VERSION: English (EN-KR) PV - Contact Us Page";
		mainURLs[99] = "https://east.msdmanuals.com/en-kr/professional/resourcespages/global-medical-knowledge-2020";
		Versions[99] = "VERSION: English (EN-KR) PV - Global Medical Knowledge Page";
		mainURLs[100] = "https://east.msdmanuals.com/en-kr/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[100] = "VERSION: English (EN-KR) PV - Normal Laboratory Values Page";
		mainURLs[101] = "https://east.msdmanuals.com/en-kr/professional/pages-with-widgets/case-studies?mode=list";
		Versions[101] = "VERSION: English (EN-KR) PV - Case Studies Page";

		// English (EN-KR) CV
		mainURLs[102] = "https://east.msdmanuals.com/en-kr/home";
		Versions[102] = "VERSION: English (EN-KR) CV - Home Page";
		mainURLs[103] = "https://east.msdmanuals.com/en-kr/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[103] = "VERSION: English (EN-KR) CV - Topic Page";
		mainURLs[104] = "https://east.msdmanuals.com/en-kr/professional/pages-with-widgets/figures?mode=list";
		Versions[104] = "VERSION: English (EN-KR) CV - Figures Page";
		mainURLs[105] = "https://east.msdmanuals.com/en-kr/home/resourcespages/about-the-manuals";
		Versions[105] = "VERSION: English (EN-KR) CV - About Page";
		mainURLs[106] = "https://east.msdmanuals.com/en-kr/home/resourcespages/disclaimer";
		Versions[106] = "VERSION: English (EN-KR) CV - Disclaimer Page";
		mainURLs[107] = "https://east.msdmanuals.com/en-kr/home/content/permissions";
		Versions[107] = "VERSION: English (EN-KR) CV - Permissions Page";
		mainURLs[108] = "https://east.msdmanuals.com/en-kr/home/resourcespages/privacy";
		Versions[108] = "VERSION: English (EN-KR) CV - Privacy Page";
		mainURLs[109] = "https://east.msdmanuals.com/en-kr/home/authors";
		Versions[109] = "VERSION: English (EN-KR) CV - Contributors Page";
		mainURLs[110] = "https://east.msdmanuals.com/en-kr/home/content/termsofuse";
		Versions[110] = "VERSION: English (EN-KR) CV - Terms Of Use Page";
		mainURLs[111] = "https://east.msdmanuals.com/en-kr/home/content/licensing";
		Versions[111] = "VERSION: English (EN-KR) CV - Licensing Page";
		mainURLs[112] = "https://east.msdmanuals.com/en-kr/home/content/contact-us";
		Versions[112] = "VERSION: English (EN-KR) CV - Contact Us Page";
		mainURLs[113] = "https://east.msdmanuals.com/en-kr/home/resourcespages/global-medical-knowledge-2020";
		Versions[113] = "VERSION: English (EN-KR) CV - Global Medical Knowledge Page";
		mainURLs[114] = "https://east.msdmanuals.com/en-kr/home/pages-with-widgets/infographics?mode=list";
		Versions[114] = "VERSION: English (EN-KR) CV - Infographics Page";
		mainURLs[115] = "https://east.msdmanuals.com/en-kr/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[115] = "VERSION: English (EN-KR) CV - The One-Page Manual of Health Page";

		// English (EN-NZ) PV
		mainURLs[116] = "https://east.msdmanuals.com/en-nz/";
		Versions[116] = "VERSION: English (EN-NZ) PV - Landing Page";
		mainURLs[117] = "https://east.msdmanuals.com/en-nz/professional";
		Versions[117] = "VERSION: English (EN-NZ) PV - Home Page";
		mainURLs[118] = "https://east.msdmanuals.com/en-nz/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[118] = "VERSION: English (EN-NZ) PV - Topic Page";
		mainURLs[119] = "https://east.msdmanuals.com/en-nz/professional/pages-with-widgets/figures?mode=list";
		Versions[119] = "VERSION: English (EN-NZ) PV - Figures Page";
		mainURLs[120] = "https://east.msdmanuals.com/en-nz/professional/resourcespages/about-the-msd-manuals";
		Versions[120] = "VERSION: English (EN-NZ) PV - About Page";
		mainURLs[121] = "https://east.msdmanuals.com/en-nz/professional/resourcespages/disclaimer";
		Versions[121] = "VERSION: English (EN-NZ) PV - Disclaimer Page";
		mainURLs[122] = "https://east.msdmanuals.com/en-nz/professional/content/permissions";
		Versions[122] = "VERSION: English (EN-NZ) PV - Permissions Page";
		mainURLs[123] = "https://east.msdmanuals.com/en-nz/professional/resourcespages/privacy";
		Versions[123] = "VERSION: English (EN-NZ) PV - Privacy Page";
		mainURLs[124] = "https://east.msdmanuals.com/en-nz/professional/authors";
		Versions[124] = "VERSION: English (EN-NZ) PV - Contributors Page";
		mainURLs[125] = "https://east.msdmanuals.com/en-nz/professional/content/termsofuse";
		Versions[125] = "VERSION: English (EN-NZ) PV - Terms Of Use Page";
		mainURLs[126] = "https://east.msdmanuals.com/en-nz/professional/content/licensing";
		Versions[126] = "VERSION: English (EN-NZ) PV - Licensing Page";
		mainURLs[127] = "https://east.msdmanuals.com/en-nz/professional/content/contact-us";
		Versions[127] = "VERSION: English (EN-NZ) PV - Contact Us Page";
		mainURLs[128] = "https://east.msdmanuals.com/en-nz/professional/resourcespages/global-medical-knowledge-2020";
		Versions[128] = "VERSION: English (EN-NZ) PV - Global Medical Knowledge Page";
		mainURLs[129] = "https://east.msdmanuals.com/en-nz/professional/resources/normal-laboratory-values/normal-laboratory-values";
		Versions[129] = "VERSION: English (EN-NZ) PV - Normal Laboratory Values Page";
		mainURLs[130] = "https://east.msdmanuals.com/en-nz/professional/pages-with-widgets/case-studies?mode=list";
		Versions[130] = "VERSION: English (EN-NZ) PV - Case Studies Page";

		// English (EN-NZ) CV
		mainURLs[131] = "https://east.msdmanuals.com/en-nz/home";
		Versions[131] = "VERSION: English (EN-NZ) CV - Home Page";
		mainURLs[132] = "https://east.msdmanuals.com/en-nz/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[132] = "VERSION: English (EN-NZ) CV - Topic Page";
		mainURLs[133] = "https://east.msdmanuals.com/en-nz/professional/pages-with-widgets/figures?mode=list";
		Versions[133] = "VERSION: English (EN-NZ) CV - Figures Page";
		mainURLs[134] = "https://east.msdmanuals.com/en-nz/home/resourcespages/about-the-manuals";
		Versions[134] = "VERSION: English (EN-NZ) CV - About Page";
		mainURLs[135] = "https://east.msdmanuals.com/en-nz/home/resourcespages/disclaimer";
		Versions[135] = "VERSION: English (EN-NZ) CV - Disclaimer Page";
		mainURLs[136] = "https://east.msdmanuals.com/en-nz/home/content/permissions";
		Versions[136] = "VERSION: English (EN-NZ) CV - Permissions Page";
		mainURLs[137] = "https://east.msdmanuals.com/en-nz/home/resourcespages/privacy";
		Versions[137] = "VERSION: English (EN-NZ) CV - Privacy Page";
		mainURLs[138] = "https://east.msdmanuals.com/en-nz/home/authors";
		Versions[138] = "VERSION: English (EN-NZ) CV - Contributors Page";
		mainURLs[139] = "https://east.msdmanuals.com/en-nz/home/content/termsofuse";
		Versions[139] = "VERSION: English (EN-NZ) CV - Terms Of Use Page";
		mainURLs[140] = "https://east.msdmanuals.com/en-nz/home/content/licensing";
		Versions[140] = "VERSION: English (EN-NZ) CV - Licensing Page";
		mainURLs[141] = "https://east.msdmanuals.com/en-nz/home/content/contact-us";
		Versions[141] = "VERSION: English (EN-NZ) CV - Contact Us Page";
		mainURLs[142] = "https://east.msdmanuals.com/en-nz/home/resourcespages/global-medical-knowledge-2020";
		Versions[142] = "VERSION: English (EN-NZ) CV - Global Medical Knowledge Page";
		mainURLs[143] = "https://east.msdmanuals.com/en-nz/home/pages-with-widgets/infographics?mode=list";
		Versions[143] = "VERSION: English (EN-NZ) CV - Infographics Page";
		mainURLs[144] = "https://east.msdmanuals.com/en-nz/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[144] = "VERSION: English (EN-NZ) CV - The One-Page Manual of Health Page";
/*
		// English (EN-PR) PV
		mainURLs[145] = "https://east.merckmanuals.com/en-pr/";
		Versions[145] = "VERSION: English (EN-PR) PV - Landing Page";
		mainURLs[146] = "https://east.merckmanuals.com/en-pr/professional";
		Versions[146] = "VERSION: English (EN-PR) PV - Home Page";
		mainURLs[147] = "https://east.merckmanuals.com/en-pr/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[147] = "VERSION: English (EN-PR) PV - Topic Page";
		mainURLs[148] = "https://east.merckmanuals.com/en-pr/professional/pages-with-widgets/figures?mode=list";
		Versions[148] = "VERSION: English (EN-PR) PV - Figures Page";
		mainURLs[149] = "https://east.merckmanuals.com/en-pr/professional/resourcespages/about-the-msd-manuals";
		Versions[149] = "VERSION: English (EN-PR) PV - About Page";
		mainURLs[150] = "https://east.merckmanuals.com/en-pr/professional/resourcespages/disclaimer";
		Versions[150] = "VERSION: English (EN-PR) PV - Disclaimer Page";
		mainURLs[151] = "https://east.merckmanuals.com/en-pr/professional/content/permissions";
		Versions[151] = "VERSION: English (EN-PR) PV - Permissions Page";
		mainURLs[152] = "https://east.merckmanuals.com/en-pr/professional/resourcespages/privacy";
		Versions[152] = "VERSION: English (EN-PR) PV - Privacy Page";
		mainURLs[153] = "https://east.merckmanuals.com/en-pr/professional/authors";
		Versions[153] = "VERSION: English (EN-PR) PV - Contributors Page";
		mainURLs[154] = "https://east.merckmanuals.com/en-pr/professional/content/termsofuse";
		Versions[154] = "VERSION: English (EN-PR) PV - Terms Of Use Page";
		mainURLs[155] = "https://east.merckmanuals.com/en-pr/professional/content/licensing";
		Versions[155] = "VERSION: English (EN-PR) PV - Licensing Page";
		mainURLs[156] = "https://east.merckmanuals.com/en-pr/professional/content/contact-us";
		Versions[156] = "VERSION: English (EN-PR) PV - Contact Us Page";
		mainURLs[157] = "https://east.merckmanuals.com/en-pr/professional/resourcespages/global-medical-knowledge-2020";
		Versions[157] = "VERSION: English (EN-PR) PV - Global Medical Knowledge Page";
		mainURLs[158] = "https://east.merckmanuals.com/en-pr/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[158] = "VERSION: English (EN-PR) PV - Normal Laboratory Values Page";
		mainURLs[159] = "https://east.merckmanuals.com/en-pr/professional/pages-with-widgets/case-studies?mode=list";
		Versions[159] = "VERSION: English (EN-PR) PV - Case Studies Page";

		// English (EN-PR) CV
		mainURLs[160] = "https://east.merckmanuals.com/en-pr/home";
		Versions[160] = "VERSION: English (EN-PR) CV - Home Page";
		mainURLs[161] = "https://east.merckmanuals.com/en-pr/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[161] = "VERSION: English (EN-PR) CV - Topic Page";
		mainURLs[162] = "https://east.merckmanuals.com/en-pr/professional/pages-with-widgets/figures?mode=list";
		Versions[162] = "VERSION: English (EN-PR) CV - Figures Page";
		mainURLs[163] = "https://east.merckmanuals.com/en-pr/home/resourcespages/about-the-manuals";
		Versions[163] = "VERSION: English (EN-PR) CV - About Page";
		mainURLs[164] = "https://east.merckmanuals.com/en-pr/home/resourcespages/disclaimer";
		Versions[164] = "VERSION: English (EN-PR) CV - Disclaimer Page";
		mainURLs[165] = "https://east.merckmanuals.com/en-pr/home/content/permissions";
		Versions[165] = "VERSION: English (EN-PR) CV - Permissions Page";
		mainURLs[166] = "https://east.merckmanuals.com/en-pr/home/resourcespages/privacy";
		Versions[166] = "VERSION: English (EN-PR) CV - Privacy Page";
		mainURLs[167] = "https://east.merckmanuals.com/en-pr/home/authors";
		Versions[167] = "VERSION: English (EN-PR) CV - Contributors Page";
		mainURLs[168] = "https://east.merckmanuals.com/en-pr/home/content/termsofuse";
		Versions[168] = "VERSION: English (EN-PR) CV - Terms Of Use Page";
		mainURLs[169] = "https://east.merckmanuals.com/en-pr/home/content/licensing";
		Versions[169] = "VERSION: English (EN-PR) CV - Licensing Page";
		mainURLs[170] = "https://east.merckmanuals.com/en-pr/home/content/contact-us";
		Versions[170] = "VERSION: English (EN-PR) CV - Contact Us Page";
		mainURLs[171] = "https://east.merckmanuals.com/en-pr/home/resourcespages/global-medical-knowledge-2020";
		Versions[171] = "VERSION: English (EN-PR) CV - Global Medical Knowledge Page";
		mainURLs[172] = "https://east.merckmanuals.com/en-pr/home/pages-with-widgets/infographics?mode=list";
		Versions[172] = "VERSION: English (EN-PR) CV - Infographics Page";
		mainURLs[173] = "https://east.merckmanuals.com/en-pr/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[173] = "VERSION: English (EN-PR) CV - The One-Page Manual of Health Page";
*/
		// English (EN-PT) PV
		mainURLs[174] = "https://east.msdmanuals.com/en-pt/";
		Versions[174] = "VERSION: English (EN-PT) PV - Landing Page";
		mainURLs[175] = "https://east.msdmanuals.com/en-pt/professional";
		Versions[175] = "VERSION: English (EN-PT) PV - Home Page";
		mainURLs[176] = "https://east.msdmanuals.com/en-pt/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[176] = "VERSION: English (EN-PT) PV - Topic Page";
		mainURLs[178] = "https://east.msdmanuals.com/en-pt/professional/pages-with-widgets/figures?mode=list";
		Versions[178] = "VERSION: English (EN-PT) PV - Figures Page";
		mainURLs[179] = "https://east.msdmanuals.com/en-pt/professional/resourcespages/about-the-msd-manuals";
		Versions[179] = "VERSION: English (EN-PT) PV - About Page";
		mainURLs[180] = "https://east.msdmanuals.com/en-pt/professional/resourcespages/disclaimer";
		Versions[180] = "VERSION: English (EN-PT) PV - Disclaimer Page";
		mainURLs[181] = "https://east.msdmanuals.com/en-pt/professional/content/permissions";
		Versions[181] = "VERSION: English (EN-PT) PV - Permissions Page";
		mainURLs[182] = "https://east.msdmanuals.com/en-pt/professional/resourcespages/privacy";
		Versions[182] = "VERSION: English (EN-PT) PV - Privacy Page";
		mainURLs[183] = "https://east.msdmanuals.com/en-pt/professional/authors";
		Versions[183] = "VERSION: English (EN-PT) PV - Contributors Page";
		mainURLs[184] = "https://east.msdmanuals.com/en-pt/professional/content/termsofuse";
		Versions[184] = "VERSION: English (EN-PT) PV - Terms Of Use Page";
		mainURLs[185] = "https://east.msdmanuals.com/en-pt/professional/content/licensing";
		Versions[185] = "VERSION: English (EN-PT) PV - Licensing Page";
		mainURLs[186] = "https://east.msdmanuals.com/en-pt/professional/content/contact-us";
		Versions[186] = "VERSION: English (EN-PT) PV - Contact Us Page";
		mainURLs[187] = "https://east.msdmanuals.com/en-pt/professional/resourcespages/global-medical-knowledge-2020";
		Versions[187] = "VERSION: English (EN-PT) PV - Global Medical Knowledge Page";
		mainURLs[188] = "https://east.msdmanuals.com/en-pt/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[188] = "VERSION: English (EN-PT) PV - Normal Laboratory Values Page";
		mainURLs[189] = "https://east.msdmanuals.com/en-pt/professional/pages-with-widgets/case-studies?mode=list";
		Versions[189] = "VERSION: English (EN-PT) PV - Case Studies Page";

		// English (EN-PT) CV
		mainURLs[190] = "https://east.msdmanuals.com/en-pt/home";
		Versions[190] = "VERSION: English (EN-PT) CV - Home Page";
		mainURLs[191] = "https://east.msdmanuals.com/en-pt/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[191] = "VERSION: English (EN-PT) CV - Topic Page";
		mainURLs[192] = "https://east.msdmanuals.com/en-pt/professional/pages-with-widgets/figures?mode=list";
		Versions[192] = "VERSION: English (EN-PT) CV - Figures Page";
		mainURLs[193] = "https://east.msdmanuals.com/en-pt/home/resourcespages/about-the-manuals";
		Versions[193] = "VERSION: English (EN-PT) CV - About Page";
		mainURLs[194] = "https://east.msdmanuals.com/en-pt/home/resourcespages/disclaimer";
		Versions[194] = "VERSION: English (EN-PT) CV - Disclaimer Page";
		mainURLs[195] = "https://east.msdmanuals.com/en-pt/home/content/permissions";
		Versions[195] = "VERSION: English (EN-PT) CV - Permissions Page";
		mainURLs[196] = "https://east.msdmanuals.com/en-pt/home/resourcespages/privacy";
		Versions[196] = "VERSION: English (EN-PT) CV - Privacy Page";
		mainURLs[197] = "https://east.msdmanuals.com/en-pt/home/authors";
		Versions[197] = "VERSION: English (EN-PT) CV - Contributors Page";
		mainURLs[198] = "https://east.msdmanuals.com/en-pt/home/content/termsofuse";
		Versions[198] = "VERSION: English (EN-PT) CV - Terms Of Use Page";
		mainURLs[199] = "https://east.msdmanuals.com/en-pt/home/content/licensing";
		Versions[199] = "VERSION: English (EN-PT) CV - Licensing Page";
		mainURLs[200] = "https://east.msdmanuals.com/en-pt/home/content/contact-us";
		Versions[200] = "VERSION: English (EN-PT) CV - Contact Us Page";
		mainURLs[201] = "https://east.msdmanuals.com/en-pt/home/resourcespages/global-medical-knowledge-2020";
		Versions[201] = "VERSION: English (EN-PT) CV - Global Medical Knowledge Page";
		mainURLs[202] = "https://east.msdmanuals.com/en-pt/home/pages-with-widgets/infographics?mode=list";
		Versions[202] = "VERSION: English (EN-PT) CV - Infographics Page";
		mainURLs[203] = "https://east.msdmanuals.com/en-pt/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[203] = "VERSION: English (EN-PT) CV - The One-Page Manual of Health Page";

		// English (EN-SG) PV
		mainURLs[204] = "https://east.msdmanuals.com/en-sg/";
		Versions[204] = "VERSION: English (EN-SG) PV - Landing Page";
		mainURLs[205] = "https://east.msdmanuals.com/en-sg/professional";
		Versions[205] = "VERSION: English (EN-SG) PV - Home Page";
		mainURLs[206] = "https://east.msdmanuals.com/en-sg/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[206] = "VERSION: English (EN-SG) PV - Topic Page";
		mainURLs[207] = "https://east.msdmanuals.com/en-sg/professional/pages-with-widgets/figures?mode=list";
		Versions[207] = "VERSION: English (EN-SG) PV - Figures Page";
		mainURLs[208] = "https://east.msdmanuals.com/en-sg/professional/resourcespages/about-the-msd-manuals";
		Versions[208] = "VERSION: English (EN-SG) PV - About Page";
		mainURLs[209] = "https://east.msdmanuals.com/en-sg/professional/resourcespages/disclaimer";
		Versions[209] = "VERSION: English (EN-SG) PV - Disclaimer Page";
		mainURLs[210] = "https://east.msdmanuals.com/en-sg/professional/content/permissions";
		Versions[210] = "VERSION: English (EN-SG) PV - Permissions Page";
		mainURLs[211] = "https://east.msdmanuals.com/en-sg/professional/resourcespages/privacy";
		Versions[211] = "VERSION: English (EN-SG) PV - Privacy Page";
		mainURLs[212] = "https://east.msdmanuals.com/en-sg/professional/authors";
		Versions[212] = "VERSION: English (EN-SG) PV - Contributors Page";
		mainURLs[213] = "https://east.msdmanuals.com/en-sg/professional/content/termsofuse";
		Versions[213] = "VERSION: English (EN-SG) PV - Terms Of Use Page";
		mainURLs[214] = "https://east.msdmanuals.com/en-sg/professional/content/licensing";
		Versions[214] = "VERSION: English (EN-SG) PV - Licensing Page";
		mainURLs[215] = "https://east.msdmanuals.com/en-sg/professional/content/contact-us";
		Versions[215] = "VERSION: English (EN-SG) PV - Contact Us Page";
		mainURLs[216] = "https://east.msdmanuals.com/en-sg/professional/resourcespages/global-medical-knowledge-2020";
		Versions[216] = "VERSION: English (EN-SG) PV - Global Medical Knowledge Page";
		mainURLs[217] = "https://east.msdmanuals.com/en-sg/professional/resources/normal-laboratory-values/normal-laboratory-values";
		Versions[217] = "VERSION: English (EN-SG) PV - Normal Laboratory Values Page";
		mainURLs[218] = "https://east.msdmanuals.com/en-sg/professional/pages-with-widgets/case-studies?mode=list";
		Versions[218] = "VERSION: English (EN-SG) PV - Case Studies Page";

		// English (EN-SG) CV
		mainURLs[219] = "https://east.msdmanuals.com/en-sg/home";
		Versions[219] = "VERSION: English (EN-SG) CV - Home Page";
		mainURLs[220] = "https://east.msdmanuals.com/en-sg/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[220] = "VERSION: English (EN-SG) CV - Topic Page";
		mainURLs[221] = "https://east.msdmanuals.com/en-sg/professional/pages-with-widgets/figures?mode=list";
		Versions[221] = "VERSION: English (EN-SG) CV - Figures Page";
		mainURLs[222] = "https://east.msdmanuals.com/en-sg/home/resourcespages/about-the-manuals";
		Versions[222] = "VERSION: English (EN-SG) CV - About Page";
		mainURLs[223] = "https://east.msdmanuals.com/en-sg/home/resourcespages/disclaimer";
		Versions[223] = "VERSION: English (EN-SG) CV - Disclaimer Page";
		mainURLs[224] = "https://east.msdmanuals.com/en-sg/home/content/permissions";
		Versions[224] = "VERSION: English (EN-SG) CV - Permissions Page";
		mainURLs[225] = "https://east.msdmanuals.com/en-sg/home/resourcespages/privacy";
		Versions[225] = "VERSION: English (EN-SG) CV - Privacy Page";
		mainURLs[226] = "https://east.msdmanuals.com/en-sg/home/authors";
		Versions[226] = "VERSION: English (EN-SG) CV - Contributors Page";
		mainURLs[227] = "https://east.msdmanuals.com/en-sg/home/content/termsofuse";
		Versions[227] = "VERSION: English (EN-SG) CV - Terms Of Use Page";
		mainURLs[228] = "https://east.msdmanuals.com/en-sg/home/content/licensing";
		Versions[228] = "VERSION: English (EN-SG) CV - Licensing Page";
		mainURLs[229] = "https://east.msdmanuals.com/en-sg/home/content/contact-us";
		Versions[229] = "VERSION: English (EN-SG) CV - Contact Us Page";
		mainURLs[230] = "https://east.msdmanuals.com/en-sg/home/resourcespages/global-medical-knowledge-2020";
		Versions[230] = "VERSION: English (EN-SG) CV - Global Medical Knowledge Page";
		mainURLs[231] = "https://east.msdmanuals.com/en-sg/home/pages-with-widgets/infographics?mode=list";
		Versions[231] = "VERSION: English (EN-SG) CV - Infographics Page";
		mainURLs[232] = "https://east.msdmanuals.com/en-sg/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[232] = "VERSION: English (EN-SG) CV - The One-Page Manual of Health Page";
/*
		// English (EN-US) PV
		mainURLs[233] = "https://east.merckmanuals.com/";
		Versions[233] = "VERSION: English (EN-US) PV - Landing Page";
		mainURLs[234] = "https://east.merckmanuals.com/professional";
		Versions[234] = "VERSION: English (EN-US) PV - Home Page";
		mainURLs[235] = "https://east.merckmanuals.com/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[235] = "VERSION: English (EN-US) PV - Topic Page";
		mainURLs[236] = "https://east.merckmanuals.com/professional/pages-with-widgets/figures?mode=list";
		Versions[236] = "VERSION: English (EN-US) PV - Figures Page";
		mainURLs[237] = "https://east.merckmanuals.com/professional/resourcespages/about-the-msd-manuals";
		Versions[237] = "VERSION: English (EN-US) PV - About Page";
		mainURLs[238] = "https://east.merckmanuals.com/professional/resourcespages/disclaimer";
		Versions[238] = "VERSION: English (EN-US) PV - Disclaimer Page";
		mainURLs[239] = "https://east.merckmanuals.com/professional/content/permissions";
		Versions[239] = "VERSION: English (EN-US) PV - Permissions Page";
		mainURLs[240] = "https://east.merckmanuals.com/professional/resourcespages/privacy";
		Versions[240] = "VERSION: English (EN-US) PV - Privacy Page";
		mainURLs[241] = "https://east.merckmanuals.com/professional/authors";
		Versions[241] = "VERSION: English (EN-US) PV - Contributors Page";
		mainURLs[242] = "https://east.merckmanuals.com/professional/content/termsofuse";
		Versions[242] = "VERSION: English (EN-US) PV - Terms Of Use Page";
		mainURLs[243] = "https://east.merckmanuals.com/professional/content/licensing";
		Versions[243] = "VERSION: English (EN-US) PV - Licensing Page";
		mainURLs[244] = "https://east.merckmanuals.com/professional/content/contact-us";
		Versions[244] = "VERSION: English (EN-US) PV - Contact Us Page";
		mainURLs[245] = "https://east.merckmanuals.com/professional/resourcespages/global-medical-knowledge-2020";
		Versions[245] = "VERSION: English (EN-US) PV - Global Medical Knowledge Page";
		mainURLs[246] = "https://east.merckmanuals.com/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[246] = "VERSION: English (EN-US) PV - Normal Laboratory Values Page";
		mainURLs[247] = "https://east.merckmanuals.com/professional/pages-with-widgets/case-studies?mode=list";
		Versions[247] = "VERSION: English (EN-US) PV - Case Studies Page";

		// English (EN-US) CV
		mainURLs[248] = "https://east.merckmanuals.com/home";
		Versions[248] = "VERSION: English (EN-US) CV - Home Page";
		mainURLs[249] = "https://east.merckmanuals.com/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[249] = "VERSION: English (EN-US) CV - Topic Page";
		mainURLs[250] = "https://east.merckmanuals.com/professional/pages-with-widgets/figures?mode=list";
		Versions[250] = "VERSION: English (EN-US) CV - Figures Page";
		mainURLs[251] = "https://east.merckmanuals.com/home/resourcespages/about-the-manuals";
		Versions[251] = "VERSION: English (EN-US) CV - About Page";
		mainURLs[252] = "https://east.merckmanuals.com/home/resourcespages/disclaimer";
		Versions[252] = "VERSION: English (EN-US) CV - Disclaimer Page";
		mainURLs[253] = "https://east.merckmanuals.com/home/content/permissions";
		Versions[253] = "VERSION: English (EN-US) CV - Permissions Page";
		mainURLs[254] = "https://east.merckmanuals.com/home/resourcespages/privacy";
		Versions[254] = "VERSION: English (EN-US) CV - Privacy Page";
		mainURLs[255] = "https://east.merckmanuals.com/home/authors";
		Versions[255] = "VERSION: English (EN-US) CV - Contributors Page";
		mainURLs[256] = "https://east.merckmanuals.com/home/content/termsofuse";
		Versions[256] = "VERSION: English (EN-US) CV - Terms Of Use Page";
		mainURLs[257] = "https://east.merckmanuals.com/home/content/licensing";
		Versions[257] = "VERSION: English (EN-US) CV - Licensing Page";
		mainURLs[258] = "https://east.merckmanuals.com/home/content/contact-us";
		Versions[258] = "VERSION: English (EN-US) CV - Contact Us Page";
		mainURLs[259] = "https://east.merckmanuals.com/home/resourcespages/global-medical-knowledge-2020";
		Versions[259] = "VERSION: English (EN-US) CV - Global Medical Knowledge Page";
		mainURLs[260] = "https://east.merckmanuals.com/home/pages-with-widgets/infographics?mode=list";
		Versions[260] = "VERSION: English (EN-US) CV - Infographics Page";
		mainURLs[261] = "https://east.merckmanuals.com/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[261] = "VERSION: English (EN-US) CV - The One-Page Manual of Health Page";
*/
		// Main Loop
		for (int j = 0; j < 262; j++) {

			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				System.out.println(Versions[j]);

				// Get all HTML Page Source code into String Variable
				try {
					String pageSource = wd.getPageSource();
					System.out.println(pageSource);

				} catch (Exception e) {
					System.out.println("Can't get source code");
				}
				
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
