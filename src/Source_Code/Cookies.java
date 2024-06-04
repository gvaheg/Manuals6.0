package Source_Code;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import java.io.BufferedWriter;		
import java.io.File;		
import java.io.FileWriter;
import java.util.Set;
import org.openqa.selenium.By;		
import org.openqa.selenium.WebDriver;		
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.Cookie;

public class Cookies {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[307];
		String[] Versions = new String[307];

		// English (EN) PV
		mainURLs[0] = "http://www.msdmanuals.com/";
		Versions[0] = "VERSION: English (EN) PV - Landing Page";
		mainURLs[1] = "http://www.msdmanuals.com/professional";
		Versions[1] = "VERSION: English (EN) PV - Home Page";
		mainURLs[2] = "http://www.msdmanuals.com/professional/clinical-pharmacology/pharmacokinetics/overview-of-pharmacokinetics";
		Versions[2] = "VERSION: English (EN) PV - Topic Page";
		mainURLs[3] = "http://www.msdmanuals.com/professional/pages-with-widgets/figures?mode=list";
		Versions[3] = "VERSION: English (EN) PV - Figures Page";
		mainURLs[4] = "http://www.msdmanuals.com/professional/resourcespages/about-the-msd-manuals";
		Versions[4] = "VERSION: English (EN) PV - About Page";
		mainURLs[5] = "http://www.msdmanuals.com/professional/resourcespages/disclaimer";
		Versions[5] = "VERSION: English (EN) PV - Disclaimer Page";
		mainURLs[6] = "http://www.msdmanuals.com/professional/content/permissions";
		Versions[6] = "VERSION: English (EN) PV - Permissions Page";
		mainURLs[7] = "http://www.msdmanuals.com/professional/resourcespages/privacy";
		Versions[7] = "VERSION: English (EN) PV - Privacy Page";
		mainURLs[8] = "http://www.msdmanuals.com/professional/authors";
		Versions[8] = "VERSION: English (EN) PV - Contributors Page";
		mainURLs[9] = "http://www.msdmanuals.com/professional/content/termsofuse";
		Versions[9] = "VERSION: English (EN) PV - Terms Of Use Page";
		mainURLs[10] = "http://www.msdmanuals.com/professional/content/licensing";
		Versions[10] = "VERSION: English (EN) PV - Licensing Page";
		mainURLs[11] = "http://www.msdmanuals.com/professional/content/contact-us";
		Versions[11] = "VERSION: English (EN) PV - Contact Us Page";
		mainURLs[12] = "http://www.msdmanuals.com/professional/resourcespages/global-medical-knowledge-2020";
		Versions[12] = "VERSION: English (EN) PV - Global Medical Knowledge Page";
		mainURLs[13] = "http://www.msdmanuals.com/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[13] = "VERSION: English (EN) PV - Normal Laboratory Values Page";
		mainURLs[14] = "http://www.msdmanuals.com/professional/pages-with-widgets/case-studies?mode=list";
		Versions[14] = "VERSION: English (EN) PV - Case Studies Page";

		// English (EN) CV
		mainURLs[15] = "http://www.msdmanuals.com/home";
		Versions[15] = "VERSION: English (EN) CV - Home Page";
		mainURLs[16] = "http://www.msdmanuals.com/home/cancer/overview-of-cancer/overview-of-cancer";
		Versions[16] = "VERSION: English (EN) CV - Topic Page";
		mainURLs[17] = "http://www.msdmanuals.com/professional/pages-with-widgets/figures?mode=list";
		Versions[17] = "VERSION: English (EN) CV - Figures Page";
		mainURLs[18] = "http://www.msdmanuals.com/home/resourcespages/about-the-manuals";
		Versions[18] = "VERSION: English (EN) CV - About Page";
		mainURLs[19] = "http://www.msdmanuals.com/home/resourcespages/disclaimer";
		Versions[19] = "VERSION: English (EN) CV - Disclaimer Page";
		mainURLs[20] = "http://www.msdmanuals.com/home/content/permissions";
		Versions[20] = "VERSION: English (EN) CV - Permissions Page";
		mainURLs[21] = "http://www.msdmanuals.com/home/resourcespages/privacy";
		Versions[21] = "VERSION: English (EN) CV - Privacy Page";
		mainURLs[22] = "http://www.msdmanuals.com/home/authors";
		Versions[22] = "VERSION: English (EN) CV - Contributors Page";
		mainURLs[23] = "http://www.msdmanuals.com/home/content/termsofuse";
		Versions[23] = "VERSION: English (EN) CV - Terms Of Use Page";
		mainURLs[24] = "http://www.msdmanuals.com/home/content/licensing";
		Versions[24] = "VERSION: English (EN) CV - Licensing Page";
		mainURLs[25] = "http://www.msdmanuals.com/home/content/contact-us";
		Versions[25] = "VERSION: English (EN) CV - Contact Us Page";
		mainURLs[26] = "http://www.msdmanuals.com/home/resourcespages/global-medical-knowledge-2020";
		Versions[26] = "VERSION: English (EN) CV - Global Medical Knowledge Page";
		mainURLs[27] = "http://www.msdmanuals.com/home/pages-with-widgets/infographics?mode=list";
		Versions[27] = "VERSION: English (EN) CV - Infographics Page";
		mainURLs[28] = "http://www.msdmanuals.com/home/resources/the-one-page-manual-of-health/one-page-manual-of-health";
		Versions[28] = "VERSION: English (EN) CV - The One-Page Manual of Health Page";

		// Spanish (ES) PV
		mainURLs[29] = "http://www.msdmanuals.com/es";
		Versions[29] = "VERSION: Spanish (ES) - Landing Page";
		mainURLs[30] = "http://www.msdmanuals.com/es/professional";
		Versions[30] = "VERSION: Spanish (ES) PV - Home Page";
		mainURLs[31] = "http://www.msdmanuals.com/es/professional/trastornos-cardiovasculares/pericarditis/pericarditis";
		Versions[31] = "VERSION: Spanish (ES) PV - Topic Page";
		mainURLs[32] = "http://www.msdmanuals.com/es/professional/pages-with-widgets/illustraciones?mode=list";
		Versions[32] = "VERSION: Spanish (ES) PV - Figures Page";
		mainURLs[33] = "http://www.msdmanuals.com/es/professional/resourcespages/about-the-manuals";
		Versions[33] = "VERSION: Spanish (ES) PV - About Page";
		mainURLs[34] = "http://www.msdmanuals.com/es/professional/resourcespages/disclaimer";
		Versions[34] = "VERSION: Spanish (ES) PV - Disclaimer Page";
		mainURLs[35] = "http://www.msdmanuals.com/es/professional/content/permissions";
		Versions[35] = "VERSION: Spanish (ES) PV - Permissions Page";
		mainURLs[36] = "http://www.msdmanuals.com/es/professional/resourcespages/privacy";
		Versions[36] = "VERSION: Spanish (ES) PV - Privacy Page";
		mainURLs[37] = "http://www.msdmanuals.com/es/professional/authors";
		Versions[37] = "VERSION: Spanish (ES) PV - Contributors Page";
		mainURLs[38] = "http://www.msdmanuals.com/es/professional/content/termsofuse";
		Versions[38] = "VERSION: Spanish (ES) PV - Terms Of Use Page";
		mainURLs[39] = "http://www.msdmanuals.com/es/professional/content/licensing";
		Versions[39] = "VERSION: Spanish (ES) PV - Licensing Page";
		mainURLs[40] = "http://www.msdmanuals.com/es/professional/content/contact-us";
		Versions[40] = "VERSION: Spanish (ES) PV - Contact Us Page";
		mainURLs[41] = "http://www.msdmanuals.com/es/professional/resourcespages/global-medical-knowledge-2020";
		Versions[41] = "VERSION: Spanish (ES) PV - Global Medical Knowledge Page";
		mainURLs[42] = "http://www.msdmanuals.com/es/professional/ap%C3%A9ndices/valores-normales-de-laboratorio/valores-normales-de-laboratorio";
		Versions[42] = "VERSION: Spanish (ES) PV - Normal Laboratory Values Page";
		mainURLs[43] = "http://www.msdmanuals.com/es/professional/pages-with-widgets/videos?mode=list";
		Versions[43] = "VERSION: Spanish (ES) PV - Videos Page";

		// Spanish (ES) CV
		mainURLs[44] = "http://www.msdmanuals.com/es/hogar";
		Versions[44] = "VERSION: Spanish (ES) CV - Home Page";
		mainURLs[45] = "http://www.msdmanuals.com/es/hogar/trastornos-de-la-piel/biolog%C3%ADa-de-la-piel/introducci%C3%B3n-a-la-piel";
		Versions[45] = "VERSION: Spanish (ES) CV - Topic Page";
		mainURLs[46] = "http://www.msdmanuals.com/es/hogar/pages-with-widgets/illustraciones?mode=list";
		Versions[46] = "VERSION: Spanish (ES) CV - Figures Page";
		mainURLs[47] = "http://www.msdmanuals.com/es/hogar/resourcespages/acerca-de-los-manuales-msd";
		Versions[47] = "VERSION: Spanish (ES) CV - About Page";
		mainURLs[48] = "http://www.msdmanuals.com/es/hogar/resourcespages/disclaimer";
		Versions[48] = "VERSION: Spanish (ES) CV - Disclaimer Page";
		mainURLs[49] = "http://www.msdmanuals.com/es/hogar/content/permissions";
		Versions[49] = "VERSION: Spanish (ES) CV - Permissions Page";
		mainURLs[50] = "http://www.msdmanuals.com/es/hogar/resourcespages/privacy";
		Versions[50] = "VERSION: Spanish (ES) CV - Privacy Page";
		mainURLs[51] = "http://www.msdmanuals.com/es/hogar/authors";
		Versions[51] = "VERSION: Spanish (ES) CV - Contributors Page";
		mainURLs[52] = "http://www.msdmanuals.com/es/hogar/content/termsofuse";
		Versions[52] = "VERSION: Spanish (ES) CV - Terms Of Use Page";
		mainURLs[53] = "http://www.msdmanuals.com/es/hogar/content/licensing";
		Versions[53] = "VERSION: Spanish (ES) CV - Licensing Page";
		mainURLs[54] = "http://www.msdmanuals.com/es/hogar/content/contact-us";
		Versions[54] = "VERSION: Spanish (ES) CV - Contact Us Page";
		mainURLs[55] = "http://www.msdmanuals.com/es/hogar/resourcespages/global-medical-knowledge-2020";
		Versions[55] = "VERSION: Spanish (ES) CV - Global Medical Knowledge Page";
		mainURLs[56] = "http://www.msdmanuals.com/es/hogar/pages-with-widgets/infographics?mode=list";
		Versions[56] = "VERSION: Spanish (ES) CV - Infographics Page";
		mainURLs[57] = "http://www.msdmanuals.com/es/hogar/recursos/el-manual-de-salud-en-una-sola-p%C3%A1gina/el-manual-de-salud-en-una-sola-p%C3%A1gina";
		Versions[57] = "VERSION: Spanish (ES) CV - The One-Page Manual of Health Page";

		// German (DE) PV
		mainURLs[58] = "http://www.msdmanuals.com/de";
		Versions[58] = "VERSION: German (DE) - Landing Page";
		mainURLs[59] = "http://www.msdmanuals.com/de/profi";
		Versions[59] = "VERSION: German (DE) PV - Home Page";
		mainURLs[60] = "http://www.msdmanuals.com/de/profi/intensivmedizin/behandlung-von-intensivpatienten/einf%C3%BChrung-zur-behandlung-von-intensivpatienten";
		Versions[60] = "VERSION: German (DE) PV - Topic Page";
		mainURLs[61] = "http://www.msdmanuals.com/de/profi/pages-with-widgets/figures?mode=list";
		Versions[61] = "VERSION: German (DE) PV - Figures Page";
		mainURLs[62] = "http://www.msdmanuals.com/de/profi/resourcespages/about-the-manuals";
		Versions[62] = "VERSION: German (DE) PV - About Page";
		mainURLs[63] = "http://www.msdmanuals.com/de/profi/resourcespages/disclaimer";
		Versions[63] = "VERSION: German (DE) PV - Disclaimer Page";
		mainURLs[64] = "http://www.msdmanuals.com/de/profi/content/permissions";
		Versions[64] = "VERSION: German (DE) PV - Permissions Page";
		mainURLs[65] = "http://www.msdmanuals.com/de/profi/resourcespages/privacy";
		Versions[65] = "VERSION: German (DE) PV - Privacy Page";
		mainURLs[66] = "http://www.msdmanuals.com/de/profi/authors";
		Versions[66] = "VERSION: German (DE) PV - Contributors Page";
		mainURLs[67] = "http://www.msdmanuals.com/de/profi/content/termsofuse";
		Versions[67] = "VERSION: German (DE) PV - Terms Of Use Page";
		mainURLs[68] = "http://www.msdmanuals.com/de/profi/content/licensing";
		Versions[68] = "VERSION: German (DE) PV - Licensing Page";
		mainURLs[69] = "http://www.msdmanuals.com/de/profi/content/contact-us";
		Versions[69] = "VERSION: German (DE) PV - Contact Us Page";
		mainURLs[70] = "http://www.msdmanuals.com/de/profi/resourcespages/global-medical-knowledge-2020";
		Versions[70] = "VERSION: German (DE) PV - Global Medical Knowledge Page";
		mainURLs[71] = "http://www.msdmanuals.com/de/profi/anh%C3%A4nge/normalwerte-in-der-labordiagnostik/normalwerte-in-der-labordiagnostik";
		Versions[71] = "VERSION: German (DE) PV - Normal Laboratory Values Page";
		mainURLs[72] = "http://www.msdmanuals.com/de/profi/pages-with-widgets/videos?mode=list";
		Versions[72] = "VERSION: German (DE) PV - Videos Page";

		// German (DE) CV
		mainURLs[73] = "http://www.msdmanuals.com/de/heim";
		Versions[73] = "VERSION: German (DE) CV - Home Page";
		mainURLs[74] = "http://www.msdmanuals.com/de/heim/hauterkrankungen/biologie-der-haut/die-haut-%E2%80%93-ein-%C3%BCberblick";
		Versions[74] = "VERSION: German (DE) CV - Topic Page";
		mainURLs[75] = "http://www.msdmanuals.com/de/heim/pages-with-widgets/figures?mode=list";
		Versions[75] = "VERSION: German (DE) CV - Figures Page";
		mainURLs[76] = "http://www.msdmanuals.com/de/heim/resourcespages/about-the-manuals";
		Versions[76] = "VERSION: German (DE) CV - About Page";
		mainURLs[77] = "http://www.msdmanuals.com/de/heim/resourcespages/disclaimer";
		Versions[77] = "VERSION: German (DE) CV - Disclaimer Page";
		mainURLs[78] = "http://www.msdmanuals.com/de/heim/content/permissions";
		Versions[78] = "VERSION: German (DE) CV - Permissions Page";
		mainURLs[79] = "http://www.msdmanuals.com/de/heim/resourcespages/privacy";
		Versions[79] = "VERSION: German (DE) CV - Privacy Page";
		mainURLs[80] = "http://www.msdmanuals.com/de/heim/authors";
		Versions[80] = "VERSION: German (DE) CV - Contributors Page";
		mainURLs[81] = "http://www.msdmanuals.com/de/heim/content/termsofuse";
		Versions[81] = "VERSION: German (DE) CV - Terms Of Use Page";
		mainURLs[82] = "http://www.msdmanuals.com/de/heim/content/licensing";
		Versions[82] = "VERSION: German (DE) CV - Licensing Page";
		mainURLs[83] = "http://www.msdmanuals.com/de/heim/content/contact-us";
		Versions[83] = "VERSION: German (DE) CV - Contact Us Page";
		mainURLs[84] = "http://www.msdmanuals.com/de/heim/resourcespages/global-medical-knowledge-2020";
		Versions[84] = "VERSION: German (DE) CV - Global Medical Knowledge Page";
		mainURLs[85] = "http://www.msdmanuals.com/de/heim/pages-with-widgets/infographics?mode=list";
		Versions[85] = "VERSION: German (DE) CV - Infographics Page";
		mainURLs[86] = "http://www.msdmanuals.com/de/heim/quellen/das-handbuch-zur-gesundheit-auf-einer-seite/gesundheitsmanual-auf-einer-seite";
		Versions[86] = "VERSION: German (DE) CV - The One-Page Manual of Health Page";

		// French (FR) PV
		mainURLs[87] = "http://www.msdmanuals.com/fr";
		Versions[87] = "VERSION: French (FR) - Landing Page";
		mainURLs[88] = "http://www.msdmanuals.com/fr/professional";
		Versions[88] = "VERSION: French (FR) PV - Home Page";
		mainURLs[89] = "http://www.msdmanuals.com/fr/professional/troubles-cardiovasculaires/p%C3%A9ricardite/p%C3%A9ricardite";
		Versions[89] = "VERSION: French (FR) PV - Topic Page";
		mainURLs[90] = "http://www.msdmanuals.com/fr/professional/pages-with-widgets/figures?mode=list";
		Versions[90] = "VERSION: French (FR) PV - Figures Page";
		mainURLs[91] = "http://www.msdmanuals.com/fr/professional/resourcespages/about-the-manuals";
		Versions[91] = "VERSION: French (FR) PV - About Page";
		mainURLs[92] = "http://www.msdmanuals.com/fr/professional/resourcespages/disclaimer";
		Versions[93] = "VERSION: French (FR) PV - Disclaimer Page";
		mainURLs[94] = "http://www.msdmanuals.com/fr/professional/content/permissions";
		Versions[94] = "VERSION: French (FR) PV - Permissions Page";
		mainURLs[95] = "http://www.msdmanuals.com/fr/professional/resourcespages/privacy";
		Versions[95] = "VERSION: French (FR) PV - Privacy Page";
		mainURLs[96] = "http://www.msdmanuals.com/fr/professional/authors";
		Versions[96] = "VERSION: French (FR) PV - Contributors Page";
		mainURLs[97] = "http://www.msdmanuals.com/fr/professional/content/termsofuse";
		Versions[97] = "VERSION: French (FR) PV - Terms Of Use Page";
		mainURLs[98] = "http://www.msdmanuals.com/fr/professional/content/licensing";
		Versions[98] = "VERSION: French (FR) PV - Licensing Page";
		mainURLs[99] = "http://www.msdmanuals.com/fr/professional/content/contact-us";
		Versions[99] = "VERSION: French (FR) PV - Contact Us Page";
		mainURLs[100] = "http://www.msdmanuals.com/fr/professional/resourcespages/global-medical-knowledge-2020";
		Versions[100] = "VERSION: French (FR) PV - Global Medical Knowledge Page";
		mainURLs[101] = "http://www.msdmanuals.com/fr/professional/annexes/valeurs-biologiques-normales/valeurs-biologiques-normales";
		Versions[101] = "VERSION: French (FR) PV - Normal Laboratory Values Page";
		mainURLs[102] = "http://www.msdmanuals.com/fr/professional/pages-with-widgets/videos?mode=list";
		Versions[102] = "VERSION: French (FR) PV - Videos Page";

		// French (FR) CV
		mainURLs[103] = "http://www.msdmanuals.com/fr/accueil";
		Versions[103] = "VERSION: French (FR) CV - Home Page";
		mainURLs[104] = "http://www.msdmanuals.com/fr/accueil/troubles-cutan%C3%A9s/biologie-de-la-peau/pr%C3%A9sentation-de-la-peau";
		Versions[104] = "VERSION: French (FR) CV - Topic Page";
		mainURLs[105] = "http://www.msdmanuals.com/fr/accueil/pages-with-widgets/figures?mode=list";
		Versions[105] = "VERSION: French (FR) CV - Figures Page";
		mainURLs[106] = "http://www.msdmanuals.com/fr/accueil/resourcespages/about-the-manuals";
		Versions[106] = "VERSION: French (FR) CV - About Page";
		mainURLs[107] = "http://www.msdmanuals.com/fr/accueil/resourcespages/disclaimer";
		Versions[107] = "VERSION: French (FR) CV - Disclaimer Page";
		mainURLs[108] = "http://www.msdmanuals.com/fr/accueil/content/permissions";
		Versions[108] = "VERSION: French (FR) CV - Permissions Page";
		mainURLs[109] = "http://www.msdmanuals.com/fr/accueil/resourcespages/privacy";
		Versions[109] = "VERSION: French (FR) CV - Privacy Page";
		mainURLs[110] = "http://www.msdmanuals.com/fr/accueil/authors";
		Versions[110] = "VERSION: French (FR) CV - Contributors Page";
		mainURLs[111] = "http://www.msdmanuals.com/fr/accueil/content/termsofuse";
		Versions[111] = "VERSION: French (FR) CV - Terms Of Use Page";
		mainURLs[112] = "http://www.msdmanuals.com/fr/accueil/content/licensing";
		Versions[112] = "VERSION: French (FR) CV - Licensing Page";
		mainURLs[113] = "http://www.msdmanuals.com/fr/accueil/content/contact-us";
		Versions[113] = "VERSION: French (FR) CV - Contact Us Page";
		mainURLs[114] = "http://www.msdmanuals.com/fr/accueil/resourcespages/global-medical-knowledge-2020";
		Versions[114] = "VERSION: French (FR) CV - Global Medical Knowledge Page";
		mainURLs[115] = "http://www.msdmanuals.com/fr/accueil/pages-with-widgets/infographics?mode=list";
		Versions[115] = "VERSION: French (FR) CV - Infographics Page";
		mainURLs[116] = "http://www.msdmanuals.com/fr/accueil/ressources/le-manuel-de-la-sant%C3%A9-en-une-page/le-manuel-de-la-sant%C3%A9-en-une-page";
		Versions[116] = "VERSION: French (FR) CV - The One-Page Manual of Health Page";

		// Italian (IT) PV
		mainURLs[117] = "http://www.msdmanuals.com/it";
		Versions[117] = "VERSION: Italian (IT) - Landing Page";
		mainURLs[118] = "http://www.msdmanuals.com/it/professionale";
		Versions[118] = "VERSION: Italian (IT) PV - Home Page";
		mainURLs[119] = "http://www.msdmanuals.com/it/professionale/medicina-di-terapia-intensiva/approccio-al-paziente-critico/introduzione-all-approccio-del-paziente-critico";
		Versions[119] = "VERSION: Italian (IT) PV - Topic Page";
		mainURLs[120] = "http://www.msdmanuals.com/it/professionale/pages-with-widgets/figures?mode=list";
		Versions[120] = "VERSION: Italian (IT) PV - Figures Page";
		mainURLs[121] = "http://www.msdmanuals.com/it/professionale/resourcespages/about-the-manuals";
		Versions[121] = "VERSION: Italian (IT) PV - About Page";
		mainURLs[122] = "http://www.msdmanuals.com/it/professionale/resourcespages/disclaimer";
		Versions[122] = "VERSION: Italian (IT) PV - Disclaimer Page";
		mainURLs[123] = "http://www.msdmanuals.com/it/professionale/content/permissions";
		Versions[123] = "VERSION: Italian (IT) PV - Permissions Page";
		mainURLs[124] = "http://www.msdmanuals.com/it/professionale/resourcespages/privacy";
		Versions[124] = "VERSION: Italian (IT) PV - Privacy Page";
		mainURLs[125] = "http://www.msdmanuals.com/it/professionale/authors";
		Versions[125] = "VERSION: Italian (IT) PV - Contributors Page";
		mainURLs[126] = "http://www.msdmanuals.com/it/professionale/content/termsofuse";
		Versions[126] = "VERSION: Italian (IT) PV - Terms Of Use Page";
		mainURLs[127] = "http://www.msdmanuals.com/it/professionale/content/licensing";
		Versions[127] = "VERSION: Italian (IT) PV - Licensing Page";
		mainURLs[128] = "http://www.msdmanuals.com/it/professionale/content/contact-us";
		Versions[128] = "VERSION: Italian (IT) PV - Contact Us Page";
		mainURLs[129] = "http://www.msdmanuals.com/it/professionale/resourcespages/global-medical-knowledge-2020";
		Versions[129] = "VERSION: Italian (IT) PV - Global Medical Knowledge Page";
		mainURLs[130] = "http://www.msdmanuals.com/it/professionale/appendici/valori-biologici-normali/valori-biologici-normali";
		Versions[130] = "VERSION: Italian (IT) PV - Normal Laboratory Values Page";
		mainURLs[131] = "http://www.msdmanuals.com/it/professionale/pages-with-widgets/videos?mode=list";
		Versions[131] = "VERSION: Italian (IT) PV - Videos Page";

		// Italian (IT) CV
		mainURLs[132] = "http://www.msdmanuals.com/it/casa";
		Versions[132] = "VERSION: Italian (IT) CV - Home Page";
		mainURLs[133] = "http://www.msdmanuals.com/it/casa/patologie-della-cute/biologia-della-pelle/panoramica-sulla-pelle";
		Versions[133] = "VERSION: Italian (IT) CV - Topic Page";
		mainURLs[134] = "http://www.msdmanuals.com/it/casa/pages-with-widgets/figures?mode=list";
		Versions[134] = "VERSION: Italian (IT) CV - Figures Page";
		mainURLs[135] = "http://www.msdmanuals.com/it/casa/resourcespages/about-the-manuals";
		Versions[135] = "VERSION: Italian (IT) CV - About Page";
		mainURLs[136] = "http://www.msdmanuals.com/it/casa/resourcespages/disclaimer";
		Versions[136] = "VERSION: Italian (IT) CV - Disclaimer Page";
		mainURLs[137] = "http://www.msdmanuals.com/it/casa/content/permissions";
		Versions[137] = "VERSION: Italian (IT) CV - Permissions Page";
		mainURLs[138] = "http://www.msdmanuals.com/it/casa/resourcespages/privacy";
		Versions[138] = "VERSION: Italian (IT) CV - Privacy Page";
		mainURLs[139] = "http://www.msdmanuals.com/it/casa/authors";
		Versions[139] = "VERSION: Italian (IT) CV - Contributors Page";
		mainURLs[140] = "http://www.msdmanuals.com/it/casa/content/termsofuse";
		Versions[140] = "VERSION: Italian (IT) CV - Terms Of Use Page";
		mainURLs[141] = "http://www.msdmanuals.com/it/casa/content/licensing";
		Versions[141] = "VERSION: Italian (IT) CV - Licensing Page";
		mainURLs[142] = "http://www.msdmanuals.com/it/casa/content/contact-us";
		Versions[142] = "VERSION: Italian (IT) CV - Contact Us Page";
		mainURLs[143] = "http://www.msdmanuals.com/it/casa/resourcespages/global-medical-knowledge-2020";
		Versions[143] = "VERSION: Italian (IT) CV - Global Medical Knowledge Page";
		mainURLs[144] = "http://www.msdmanuals.com/it/casa/pages-with-widgets/infographics?mode=list";
		Versions[144] = "VERSION: Italian (IT) CV - Infographics Page";
		mainURLs[145] = "http://www.msdmanuals.com/it/casa/risorse/il-manuale-della-salute-in-una-pagina/manuale-monopagina-sulla-salute";
		Versions[145] = "VERSION: Italian (IT) CV - The One-Page Manual of Health Page";

		// Japanese (JA) PV
		mainURLs[146] = "http://www.msdmanuals.com/ja";
		Versions[146] = "VERSION: Japanese (JA) - Landing Page";
		mainURLs[147] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB";
		Versions[147] = "VERSION: Japanese (JA) PV - Home Page";
		mainURLs[148] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/21-%E6%95%91%E5%91%BD%E5%8C%BB%E7%99%82/%E9%87%8D%E7%97%87%EF%BC%88critically-ill%EF%BC%89%E6%82%A3%E8%80%85%E3%81%B8%E3%81%AE%E3%82%A2%E3%83%97%E3%83%AD%E3%83%BC%E3%83%81/%E9%87%8D%E7%97%87%EF%BC%88critically-ill%EF%BC%89%E6%82%A3%E8%80%85%E3%81%B8%E3%81%AE%E3%82%A2%E3%83%97%E3%83%AD%E3%83%BC%E3%83%81%E3%81%AB%E9%96%A2%E3%81%99%E3%82%8B%E5%BA%8F%E8%AB%96";
		Versions[148] = "VERSION: Japanese (JA) PV - Topic Page";
		mainURLs[149] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/pages-with-widgets/%E5%9B%B3?mode=list";
		Versions[149] = "VERSION: Japanese (JA) PV - Figures Page";
		mainURLs[150] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/resourcespages/about-the-manuals";
		Versions[150] = "VERSION: Japanese (JA) PV - About Page";
		mainURLs[151] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/resourcespages/disclaimer";
		Versions[151] = "VERSION: Japanese (JA) PV - Disclaimer Page";
		mainURLs[152] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/content/permissions";
		Versions[152] = "VERSION: Japanese (JA) PV - Permissions Page";
		mainURLs[153] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/resourcespages/privacy";
		Versions[153] = "VERSION: Japanese (JA) PV - Privacy Page";
		mainURLs[154] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/authors";
		Versions[154] = "VERSION: Japanese (JA) PV - Contributors Page";
		mainURLs[155] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/content/termsofuse";
		Versions[155] = "VERSION: Japanese (JA) PV - Terms Of Use Page";
		mainURLs[156] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/content/licensing";
		Versions[156] = "VERSION: Japanese (JA) PV - Licensing Page";
		mainURLs[157] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/content/contact-us";
		Versions[157] = "VERSION: Japanese (JA) PV - Contact Us Page";
		mainURLs[158] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/resourcespages/global-medical-knowledge-2020";
		Versions[158] = "VERSION: Japanese (JA) PV - Global Medical Knowledge Page";
		mainURLs[159] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/%E4%BB%98%E9%8C%B2/%E6%AD%A3%E5%B8%B8%E6%A4%9C%E6%9F%BB%E5%80%A4%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/%E6%AD%A3%E5%B8%B8%E6%A4%9C%E6%9F%BB%E5%80%A4%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6";
		Versions[159] = "VERSION: Japanese (JA) PV - Normal Laboratory Values Page";
		mainURLs[160] = "http://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB/pages-with-widgets/%E5%8B%95%E7%94%BB?mode=list&order=bysection";
		Versions[160] = "VERSION: Japanese (JA) PV - Videos Page";

		// Japanese (JA) CV
		mainURLs[161] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0";
		Versions[161] = "VERSION: Japanese (JA) CV - Home Page";
		mainURLs[162] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/17-%E7%9A%AE%E8%86%9A%E3%81%AE%E7%97%85%E6%B0%97/%E7%9A%AE%E8%86%9A%E3%81%AE%E7%94%9F%E7%89%A9%E5%AD%A6/%E7%9A%AE%E8%86%9A%E3%81%AE%E6%A6%82%E8%A6%81";
		Versions[162] = "VERSION: Japanese (JA) CV - Topic Page";
		mainURLs[163] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/pages-with-widgets/%E5%9B%B3?mode=list";
		Versions[163] = "VERSION: Japanese (JA) CV - Figures Page";
		mainURLs[164] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/resourcespages/about-the-manuals";
		Versions[164] = "VERSION: Japanese (JA) CV - About Page";
		mainURLs[165] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/resourcespages/disclaimer";
		Versions[165] = "VERSION: Japanese (JA) CV - Disclaimer Page";
		mainURLs[166] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/content/permissions";
		Versions[166] = "VERSION: Japanese (JA) CV - Permissions Page";
		mainURLs[167] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/resourcespages/privacy";
		Versions[167] = "VERSION: Japanese (JA) CV - Privacy Page";
		mainURLs[168] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/authors";
		Versions[168] = "VERSION: Japanese (JA) CV - Contributors Page";
		mainURLs[169] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/content/termsofuse";
		Versions[169] = "VERSION: Japanese (JA) CV - Terms Of Use Page";
		mainURLs[170] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/content/licensing";
		Versions[170] = "VERSION: Japanese (JA) CV - Licensing Page";
		mainURLs[171] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/content/contact-us";
		Versions[171] = "VERSION: Japanese (JA) CV - Contact Us Page";
		mainURLs[172] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/resourcespages/global-medical-knowledge-2020";
		Versions[172] = "VERSION: Japanese (JA) CV - Global Medical Knowledge Page";
		mainURLs[173] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/pages-with-widgets/%E5%8B%95%E7%94%BB?mode=list";
		Versions[173] = "VERSION: Japanese (JA) CV - Videos Page";
		mainURLs[174] = "http://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0/%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9/%E5%81%A5%E5%BA%B7%E3%81%AA%E7%94%9F%E6%B4%BB%E3%82%92%E9%80%81%E3%82%8B%E3%81%9F%E3%82%81%E3%81%AE%E3%83%AF%E3%83%B3%E3%83%9A%E3%83%BC%E3%82%B8%E3%83%9E%E3%83%8B%E3%83%A5%E3%82%A2%E3%83%AB/%E5%81%A5%E5%BA%B7%E3%81%AA%E7%94%9F%E6%B4%BB%E3%82%92%E9%80%81%E3%82%8B%E3%81%9F%E3%82%81%E3%81%AE%E3%83%AF%E3%83%B3%E3%83%9A%E3%83%BC%E3%82%B8%E3%83%9E%E3%83%8B%E3%83%A5%E3%82%A2%E3%83%AB";
		Versions[174] = "VERSION: Japanese (JA) CV - The One-Page Manual of Health Page";

		// Korean (KO) PV
		mainURLs[175] = "http://www.msdmanuals.com/ko";
		Versions[175] = "VERSION: Korean (KO) - Landing Page";
		mainURLs[176] = "http://www.msdmanuals.com/en-kr/professional";
		Versions[176] = "VERSION: Korean (KO) PV - Home Page";
		mainURLs[177] = "http://www.msdmanuals.com/en-kr/professional/genitourinary-disorders/obstructive-uropathy/obstructive-uropathy";
		Versions[177] = "VERSION: Korean (KO) PV - Topic Page";
		mainURLs[178] = "http://www.msdmanuals.com/en-kr/professional/pages-with-widgets/figures?mode=list";
		Versions[178] = "VERSION: Korean (KO) PV - Figures Page";
		mainURLs[179] = "http://www.msdmanuals.com/en-kr/professional/resourcespages/about-the-msd-manuals";
		Versions[179] = "VERSION: Korean (KO) PV - About Page";
		mainURLs[180] = "http://www.msdmanuals.com/en-kr/professional/resourcespages/disclaimer";
		Versions[180] = "VERSION: Korean (KO) PV - Disclaimer Page";
		mainURLs[181] = "http://www.msdmanuals.com/en-kr/professional/content/permissions";
		Versions[181] = "VERSION: Korean (KO) PV - Permissions Page";
		mainURLs[182] = "http://www.msdmanuals.com/en-kr/professional/resourcespages/privacy";
		Versions[182] = "VERSION: Korean (KO) PV - Privacy Page";
		mainURLs[183] = "http://www.msdmanuals.com/en-kr/professional/authors";
		Versions[183] = "VERSION: Korean (KO) PV - Contributors Page";
		mainURLs[184] = "http://www.msdmanuals.com/en-kr/professional/content/termsofuse";
		Versions[184] = "VERSION: Korean (KO) PV - Terms Of Use Page";
		mainURLs[185] = "http://www.msdmanuals.com/en-kr/professional/content/licensing";
		Versions[185] = "VERSION: Korean (KO) PV - Licensing Page";
		mainURLs[186] = "http://www.msdmanuals.com/en-kr/professional/content/contact-us";
		Versions[186] = "VERSION: Korean (KO) PV - Contact Us Page";
		mainURLs[187] = "http://www.msdmanuals.com/en-kr/professional/resourcespages/global-medical-knowledge-2020";
		Versions[187] = "VERSION: Korean (KO) PV - Global Medical Knowledge Page";
		mainURLs[188] = "http://www.msdmanuals.com/en-kr/professional/appendixes/normal-laboratory-values/normal-laboratory-values";
		Versions[188] = "VERSION: Korean (KO) PV - Normal Laboratory Values Page";
		mainURLs[189] = "http://www.msdmanuals.com/en-kr/professional/pages-with-widgets/videos?mode=list";
		Versions[189] = "VERSION: Korean (KO) PV - Videos Page";

		// Korean (KO) CV
		mainURLs[190] = "http://www.msdmanuals.com/ko/%ED%99%88";
		Versions[190] = "VERSION: Korean (KO) CV - Home Page";
		mainURLs[191] = "http://www.msdmanuals.com/ko/%ED%99%88/%EB%85%B8%EC%9D%B8%EC%9D%98-%EA%B1%B4%EA%B0%95-%EB%AC%B8%EC%A0%9C/%EA%B3%A0%EB%A0%B9%EC%9E%90-%EA%B1%B4%EA%B0%95-%EB%B3%B4%ED%97%98%EC%9D%98-%EC%A0%81%EC%9A%A9/%EB%85%B8%EC%9D%B8-%EA%B1%B4%EA%B0%95-%EB%B3%B4%ED%97%98-%EC%A0%81%EC%9A%A9%EC%9D%98-%EA%B0%9C%EC%9A%94";
		Versions[191] = "VERSION: Korean (KO) CV - Topic Page";
		mainURLs[192] = "http://www.msdmanuals.com/ko/%ED%99%88/pages-with-widgets/%EA%B7%B8%EB%A6%BC?mode=list";
		Versions[192] = "VERSION: Korean (KO) CV - Figures Page";
		mainURLs[193] = "http://www.msdmanuals.com/ko/%ED%99%88/resourcespages/about-the-manuals";
		Versions[193] = "VERSION: Korean (KO) CV - About Page";
		mainURLs[194] = "http://www.msdmanuals.com/ko/%ED%99%88/resourcespages/disclaimer";
		Versions[194] = "VERSION: Korean (KO) CV - Disclaimer Page";
		mainURLs[195] = "http://www.msdmanuals.com/ko/%ED%99%88/content/permissions";
		Versions[195] = "VERSION: Korean (KO) CV - Permissions Page";
		mainURLs[196] = "http://www.msdmanuals.com/ko/홈/resourcespages/privacy";
		Versions[196] = "VERSION: Korean (KO) CV - Privacy Page";
		mainURLs[197] = "http://www.msdmanuals.com/ko/%ED%99%88/authors";
		Versions[197] = "VERSION: Korean (KO) CV - Contributors Page";
		mainURLs[198] = "http://www.msdmanuals.com/ko/홈/content/termsofuse";
		Versions[198] = "VERSION: Korean (KO) CV - Terms Of Use Page";
		mainURLs[199] = "http://www.msdmanuals.com/ko/%ED%99%88/content/licensing";
		Versions[199] = "VERSION: Korean (KO) CV - Licensing Page";
		mainURLs[200] = "http://www.msdmanuals.com/ko/%ED%99%88/content/contact-us";
		Versions[200] = "VERSION: Korean (KO) CV - Contact Us Page";
		mainURLs[201] = "http://www.msdmanuals.com/ko/%ED%99%88/resourcespages/global-medical-knowledge-2020";
		Versions[201] = "VERSION: Korean (KO) CV - Global Medical Knowledge Page";
		mainURLs[202] = "http://www.msdmanuals.com/ko/%ED%99%88/pages-with-widgets/infographics?mode=list";
		Versions[202] = "VERSION: Korean (KO) CV - Infographics Page";
		mainURLs[203] = "http://www.msdmanuals.com/ko/%ED%99%88/%EC%B0%B8%EA%B3%A0-%EC%9E%90%EB%A3%8C/%ED%95%9C-%ED%8E%98%EC%9D%B4%EC%A7%80-%EA%B1%B4%EA%B0%95-%EB%A7%A4%EB%89%B4%EC%96%BC/%ED%95%9C-%ED%8E%98%EC%9D%B4%EC%A7%80-%EA%B1%B4%EA%B0%95-%EB%A7%A4%EB%89%B4%EC%96%BC";
		Versions[203] = "VERSION: Korean (KO) CV - The One-Page Manual of Health Page";

		// Portuguese (PT) PV
		mainURLs[204] = "http://www.msdmanuals.com/pt";
		Versions[204] = "VERSION: Portuguese (PT) - Landing Page";
		mainURLs[205] = "http://www.msdmanuals.com/pt/profissional";
		Versions[205] = "VERSION: Portuguese (PT) PV - Home Page";
		mainURLs[206] = "http://www.msdmanuals.com/pt/profissional/dist%C3%BArbios-odontol%C3%B3gicos/dist%C3%BArbios-periodontais/periodontite";
		Versions[206] = "VERSION: Portuguese (PT) PV - Topic Page";
		mainURLs[207] = "http://www.msdmanuals.com/pt/profissional/pages-with-widgets/figures?mode=list";
		Versions[207] = "VERSION: Portuguese (PT) PV - Figures Page";
		mainURLs[208] = "http://www.msdmanuals.com/pt/profissional/resourcespages/about-the-manuals";
		Versions[208] = "VERSION: Portuguese (PT) PV - About Page";
		mainURLs[209] = "http://www.msdmanuals.com/pt/profissional/resourcespages/disclaimer";
		Versions[209] = "VERSION: Portuguese (PT) PV - Disclaimer Page";
		mainURLs[210] = "http://www.msdmanuals.com/pt/profissional/content/permissions";
		Versions[210] = "VERSION: Portuguese (PT) PV - Permissions Page";
		mainURLs[211] = "http://www.msdmanuals.com/pt/profissional/resourcespages/privacy";
		Versions[211] = "VERSION: Portuguese (PT) PV - Privacy Page";
		mainURLs[212] = "http://www.msdmanuals.com/pt/profissional/authors";
		Versions[212] = "VERSION: Portuguese (PT) PV - Contributors Page";
		mainURLs[213] = "http://www.msdmanuals.com/pt/profissional/content/termsofuse";
		Versions[213] = "VERSION: Portuguese (PT) PV - Terms Of Use Page";
		mainURLs[214] = "http://www.msdmanuals.com/pt/profissional/content/licensing";
		Versions[214] = "VERSION: Portuguese (PT) PV - Licensing Page";
		mainURLs[215] = "http://www.msdmanuals.com/pt/profissional/content/contact-us";
		Versions[215] = "VERSION: Portuguese (PT) PV - Contact Us Page";
		mainURLs[216] = "http://www.msdmanuals.com/pt/profissional/resourcespages/global-medical-knowledge-2020";
		Versions[216] = "VERSION: Portuguese (PT) PV - Global Medical Knowledge Page";
		mainURLs[217] = "http://www.msdmanuals.com/pt/profissional/ap%C3%AAndices/valores-laboratoriais-normais/valores-laboratoriais-normais";
		Versions[217] = "VERSION: Portuguese (PT) PV - Normal Laboratory Values Page";
		mainURLs[218] = "http://www.msdmanuals.com/pt/profissional/pages-with-widgets/videos?mode=list";
		Versions[218] = "VERSION: Portuguese (PT) PV - Videos Page";

		// Portuguese (PT) CV
		mainURLs[219] = "http://www.msdmanuals.com/pt/casa";
		Versions[219] = "VERSION: Portuguese (PT) CV - Home Page";
		mainURLs[220] = "http://www.msdmanuals.com/pt/casa/dist%C3%BArbios-hep%C3%A1ticos-e-da-ves%C3%ADcula-biliar/hepatite/vis%C3%A3o-geral-da-hepatite";
		Versions[220] = "VERSION: Portuguese (PT) CV - Topic Page";
		mainURLs[221] = "http://www.msdmanuals.com/pt/casa/pages-with-widgets/figures?mode=list";
		Versions[221] = "VERSION: Portuguese (PT) CV - Figures Page";
		mainURLs[222] = "http://www.msdmanuals.com/pt/casa/resourcespages/about-the-manuals";
		Versions[222] = "VERSION: Portuguese (PT) CV - About Page";
		mainURLs[223] = "http://www.msdmanuals.com/pt/casa/resourcespages/disclaimer";
		Versions[223] = "VERSION: Portuguese (PT) CV - Disclaimer Page";
		mainURLs[224] = "http://www.msdmanuals.com/pt/casa/content/permissions";
		Versions[224] = "VERSION: Portuguese (PT) CV - Permissions Page";
		mainURLs[225] = "http://www.msdmanuals.com/pt/casa/resourcespages/privacy";
		Versions[225] = "VERSION: Portuguese (PT) CV - Privacy Page";
		mainURLs[226] = "http://www.msdmanuals.com/pt/casa/authors";
		Versions[226] = "VERSION: Portuguese (PT) CV - Contributors Page";
		mainURLs[227] = "http://www.msdmanuals.com/pt/casa/content/termsofuse";
		Versions[227] = "VERSION: Portuguese (PT) CV - Terms Of Use Page";
		mainURLs[228] = "http://www.msdmanuals.com/pt/casa/content/licensing";
		Versions[228] = "VERSION: Portuguese (PT) CV - Licensing Page";
		mainURLs[229] = "http://www.msdmanuals.com/pt/casa/content/contact-us";
		Versions[229] = "VERSION: Portuguese (PT) CV - Contact Us Page";
		mainURLs[230] = "http://www.msdmanuals.com/pt/casa/resourcespages/global-medical-knowledge-2020";
		Versions[230] = "VERSION: Portuguese (PT) CV - Global Medical Knowledge Page";
		mainURLs[231] = "http://www.msdmanuals.com/pt/casa/resourcespages/one-page-manual-of-health";
		Versions[231] = "VERSION: Portuguese (PT) CV - Infographics Page";
		mainURLs[232] = "http://www.msdmanuals.com/pt/casa/pages-with-widgets/infographics?mode=list";
		Versions[232] = "VERSION: Portuguese (PT) CV - The One-Page Manual of Health Page";

		// Russian (RU) PV
		mainURLs[233] = "http://www.msdmanuals.com/ru";
		Versions[233] = "VERSION: Russian (RU) - Landing Page";
		mainURLs[234] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9";
		Versions[234] = "VERSION: Russian (RU) PV - Home Page";
		mainURLs[235] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/%D0%B8%D0%BD%D1%82%D0%B5%D0%BD%D1%81%D0%B8%D0%B2%D0%BD%D0%B0%D1%8F-%D1%82%D0%B5%D1%80%D0%B0%D0%BF%D0%B8%D1%8F-%D0%B8-%D1%80%D0%B5%D0%B0%D0%BD%D0%B8%D0%BC%D0%B0%D1%86%D0%B8%D1%8F/%EF%BB%BF%D0%BF%D0%BE%D0%B4%D1%85%D0%BE%D0%B4-%D0%BA-%D1%82%D1%8F%D0%B6%D0%B5%D0%BB%D0%BE%D0%B1%D0%BE%D0%BB%D1%8C%D0%BD%D1%8B%D0%BC-%D0%BF%D0%B0%D1%86%D0%B8%D0%B5%D0%BD%D1%82%D0%B0%D0%BC/%EF%BB%BF%D0%BF%D0%BE%D0%B4%D1%85%D0%BE%D0%B4-%D0%BA-%D1%82%D1%8F%D0%B6%D0%B5%D0%BB%D0%BE%D0%B1%D0%BE%D0%BB%D1%8C%D0%BD%D1%8B%D0%BC-%D0%BF%D0%B0%D1%86%D0%B8%D0%B5%D0%BD%D1%82%D0%B0%D0%BC-%D0%B2%D0%B2%D0%B5%D0%B4%D0%B5%D0%BD%D0%B8%D0%B5";
		Versions[235] = "VERSION: Russian (RU) PV - Topic Page";
		mainURLs[236] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/pages-with-widgets/figures?mode=list";
		Versions[236] = "VERSION: Russian (RU) PV - Figures Page";
		mainURLs[237] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/resourcespages/about-the-manuals";
		Versions[237] = "VERSION: Russian (RU) PV - About Page";
		mainURLs[238] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/resourcespages/disclaimer";
		Versions[238] = "VERSION: Russian (RU) PV - Disclaimer Page";
		mainURLs[239] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/content/permissions";
		Versions[239] = "VERSION: Russian (RU) PV - Permissions Page";
		mainURLs[240] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/resourcespages/privacy";
		Versions[240] = "VERSION: Russian (RU) PV - Privacy Page";
		mainURLs[241] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/authors";
		Versions[241] = "VERSION: Russian (RU) PV - Contributors Page";
		mainURLs[242] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/content/termsofuse";
		Versions[242] = "VERSION: Russian (RU) PV - Terms Of Use Page";
		mainURLs[243] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/content/licensing";
		Versions[243] = "VERSION: Russian (RU) PV - Licensing Page";
		mainURLs[244] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/content/contact-us";
		Versions[244] = "VERSION: Russian (RU) PV - Contact Us Page";
		mainURLs[245] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/resourcespages/global-medical-knowledge-2020";
		Versions[245] = "VERSION: Russian (RU) PV - Global Medical Knowledge Page";
		mainURLs[246] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/%D0%BF%D1%80%D0%B8%D0%BB%D0%BE%D0%B6%D0%B5%D0%BD%D0%B8%D1%8F/%D0%BD%D0%BE%D1%80%D0%BC%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B5-%D0%BB%D0%B0%D0%B1%D0%BE%D1%80%D0%B0%D1%82%D0%BE%D1%80%D0%BD%D1%8B%D0%B5-%D0%BF%D0%BE%D0%BA%D0%B0%D0%B7%D0%B0%D1%82%D0%B5%D0%BB%D0%B8/%D0%BD%D0%BE%D1%80%D0%BC%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B5-%D0%BB%D0%B0%D0%B1%D0%BE%D1%80%D0%B0%D1%82%D0%BE%D1%80%D0%BD%D1%8B%D0%B5-%D0%BF%D0%BE%D0%BA%D0%B0%D0%B7%D0%B0%D1%82%D0%B5%D0%BB%D0%B8";
		Versions[246] = "VERSION: Russian (RU) PV - Normal Laboratory Values Page";
		mainURLs[247] = "http://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9/pages-with-widgets/videos?mode=list";
		Versions[247] = "VERSION: Russian (RU) PV - Videos Page";

		// Russian (RU) CV
		mainURLs[248] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0";
		Versions[248] = "VERSION: Russian (RU) CV - Home Page";
		mainURLs[249] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/%D0%BF%D1%80%D0%BE%D0%B1%D0%BB%D0%B5%D0%BC%D1%8B-%D1%81%D0%BE-%D0%B7%D0%B4%D0%BE%D1%80%D0%BE%D0%B2%D1%8C%D0%B5%D0%BC-%D1%83-%D0%BF%D0%BE%D0%B6%D0%B8%D0%BB%D1%8B%D1%85-%D0%BB%D1%8E%D0%B4%D0%B5%D0%B9/%D0%BC%D0%B5%D0%B4%D0%B8%D1%86%D0%B8%D0%BD%D1%81%D0%BA%D0%BE%D0%B5-%D1%81%D1%82%D1%80%D0%B0%D1%85%D0%BE%D0%B2%D0%BE%D0%B5-%D0%BF%D0%BE%D0%BA%D1%80%D1%8B%D1%82%D0%B8%D0%B5-%D0%B4%D0%BB%D1%8F-%D0%BF%D0%BE%D0%B6%D0%B8%D0%BB%D1%8B%D1%85-%D0%BB%D1%8E%D0%B4%D0%B5%D0%B9/%D0%BE%D0%B1%D0%B7%D0%BE%D1%80-%D0%BC%D0%B5%D0%B4%D0%B8%D1%86%D0%B8%D0%BD%D1%81%D0%BA%D0%BE%D0%B3%D0%BE-%D1%81%D1%82%D1%80%D0%B0%D1%85%D0%BE%D0%B2%D0%BE%D0%B3%D0%BE-%D0%BF%D0%BE%D0%BA%D1%80%D1%8B%D1%82%D0%B8%D1%8F-%D0%B4%D0%BB%D1%8F-%D0%BF%D0%BE%D0%B6%D0%B8%D0%BB%D1%8B%D1%85-%D0%BB%D1%8E%D0%B4%D0%B5%D0%B9";
		Versions[249] = "VERSION: Russian (RU) CV - Topic Page";
		mainURLs[250] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/pages-with-widgets/figures?mode=list";
		Versions[250] = "VERSION: Russian (RU) CV - Figures Page";
		mainURLs[251] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/resourcespages/about-the-manuals";
		Versions[251] = "VERSION: Russian (RU) CV - About Page";
		mainURLs[252] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/resourcespages/disclaimer";
		Versions[252] = "VERSION: Russian (RU) CV - Disclaimer Page";
		mainURLs[253] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/content/permissions";
		Versions[253] = "VERSION: Russian (RU) CV - Permissions Page";
		mainURLs[254] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/resourcespages/privacy";
		Versions[254] = "VERSION: Russian (RU) CV - Privacy Page";
		mainURLs[255] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/authors";
		Versions[255] = "VERSION: Russian (RU) CV - Contributors Page";
		mainURLs[256] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/content/termsofuse";
		Versions[256] = "VERSION: Russian (RU) CV - Terms Of Use Page";
		mainURLs[257] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/content/licensing";
		Versions[257] = "VERSION: Russian (RU) CV - Licensing Page";
		mainURLs[258] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/content/contact-us";
		Versions[258] = "VERSION: Russian (RU) CV - Contact Us Page";
		mainURLs[259] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/resourcespages/global-medical-knowledge-2020";
		Versions[259] = "VERSION: Russian (RU) CV - Global Medical Knowledge Page";
		mainURLs[260] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/pages-with-widgets/infographics?mode=list";
		Versions[260] = "VERSION: Russian (RU) CV - Infographics Page";
		mainURLs[261] = "http://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0/%D1%80%D0%B5%D1%81%D1%83%D1%80%D1%81%D1%8B/%D0%BE%D0%B4%D0%BD%D0%BE%D1%81%D1%82%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%BD%D0%BE%D0%B5-%D0%BC%D0%B5%D0%B4%D0%B8%D1%86%D0%B8%D0%BD%D1%81%D0%BA%D0%BE%D0%B5-%D1%80%D1%83%D0%BA%D0%BE%D0%B2%D0%BE%D0%B4%D1%81%D1%82%D0%B2%D0%BE/%D0%BE%D0%B4%D0%BD%D0%BE%D1%81%D1%82%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%BD%D1%8B%D0%B9-%D0%BC%D0%B5%D0%B4%D0%B8%D1%86%D0%B8%D0%BD%D1%81%D0%BA%D0%B8%D0%B9-%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BE%D1%87%D0%BD%D0%B8%D0%BA";
		Versions[261] = "VERSION: Russian (RU) CV - The One-Page Manual of Health Page";

		// Chinese (ZH) PV
		mainURLs[262] = "http://www.msdmanuals.com/zh";
		Versions[262] = "VERSION: Chinese (ZH) - Landing Page";
		mainURLs[263] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A";
		Versions[263] = "VERSION: Chinese (ZH) PV - Home Page";
		mainURLs[264] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/%E6%80%A5%E6%95%91%E5%8C%BB%E5%AD%A6/%E5%8D%B1%E9%87%8D%E7%97%85%E4%BA%BA%E7%9A%84%E5%A4%84%E7%BD%AE/%E5%BC%95%E8%A8%80";
		Versions[264] = "VERSION: Chinese (ZH) PV - Topic Page";
		mainURLs[265] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/pages-with-widgets/figures?mode=list";
		Versions[265] = "VERSION: Chinese (ZH) PV - Figures Page";
		mainURLs[266] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/resourcespages/about-the-manuals";
		Versions[266] = "VERSION: Chinese (ZH) PV - About Page";
		mainURLs[267] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/resourcespages/disclaimer";
		Versions[267] = "VERSION: Chinese (ZH) PV - Disclaimer Page";
		mainURLs[268] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/content/permissions";
		Versions[268] = "VERSION: Chinese (ZH) PV - Permissions Page";
		mainURLs[269] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/resourcespages/privacy";
		Versions[269] = "VERSION: Chinese (ZH) PV - Privacy Page";
		mainURLs[270] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/authors";
		Versions[270] = "VERSION: Chinese (ZH) PV - Contributors Page";
		mainURLs[271] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/content/termsofuse";
		Versions[271] = "VERSION: Chinese (ZH) PV - Terms Of Use Page";
		mainURLs[272] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/content/licensing";
		Versions[272] = "VERSION: Chinese (ZH) PV - Licensing Page";
		mainURLs[273] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/content/contact-us";
		Versions[273] = "VERSION: Chinese (ZH) PV - Contact Us Page";
		mainURLs[274] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/resourcespages/global-medical-knowledge-2020";
		Versions[274] = "VERSION: Chinese (ZH) PV - Global Medical Knowledge Page";
		mainURLs[275] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/%E9%99%84%E5%BD%95/%E5%AE%9E%E9%AA%8C%E5%AE%A4%E6%AD%A3%E5%B8%B8%E5%80%BC/%E5%AE%9E%E9%AA%8C%E5%AE%A4%E6%AD%A3%E5%B8%B8%E5%80%BC";
		Versions[275] = "VERSION: Chinese (ZH) PV - Normal Laboratory Values Page";
		mainURLs[276] = "http://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A/pages-with-widgets/videos?mode=list";
		Versions[276] = "VERSION: Chinese (ZH) PV - Videos Page";

		// Chinese (ZH) CV
		mainURLs[277] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5";
		Versions[277] = "VERSION: Chinese (ZH) CV - Home Page";
		mainURLs[278] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/%E8%80%81%E5%B9%B4%E4%BA%BA%E7%9A%84%E5%81%A5%E5%BA%B7%E9%97%AE%E9%A2%98/%E8%80%81%E5%B9%B4%E4%BA%BA%E7%9A%84%E5%8C%BB%E4%BF%9D%E6%94%BF%E7%AD%96/%E8%80%81%E5%B9%B4%E4%BA%BA%E5%8C%BB%E7%96%97%E4%BF%9D%E9%99%A9%E6%A6%82%E8%BF%B0";
		Versions[278] = "VERSION: Chinese (ZH) CV - Topic Page";
		mainURLs[279] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/pages-with-widgets/figures?mode=list";
		Versions[279] = "VERSION: Chinese (ZH) CV - Figures Page";
		mainURLs[280] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/resourcespages/about-the-manuals";
		Versions[280] = "VERSION: Chinese (ZH) CV - About Page";
		mainURLs[281] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/resourcespages/disclaimer";
		Versions[281] = "VERSION: Chinese (ZH) CV - Disclaimer Page";
		mainURLs[282] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/content/permissions";
		Versions[282] = "VERSION: Chinese (ZH) CV - Permissions Page";
		mainURLs[283] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/resourcespages/privacy";
		Versions[283] = "VERSION: Chinese (ZH) CV - Privacy Page";
		mainURLs[284] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/authors";
		Versions[284] = "VERSION: Chinese (ZH) CV - Contributors Page";
		mainURLs[285] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/content/termsofuse";
		Versions[285] = "VERSION: Chinese (ZH) CV - Terms Of Use Page";
		mainURLs[286] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/content/licensing";
		Versions[286] = "VERSION: Chinese (ZH) CV - Licensing Page";
		mainURLs[287] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/content/contact-us";
		Versions[287] = "VERSION: Chinese (ZH) CV - Contact Us Page";
		mainURLs[288] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/resourcespages/global-medical-knowledge-2020";
		Versions[288] = "VERSION: Chinese (ZH) CV - Global Medical Knowledge Page";
		mainURLs[289] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/pages-with-widgets/infographics?mode=list";
		Versions[289] = "VERSION: Chinese (ZH) CV - Infographics Page";
		mainURLs[290] = "http://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5/%E8%B5%84%E6%BA%90/%E5%8D%95%E9%A1%B5%E9%BB%98%E5%85%8B%E5%81%A5%E5%BA%B7%E6%89%8B%E5%86%8C/%E5%8D%95%E9%A1%B5";
		Versions[290] = "VERSION: Chinese (ZH) CV - The One-Page Manual of Health Page";

		// Vet
		mainURLs[291] = "http://www.msdvetmanual.com";
		Versions[291] = "VERSION: MSD Vet Manual - Home Page";
		mainURLs[292] = "http://www.msdvetmanual.com/immune-system/immunologic-diseases/overview-of-immunologic-diseases";
		Versions[292] = "VERSION: MSD Vet Manual - Topic Page";
		mainURLs[293] = "http://www.msdvetmanual.com/pages-with-widgets/figures?mode=list";
		Versions[292] = "VERSION: MSD Vet Manual - Figures Page";
		mainURLs[294] = "http://www.msdvetmanual.com/resourcespages/about";
		Versions[294] = "VERSION: MSD Vet Manual - Foreword Page";
		mainURLs[295] = "http://www.msdvetmanual.com/resourcespages/disclaimer";
		Versions[295] = "VERSION: MSD Vet Manual - Disclaimer Page";
		mainURLs[296] = "http://www.msdvetmanual.com/content/permissions";
		Versions[296] = "VERSION: MSD Vet Manual - Permissions Page";
		mainURLs[297] = "http://www.msdvetmanual.com/resourcespages/partners";
		Versions[297] = "VERSION: MSD Vet Manual - Partners Page";
		mainURLs[298] = "http://www.msdvetmanual.com/resourcespages/content-providers";
		Versions[298] = "VERSION: MSD Vet Manual - Content Providers Page";
		mainURLs[299] = "http://www.msdvetmanual.com/content/termsofuse";
		Versions[299] = "VERSION: MSD Vet Manual - Terms Of Use Page";
		mainURLs[300] = "http://www.msdvetmanual.com/content/licensing";
		Versions[300] = "VERSION: MSD Vet Manual - Licensing Page";
		mainURLs[301] = "http://www.msdvetmanual.com/resourcespages/contactus";
		Versions[301] = "VERSION: MSD Vet Manual - Contact Us Page";
		mainURLs[302] = "http://www.msdvetmanual.com/resourcespages/contributors";
		Versions[302] = "VERSION: MSD Vet Manual - Contributors Page";
		mainURLs[303] = "http://www.msdvetmanual.com/resourcespages/history";
		Versions[303] = "VERSION: MSD Vet Manual - History Page";
		mainURLs[304] = "http://www.msdvetmanual.com/resourcespages/user-guide";
		Versions[304] = "VERSION: MSD Vet Manual - MVM User Guide Page";
		mainURLs[305] = "http://www.msdvetmanual.com/resourcespages/staff";
		Versions[305] = "VERSION: MSD Vet Manual - Production and Publishing Staff Page";
		mainURLs[306] = "http://www.msdvetmanual.com/resourcespages/editorial-board";
		Versions[306] = "VERSION: MSD Vet Manual - Editorial Board Page";

		// Main Loop
		for (int j = 0; j < 307; j++) {

			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				String CurrentVersion = Versions[j];
				// Handle modal windows
				if (CurrentVersion.contains("Russian CV")) {
					wd.findElement(By.xpath("html/body/div[14]/div[2]/div/a[2]")).click();

				} else if (CurrentVersion.contains("Chinese PV")) {
					wd.findElement(By.xpath("html/body/div[14]/div[2]/div/a[2]")).click();
				}
				System.out.println(Versions[j]);


				// Get all HTML Page Source code into String Variable
			      // create file named Cookies to store Login Information		
		        File file = new File("Cookies.data");							
		        try		
		        {	  
		            // Delete old file if exists
					file.delete();		
		            file.createNewFile();			
		            FileWriter fileWrite = new FileWriter(file);							
		            BufferedWriter Bwrite = new BufferedWriter(fileWrite);							
		            // loop for getting the cookie information 		
		            	
		            // loop for getting the cookie information 		
		            for(Cookie ck : wd.manage().getCookies())							
		            {			
		                Bwrite.write((ck.getName()+";"+ck.getValue()+";"+ck.getDomain()+";"+ck.getPath()+";"+ck.getExpiry()+";"+ck.isSecure()));																									
		                Bwrite.newLine();
		                System.out.println(ck.getName());
		                System.out.println(ck.getValue());
		                System.out.println(ck.getDomain());
		                System.out.println(ck.getPath());
		                System.out.println(ck.getExpiry());
		                System.out.println(ck.isSecure());
		            }			
		            Bwrite.close();			
		            fileWrite.close();	
		            
		        }
		        catch(Exception ex)					
		        {		
		            ex.printStackTrace();			
		        }

				
			} catch (Exception e) {
				System.out.println("Page Error!");
			}
		}
		// Close excel and browser
		wd.close();
		System.out.println("Browser Closed!");
	}

}
