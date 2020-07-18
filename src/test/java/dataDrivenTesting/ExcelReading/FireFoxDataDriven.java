package dataDrivenTesting.ExcelReading;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import junit.framework.Assert;
import utilities.ExcelUtils;

public class FireFoxDataDriven {

	Logger logger;

	WebDriver driver;

	@BeforeClass
	void setUp() {

		logger = Logger.getLogger(FireFoxDataDriven.class);
		PropertyConfigurator.configure("log4j.properties");
		logger.setLevel(Level.DEBUG);

	}

	@Test(dataProvider = "userDetails")
	void readFromExcel(String uName, String uPassword) throws Exception {

		logger.info("..............This is Data driven test Start");

		System.setProperty("webdriver.gecko.driver",
				"C:\\Users\\Owner\\eclipse-workspace\\ExcelReading\\drivers\\geckodriver.exe");

		driver = new FirefoxDriver();

		driver.manage().window().maximize();

		driver.get(
				"https://accounts.google.com/signin/v2/identifier?continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&service=mail&sacu=1&rip=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin");

		System.out.println("INSIDE LOGIN");

		System.out.println("Username is:" + uName + "....." + "Password is : " + uPassword);

		Thread.sleep(1500);

		driver.findElement(By.id("identifierId")).sendKeys(uName);

		driver.findElement(
				By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/span"))
				.click();

		Thread.sleep(1000);

		driver.findElement(By.xpath(
				"/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/form/span/section/div/div/div[1]/div[1]/div/div/div/div/div[1]/div/div[1]/input"))
				.sendKeys(uPassword);
		Thread.sleep(1000);

		driver.findElement(
				By.xpath("/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/span"))
				.click();

		Thread.sleep(1500);

		String title = driver.getTitle().toString();

		logger.info("......................This is Title......................" + title);

		if (title.contains("Gmail")) {

			String filePath = "C:\\Users\\Owner\\eclipse-workspace\\ExcelReading\\src\\email.xlsx";

			int row = ExcelUtils.getRowCount(filePath, "Sheet1");

			for (int i = 1; i <= row; i++) {

				logger.info(".................Inside For loop.....................");

				System.out.println("Inside i Loop");

				ExcelUtils.setCellData(filePath, "sheet1", i, 2, "Sucess");

			}

			driver.findElement(By.xpath("//a[@role='button' and @tabindex='0' and @class='gb_D gb_Ra gb_i']")).click();

			driver.findElement(By.xpath("//a[text()='Sign out']")).click();

			logger.info(".................Signout Complete.....................");

			Thread.sleep(2000);

		}

	}

	@DataProvider(name = "userDetails")
	String[][] dataProvider() throws Exception {

		System.out.println("We are inside Data Provider");
		String filePath = "C:\\Users\\Owner\\eclipse-workspace\\ExcelReading\\src\\test\\java\\utilities\\email.xlsx";

		int row = ExcelUtils.getRowCount(filePath, "Sheet1");
		int col = ExcelUtils.getCellCount(filePath, "Sheet1", 1);

		String[][] userData = new String[row][col];

		for (int i = 1; i <= row; i++) {

			for (int j = 0; j < col; j++) {

				userData[i - 1][j] = ExcelUtils.getCellData(filePath, "Sheet1", i, j);
			}

		}

		System.out.println("The data provided is: " + userData.toString());

		return userData;

	}

	@AfterTest
	void tearDown() throws Exception {

		System.out.println("TEAR DOWN");

		Thread.sleep(1500);

		driver.quit();

	}
}
