package UploadDownloadExcelFiles;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class UploadDownload extends ParseExcel{
	@Test
	public void uploadOrDownloadFile() throws IOException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
		Actions actions = new Actions(driver);
		actions.moveToElement(driver.findElement(By.id("downloadButton"))).click().build().perform();
		driver.findElement(By.cssSelector(".upload")).sendKeys("C:\\Users\\yashn\\Downloads\\download.xlsx");
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("(//div[@class=\"Toastify__toast-body\"]/div)[2]")));
		String ActualMessage = driver.findElement(By.xpath("(//div[@class=\"Toastify__toast-body\"]/div)[2]")).getText();
		Assert.assertTrue(ActualMessage.equalsIgnoreCase("Updated Excel Data Successfully."));
		String fruitName = "Papaya";
		int row = getRow("Sheet1", fruitName);
		System.out.println(row);
		int column = getColumnNumber("Sheet1", "price");
		updateSheet("sheet1", row, column,"187");
		String cloumnNumber =  driver.findElement(By.xpath("//div[text()=\"Price\"]")).getDomAttribute("data-column-id");
		int priceOnExcel = (int) getData("Sheet1", fruitName);
		int priceOnWeb = Integer.parseInt(driver.findElement(By.xpath("//div[text()='"+fruitName+"']/parent::div/following-sibling::div[@data-column-id="+cloumnNumber+"]")).getText());
		Assert.assertEquals(priceOnExcel, priceOnWeb);
	}
	

}
