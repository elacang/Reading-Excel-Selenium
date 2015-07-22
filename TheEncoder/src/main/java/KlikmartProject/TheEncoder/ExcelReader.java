package KlikmartProject.TheEncoder;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.thoughtworks.selenium.Wait;

public class ExcelReader {
	public static void main(String[] args) throws BiffException, IOException,
			WriteException, InterruptedException {
		/*
		 * WritableWorkbook wworkbook; wworkbook = Workbook.createWorkbook(new
		 * File("output.xls")); WritableSheet wsheet =
		 * wworkbook.createSheet("First Sheet", 0); Label label = new Label(0,
		 * 3, "A label record"); wsheet.addCell(label); Number number = new
		 * Number(3, 4, 3.1459); wsheet.addCell(number); wworkbook.write();
		 * wworkbook.close();
		 */

		Workbook workbook = Workbook.getWorkbook(new File("output.xls"));
		Sheet sheet = workbook.getSheet(0);

		WebDriver driver = new FirefoxDriver();
		driver.get("http://thetestroom.com/webapp/index.html");

		for (int i = 1; i <= sheet.getRows(); i++) {
			Cell col1 = sheet.getCell(0, i);
			//System.out.println(col1.getContents());
			Cell col2 = sheet.getCell(1, i);
			//System.out.println(col2.getContents());
			Cell col3 = sheet.getCell(2, i);
			//System.out.println(col3.getContents());
			Cell col4 = sheet.getCell(3, i);
			//System.out.println(col4.getContents());

			// input data
			
			driver.findElement(By.id("contact_link")).click();
			driver.findElement(By.name("name_field")).sendKeys(
					col1.getContents());
			driver.findElement(By.name("address_field")).sendKeys(
					col2.getContents());
			driver.findElement(By.name("postcode_field")).sendKeys(
					col3.getContents());
			driver.findElement(By.name("email_field")).sendKeys(
					col4.getContents());
			driver.findElement(By.id("submit_message")).click();
		}

		driver.close();
		workbook.close();
	}
}