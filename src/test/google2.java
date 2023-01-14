package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class google2 {

	public static void main(String[] args) throws InterruptedException, IOException {
		
		// Set Driver path
		System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		//Read start
		//Path of the excel file
		File src=new File("F:\\test\\test.xlsx");
		FileInputStream fs = new FileInputStream(src);
		//Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		//find day
		LocalDate today = LocalDate.now();
		DayOfWeek dayOfWeek = today.getDayOfWeek();
		String day = dayOfWeek.toString();
		//System.out.println("Day of the Week :: " + day);
		
		XSSFSheet sheet = workbook.getSheet(day);
		int rc=sheet.getLastRowNum();

		for(int i=2;i<rc+1;i++)
		{
			String data0=sheet.getRow(i).getCell(2).getStringCellValue();
			//System.out.println("Input: "+data0);
			//open google
		    driver.get("https://www.google.com");
		    driver.manage().window().maximize();
		    Thread.sleep(2000);

		    //enter techlistic tutorials in search box
		    driver.findElement(By.name("q")).sendKeys(data0);
		    //wait for suggestions
		    driver.findElement(By.name("q")).submit();
		    driver.findElement(By.name("q")).click();

		    
		    //Xpath
	        List<WebElement> option = driver.findElements(By.xpath("//ul[@role='listbox']/li/descendant::div[@class='wM6W7d']"));
	        int n=option.size();
	        //System.out.println("number:"+n);
	        String st= option.get(0).getText();
	        String lt= option.get(0).getText();
	        //Sort start
	        for(int j=0;j<n;j++)
	        {
	        	String a= option.get(j).getText();
	        	if(a.length()<st.length())
	        	{
	        		st=a;
	        	}
	        	if(a.length()>lt.length())
	        	{
	        		lt=a;
	        	}
	        	//System.out.println("option:"+a);
	        }
	        
	        //System.out.println("textst:"+st);
	        //System.out.println("textlt:"+lt);
	        sheet.getRow(i).createCell(3).setCellValue(st);
	        sheet.getRow(i).createCell(4).setCellValue(lt);
	        FileOutputStream fout= new FileOutputStream(src); 
	        workbook.write(fout);

		}
		//Read and sort finish
		
		//excel print[have some problem]
		driver.close();
		

	}

}
