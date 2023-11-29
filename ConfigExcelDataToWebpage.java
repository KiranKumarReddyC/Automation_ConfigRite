package cRite_Automation;



	import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.apache.poi.xssf.usermodel.XSSFSheet;

	import java.io.File;
	import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

	public class ConfigExcelDataToWebpage {
		String pass ="hw*Y^7M3";
		String LKP_TYPE_MNG ="Lookup Type";
		@Test
			public void Test() throws IOException, InterruptedException {
	        // Initialize the WebDriver (Assuming you are using ChromeDriver)
	        WebDriver driver = new ChromeDriver();

	        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
	        
	        // Read data from the Excel file
	        FileInputStream file = new FileInputStream(new File("C:\\Users\\KiranKumarReddyC\\Documents\\Payables-Lookups_Test_Script.xlsx"));
	        Workbook workbook = new XSSFWorkbook(file);
	        Sheet sheet = workbook.getSheet("Sheet1");

	        // Get the headers from the first row (assuming they are in the first row)
	        Row headerRow = sheet.getRow(0);
	       try { 
	        driver.get("https://fa-etao-dev20-saasfademo1.ds-fa.oraclepdemos.com/");
			driver.manage().window().maximize();
			WebElement Sign_in = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@name='btnActive' and @id='btnActive']")));
			driver.findElement(By.xpath("//input[@name='userid' and @id='userid']")).sendKeys("casey.brown");
			driver.findElement(By.xpath("//input[@name='password' and @id='password']")).sendKeys(pass);
			
			driver.findElement(By.xpath("//button[@name='btnActive' and @id='btnActive']")).click();
			
			WebElement HomeIcon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='pt1:_UIShome' and @title='Home']")));
			driver.findElement(By.xpath("//a[@id='pt1:_UIShome' and @title='Home']")).click();
			
			Thread.sleep(5000);
			WebElement NavigationMenu = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='pt1:_UISmmLink'  and @title='Navigator']")));
			driver.findElement(By.xpath("//a[@id='pt1:_UISmmLink'  and @title='Navigator']")).click();

			WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='pt1:_UISnvr:0:nv_pgl3']")));
			Thread.sleep(5000);

			// Locate the scroll bar element using its XPath
			WebElement scrollBar = driver.findElement(By.xpath("//*[@id='pt1:_UISnvr:0:nv_pgl3']"));
            Thread.sleep(5000);
			// Use JavaScript to scroll the element to the bottom
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollTop = arguments[0].scrollHeight;", scrollBar);
			
		//	WebElement NavigationMenu = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='pt1:_UISmmLink'  and @title='Navigator']")));

			//Thread.sleep(20000);
	        
			driver.findElement(By.cssSelector("#pt1\\:_UISnvr\\:0\\:nvgpgl2_nvmOthersCustomGrp")).click();
			
		    WebElement OthersE = wait.until(ExpectedConditions.presenceOfElementLocated(By.partialLinkText("Setup and Maintenance")));

			driver.findElement(By.partialLinkText("Setup and Maintenance")).click();
			
		    WebElement GlobalSearch = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@id='pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content' and @type='text' ]")));

			
//			driver.findElement(By.xpath("//input[@id='pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content' and @type='text' ]")).sendKeys("Manage Set Enabled Lookups");
//			
//			WebElement SearchButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@role='button' and @class='xrg' ]//img[@title='Search']")));
//
//			driver.findElement(By.xpath("//a[@role='button' and @class='xrg' ]//img[@title='Search']")).click();
		    
  		    WebElement TaskBar = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*//a[@id='pt1:r1:0:r0:0:r1:0:AP1:sdi10::disAcr']")));

	        driver.findElement(By.xpath("//*//a[@id='pt1:r1:0:r0:0:r1:0:AP1:sdi10::disAcr']")).click();
		    
  		    WebElement Search = wait.until(ExpectedConditions.elementToBeClickable(By.partialLinkText("Search")));
	        driver.findElement(By.partialLinkText("Search")).click();
	        
	        
  		    WebElement SearchButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='pt1:r1:0:r0:1:AP1:s9:ctb1']//a[@role='button' and @class='xrg' ]")));
  		    driver.findElement(By.xpath("//input[@id='pt1:r1:0:r0:1:AP1:s9:it1::content']")).sendKeys("Manage Set Enabled Lookups");
	        driver.findElement(By.xpath("//div[@id='pt1:r1:0:r0:1:AP1:s9:ctb1']//a[@role='button' and @class='xrg' ]")).click();
	      
  		    WebElement Lookuptype = wait.until(ExpectedConditions.elementToBeClickable(By.partialLinkText("Manage Set Enabled Lookups")));
	        driver.findElement(By.partialLinkText("Manage Set Enabled Lookups")).click();
	        
  		    WebElement New_LOOKUP_TYPE = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@title='New' and @class='xeq p_AFIconOnly']//a[@role='button' and @class='xrg' ]")));
	        driver.findElement(By.xpath("//div[@title='New' and @class='xeq p_AFIconOnly']//a[@role='button' and @class='xrg' ]")).click();
	       
  		    WebElement New_LOOKUP_TYPE_Name = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@title='New' and @class='xeq p_AFIconOnly']//a[@role='button' and @class='xrg' ]")));
	        //driver.findElement(By.xpath("//input[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:it8::content']")).sendKeys("");
	        Thread.sleep(2000);
  		    Actions actions =new Actions(driver);
  		            actions.moveToElement(New_LOOKUP_TYPE_Name);
  		            actions.click();
//  		            actions.moveByOffset(5, 0);
//  		            actions.moveByOffset(-10, 0);
//  		            actions.moveByOffset(10, 0);
//  		            actions.moveByOffset(-5, 0);
//  		            actions.perform();
  		    
	        By L_C_T_N_XPath = By.xpath("//input[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:it8::content']");

	        int rowIndex = 3;
	        int columnIndex=1;
	        Cell LOOKUP_TYPE_cell = sheet.getRow(rowIndex).getCell(columnIndex);
	        String cellData=LOOKUP_TYPE_cell.getStringCellValue();
	        file.close();
	       
	        WebElement LOOKUP_TYPE_textBox = driver.findElement(L_C_T_N_XPath);
	        LOOKUP_TYPE_textBox.sendKeys(cellData);
	        
	        
	     // WebElement Meaning_TextBox
            WebElement L_K_TYPE_Meaning =wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:APscl']")));
	        By MEANING_XPATH =By.xpath("//input[@class='x25' and @id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:it2::content']");
            
	        Cell Module_cell=sheet.getRow(4).getCell(1);
	        String Module_CELL_DATA=Module_cell.getStringCellValue();
	        file.close();
	        
	       
	        WebElement MNG_TXT_BOX=driver.findElement(MEANING_XPATH);
	        MNG_TXT_BOX.sendKeys(Module_CELL_DATA);
	        
	        
	     // WebElement DESCRIPTION
	        
	        By DESCRIPTION = By.xpath("//input[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:it6::content']");
	        WebElement DESCRIPTION_TXT=driver.findElement(DESCRIPTION);
	        DESCRIPTION_TXT.sendKeys(LKP_TYPE_MNG);
	        
	        
		     // WebElement module_dropdown
	           //click
               driver.findElement(By.xpath("//span//a[@title='Search: Module']")).click();
            
            WebElement Module_Search=wait.until(ExpectedConditions.elementToBeClickable(By.partialLinkText("Search...")));
            Module_Search.click();
            
            WebElement Search_TXT_B=wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId::_afrLovInternalQueryId::search']")));
            driver.findElement(By.xpath("//input[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId::_afrLovInternalQueryId:value00::content']")).sendKeys(Module_CELL_DATA);	
            Search_TXT_B.click();
            
            WebElement Click_Module_LINK=wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::db']/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td[1]")));
            //driver.findElement(By.partialLinkText("Payables")).click();
            
            driver.findElement(By.xpath("//*[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::db']/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td[1]")).click();
            
           WebElement ok=wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='xux p_AFTextOnly' and @_afrpdo='ok' and@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId::lovDialogId::ok']")));
           ok.click();
           
           //save button for lookup type
           driver.findElement(By.xpath("//button[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:APsv']")).click();
           driver.findElement(By.xpath("//button[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:APscl']")).click();
           
           driver.findElement(By.xpath("//select[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:soc2::content']")).click();
           
           
            //select relevant module
//            
//            WebElement table = driver.findElement(By.xpath("//div[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId']//table[@class='x1hp' and @id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::ch::d2::t2']"));
//
//            // Find all the rows in the table
//            List<WebElement> rows = table.findElements(By.xpath("//*[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::db']/table/tbody/tr/td/div/table/tbody/tr"));
//
//            
//            // Loop through the rows to find the one with "Payables"
//            for (WebElement row : rows) {
//                // Get the cell(s) within the row
//                List<WebElement> cells_ST = row.findElements(By.xpath("//*[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::db']/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]"));
//
//                
//                // Check if the second cell contains "Payables"
//                for (WebElement cellS_T : cells_ST) {
//                    if (cellS_T.getText().equals("Payables")) {
//                        // Click the row or perform the desired action
//                        row.click();
//                        System.out.println(cellS_T.getText());
//                        break;  // Exit the loop
//                    }
//                }
//            }

            
            
            
//            
//         // Find the table element
//            WebElement table = driver.findElement(By.xpath("//div[@id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId']//table[@class='x1hp' and @id='pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:ATt1:0:userModuleNameId_afrLovInternalTableId::ch::d2::t2']"));
//
//            // Find all the rows in the table
//            List<WebElement> rows = table.findElements(By.xpath(".//tbody/tr"));
//
//            String payablesXPath = "";
//			for (WebElement row : rows) {
//                // Find the 1st column (1st cell) in the current row
//                WebElement firstCell = row.findElement(By.xpath(".//td[1]"));
//
//                // Check if the text in the 1st column is equal to "Payables"
//                if (firstCell.getText().equals("Payables")) {
//                    // Perform your desired action (e.g., click on the row)
//                	 payablesXPath = (String) ((JavascriptExecutor) driver).executeScript("return arguments[0].outerHTML;", row);
//                	row.click();
//
//                    // Optionally, you can break the loop if you only need to find the first occurrence
//                    break;
//                }
//            }
//            System.out.println("XPath of Payables element: " + payablesXPath);

            
            
	        // Iterate through rows, starting from the second row (index 1)
	        for (int rowIndex1 = 1; rowIndex1 <= sheet.getLastRowNum(); rowIndex1++) {
	            Row row1 = sheet.getRow(rowIndex1);

	            // Iterate through cells in the current row
	            for (int cellIndex1 = 0; cellIndex1 < row1.getLastCellNum(); cellIndex1++) {
	                Cell cell1 = row1.getCell(cellIndex1);
	                String header = headerRow.getCell(cellIndex1).getStringCellValue();
	                String cellValue = cell1.getStringCellValue();

	                // Perform actions based on the header and cell value
	                if (header.equals("Action")) {
	                    // Do something with the "Action" column value
	                } else if (header.equals("Lookup Code")) {
	                    // Do something with the "Lookup Code" column value
	                }
	                // Add more conditions for other headers as needed
	            }
	        }
	        // Close the WebDriver when done
	        
	        
	       }catch(Exception e ){
	    	   try {
	    		   TakesScreenshot screenshot = ((TakesScreenshot) driver);
	    		   File sourceFile = screenshot.getScreenshotAs(OutputType.FILE);
	    		   FileUtils.copyFile(sourceFile, new File("C:\\Users\\KiranKumarReddyC\\eclipse-workspace\\com.P_O_M\\src\\Test_Results_Configrite"));
	    	   }catch(Exception screenshotException) {
	    		   System.out.println("Failed to capture screenshot :" + screenshotException.getMessage());
	    		   
	    	   }
	    	   System.out.println("Exception occurred: " +e.getMessage());
	    	   
	       }//finally{
	    	//   if (driver!=null) {
	    	//	   driver.quit();
	    	//   }
	      // }
	    }
	}
	
	
	

