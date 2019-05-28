package project;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class summarize  extends others{

	static WebDriver driver;
	public static void main(String[] args) {
		
        System.setProperty("webdriver.chrome.driver", "C:\\\\Users\\\\Ashik\\\\Downloads\\\\chromedriver_win32\\\\chromedriver.exe");
        driver = new ChromeDriver();
        JavascriptExecutor js = (JavascriptExecutor) driver;
        driver.get("https://www.internetworldstats.com/top20.htm");
        List<String> Countrylist=new ArrayList<String>();
        for(int i=3;i<23;i++) {
        	WebElement cell=driver.findElement(By.xpath("/html/body/table[4]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr["+i+"]/td[2]"));
        	//System.out.println("Country name : "+cell.getText());
        	Countrylist.add(cell.getText());
        	
         }
        

        
        Iterator<String>it=Countrylist.iterator();
     	
     	
     		String s=Countrylist.get(0);
     		System.out.println("Country Name is : "+s);
     		driver.get("https://www.google.com");
     		
            WebElement google=driver.findElement(By.name("q"));
            google.sendKeys(s);;
            google.submit();
            google=driver.findElement(By.className("LC20lb"));
           
            google.click();
            
            List <String> At=new ArrayList<String>();
           WebElement dateBox = driver.findElement(By.xpath("//*[contains(text(),'Largest city')]/parent::tr//td/a[text()]"));
           WebElement dateBox2 = driver.findElement(By.xpath("//*[contains(text(),'Capital')]/parent::tr//td/a[text()]"));
           WebElement dateBox3 = driver.findElement(By.xpath("//*[contains(text(),'Official script')]/ancestor::tr//td/a[text()] "));
           WebElement dateBox5 = driver.findElement(By.xpath("//*[contains(text(),'Calling code')]/ancestor::tr//td/a[text()]"));
           WebElement dateBox6 = driver.findElement(By.xpath("//*[contains(text(),'President')]/ancestor::tr//td/a[text()]"));
           WebElement dateBox7 = driver.findElement(By.xpath("//*[contains(.,'Population')]/parent::tr/following-sibling::tr[1]/td"));
           WebElement dateBox8 = driver.findElement(By.xpath("//*[contains(.,'Currency')]/ancestor::tr//td/a[text()]"));
          
               // System.out.println(dateBox7.getText());
                At.add(dateBox.getText());
                 At.add(dateBox2.getText());
                
            System.out.println(dateBox3.getText()+dateBox5.getText());
            
            At.add(dateBox3.getText());
            At.add(dateBox5.getText());
            At.add(dateBox6.getText());
            At.add(dateBox7.getText());
            At.add(dateBox8.getText());
            for(int ii=0;ii<At.size();ii++)
            
            {
            	System.out.println(At.get(ii)+" ");
            }
            
            
            XSSFWorkbook workbook = new XSSFWorkbook(); 

            XSSFSheet sheet = workbook.createSheet("Countries"); 
            summarize ot=new summarize();
            
            Map<String, Object[]> data = new TreeMap<String, Object[]>();
           
           data.put("1", new Object[]{ s }); 
            data.put("2", new Object[]{ "Largest City : "+At.get(0)}); 
            data.put("3", new Object[]{ "Capital :"+At.get(1) }); 
            data.put("4", new Object[]{  "Official Script : "+At.get(2)}); 
            data.put("5", new Object[]{  "Calling Code : "+At.get(3)});
            data.put("6", new Object[]{  "President : "+At.get(4)});
            data.put("7", new Object[]{  "Population : "+At.get(5)});
            data.put("8", new Object[]{  "Currency : "+At.get(6)});
             
            summarize ss=new summarize();
            ss.excel(data);
//            Set<String> keyset = data.keySet(); 
//           int rownum = 0; 
//            for (String key : keyset) { 
//         
//                Row row = sheet.createRow(rownum++); 
//                Object[] objArr = data.get(key); 
//                int cellnum = 0; 
//                for (Object obj : objArr) { 
//
//                    Cell cell = row.createCell(cellnum++); 
//                    if (obj instanceof String) 
//                        cell.setCellValue((String)obj); 
//                    else if (obj instanceof Integer) 
//                        cell.setCellValue((Integer)obj); 
//                } 
//            } 
//            try { 
//              
//                FileOutputStream out = new FileOutputStream(new File("Internate_Countries.xlsx")); 
//                workbook.write(out); 
//                out.close(); 
//                System.out.println("Internate_Countries.xlsx written successfully ."); 
//            } 
//            catch (Exception e) { 
//                e.printStackTrace(); 
//            } 
        	
            driver.close();
       
     	}

}
	
	

