package project;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class others {

	
	public void excel(Map<String, Object[]> data2 ) {
		
                 	

    XSSFWorkbook workbook = new XSSFWorkbook(); 

    XSSFSheet sheet = workbook.createSheet("Countries"); 
    
    Map<String, Object[]> data = data2;
    

    Set<String> keyset = data.keySet(); 
   int rownum = 0; 
    for (String key : keyset) { 
 
        Row row = sheet.createRow(rownum++); 
        Object[] objArr = data.get(key); 
        int cellnum = 0; 
        for (Object obj : objArr) { 

            Cell cell = row.createCell(cellnum++); 
            if (obj instanceof String) 
                cell.setCellValue((String)obj); 
            else if (obj instanceof Integer) 
                cell.setCellValue((Integer)obj); 
        } 
    } 
    try { 
      
        FileOutputStream out = new FileOutputStream(new File("Internate_Countries.xlsx")); 
        workbook.write(out); 
        out.close(); 
        System.out.println("Internate_Countries.xlsx written successfully on disk."); 
    } 
    catch (Exception e) { 
        e.printStackTrace(); 
    } 
	

	
	
	}
	
}
	
	
	
	

