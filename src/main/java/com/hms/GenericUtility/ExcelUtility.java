package com.hms.GenericUtility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.WebDriverWait;

public class ExcelUtility extends JavaUtility {
	
	/**
	 * This method is used  to read the data from excel
	 * @author Admin
	 * @param SheetName
	 * @param RowNo
	 * @param ColumnNo
	 * @return
	 * @throws Throwable
	 */
	public String readDataFromExcel(String SheetName,int RowNo, int ColumnNo) throws Throwable
	{
		FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
		 Workbook wb = WorkbookFactory.create(fi);
		Sheet sh = wb.getSheet(SheetName);
		 Row ro = sh.getRow(RowNo);
		  Cell cel = ro.getCell(ColumnNo);
		 String value = cel.getStringCellValue();
		 return value;
	}
	
	
	
	/**
	 * This method is used to read the data from Excel
	 * @author Admin
	 * @param SheetName
	 * @param RowNo
	 * @param ColumnNo
	 * @param data
	 * @throws Throwable
	 */
	public void writeDataIntoExcel(String SheetName,int RowNo, int ColumnNo,String data ) throws Throwable
	{
		FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
		//create workbook
		 Workbook wb = WorkbookFactory.create(fi);
		//get sheet
		Sheet sh = wb.getSheet(SheetName);
		//get row
		 Row ro = sh.createRow(RowNo);
		//create cell
		  Cell cel = ro.createCell(ColumnNo);
		  cel.setCellValue(data);
	FileOutputStream fos = new FileOutputStream(IPathConstants.ExcelPath);	  
	wb.write(fos);
	}
	
	/**
	 * This method is used to get last row count 
	 * @param SheetName
	 * @return
	 * @throws Throwable
	 */
	
	public int getLastRowNo(String SheetName) throws Throwable
	{
		 FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
         Workbook wb = WorkbookFactory.create(fi);
		    Sheet sh = wb.getSheet(SheetName);
		     int count = sh.getLastRowNum();
			return count;     
		     
	}
	/*for arraylist*/
	public void getList( String sheetName, WebDriver driver) throws Throwable{
		
		
		 FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
         Workbook wb = WorkbookFactory.create(fi);
		    Sheet sh = wb.getSheet(sheetName);
		     int count = sh.getLastRowNum();
		   for(int i=0; i<=count; i++)
 		     {
 		    	  String key = sh.getRow(i).getCell(0).getStringCellValue();
 		    	  String value = sh.getRow(i).getCell(1).getStringCellValue();
 		    	  if(key.equals("docemail")) {
 		    		  value=value+"a"+getRandomNumber();
 		    	  }
 		    	  if(key.equals("patemail")) {
 		    		 value=value+"a"+getRandomNumber();
 		    	  }
 		    	  driver.findElement(By.name(key)).sendKeys(value);
 		     }
	}
		  /*arraylist*/ 
		   public void getList1( String sheetName, WebDriver driver) throws Throwable{
				
				
				 FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
		         Workbook wb = WorkbookFactory.create(fi);
				    Sheet sh = wb.getSheet(sheetName);
				     int count = sh.getLastRowNum();
				   for(int i=0; i<=count; i++)
		 		     {
		 		    	  String key = sh.getRow(i).getCell(0).getStringCellValue();
		 		    	  String value = sh.getRow(i).getCell(1).getStringCellValue();
		 		    	 driver.findElement(By.name(key)).clear();
		 		    	  if(key.equals("city")) {
		 		    		 value=value+"a"+getRandomNumber();
		 		    	  }
		 		    	 driver.findElement(By.name(key)).clear();
		 		    	  driver.findElement(By.name(key)).sendKeys(value);
		 		     }
		   }	    
			
	
		   public Map<String , String>getList(String sheetName) throws Throwable
		   {
			   FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
		         Workbook wb = WorkbookFactory.create(fi);
				    Sheet sh = wb.getSheet(sheetName);
				     int count = sh.getLastRowNum();
				  Map<String,String> map = new HashMap<String, String>();
				  for(int i=0;i<=count;i++)
				  {
					  String key = sh.getRow(i).getCell(0).getStringCellValue();
	 		    	  String value = sh.getRow(i).getCell(1).getStringCellValue();
	 		    	  map.put(key, value);
					  
				  }
				  return map;
				     
		   }
		   
		   public Object[][] readMultipleData(String sheetName) throws Throwable
		   {
			   FileInputStream fi = new FileInputStream(IPathConstants.ExcelPath);
			   Workbook wb = WorkbookFactory.create(fi);
			    Sheet sh = wb.getSheet(sheetName);
			   int lastRow = sh.getLastRowNum()+1;
			     int lastCell = sh.getRow(0).getLastCellNum();
			     
			     Object[][] obj = new Object[lastRow][lastCell];
			     for(int i=0;i<lastRow;i++) {
			    	 
			    	 for(int j=0;j<lastCell;j++) {
			    		  obj[i][j]=sh.getRow(i).getCell(j).getStringCellValue();
			    	 }
			    	 
			     }
			     return obj;
			    	 
			    		
		   }
		   
		   
	}
	
	
	


