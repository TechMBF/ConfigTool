package com.att.blackflag.json.util;
 
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
public class ExcelReader { 
 
   public static void readExcelData(String fileName) { 
       List<ExcelPojo> excelRowList = new ArrayList<ExcelPojo>(); 
         
       try { 
    	   String[] str = new String[5];
    	  
           //Create the input stream from the xlsx/xls file  
			FileInputStream fis = new FileInputStream(fileName); 
			             
           //Create Workbook instance for xlsx/xls file input stream  
			Workbook workbook = null; 
           if(fileName.toLowerCase().endsWith("xlsx")){ 
               workbook = new HSSFWorkbook(fis); 
           }else if(fileName.toLowerCase().endsWith("xls")){ 
               workbook = new HSSFWorkbook(fis); 
           } 
             
           //Get the number of sheets in the xlsx file 
           int numberOfSheets = workbook.getNumberOfSheets(); 
			       
           //loop through each of the sheets    
			for(int i=0; i < numberOfSheets; i++){ 
			                
               //Get the nth sheet from the workbook  
				Sheet sheet = workbook.getSheetAt(i); 
				
               //every sheet has rows, iterate over them  
				Iterator<Row> rowIterator = sheet.iterator(); 
               
               while (rowIterator.hasNext())  
               { 
                   //Get the row object       
					Row row = rowIterator.next(); 
					                     
                   //Every row has columns, get the column iterator and iterate over them 
                   Iterator<Cell> cellIterator = row.cellIterator(); 
                    int k=0;
                    ExcelPojo excelPojo = new ExcelPojo();
                   while (cellIterator.hasNext())  
                   { 
                	   
                       //Get the Cell object         
						Cell cell = cellIterator.next();

						if(cell.getCellType() == 1){
							str[k] = String.valueOf(cell.getStringCellValue()).trim();
						}else{
							str[k] = String.valueOf(cell.getNumericCellValue()).trim();
						}
					
						//System.out.println("The values are ="+str[k]);
						k++;
                   }
                   excelPojo.setProxyName(str[1]);  System.out.println("the value of 1 column = "+excelPojo.getProxyName());
                   excelPojo.setPolicyName(str[2]); System.out.println("the value of 1 column = "+excelPojo.getPolicyName());
                   excelPojo.setXpathParam(str[3]);System.out.println("the value of 1 column = "+excelPojo.getXpathParam());
                   excelPojo.setValueParam(str[4]);System.out.println("the value of 1 column = "+excelPojo.getValueParam()); 
                   excelRowList.add(excelPojo);
                 }
			}
			fis.close(); 
            
       } catch (IOException e) { 
           e.printStackTrace(); 
       } 
         
       
   } 
       
                 
   public static void main(String args[]){ 
	   readExcelData("C:\\Users\\MP00335779.ATT\\Desktop\\Json_Automation\\prod2.xls"); 
	   
      // List<Country> list = readExcelData("Sample.xlsx"); 
       //System.out.println("Country List\n"+list); 
   } 
 
}
