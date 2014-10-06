package com.att.blackflag.json.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.map.ObjectMapper;

import com.att.blackflag.json.vo.Configurations;

public class JSONUtil {
	
	private static final String XLS_EXT = "xls";
	
	private static final String XLSX_EXT = "xlsx";
	private static String newFile = "\\config.json";
	/**
	 * 
	 * @param config
	 * @param mapper
	 * @throws IOException
	 */
	@SuppressWarnings("deprecation")
	public static void generarePrettyPrintJSON(String dirName,String apiName,Configurations config) throws JsonGenerationException,IOException
	{
		ObjectMapper mapper = new ObjectMapper();
		mapper.defaultPrettyPrintingWriter().writeValue(new File(dirName+apiName+newFile), config);
	}
	
	/**
	 * 
	 * @param envName
	 * @param excelEnv
	 * @return
	 */
	public static boolean isValidEnvironment(String envName, String excelEnv)
	{
		boolean IsValidEnv= false;
		
		if(excelEnv.equalsIgnoreCase(envName))
		{
			IsValidEnv = true;
		}
		return IsValidEnv;
	}
	
	/**
	 * 
	 * @param envName
	 * @param excelEnv
	 * @return
	 */
	public static boolean isNewEnvironment(Configurations config, String envName) 
	{
		boolean isNewEnv= true;
		for(int c=0;c<config.getConfigurations().size();c++){
			String sEnvName = config.getConfigurations().get(c).getName().toString(); 
			if(sEnvName.equalsIgnoreCase(envName))
			{
				isNewEnv = false;
			}
		}
		return isNewEnv;
	}
	
	/**
	 * 
	 * @param fileName
	 * @return
	 */
	 public static ArrayList<ExcelPojo> readExcelData(String fileName) throws IOException{ 
	       ArrayList<ExcelPojo> excelRowList = new ArrayList<ExcelPojo>(); 
	         
	    	   String[] str = new String[5];
	           //Create the input stream from the xlsx/xls file  
			   FileInputStream fis = new FileInputStream(fileName); 
				             
	           //Create Workbook instance for xlsx/xls file input stream  
				Workbook workbook = null; 
	           if(fileName.toLowerCase().endsWith(XLSX_EXT)){ 
	        	   workbook = new XSSFWorkbook(fis);
	           }else if(fileName.toLowerCase().endsWith(XLS_EXT)){ 
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
						
							k++;
	                   }
	
	                   excelPojo.setProxyName(str[1]);
	                   excelPojo.setPolicyName(str[2]);
	                   excelPojo.setXpathParam(str[3]);
	                   excelPojo.setValueParam(str[4]); 
	                   excelRowList.add(excelPojo);
	                 }
				}
				fis.close(); 
	      return excelRowList;
	   }
}
