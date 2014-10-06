package com.att.blackflag.json.util;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.LineNumberReader;
import java.util.LinkedHashMap;
import java.util.ResourceBundle;

public class JSONComparator {
	
	ResourceBundle resource = null;
	
	String apiName = null;
	
	public JSONComparator(String apiName) {
		this.apiName = apiName;
	}

	    /**
	     * 
	     * @return
	     * @throws FileNotFoundException
	     * @throws IOException
	     */
	    public LinkedHashMap<String, LinkedHashMap<String,String>> quickCompare() throws FileNotFoundException, IOException {

	    	resource = ResourceBundle.getBundle("com.att.blackflag.json.resource.ApplicationProperties");
	    	File f1 = new File(resource.getString("CONFIG_BASE_DIR")+apiName+"\\old-config.json");// Old Config 
	    	File f2 = new File(resource.getString("CONFIG_BASE_DIR")+apiName+"\\config.json");// New Config
	        
	        LinkedHashMap<String, LinkedHashMap<String,String>> dispMap = new LinkedHashMap<String,LinkedHashMap<String,String>>();
	        String line1 = null;
	        String line2 = null;

	    	LineNumberReader lr1 = new LineNumberReader(new FileReader(f1)); 
	        LineNumberReader lr2 = new LineNumberReader(new FileReader(f2)); 
	        LinkedHashMap<String, String> diffStrMap = null;
	        while (((line1 = lr1.readLine()) != null)
	                && ((line2 = lr2.readLine()) != null)) {
	        		        	
	            if (!line1.trim().equalsIgnoreCase(line2.trim()))
	            {	
	            	diffStrMap = new LinkedHashMap<String,String>(); 
	            	diffStrMap.put("<td nowrap style='background-color:#088A29;'>"+line1+"</td>", "<td nowrap style='background-color:#088A29;'>"+line2+"</td>");
	 	            dispMap.put("<td nowrap style='background-color:#088A29;'>"+Integer.toString(lr1.getLineNumber())+"</td>", diffStrMap);
	            }
	        }
	        lr1.close();
	        lr2.close();
	     return dispMap; 
	    }
	    
	    /**
	     * 
	     * @return
	     * @throws FileNotFoundException
	     * @throws IOException
	     */
	    public LinkedHashMap<String, LinkedHashMap<String,String>> detailCompare() throws FileNotFoundException, IOException {

	        
			resource = ResourceBundle.getBundle("com.att.blackflag.json.resource.ApplicationProperties");
	    	File f1 = new File(resource.getString("CONFIG_BASE_DIR")+apiName+"\\old-config.json");// Old Config 
	    	File f2 = new File(resource.getString("CONFIG_BASE_DIR")+apiName+"\\config.json");// New Config
	        
	        LinkedHashMap<String, LinkedHashMap<String,String>> dispMap = new LinkedHashMap<String,LinkedHashMap<String,String>>();
	        String line1 = null;
	        String line2 = null;

	    	LineNumberReader lr1 = new LineNumberReader(new FileReader(f1)); 
	        LineNumberReader lr2 = new LineNumberReader(new FileReader(f2)); 
	        LinkedHashMap<String, String> diffStrMap = null;
	        while (((line2 = lr2.readLine()) != null)) {
	        	line1 = lr1.readLine();
	            if (line1!=null && line2!=null && !line1.trim().equalsIgnoreCase(line2.trim()))
	            {	
	            	diffStrMap = new LinkedHashMap<String,String>(); 
	            	diffStrMap.put("<td nowrap style='background-color:#088A29;'>"+line1+"</td>", "<td nowrap style='background-color:#088A29;'>"+line2+"</td>");
	 	            dispMap.put("<td nowrap style='background-color:#088A29;'>"+Integer.toString(lr1.getLineNumber())+"</td>", diffStrMap);
	            }else if(line1==null && line2!=null){
	            	diffStrMap = new LinkedHashMap<String,String>(); 
	            	diffStrMap.put("<td nowrap style='background-color:#088A29;'></td>", "<td nowrap style='background-color:#088A29;'>"+line2+"</td>");
	 	            dispMap.put("<td nowrap style='background-color:#088A29;'>"+Integer.toString(lr2.getLineNumber())+"</td>", diffStrMap);
	            }else{
	            	diffStrMap = new LinkedHashMap<String,String>(); 
	            	diffStrMap.put("<td nowrap>"+line1+"</td>", "<td nowrap>"+line2+"</td>");
	 	            dispMap.put("<td nowrap>"+Integer.toString(lr2.getLineNumber())+"</td>", diffStrMap);
	            }
	      
	        }
	        lr1.close();
	        lr2.close();
	     return dispMap; 
	    }
	}	
		
