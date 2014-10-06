package com.att.blackflag.json.util;

import java.util.*;
import java.io.*;

	public class DirReader {
	    public static void main(String[] args) { 
	        
	    	//String relativePath = "\\applications\\AccountDetailsAPI\\";
	    	
	    	if (args.length == 0) { 
	            args = new String[] { "D:\\workspace\\JSON\\applications" }; 
	        }

	        List<String> nextDir = new ArrayList<String>(); 
	        nextDir.add(args[0]); // either the one file, or the directory
	        try { 
	            while(nextDir.size() > 0) {     // size() is num of elements in List 
	                File pathName = new File(nextDir.get(0)); // gets the element at the index of the List 
	                String[] fileNames = pathName.list();  // lists all files in the directory
	                for(int i = 0; i < fileNames.length; i++) { 
	                  File f = new File(pathName.getPath(), fileNames[i]); // getPath converts abstract path to path in String, 
	                                                                    // constructor creates new File object with fileName name   
	                  if (f.isDirectory()) { 
	                     System.out.println(f.getCanonicalPath()); 
	                     nextDir.add(f.getPath()); 
	                  } 
	                  else {
	                      System.out.println(f);
	                  }
	               } 
	               nextDir.remove(0); 
	            } 
	        } 
	        catch(IOException e) { 
	            e.printStackTrace();  
	        }       
	    } 
	}
