package com.att.blackflag.json.helper;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.ResourceBundle;

import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

import com.att.blackflag.json.util.ExcelPojo;
import com.att.blackflag.json.util.JSONUtil;
import com.att.blackflag.json.vo.Config;
import com.att.blackflag.json.vo.Configurations;

public class JSONObjectMapper {
	
	private String envName = null;
	ResourceBundle resource = null;
	
	private String NEW_CONFIG_FILE_NAME = "\\config.json";
	private String OLD_CONFIG_FILE_NAME = "\\old-config.json";
	private String TEMPLATE_FILE_NAME = "\\template-config.json";
	private String inputFileName = null;

	ArrayList<HashMap<String, String>> policyList = new ArrayList<HashMap<String, String>>();
	LinkedHashMap<String, String> resultApi =new LinkedHashMap<String, String>();
	public JSONObjectMapper(String envName, String inputFileName)
	{
		this.inputFileName = inputFileName;
		this.envName = envName;
		resource = ResourceBundle.getBundle("com.att.blackflag.json.resource.ApplicationProperties");
	}

	/**
	 */
	@SuppressWarnings("deprecation")
	public HashMap<String, String> addNewEnvConfig(){
		String configBaseDirectory = resource.getString("CONFIG_BASE_DIR");
		String newEnvTemplateConfig = resource.getString("TEMPLATE_BASE_DIR");
		String[] apiNames = resource.getString("API-NAMES").split(",");
			try {
				for(String apiName : apiNames){
					try{
						ObjectMapper mapper = new ObjectMapper();
						Configurations config = mapper.readValue(new File(configBaseDirectory+apiName+NEW_CONFIG_FILE_NAME), Configurations.class);
						boolean isNewEnv = JSONUtil.isNewEnvironment(config,envName);
						if(isNewEnv){
							mapper.defaultPrettyPrintingWriter().writeValue(new File(configBaseDirectory+apiName+OLD_CONFIG_FILE_NAME), config);
							
							ObjectMapper newEnvConfigMapper = new ObjectMapper();
							newEnvConfigMapper.writeValue(new File(configBaseDirectory+apiName+NEW_CONFIG_FILE_NAME), config);
							Configurations newConfig = newEnvConfigMapper.readValue(new File(configBaseDirectory+apiName+NEW_CONFIG_FILE_NAME), Configurations.class);
							
							ObjectMapper templateMapper = new ObjectMapper();
							Config newEnvConfig = templateMapper.readValue(new File(newEnvTemplateConfig+apiName+TEMPLATE_FILE_NAME), Config.class);
							newEnvConfig.setName(envName);
							newConfig.getConfigurations().add(newEnvConfig);
	
							JSONUtil.generarePrettyPrintJSON(configBaseDirectory,apiName,newConfig);
							resultApi.put(apiName, "<p bgcolor='red'> New Environment Added Successfully </p>");
						}else{
							resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! The environment already exist in this API</p>");
						}
					}catch (JsonGenerationException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Problem occurred while write JSON into config.json </p>");
						e.printStackTrace();
					} catch (JsonMappingException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Invalid JSON format,Please validate config.json </p>");
						e.printStackTrace();
					} catch (IOException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Config file not found for the  API to add/update</p>");
						e.printStackTrace();
					}
				}
				return resultApi;
			}catch (Exception e) {
				e.printStackTrace();
				return resultApi;
			} 
	}
	
	
	/**
	 */
	@SuppressWarnings("deprecation")
	public HashMap<String, String> updateEnvConfig(){
		LinkedHashSet<String> apiNames =new LinkedHashSet<String>();
		ArrayList<ExcelPojo> excelPojoList = new ArrayList<ExcelPojo>();
		String configBaseDirectory = resource.getString("CONFIG_BASE_DIR");
			try {
				ObjectMapper mapper = new ObjectMapper();
				excelPojoList = JSONUtil.readExcelData(new File(inputFileName).toString());
				if(excelPojoList!=null)
					excelPojoList.remove(0);
				apiNames = getAPINames(excelPojoList);
				String apiName = null;
				Iterator<String> iter = apiNames.iterator();
				while (iter.hasNext()) {
					apiName = (String)iter.next();
					try{
						Configurations config = mapper.readValue(new File(configBaseDirectory+apiName+NEW_CONFIG_FILE_NAME), Configurations.class);
						mapper.defaultPrettyPrintingWriter().writeValue(new File(configBaseDirectory+apiName+OLD_CONFIG_FILE_NAME), config);
							if(config!=null && config.getConfigurations().size()>0){
								for(int c=0;c<config.getConfigurations().size();c++){
									String sEnvName = config.getConfigurations().get(c).getName().toString(); 
									boolean bValidEnv = JSONUtil.isValidEnvironment(sEnvName,envName);
									if(bValidEnv){
										Iterator<ExcelPojo> itr = excelPojoList.iterator();
										while(itr.hasNext()){
											ExcelPojo excelPojo = new ExcelPojo();
											excelPojo = (ExcelPojo)itr.next();
											if(excelPojo.getProxyName().equalsIgnoreCase(apiName)){
												updatePolicies(config,c,excelPojo);
												updateProxies(config,c,excelPojo);
												updateTargets(config,c,excelPojo);
												resultApi.put(apiName, "Environment Updated Successfully");
											}
										}
									}
								}
							}
							JSONUtil.generarePrettyPrintJSON(configBaseDirectory,apiName,config);
					}catch (JsonGenerationException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Problem occurred while write JSON into config.json </p>");
						e.printStackTrace();
					} catch (JsonMappingException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Invalid JSON format,Please validate config.json </p>");
						e.printStackTrace();
					} catch (IOException e) {
						resultApi.put(apiName, "<p bgcolor='red'> We're Sorry! Config file not found for the  API to add/update</p>");
						e.printStackTrace();
					}
				}
				return resultApi;
			}catch (Exception e) {
				e.printStackTrace();
				return resultApi;
			} 
	}
	
	/**
	 * 
	 * @param apiList
	 */
	private LinkedHashSet<String> getAPINames(ArrayList<ExcelPojo> excelList){
		LinkedHashSet<String> apiNames =new LinkedHashSet<String>();
		if(excelList!=null){
			Iterator<ExcelPojo> itr = excelList.iterator();
			ExcelPojo excelPojo = null;
			while(itr.hasNext())
			{
				excelPojo = (ExcelPojo)itr.next();
				if(excelPojo!=null){
					apiNames.add(excelPojo.getProxyName());
				}
			}
		}
		return apiNames;
	}
	
	/*
	 *  This is used to update policies values according the matches
	 */
	private void updatePolicies(Configurations config,int c,ExcelPojo excelPojo)
	{
		int policyCount = config.getConfigurations().get(c).getPolicies().size();
		
		for(int i=0;i<policyCount;i++)
		{
			if(excelPojo.getPolicyName().equalsIgnoreCase(config.getConfigurations().get(c).getPolicies().get(i).getName()))
			{
				int pTokenCount = config.getConfigurations().get(c).getPolicies().get(i).getTokens().size();
	
				for(int j=0;j<pTokenCount;j++)
				{
					if(excelPojo.getXpathParam().equalsIgnoreCase(config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).getXpath()))
					{
						config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).setValue(excelPojo.getValueParam());
					}
				}
			}
		}
	}
	
	/*
	 * This is used to update proxies values according the matches
	 */
	private void updateProxies(Configurations config,int c,ExcelPojo excelPojo)
	{
		int proxyCount = config.getConfigurations().get(c).getProxies().size();
		
		for(int i=0;i<proxyCount;i++)
		{
			if(excelPojo.getPolicyName().equalsIgnoreCase(config.getConfigurations().get(c).getProxies().get(i).getName()))
			{
				int pTokenCount = config.getConfigurations().get(c).getProxies().get(i).getTokens().size();
	
				for(int j=0;j<pTokenCount;j++)
				{
					if(excelPojo.getXpathParam().equalsIgnoreCase(config.getConfigurations().get(c).getProxies().get(i).getTokens().get(j).getXpath()))
					{
						config.getConfigurations().get(c).getProxies().get(i).getTokens().get(j).setValue(excelPojo.getValueParam());
					}
				}
			}
		}
	}
	
	/*
	 * This is used to update targets values according the matches
	 */
	private void updateTargets(Configurations config,int c,ExcelPojo excelPojo)
	{
		int targetCount = config.getConfigurations().get(c).getTargets().size();
		//System.out.println("i am in target = "+targetCount);
		for(int i=0;i<targetCount;i++)
		{
			//System.out.println("i am in target with excel policy name = "+excelPojo.getPolicyName());
			//System.out.println("target's name = "+config.getConfigurations().get(c).getTargets().get(i).getName());
			if(excelPojo.getPolicyName().equalsIgnoreCase(config.getConfigurations().get(c).getTargets().get(i).getName()))
			{
				int pTokenCount = config.getConfigurations().get(c).getTargets().get(i).getTokens().size();
	
				for(int j=0;j<pTokenCount;j++)
				{
					if(excelPojo.getXpathParam().equalsIgnoreCase(config.getConfigurations().get(c).getTargets().get(i).getTokens().get(j).getXpath()))
					{
						config.getConfigurations().get(c).getTargets().get(i).getTokens().get(j).setValue(excelPojo.getValueParam());
					}
				}
			}
		}
	}
	
}










































































/*package com.att.blackflag.json.helper;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.codehaus.jackson.JsonParseException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

import com.att.blackflag.json.util.ExcelPojo;
import com.att.blackflag.json.util.JSONUtil;
import com.att.blackflag.json.vo.Configurations;

public class JSONObjectMapper {
	
	private String envName = null;
	private String inputFileName = null;
	private String jsonFileName = null;
	private String apiName = null;
	private String INPUT_BASE_DIR = "C:\\Users\\MP00335779.ATT\\Desktop\\JSON_PROJECT\\WebContent\\CONFIG_DIR\\";
	private String CONFIG_FILE_NAME = "\\config.json";
	ArrayList<ExcelPojo> excelPojoList = new ArrayList<ExcelPojo>();
	HashSet<String> apiNames =new HashSet<String>();
	ArrayList<HashMap<String, String>> policyList = new ArrayList<HashMap<String, String>>();
	//ArrayList<HashMap<String, String>> xpathList = new ArrayList<HashMap<String, String>>();
	HashMap<String, String> policyParamMap = new HashMap<String, String>();
	public JSONObjectMapper(String envName, String inputFileName) {
		this.envName = envName;
		this.inputFileName = inputFileName;
		excelPojoList = JSONUtil.readExcelData(new File(inputFileName).toString());
		getAPINames(excelPojoList);
	}

	public String updateEnvConfig(){
		ObjectMapper mapper = new ObjectMapper();
		try {
			
			String apiName = null;
			Iterator iter = apiNames.iterator();
			while (iter.hasNext()) {
				apiName = (String)iter.next();
				
				
				Configurations config = mapper.readValue(new File(INPUT_BASE_DIR+apiName+CONFIG_FILE_NAME), Configurations.class);
				for(int c=0;c<config.getConfigurations().size();c++)
				{
					String sEnvName = config.getConfigurations().get(c).getName().toString();
					if(sEnvName.equals(envName)){
						int policyCount = config.getConfigurations().get(c).getPolicies().size();
						for(int i=0;i<policyCount;i++)
						{
							//Iterator policyIterator = policyList.iterator();
							//while(policyIterator.hasNext()){
								//apiName = (String)policyIterator.next();
								HashMap<String, String> policyAndParam = getPolicyAndParamNames(apiName, excelPojoList);
								Iterator it = policyAndParam.entrySet().iterator();
							    while (it.hasNext()) {
							        Map.Entry pairs = (Map.Entry)it.next();
							        String key = (String)pairs.getKey();
									if(key.equals((config.getConfigurations().get(c).getPolicies().get(i).getName())))
									{
										int pTokenCount = config.getConfigurations().get(c).getPolicies().get(i).getTokens().size();
							
										for(int j=0;j<pTokenCount;j++)
										{
										/*	if(policyAndParam.getXpathParam().equalsIgnoreCase(config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).getXpath()))
											{
												config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).setValue(excelPojo.getValueParam());
												System.out.println(config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).getValue());
											}
										}
									}
							}
						}
					}
				}
			}
			return null;
		}
		catch (JsonParseException e) {
			e.printStackTrace();
			return null;
		} catch (JsonMappingException e) {
			e.printStackTrace();
			return null;
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}
	}

	public String updateJSONConfig(){
		

		System.out.println(inputFileName +"=="+jsonFileName+"=="+envName);
		excelPojoList = JSONUtil.readExcelData(new File(inputFileName).toString()); 
		System.out.println(apiNames);
		ObjectMapper mapper = new ObjectMapper();
			try {
					Configurations config = mapper.readValue(new File(jsonFileName), Configurations.class);
					
					if(config!=null && config.getConfigurations().size()>0){
						for(int c=0;c<config.getConfigurations().size();c++)
						{
							String sEnvName = config.getConfigurations().get(c).getName().toString(); 
							boolean bValidEnv = JSONUtil.isValidEnvironment(sEnvName,envName);
							if(bValidEnv)
							{
		
								Iterator<ExcelPojo> itr = excelPojoList.iterator();
								while(itr.hasNext())
								{
									ExcelPojo excelPojo = new ExcelPojo();
									excelPojo = (ExcelPojo)itr.next();
									if(excelPojo.getProxyName().equalsIgnoreCase(apiName))
									{
							
										int policyCount = config.getConfigurations().get(c).getPolicies().size();
										System.out.println("The policy count = "+policyCount);
									
										for(int i=0;i<policyCount;i++)
										{
											if(excelPojo.getPolicyName().equalsIgnoreCase(config.getConfigurations().get(c).getPolicies().get(i).getName()))
											{
												int pTokenCount = config.getConfigurations().get(c).getPolicies().get(i).getTokens().size();
									
												for(int j=0;j<pTokenCount;j++)
												{
													if(excelPojo.getXpathParam().equalsIgnoreCase(config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).getXpath()))
													{
														config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).setValue(excelPojo.getValueParam());
														System.out.println(config.getConfigurations().get(c).getPolicies().get(i).getTokens().get(j).getValue());
													}
												}
											}
										}
									}
								}
							}
						}
					}	
				return JSONUtil.JSONGenerator(config, mapper);
			}
			catch (JsonParseException e) {
				e.printStackTrace();
				return null;
			} catch (JsonMappingException e) {
				e.printStackTrace();
				return null;
			} catch (IOException e) {
				e.printStackTrace();
				return null;
			}
	}
	
	/**
	 * 
	 * @param apiList
	 
	private void getAPINames(ArrayList<ExcelPojo> apiList){
		if(apiList!=null){
			Iterator<ExcelPojo> itr = apiList.iterator();
			ExcelPojo excelPojo = null;
			while(itr.hasNext())
			{
				excelPojo = (ExcelPojo)itr.next();
				if(excelPojo!=null){
					apiNames.add(excelPojo.getProxyName());
				}
			}
		}
	}
	
	/**
	 * 
	 * @param apiList
	 
	private HashMap<String, String> getPolicyAndParamNames(String apiName,ArrayList<ExcelPojo> apiList){
		if(apiList!=null){
			Iterator<ExcelPojo> itr = apiList.iterator();
			ExcelPojo excelPojo = null;
			while(itr.hasNext())
			{
				excelPojo = (ExcelPojo)itr.next();
				if(excelPojo.getProxyName().equals(apiName)){
					policyParamMap.put(excelPojo.getPolicyName(),excelPojo.getXpathParam());
				}
			}
		return policyParamMap;
	}
}*/
