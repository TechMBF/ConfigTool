package com.att.blackflag.json.vo;

import java.io.Serializable;

import org.codehaus.jackson.annotate.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Tokens implements Serializable{

	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private String xpath= null;
	
	private String value = null;

	/**
	 * @return the xpath
	 */
	public String getXpath() {
		return xpath;
	}

	/**
	 * @param xpath the xpath to set
	 */
	public void setXpath(String xpath) {
		this.xpath = xpath;
	}

	/**
	 * @return the value
	 */
	public String getValue() {
		return value;
	}

	/**
	 * @param value the value to set
	 */
	public void setValue(String value) {
		this.value = value;
	}

	/*
	@SuppressWarnings("unused")
	private HashMap<String, String> properties = null;*/
}
