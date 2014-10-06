package com.att.blackflag.json.vo;

import java.io.Serializable;
import java.util.List;

import org.codehaus.jackson.annotate.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Configurations implements Serializable{
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 5325668313370323511L;
	
	private List<Config> configurations = null;

	public List<Config> getConfigurations() {
		return configurations;
	}

	public void setConfigurations(List<Config> configurations) {
		this.configurations = configurations;
	}
 }
