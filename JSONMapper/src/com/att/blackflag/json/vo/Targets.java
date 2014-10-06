package com.att.blackflag.json.vo;

import java.io.Serializable;
import java.util.ArrayList;

import org.codehaus.jackson.annotate.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Targets implements Serializable{
	
/**
	 * 
	 */
	private static final long serialVersionUID = 7526938776799658243L;

private String name = null;
	
private ArrayList<Tokens> tokens;
	public ArrayList<Tokens> getTokens() {
	return tokens;
}


public void setTokens(ArrayList<Tokens> tokens) {
	this.tokens = tokens;
}


	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}


	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}



	@Override
	public String toString() {
		return "targets [tokens=" + tokens.toString()+ "]";
	}
	
}
