package com.att.blackflag.json.vo;

import java.io.Serializable;
import java.util.ArrayList;

import org.codehaus.jackson.annotate.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown=true)
public class Policies implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 6469452507833147881L;
	private String name = null;
	private ArrayList<Tokens> tokens;

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


	/**
	 * @return the tokens
	 */
	public ArrayList<Tokens> getTokens() {
		return tokens;
	}


	/**
	 * @param tokens the tokens to set
	 */
	public void setTokens(ArrayList<Tokens> tokens) {
		this.tokens = tokens;
	}


	@Override
	public String toString() {
		return "policies [tokens=" + tokens.toString()+ "]";
	}
}
