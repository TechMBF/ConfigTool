package com.att.blackflag.json.vo;
import java.util.ArrayList;
import java.util.List;

import org.codehaus.jackson.annotate.JsonIgnoreProperties;

@JsonIgnoreProperties(ignoreUnknown=true)  
public class Config {
	
	 private String name = null;
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

	
	private ArrayList<Proxies> proxies= null;
	private ArrayList<Policies> policies = null;
	private ArrayList<Policies> targets = null;
	
	
	
	/**
	 * @return the targets
	 */
	public ArrayList<Policies> getTargets() {
		return targets;
	}

	/**
	 * @param targets the targets to set
	 */
	public void setTargets(ArrayList<Policies> targets) {
		this.targets = targets;
	}

	/**
	 * @return the proxies
	 */
	public ArrayList<Proxies> getProxies() {
		return proxies;
	}

	/**
	 * @param proxies the proxies to set
	 */
	public void setProxies(ArrayList<Proxies> proxies) {
		this.proxies = proxies;
	}

	
	
		
	/**
	 * @return the policies
	 */
	public List<Policies> getPolicies() {
		return policies;
	}

	/**
	 * @param policies the policies to set
	 */
	public void setPolicies(ArrayList<Policies> policies) {
		this.policies = policies;
	}

}
