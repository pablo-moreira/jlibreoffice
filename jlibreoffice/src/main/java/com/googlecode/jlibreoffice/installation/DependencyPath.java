package com.googlecode.jlibreoffice.installation;


public class DependencyPath {

	private String dependencyName;
	private boolean required;
	
	public DependencyPath(String dependencyName, boolean required) {
		super();
		this.dependencyName = dependencyName;
		this.required = required;
	}

	public String getDependencyName() {
		return dependencyName;
	}
	
	public boolean isJar() {
		return getDependencyName().endsWith(".jar");
	}

	public boolean isRequired() {
		return required;
	}

	public boolean isLib() {
		return !isJar();
	}
}