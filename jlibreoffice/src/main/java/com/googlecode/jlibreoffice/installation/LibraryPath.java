package com.googlecode.jlibreoffice.installation;


public class LibraryPath {

	private String libraryName;
	private boolean required;
	
	public LibraryPath(String lib, boolean required) {
		super();
		this.libraryName = lib;
		this.required = required;
	}

	public String getLibraryName() {
		return libraryName;
	}

	public boolean isRequired() {
		return required;
	}
}