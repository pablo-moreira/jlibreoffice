package com.googlecode.jlibreoffice.util;

import java.io.File;
import java.lang.reflect.Field;

public class AdicionaPath {

//	public static void main(String[] args) {
	public void adicioPathRT(){
		try {
			Class clazz = ClassLoader.class;
			Field field = clazz.getDeclaredField("sys_paths");
			
			boolean accessible = field.isAccessible();
			if (!accessible)
				field.setAccessible(true);
			Object original = field.get(clazz);

			field.set(clazz, null);

			try {
				  System.setProperty("java.library.path", System.getProperty("java.library.path")
						+ ";"  
//						+ new File("C:/Arquivos de programas/BrOffice.org 3/program/;" )  
						+ new File("C:/Arquivos de programas/BrOffice.org 3/Basis/program/;" ) 
				  		+ new File("C:/Program Files/BrOffice.org 3/URE/java;" ) 
						+ new File("C:/Arquivos de programas/BrOffice.org 3/URE/bin ") );

				     //   + new File("C:/Program Files/BrOffice.org 2.2/program") );
//C:/Arquivos de programas/BrOffice.org 3/program;C:/Arquivos de programas/BrOffice.org 3/Basis/program;C:/Arquivos de programas/BrOffice.org 3/URE/bin
//				System.loadLibrary("jurt");
				
			} finally {
				field.set(clazz, original);
				field.setAccessible(accessible);
			}
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (NoSuchFieldException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
	}
}
