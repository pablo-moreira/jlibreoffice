package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.Frame;
import java.lang.reflect.Array;
import java.lang.reflect.Method;

public class OOoBeanProxy {

	private ClassLoader classLoader;
	private Object bean;
	private Class<?> beanClass;

	public OOoBeanProxy(ClassLoader classLoader) throws Exception {
		
		super();
		
		this.classLoader = classLoader;
		
		try {
			beanClass = classLoader.loadClass("com.sun.star.comp.beans.OOoBean");
		}
		catch (Exception e) {
			e.printStackTrace();
			throw new Exception("A classe OOoBean não foi encontrada, mensagem interna: " + e.getMessage());
		}
		
		bean = beanClass.newInstance();
	}	
		
	public boolean isOOoConnected() {		
		try {
			Method method = beanClass.getMethod("isOOoConnected", new Class[]{});
		
			Object resultado = method.invoke(bean, new Object[] {});
			
			return (Boolean) resultado;
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.isOOoConnected, mensagem interna: " + e.getMessage());
		}		
	}
	
	public void loadFromURL() {
		
		try {
			Object arProp = Array.newInstance(classLoader.loadClass("com.sun.star.beans.PropertyValue"), 1);
            
			//Class<? extends Object> clazz = arProp.getClass();

            Method methLoad = beanClass.getMethod("loadFromURL", new Class[] { String.class, arProp.getClass() });

            methLoad.invoke(bean, new Object[] {"private:factory/swriter", null});
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.loadFromURL, mensagem interna: " + e.getMessage());
		}		
	}
	
	public void insertInto(Container container) {
		container.add((java.awt.Container) bean, BorderLayout.CENTER);
	}
	
	
}
