package com.googlecode.jlibreoffice;

import java.awt.Container;
import java.awt.LayoutManager;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.net.URLDecoder;

import org.apache.log4j.Logger;

public class OOoBeanProxy {

	private static final Logger log = Logger.getLogger(OOoBeanProxy.class);
	
	private ClassLoader classLoader;
	private Object bean;
	private Class<?> beanClass;

	public OOoBeanProxy(ClassLoader classLoader) throws Exception {
		
		this.classLoader = classLoader;
		
		try {
			beanClass = classLoader.loadClass("com.sun.star.comp.beans.OOoBean");
		}
		catch (Exception e) {
			String msg = "A classe OOoBean não foi encontrada, mensagem interna: " + e.getMessage(); 
			log.error(msg, e);
			throw new Exception(msg);
		}
		
		try {
			bean = beanClass.newInstance();
		}
		catch (Exception e) {
			String msg = "Erro ao instanciar um objeto da classe OOoBean, mensagem interna: " + e.getMessage(); 
			log.error(msg, e);
			throw new Exception(msg);
		}
	}

	public void stopOOoConnection() {
		try {
			invoke(bean, "stopOOoConnection");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.stopOOoConnection(), mensagem interna: " + e.getMessage());
		}
	}
	
	public void setMenuBarVisible(boolean value) {
		try {			  
			invoke(bean, "setMenuBarVisible", value);
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.setMenuBarVisible(), mensagem interna: " + e.getMessage());
		}
	}
	
	public void setStandardBarVisible(boolean value) {
		try {
			invoke(bean, "setStandardBarVisible", value);
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.setStandardBarVisible(), mensagem interna: " + e.getMessage());
		}
	}
	
	public void setToolBarVisible(boolean value) {
		try {
			invoke(bean, "setToolBarVisible", value);
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.setToolBarVisible(), mensagem interna: " + e.getMessage());
		}
	}
		
	public boolean isOOoConnected() {		
		try {			
			return (Boolean) invoke(bean, "isOOoConnected");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.isOOoConnected(), mensagem interna: " + e.getMessage());
		}
	}
	
	public Object getDocument() {
		try {
			return invoke(bean, "getDocument");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.getDocument(), mensagem interna: " + e.getMessage());
		}
	}

	public String getDocumentURL() {
		try {
			return (String) invoke(getDocument(), "getURL");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.getDocumentURL(), mensagem interna: " + e.getMessage());
		}
	}
	
	public boolean isModifiedDocument() {
		try {
			return (Boolean) invoke(getDocument(), "isModified");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.isModifiedDocument(), mensagem interna: " + e.getMessage());
		}
	}
	
	public void loadFromStream(InputStream is) {
		try {
            Method methLoad = beanClass.getMethod("loadFromStream", InputStream.class, getPropertyValueArrayClass());
            
            methLoad.invoke(bean, is, null);
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.loadFromStream(), mensagem interna: " + e.getMessage());
		}		
	}
	
	public void loadFromURL(String url) {
		try {
            Method methLoad = beanClass.getMethod("loadFromURL", String.class, getPropertyValueArrayClass());

            methLoad.invoke(bean, url, null);
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.loadFromURL(), mensagem interna: " + e.getMessage());
		}
	}
	
	public void aquireSystemWindow() {
		try {
			invoke(bean, "aquireSystemWindow");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.aquireSystemWindow(), mensagem interna: " + e.getMessage());
		}
	}
	
	public Container getContainer() {
		return (Container) bean;
	}
	
	public void clear() {
		try {
			invoke(bean, "clear");
		}
		catch (Exception e) {
			throw new RuntimeException("Erro ao executar OOoBean.clear(), mensagem interna: " + e.getMessage());
		}
		
	}
	
	private Object invoke(Object object, String method, boolean value) throws Exception {
		
		Method m;
		
		try {
			m = object.getClass().getMethod(method, boolean.class);
			return m.invoke(object, value);
		}
		catch (Exception e) {}
		
		try {
			m = object.getClass().getDeclaredMethod(method, boolean.class);
			return m.invoke(object, value);
		}
		catch (Exception e) {
			throw e;
		}	
	}
	
	private Object invoke(Object object, String method, Object ... params) throws Exception {
		
		Class<?>[] parameterTypes = new Class<?>[params.length];
		
		int i=0;
		for (Object param : params) {
			parameterTypes[i++] = param.getClass();
		}
		
		Method m;
		
		try {
			m = object.getClass().getMethod(method, parameterTypes);
			return m.invoke(object, params);
		}
		catch (Exception e) {}
		
		try {
			m = object.getClass().getDeclaredMethod(method, parameterTypes);
			return m.invoke(object, params);
		}
		catch (Exception e) {
			throw e;
		}
	}
	
	private void set(Object object, String fieldName, Object value) throws Exception {		
		try {
			object.getClass().getField(fieldName).set(fieldName, value);		
		}
		catch (Exception e) {}
		
		try {
			object.getClass().getDeclaredField(fieldName).set(object, value);
		}
		catch (Exception e) {
			throw e;
		}		
	}
	
	private Class<?> getPropertyValueArrayClass() throws Exception {
		return Array.newInstance(classLoader.loadClass("com.sun.star.beans.PropertyValue"), 1).getClass();
	}
		
	public void execute(String cmd, Object[] propertyValues) throws Exception {
		
		try {			
			Class<?> componentContextClass = classLoader.loadClass("com.sun.star.uno.XComponentContext");
			Class<?> dispatchHelperClass = classLoader.loadClass("com.sun.star.frame.XDispatchHelper");
			Class<?> dispatchProviderClass = classLoader.loadClass("com.sun.star.frame.XDispatchProvider");
			Class<?> unoRuntimeClass = classLoader.loadClass("com.sun.star.uno.UnoRuntime");
			
			Object xCc = invoke(invoke(bean, "getOOoConnection"), "getComponentContext");
			Object xFrame = invoke(invoke(getDocument(),"getCurrentController"),"getFrame");
			Object sm = invoke(xCc, "getServiceManager");
			Object dispatchHelperObject = sm.getClass().getMethod("createInstanceWithContext", String.class, componentContextClass).invoke(sm, "com.sun.star.frame.DispatchHelper", xCc);
					
			Object xDh = unoRuntimeClass.getMethod("queryInterface", Class.class, Object.class).invoke(null, dispatchHelperClass, dispatchHelperObject);
			Object xDispatchProvider = unoRuntimeClass.getMethod("queryInterface", Class.class, Object.class).invoke(null, dispatchProviderClass, xFrame);
			Object xWindow = invoke(xFrame, "getComponentWindow");

			invoke(xWindow, "setFocus");

			Method excuteDispatchMethod = xDh.getClass().getMethod("executeDispatch", dispatchProviderClass, String.class, String.class, int.class, getPropertyValueArrayClass());
				
			excuteDispatchMethod.invoke(xDh, xDispatchProvider, cmd, "", 0, propertyValues);
		}
		catch (Exception e) {
			throw new Exception(e.getMessage());			
		}
	}

	public void setLayout(LayoutManager layout) {
		getContainer().setLayout(layout);
	}

	public void setVisible(boolean value) {
		getContainer().setVisible(value);
	}


	public String exportToPdf() throws Exception {

		/*
		String urlFile = o3Bean.getDocument().getURL();
		String urlPdf = urlFile + ".pdf";

		PropertyValue[] propertyValue = new PropertyValue[3];
		propertyValue[0] = new PropertyValue();
		propertyValue[0].Name = "URL";
		propertyValue[0].Value = urlPdf;
		propertyValue[1] = new PropertyValue();
		propertyValue[1].Name = "FilterName";
		propertyValue[1].Value = "writer_pdf_Export";
		propertyValue[2] = new PropertyValue();
		propertyValue[2].Name = "SelectionOnly";
		propertyValue[2].Value = false;

		execute(JLibreOffice.UNO_EXPORT_DIRECT_TO_PDF, propertyValue);
		 */
		
		String urlFile = getDocumentURL();		
		String urlPdf = urlFile + ".pdf";

		Class<?> propertyValueClass = classLoader.loadClass("com.sun.star.beans.PropertyValue");
		
		Object[] propertyValues = (Object[]) Array.newInstance(propertyValueClass, 3);

		propertyValues[0] = propertyValueClass.newInstance();
		set(propertyValues[0], "Name", "URL");
		set(propertyValues[0], "Value", urlPdf);

		propertyValues[1] = propertyValueClass.newInstance();
		set(propertyValues[1], "Name", "FilterName");
		set(propertyValues[1], "Value", "writer_pdf_Export");

		propertyValues[2] = propertyValueClass.newInstance();
		set(propertyValues[2], "Name", "SelectionOnly");
		set(propertyValues[2], "Value", false);
		
		execute(JLibreOffice.UNO_EXPORT_DIRECT_TO_PDF, propertyValues);
			
		String path = urlPdf.substring(8, urlPdf.length());
		
		try {
			return URLDecoder.decode(path, "utf-8");
		} 
		catch (UnsupportedEncodingException e) {
			return path;
		}
	}
}