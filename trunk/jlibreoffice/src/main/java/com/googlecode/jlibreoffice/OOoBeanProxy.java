package com.googlecode.jlibreoffice;

import java.awt.Container;
import java.awt.LayoutManager;
import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.lang.reflect.Type;

public class OOoBeanProxy {

	private ClassLoader classLoader;
	private Object bean;
	private Class<?> beanClass;

	public OOoBeanProxy(ClassLoader classLoader) throws Exception {
		
		this.classLoader = classLoader;
		
		try {
			beanClass = classLoader.loadClass("com.sun.star.comp.beans.OOoBean");
		}
		catch (Exception e) {
			e.printStackTrace();
			throw new Exception("A classe OOoBean não foi encontrada, mensagem interna: " + e.getMessage());
		}
		
		try {
			bean = beanClass.newInstance();
		}
		catch (Exception e) {
			e.printStackTrace();
			throw new Exception("Erro ao instanciar um objeto da classe OOoBean, mensagem interna: " + e.getMessage());
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
	
	public void loadFromURL(String url) {
		try {
			Object arProp = Array.newInstance(classLoader.loadClass("com.sun.star.beans.PropertyValue"), 1);

            Method methLoad = beanClass.getMethod("loadFromURL", new Class[] { String.class, arProp.getClass() });

            methLoad.invoke(bean, new Object[] { url, null });
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
		
	public void execute(String cmd, Object[] propertyValues) throws Exception {
		
		try {
			/*
			com.sun.star.frame.XDispatchHelper
			com.sun.star.frame.XDispatchProvider
			com.sun.star.frame.XFrame
			com.sun.star.uno.UnoRuntime
			com.sun.star.uno.XComponentContext
			
			XComponentContext xCc = o3Bean.getOOoConnection().getComponentContext();
			XFrame xFrame = o3Bean.getDocument().getCurrentController().getFrame();
			Object dispatchHelperObject = xCc.getServiceManager().createInstanceWithContext("com.sun.star.frame.DispatchHelper", xCc);
			XDispatchHelper xDh = (XDispatchHelper) UnoRuntime.queryInterface(XDispatchHelper.class, dispatchHelperObject);
			XDispatchProvider xDispatchProvider = (XDispatchProvider) UnoRuntime.queryInterface(XDispatchProvider.class, xFrame);
			XWindow xWindow = xFrame.getComponentWindow();

			xWindow.setFocus();
			xDh.executeDispatch(xDispatchProvider, cmd, "", 0, propertyValue);
			*/
			
			Class<?> componentContextClass = classLoader.loadClass("com.sun.star.uno.XComponentContext");
			Class<?> dispatchHelperClass = classLoader.loadClass("com.sun.star.frame.XDispatchHelper");
			Class<?> dispatchProviderClass = classLoader.loadClass("com.sun.star.frame.XDispatchProvider");
			Class<?> unoRuntimeClass = classLoader.loadClass("com.sun.star.uno.UnoRuntime");
			
			Object xCc = invoke(invoke(bean, "getOOoConnection"), "getComponentContext");					
			Object xFrame = invoke(invoke(getDocument(),"getCurrentController"),"getFrame");
			Object sm = invoke(xCc, "getServiceManager");
			Object dispatchHelperObject = sm.getClass().getMethod("createInstanceWithContext", String.class, componentContextClass).invoke(sm, "com.sun.star.frame.DispatchHelper", xCc);
					
			Method urQueryInterfaceMethod = unoRuntimeClass.getMethod("queryInterface", Class.class, Object.class);
			
			Object xDh = urQueryInterfaceMethod.invoke(null, dispatchHelperClass, dispatchHelperObject);
			Object xDispatchProvider = unoRuntimeClass.getMethod("queryInterface", Class.class, Object.class).invoke(null, dispatchProviderClass, xFrame);
			Object xWindow = invoke(xFrame, "getComponentWindow");

			invoke(xWindow, "setFocus");
			invoke(xDh, "executeDispatch", xDispatchProvider, cmd, "", 0, propertyValues);
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
}