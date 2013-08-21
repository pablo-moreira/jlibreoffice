package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.io.File;
import java.lang.reflect.Array;
import java.lang.reflect.Method;
import java.net.URL;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;
import com.googlecode.jlibreoffice.util.CustomURLClassLoader;


public class JLibreOfficeTestDireto {
	
	public static void main(String[] args) throws Exception {
	
		InstallationConfigs.iniciar();
				
		JFrame frame = new JFrame();
		frame.setTitle("Teste");
		frame.setSize(800, 600);
		
		try {
			String unoPath = InstallationConfigs.getInstance().getUnoPath();
			String s = unoPath.replace("program", "");
			
            System.out.println("sun.awt.noxembed: " + System.getProperty("sun.awt.noxembed"));
            System.setProperty("sun.awt.xembedserver", "true");

            File f = new File(s);
            URL url = f.toURI().toURL();
            String officeURL = url.toString();
            
            URL[] arURL;
            
            if (InstallationConfigs.isOsWindows()) {
            	arURL = new URL[] {            
		            new URL(officeURL + "program\\classes\\officebean.jar"),
		            new URL(officeURL + "program\\classes\\unoil.jar"),
		            new URL(officeURL + "URE\\java\\jurt.jar"),
		            new URL(officeURL + "URE\\java\\ridl.jar"),		            
		            new URL(officeURL + "URE\\java\\java_uno.jar"),
		            new URL(officeURL + "URE\\java\\juh.jar")
            	};
            }
            else {
            	arURL = new URL[] {            
		            new URL(officeURL + "/program/classes/officebean.jar"),
		            new URL(officeURL + "/program/classes/jurt.jar"),
		            new URL(officeURL + "/program/classes/ridl.jar"),
		            new URL(officeURL + "/program/classes/unoil.jar"),
		            new URL(officeURL + "/program/classes/java_uno.jar"),
		            new URL(officeURL + "/program/classes/juh.jar")
                };
            }
            CustomURLClassLoader m_loader = new CustomURLClassLoader(arURL, InstallationConfigs.getInstance().getClassLoader());
            File fileProg = new File(s + "/program");
            m_loader.addResourcePath(fileProg.toURI().toURL());
			
            Class beanClass = m_loader.loadClass("com.sun.star.comp.beans.OOoBean");
            Object m_objBean = beanClass.newInstance();
            
            frame.add((java.awt.Container)m_objBean, BorderLayout.CENTER);
            
            frame.setVisible(true);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            
			Object arProp = Array.newInstance(
                m_loader.loadClass("com.sun.star.beans.PropertyValue"), 1);
            Class<? extends Object> clazz = arProp.getClass();

            Method methLoad = beanClass.getMethod(
                "loadFromURL", new Class[] {
                    String.class, arProp.getClass() });

            methLoad.invoke(m_objBean, new Object[] {"private:factory/swriter", null});

            frame.validate();
			
			
			
		}
		catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		catch (Error e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		
	}
}
