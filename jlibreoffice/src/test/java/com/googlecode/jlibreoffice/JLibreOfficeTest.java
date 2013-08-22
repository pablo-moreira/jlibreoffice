package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.io.File;
import java.net.URL;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;
import com.googlecode.jlibreoffice.util.CustomURLClassLoader;
import com.googlecode.jlibreoffice.util.SystemUtils;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
	
		InstallationConfigs.iniciar();
				
		JFrame frame = new JFrame();
		frame.setTitle("Teste");
		frame.setSize(800, 600);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		try {
			String unoPath = InstallationConfigs.getInstance().getUnoPath();
			String s = unoPath.replace("program", "");
			
            System.out.println("sun.awt.noxembed: " + System.getProperty("sun.awt.noxembed"));
            System.setProperty("sun.awt.xembedserver", "true");

            File f = new File(s);
            URL url = f.toURI().toURL();
            String officeURL = url.toString();
            
            URL[] arURL;
            
            if (SystemUtils.isOsWindows()) {
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
		            new URL(officeURL + "/program/classes/ridl.jar"),
		            new URL(officeURL + "/ure-link/share/java/jurt.jar"),		            
		            new URL(officeURL + "/ure-link/share/java/unoil.jar"),
		            new URL(officeURL + "/ure-link/share/java/java_uno.jar"),
		            new URL(officeURL + "/ure-link/share/java/juh.jar")
                };
            }
            //CustomURLClassLoader cl = new CustomURLClassLoader(arURL, InstallationConfigs.getInstance().getClassLoader());
            //cl.addResourcePath(new File(unoPath).toURI().toURL());
            
			
            JLibreOffice jLibreOffice = new JLibreOffice(InstallationConfigs.getInstance().getClassLoader());
            
            frame.add(jLibreOffice.getBean().getContainer(), BorderLayout.CENTER);            
            frame.setVisible(true);
            
            jLibreOffice.newWriterDocument();
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
