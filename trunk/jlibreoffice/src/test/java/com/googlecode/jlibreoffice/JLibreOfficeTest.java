package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
	
		InstallationConfigs.iniciar();
		
		try {
			JLibreOffice jLibreOffice = new JLibreOffice(InstallationConfigs.getInstance().getClassLoader());
			
			JFrame frame = new JFrame();
			frame.setTitle("Teste");
			frame.setSize(800, 600);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            
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
