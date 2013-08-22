package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
	
		InstallationConfigs.iniciar();
		
		try {
			final JLibreOffice jLibreOffice = new JLibreOffice(InstallationConfigs.getInstance().getClassLoader());
			
			JFrame frame = new JFrame();
			frame.setTitle("Teste");
			frame.setSize(800, 600);
			frame.addWindowListener(new WindowListener() {
				
				public void windowOpened(WindowEvent e) {		
				}
				
				public void windowIconified(WindowEvent e) {
				}
				
				public void windowDeiconified(WindowEvent e) {
				}
				
				public void windowDeactivated(WindowEvent e) {				
				}
				
				public void windowClosing(WindowEvent e) {					
				}
				
				public void windowClosed(WindowEvent e) {
					jLibreOffice.closeConnection();
					System.exit(0);
				}
				
				public void windowActivated(WindowEvent e) {
				}
			});
            
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
