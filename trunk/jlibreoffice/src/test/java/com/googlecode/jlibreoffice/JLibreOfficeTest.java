package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.MenuBar;
import java.awt.MenuItem;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.net.URLClassLoader;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;
import com.googlecode.jlibreoffice.menu.MenuBuilder;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
		
		InstallationConfigs.iniciar();
		URLClassLoader classLoader = InstallationConfigs.getInstance().getClassLoader();
		
		try {
			final JLibreOffice jLibreOffice = new JLibreOffice(classLoader);
			jLibreOffice.setMenuBarVisible(false);
			jLibreOffice.setToolBarVisible(false);
			jLibreOffice.setStandardBarVisible(false);
		
			JFrame frame = new JFrame();
			frame.setTitle("JLibreOffice");
			frame.setSize(800, 600);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
			MenuBuilder mb = new MenuBuilder(jLibreOffice);

			MenuBar menuBar = mb.buildCompleteMenuBar();
			
			MenuItem menuItem = new MenuItem("PDF");
			menuItem.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					try {
						jLibreOffice.exportToPdf();
					} 
					catch (Exception e) {
						e.printStackTrace();
					}
				}
			});
			menuBar.getMenu(0).add(menuItem);
			
			frame.setMenuBar(menuBar);
			
						
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
