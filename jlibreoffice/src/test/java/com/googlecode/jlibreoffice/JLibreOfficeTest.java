package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.MenuBar;
import java.awt.MenuItem;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.JLibreOfficeInstallation;
import com.googlecode.jlibreoffice.menu.MenuBuilder;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
		
		try {
			JLibreOfficeInstallation.iniciar();
									
			final JLibreOffice jLibreOffice = new JLibreOffice(JLibreOfficeInstallation.getInstance().getClassLoader());
			jLibreOffice.setMenuBarVisible(false);
			jLibreOffice.setToolBarVisible(false);
			jLibreOffice.setStandardBarVisible(false);
		
			JFrame frame = new JFrame();
			frame.setTitle("JLibreOffice");
			frame.setSize(800, 600);
			frame.addWindowListener(getWindowListener(jLibreOffice));
		
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
	
	private static WindowListener getWindowListener(final JLibreOffice jLibreOffice) {
		return new WindowListener() {
			
			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void windowClosing(WindowEvent arg0) {
				new Thread(new Runnable() {
					
					@Override
					public void run() {
						jLibreOffice.closeConnection();
					}
				}).start();				
			}
			
			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub
				
			}
		};
	}
}
