package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
		
		InstallationConfigs.iniciar();
		
		try {
			final JLibreOffice jLibreOffice = new JLibreOffice(InstallationConfigs.getInstance().getClassLoader());
			
			JFrame frame = new JFrame();
			frame.setTitle("Teste");
			frame.setSize(800, 600);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);            

			initPnButtons(frame, jLibreOffice);
						
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

	private static void initPnButtons(JFrame frame, final JLibreOffice jLibreOffice) {
		
		JPanel pnButtons = new JPanel(new FlowLayout());
		
		JButton btnSalvar = new JButton("Salvar");
		btnSalvar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					jLibreOffice.save();
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
		
		JButton btnSalvarAs = new JButton("SalvarAs");
		btnSalvarAs.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					jLibreOffice.saveAs();
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
			}
		});

		JButton btnOpen = new JButton("Open");
		btnOpen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					File f = new File("c:/teste.odt");					
					jLibreOffice.open(f.toURL().toString());
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
		

		JButton btnPdf = new JButton("PDF");
		btnPdf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					jLibreOffice.open("file:/home/pablo-moreira/Desktop/teste.odt");
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
		
		pnButtons.add(btnOpen);
		pnButtons.add(btnSalvar);
		pnButtons.add(btnSalvarAs);
		frame.add(pnButtons, BorderLayout.NORTH);	
	}
}
