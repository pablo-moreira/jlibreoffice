package com.googlecode.jlibreoffice;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.JLibreOffice;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
	
		JFrame frame = new JFrame();
		frame.setTitle("Teste");
		frame.setSize(800, 600);
		
		JLibreOffice jLibreOffice = new JLibreOffice(frame);
		
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		jLibreOffice.newWriterDocument();
	}

}
