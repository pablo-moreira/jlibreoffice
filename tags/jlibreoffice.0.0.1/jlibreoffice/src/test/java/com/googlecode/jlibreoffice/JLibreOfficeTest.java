package com.googlecode.jlibreoffice;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.JLibreOffice;


public class JLibreOfficeTest {
	
	public static void main(String[] args) throws Exception {
	
		JFrame frame = new JFrame();
		frame.setTitle("Teste");
		frame.setSize(800, 600);
		
		JLibreOffice jLibreOffice = new JLibreOffice(frame);
		
		try {
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
		
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		
	}

}
