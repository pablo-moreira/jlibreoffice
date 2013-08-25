package com.googlecode.jlibreoffice.menu;

import java.awt.Menu;
import java.awt.MenuBar;
import java.awt.MenuItem;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ResourceBundle;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import com.googlecode.jlibreoffice.JLibreOffice;

public class MenuBuilder {

	private JLibreOffice jLibreOffice;
	
	public MenuBuilder(JLibreOffice jLibreOffice) {
		super();
		this.jLibreOffice = jLibreOffice;
	}
	
	public JLibreOffice getJLibreOffice() {
		return jLibreOffice;
	}

	private String getMessageBundleLabel(String name) {
		return getMessageBundle().getString("viewer.menu." + name + ".label");
	}
	
	private Menu buildMenu(String name) {
    	
		Menu menu = new Menu(getMessageBundleLabel(name));
		    	
    	return menu;
	}
	
	private MenuItem buildMenuItem(String name, final String comando) {
		
		MenuItem menuItem = new MenuItem(getMessageBundleLabel(name));	
		menuItem.addActionListener(new ActionListener(){	
			public void actionPerformed(ActionEvent evt) {
				try {
					getJLibreOffice().execute(comando, null);
				}
				catch (Exception e) {
					e.printStackTrace();
					JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(),"jLibreOffice - Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		
		return menuItem;
	}
	
	public void buildMenuBar(MenuBar menuBar) {
		menuBar.add(buildEditMenu());
		menuBar.add(buildInsertMenu());
		menuBar.add(buildFormatMenu());
		menuBar.add(buildTableMenu());
	}
	
	public MenuBar buildCompleteMenuBar() {
		
		MenuBar menuBar = new MenuBar();
		
		menuBar.add(buildFileMenu());
		menuBar.add(buildInsertMenu());
		menuBar.add(buildFormatMenu());
		menuBar.add(buildTableMenu());
		
		return menuBar;
	}
	
	public Menu buildFileMenu() {
		
		Menu menu = buildMenu("file");
		
		menu.add(buildNewWriterMenuItem());
		menu.add(buildOpenMenuItem());
		menu.add(buildSaveMenuItem());
		menu.add(buildSaveAsMenuItem());
		
		return menu;
	}

	public Menu buildEditMenu() {
		
    	Menu menu = buildMenu("edit");
    	
    	menu.add(buildMenuItem("edit.undo", JLibreOffice.UNO_UNDO));
    	menu.add(buildMenuItem("edit.redo", JLibreOffice.UNO_REDO));
    	menu.add(buildMenuItem("edit.repeat", JLibreOffice.UNO_REPEAT));
    	menu.addSeparator();
    	menu.add(buildMenuItem("edit.cut", JLibreOffice.UNO_CUT));
    	menu.add(buildMenuItem("edit.copy", JLibreOffice.UNO_COPY));
    	menu.add(buildMenuItem("edit.paste", JLibreOffice.UNO_PASTE));
    	menu.add(buildMenuItem("edit.pasteSpecial", JLibreOffice.UNO_PASTE_SPECIAL));
    	menu.add(buildMenuItem("edit.selectTextMode", JLibreOffice.UNO_SELECT_TEXT_MODE));
    	menu.add(buildMenuItem("edit.selectAll", JLibreOffice.UNO_SELECT_ALL));
    	
    	return menu;
	}
	
	public Menu buildInsertMenu() {
		
    	Menu menu = buildMenu("insert");
    	
    	menu.add(buildMenuItem("insert.insertBreak", JLibreOffice.UNO_INSERT_BREAK));
    			
			Menu menuFields = buildMenu("insert.fields");
	
			menuFields.add(buildMenuItem("insert.fields.insertDateField", JLibreOffice.UNO_INSERT_DATE_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertTimeField", JLibreOffice.UNO_INSERT_TIME_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertPageNumberField", JLibreOffice.UNO_INSERT_PAGE_NUMBER_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertPageCountField", JLibreOffice.UNO_INSERT_PAGE_COUNT_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertTopicField", JLibreOffice.UNO_INSERT_TOPIC_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertTitleField", JLibreOffice.UNO_INSERT_TITLE_FIELD));
			menuFields.add(buildMenuItem("insert.fields.insertAuthorField", JLibreOffice.UNO_INSERT_AUTHOR_FIELD));
			menuFields.addSeparator();
			menuFields.add(buildMenuItem("insert.fields.insertField", JLibreOffice.UNO_INSERT_FIELD));

			menu.add(menuFields);
			
		menu.add(buildMenuItem("insert.insertSymbol", JLibreOffice.UNO_INSERT_SYMBOL));		
		menu.addSeparator();
		menu.add(buildMenuItem("insert.insertSection", JLibreOffice.UNO_INSERT_SECTION));
		menu.add(buildMenuItem("insert.hyperlinkDialog", JLibreOffice.UNO_HYPERLINK_DIALOG));
		menu.addSeparator();
		menu.add(buildMenuItem("insert.headerPageDialog", JLibreOffice.UNO_PAGE_DIALOG));
		menu.add(buildMenuItem("insert.footerPageDialog", JLibreOffice.UNO_PAGE_DIALOG));
		menu.add(buildMenuItem("insert.insertFootnoteDialog", JLibreOffice.UNO_INSERT_FOOTNOTE_DIALOG));
		menu.add(buildMenuItem("insert.insertBookmark", JLibreOffice.UNO_INSERT_BOOKMARK));
		menu.add(buildMenuItem("insert.insertReferenceField", JLibreOffice.UNO_INSERT_REFERENCE_FIELD));
		menu.add(buildMenuItem("insert.insertAnnotation", JLibreOffice.UNO_INSERT_ANNOTATION));
		menu.add(buildMenuItem("insert.insertScript", JLibreOffice.UNO_INSERT_SCRIPT));
		
			Menu menuIndexes = buildMenu("insert.indexes");
			
			menuIndexes.add(buildMenuItem("insert.indexes.insertIndexesEntry", JLibreOffice.UNO_INSERT_INDEXES_ENTRY));
			menuIndexes.add(buildMenuItem("insert.indexes.insertMultiIndex", JLibreOffice.UNO_INSERT_MULTI_INDEX));
			menuIndexes.add(buildMenuItem("insert.indexes.insertAuthoritiesEntry", JLibreOffice.UNO_INSERT_AUTHORITIES_ENRTY));

			menu.add(menuIndexes);
			
		menu.addSeparator();
		menu.add(buildMenuItem("insert.insertEnvelope", JLibreOffice.UNO_INSERT_ENVELOPE));
		menu.addSeparator();
		menu.add(buildMenuItem("insert.insertFrame", JLibreOffice.UNO_INSERT_FRAME));
		menu.add(buildMenuItem("insert.insertTable", JLibreOffice.UNO_INSERT_TABLE));
		menu.add(buildMenuItem("insert.insertGraphicRuler", JLibreOffice.UNO_INSERT_GRAPHIC_RULER));
		
			Menu menuGraphic = buildMenu("insert.graphic");
			
			menuGraphic.add(buildMenuItem("insert.graphic.insertGraphic", JLibreOffice.UNO_INSERT_GRAPHIC));
				
				Menu menuGraphicTwain = buildMenu("insert.graphic.twain");
				
				menuGraphicTwain.add(buildMenuItem("insert.graphic.twain.twainSelect", JLibreOffice.UNO_TWAIN_SELECT));
				menuGraphicTwain.add(buildMenuItem("insert.graphic.twain.twainTransfer", JLibreOffice.UNO_TWAIN_TRANSFER));
				
				menuGraphic.add(menuGraphicTwain);
	
			menu.add(menuGraphic);
			
		menu.add(buildMenuItem("insert.insertAVMedia", JLibreOffice.UNO_INSERT_AVMEDIA));
		
			Menu menuObjeto = buildMenu("insert.object");
			
			menuObjeto.add(buildMenuItem("insert.object.insertObject", JLibreOffice.UNO_INSERT_OBJECT));
			menuObjeto.add(buildMenuItem("insert.object.insertPlugin", JLibreOffice.UNO_INSERT_PLUGIN));
			menuObjeto.add(buildMenuItem("insert.object.insertSound", JLibreOffice.UNO_INSERT_SOUND));
			menuObjeto.add(buildMenuItem("insert.object.insertVactioneo", JLibreOffice.UNO_INSERT_VACTIONEO));
			menuObjeto.add(buildMenuItem("insert.object.insertApplet", JLibreOffice.UNO_INSERT_APPLET));
			menuObjeto.add(buildMenuItem("insert.object.insertObjectStarMath", JLibreOffice.UNO_INSERT_OBJECT_STAR_MATH));
			menuObjeto.add(buildMenuItem("insert.object.insertObjectChart", JLibreOffice.UNO_INSERT_OBJECT_CHART));

			menu.add(menuObjeto);
			
		menu.add(buildMenuItem("insert.insertObjectFloatingFrame", JLibreOffice.UNO_INSERT_OBJECT_FLOATING_FRAME));
		menu.addSeparator();
		menu.add(buildMenuItem("insert.insertDoc", JLibreOffice.UNO_INSERT_DOC));
		
		return menu;
	}
	
	
    public Menu buildTableMenu() {
    	
    	Menu menu = buildMenu("table");
    	
    		Menu menuInserir = buildMenu("table.insert");
    		
    		menuInserir.add(buildMenuItem("table.insert.insertTable", JLibreOffice.UNO_INSERT_TABLE));
    		menuInserir.add(buildMenuItem("table.insert.insertRowDialog", JLibreOffice.UNO_INSERT_ROW_DIALOG));
    		menuInserir.add(buildMenuItem("table.insert.insertColumnDialog", JLibreOffice.UNO_INSERT_COLUMN_DIALOG));    	
    		
    	menu.add(menuInserir);    				
		
    		Menu menuExcluir = buildMenu("table.delete");
			
    		menuExcluir.add(buildMenuItem("table.delete.deleteTable", JLibreOffice.UNO_DELETE_TABLE));
			menuExcluir.add(buildMenuItem("table.delete.deleteRows", JLibreOffice.UNO_DELETE_ROWS));
			menuExcluir.add(buildMenuItem("table.delete.deleteColumns", JLibreOffice.UNO_DELETE_COLUMNS));
	
		menu.add(menuExcluir);	
		
			Menu menuSelecionar = buildMenu("table.select");
			
			menuSelecionar.add(buildMenuItem("table.select.selectTable", JLibreOffice.UNO_SELECT_TABLE));
			menuSelecionar.add(buildMenuItem("table.select.entireRow", JLibreOffice.UNO_ENTIRE_ROW));
			menuSelecionar.add(buildMenuItem("table.select.entireColumn", JLibreOffice.UNO_ENTIRE_COLUMN));
			menuSelecionar.add(buildMenuItem("table.select.entireCell", JLibreOffice.UNO_ENTIRE_CELL));

		menu.add(menuSelecionar);
		
    	menu.add(buildMenuItem("table.mergeCells", JLibreOffice.UNO_MERGE_CELLS));
    	menu.add(buildMenuItem("table.splitCell", JLibreOffice.UNO_SPLIT_CELL));
    	menu.add(buildMenuItem("table.protect", JLibreOffice.UNO_PROTECT));
    	menu.add(buildMenuItem("table.mergeTable", JLibreOffice.UNO_MERGE_TABLE));
    	menu.add(buildMenuItem("table.splitTable", JLibreOffice.UNO_SPLIT_TABLE));
    	menu.addSeparator();
    	menu.add(buildMenuItem("table.tableDialog", JLibreOffice.UNO_TABLE_DIALOG));
    	
    	return menu;
    }
       
    public MenuItem buildMenuItem(String name, ActionListener actionListener) {
    
    	MenuItem menuItem = new MenuItem(getMessageBundleLabel(name));
    	menuItem.addActionListener(actionListener);
		
		return menuItem;
    }
    	
	public MenuItem buildNewWriterMenuItem() {
		return buildMenuItem("file.new.writer", new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					getJLibreOffice().newWriterDocument();
				} 
				catch (Exception e) {
					e.printStackTrace();
					JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(),"jLibreOffice - Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
	}
    
    public MenuItem buildSaveAsMenuItem() {
    	return buildMenuItem("file.saveAs", new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					getJLibreOffice().saveAs();
				} 
				catch (Exception e) {
					e.printStackTrace();
					JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(),"jLibreOffice - Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
    }
    
    public MenuItem buildSaveMenuItem() {
    	return buildMenuItem("file.save", new ActionListener(){	
			public void actionPerformed(ActionEvent evt) {
				try {
					getJLibreOffice().save();
				}
				catch (Exception e) {
					e.printStackTrace();
					JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(),"jLibreOffice - Erro", JOptionPane.ERROR_MESSAGE);
				}			
			}
		});		
    }
    
    public MenuItem buildOpenMenuItem() {
    	return buildMenuItem("file.open", new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try { 				
					JFileChooser fc = new JFileChooser();
					
					fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
					
					int retorno = fc.showOpenDialog(null);				
					
					if(retorno == JFileChooser.APPROVE_OPTION) {
					
						File file = fc.getSelectedFile();
								
						String url = "file:///" + file.getAbsolutePath().replace('\\', '/');
						
						getJLibreOffice().open(url);
					}
				} 
				catch (Exception e) {					
					e.printStackTrace();
					JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(),"jLibreOffice - Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
    }    
    
    public Menu buildFormatMenu() {
    	
    	Menu menu = buildMenu("format");
    	
    	menu.add(buildMenuItem("format.resetAttributes", JLibreOffice.UNO_RESET_ATTRIBUTES));
    	menu.addSeparator();
		menu.add(buildMenuItem("format.fontDialog", JLibreOffice.UNO_FONT_DIALOG));
		menu.add(buildMenuItem("format.paragraphDialog", JLibreOffice.UNO_PARAGRAPH_DIALOG));
		menu.add(buildMenuItem("format.bulletsAndNumberingDialog", JLibreOffice.UNO_BULLETS_AND_NUMBERING_DIALOG));
		menu.add(buildMenuItem("format.pageDialog", JLibreOffice.UNO_PAGE_DIALOG));
		menu.addSeparator();
		menu.add(buildMenuItem("format.formatColumns", JLibreOffice.UNO_FORMAT_COLUMNS));
		menu.add(buildMenuItem("format.editRegion", JLibreOffice.UNO_EDIT_REGION));
		menu.addSeparator();
		menu.add(buildMenuItem("format.designerDialog", JLibreOffice.UNO_DESIGNER_DIALOG));
		
			Menu menuAutoFormat = buildMenu("format.autoFormat");
			
			menuAutoFormat.add(buildMenuItem("format.autoFormat.onlineAutoFormat", JLibreOffice.UNO_ONLINE_AUTO_FORMAT));
			menuAutoFormat.add(buildMenuItem("format.autoFormat.autoFormatApply", JLibreOffice.UNO_AUTO_FORMAT_APPLY));
			menuAutoFormat.add(buildMenuItem("format.autoFormat.autoFormatRedlineApply", JLibreOffice.UNO_AUTO_FORMAT_REDLINE_APPLY));
			
		menu.add(menuAutoFormat);
		menu.addSeparator();
			
			Menu menuAlign = buildMenu("format.align");
			
			menuAlign.add(buildMenuItem("format.align.commonAlignLeft", JLibreOffice.UNO_COMMON_ALIGN_LEFT));
			menuAlign.add(buildMenuItem("format.align.commonAlignHorizontalCenter", JLibreOffice.UNO_COMMON_ALIGN_HORIZONTAL_CENTER));
			menuAlign.add(buildMenuItem("format.align.commonAlignRight", JLibreOffice.UNO_COMMON_ALIGN_RIGHT));
			menuAlign.add(buildMenuItem("format.align.commonAlignJustified", JLibreOffice.UNO_COMMON_ALIGN_JUSTIFIED));
			menuAlign.addSeparator();		
			menuAlign.add(buildMenuItem("format.align.commonAlignTop", JLibreOffice.UNO_COMMON_ALIGN_TOP));
			menuAlign.add(buildMenuItem("format.align.commonAlignVerticalCenter", JLibreOffice.UNO_COMMON_ALIGN_VERTICAL_CENTER));
			menuAlign.add(buildMenuItem("format.align.commonAlignBottom", JLibreOffice.UNO_COMMON_ALIGN_BOTTOM));
		
		menu.add(menuAlign);
			
		return menu;
    }
    
	public ResourceBundle getMessageBundle() {
		return jLibreOffice.getMessageBundle();
	}    
}
