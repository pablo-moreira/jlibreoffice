package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.PropertyResourceBundle;
import java.util.ResourceBundle;

import org.apache.log4j.Logger;

import com.googlecode.jlibreoffice.util.SystemUtils;

public class JLibreOffice {
	
	private static final Logger log = Logger.getLogger(JLibreOffice.class);
	
	public static final String NEW_WRITER = "private:factory/swriter";
	public static final String UNO_AUTO_FORMAT_APPLY = ".uno:AutoFormatApply";
	public static final String UNO_AUTO_FORMAT_REDLINE_APPLY = ".uno:AutoFormatRedlineApply";
	public static final String UNO_BULLETS_AND_NUMBERING_DIALOG = ".uno:BulletsAndNumberingDialog";
	public static final String UNO_COMMON_ALIGN_BOTTOM = ".uno:CommonAlignBottom";
	public static final String UNO_COMMON_ALIGN_HORIZONTAL_CENTER = ".uno:CommonAlignHorizontalCenter";
	public static final String UNO_COMMON_ALIGN_JUSTIFIED = ".uno:CommonAlignJustified";
	public static final String UNO_COMMON_ALIGN_LEFT = ".uno:CommonAlignLeft";
	public static final String UNO_COMMON_ALIGN_RIGHT = ".uno:CommonAlignRight";			
	public static final String UNO_COMMON_ALIGN_TOP = ".uno:CommonAlignTop";
	public static final String UNO_COMMON_ALIGN_VERTICAL_CENTER = ".uno:CommonAlignVerticalCenter";
	public static final String UNO_COPY =  ".uno:Copy";
	public static final String UNO_CUT =  ".uno:Cut";
	public static final String UNO_DELETE_COLUMNS = ".uno:DeleteColumns";
	public static final String UNO_DELETE_ROWS = ".uno:DeleteRows";
	public static final String UNO_DELETE_TABLE = ".uno:DeleteTable";
	public static final String UNO_DESIGNER_DIALOG = ".uno:DesignerDialog";	
	public static final String UNO_EDIT_REGION = ".uno:EditRegion";
	public static final String UNO_ENTIRE_CELL = ".uno:EntireCell";
	public static final String UNO_ENTIRE_COLUMN = ".uno:EntireColumn";
	public static final String UNO_ENTIRE_ROW = ".uno:EntireRow";
	public static final String UNO_FONT_DIALOG = ".uno:FontDialog";
	public static final String UNO_FORMAT_COLUMNS = ".uno:FormatColumns";
	public static final String UNO_HYPERLINK_DIALOG =  ".uno:HyperlinkDialog";
	public static final String UNO_INSERT_ANNOTATION =  ".uno:InsertAnnotation";
	public static final String UNO_INSERT_APPLET = ".uno:InsertApplet";
	public static final String UNO_INSERT_AUTHOR_FIELD =  ".uno:InsertAuthorField";
	public static final String UNO_INSERT_AUTHORITIES_ENRTY = ".uno:InsertAuthoritiesEntry";
	public static final String UNO_INSERT_AVMEDIA =  ".uno:InsertAVMedia";
	public static final String UNO_INSERT_BOOKMARK =  ".uno:InsertBookmark";
	public static final String UNO_INSERT_BREAK =  ".uno:InsertBreak";
	public static final String UNO_INSERT_COLUMN_DIALOG = ".uno:InsertColumnDialog";
	public static final String UNO_INSERT_DATE_FIELD =  ".uno:InsertDateField";
	public static final String UNO_INSERT_DOC = ".uno:InsertDoc";
	public static final String UNO_INSERT_ENVELOPE = ".uno:InsertEnvelope";
	public static final String UNO_INSERT_FIELD =  ".uno:InsertField";	
	public static final String UNO_INSERT_FOOTNOTE_DIALOG =  ".uno:InsertFootnoteDialog";
	public static final String UNO_INSERT_FRAME =  ".uno:InsertFrame";
	public static final String UNO_INSERT_GRAPHIC =  ".uno:InsertGraphic";
	public static final String UNO_INSERT_GRAPHIC_RULER =  ".uno:InsertGraphicRuler";
	public static final String UNO_INSERT_INDEXES_ENTRY = ".uno:InsertIndexesEntry";
	public static final String UNO_INSERT_MULTI_INDEX = ".uno:InsertMultiIndex";
	public static final String UNO_INSERT_OBJECT =  ".uno:InsertObject";
	public static final String UNO_INSERT_OBJECT_CHART = ".uno:InsertObjectChart";
	public static final String UNO_INSERT_OBJECT_FLOATING_FRAME = ".uno:InsertObjectFloatingFrame";	
	public static final String UNO_INSERT_OBJECT_STAR_MATH = ".uno:InsertObjectStarMath";
	public static final String UNO_INSERT_PAGE_COUNT_FIELD =  ".uno:InsertPageCountField";
	public static final String UNO_INSERT_PAGE_NUMBER_FIELD =  ".uno:InsertPageNumberField";
	public static final String UNO_INSERT_PLUGIN =  ".uno:InsertPlugin";
	public static final String UNO_INSERT_REFERENCE_FIELD =  ".uno:InsertReferenceField";
	public static final String UNO_INSERT_ROW_DIALOG = ".uno:InsertRowDialog";
	public static final String UNO_INSERT_SCRIPT = ".uno:InsertScript";
	public static final String UNO_INSERT_SECTION =  ".uno:InsertSection";
	public static final String UNO_INSERT_SOUND =  ".uno:InsertSound";
	public static final String UNO_INSERT_SYMBOL =  ".uno:InsertSymbol";
	public static final String UNO_INSERT_TABLE = ".uno:InsertTable";
	public static final String UNO_INSERT_TIME_FIELD = ".uno:InsertTimeField";
	public static final String UNO_INSERT_TITLE_FIELD =  ".uno:InsertTitleField";
	public static final String UNO_INSERT_TOPIC_FIELD =  ".uno:InsertTopicField";
	public static final String UNO_INSERT_VACTIONEO =  ".uno:InsertVactioneo";
	public static final String UNO_MERGE_CELLS = ".uno:MergeCells";
	public static final String UNO_MERGE_TABLE = ".uno:MergeTable";
	public static final String UNO_ONLINE_AUTO_FORMAT = ".uno:OnlineAutoFormat";			
	public static final String UNO_PAGE_DIALOG = ".uno:PageDialog";
	public static final String UNO_PARAGRAPH_DIALOG = ".uno:ParagraphDialog";
	public static final String UNO_PASTE =  ".uno:Paste";
	public static final String UNO_PASTE_SPECIAL =  ".uno:PasteSpecial";
	public static final String UNO_PROTECT = ".uno:Protect";
	public static final String UNO_REDO =  ".uno:Redo";
	public static final String UNO_REPEAT =  ".uno:Repeat";
	public static final String UNO_RESET_ATTRIBUTES = ".uno:ResetAttributes";
	public static final String UNO_SAVE = ".uno:Save";
	public static final String UNO_SAVE_AS = ".uno:SaveAs";
	public static final String UNO_SELECT_ALL =  ".uno:SelectAll";
	public static final String UNO_SELECT_TABLE = ".uno:SelectTable";
	public static final String UNO_SELECT_TEXT_MODE =  ".uno:SelectTextMode";
	public static final String UNO_SPLIT_CELL = ".uno:SplitCell";
	public static final String UNO_SPLIT_TABLE = ".uno:SplitTable";
	public static final String UNO_TABLE_DIALOG =  ".uno:TableDialog";
	public static final String UNO_TWAIN_SELECT =  ".uno:TwainSelect";
	public static final String UNO_TWAIN_TRANSFER =  ".uno:TwainTransfer";
	public static final String UNO_UNDO =  ".uno:Undo";
	public static final String UNO_EXPORT_DIRECT_TO_PDF = ".uno:ExportDirectToPDF";

	private OOoBeanProxy bean;
	private ResourceBundle messageBundle;

	public JLibreOffice(ClassLoader classLoader) {		
		initialize(classLoader);
	}
		
	private void initialize(ClassLoader classLoader) {
		
		try {
			// !!! Importante - Faz nao aparecer tela de restauracao de arquivos
			if (SystemUtils.isOsLinux()) {
				System.setProperty("com.sun.star.officebean.Options", "--norestore");
			}
			else {
				System.setProperty("com.sun.star.officebean.Options", "-norestore");
			}
			
			try {
				messageBundle = ResourceBundle.getBundle(JLibreOfficeConstants.DEFAULT_MESSAGE_BUNDLE);
			}
			catch (Exception e) {				
				log.error(e);
				try {
					InputStream is = getClass().getResourceAsStream("/com/googlecode/jlibreoffice/resources/MessageBundle.properties");				
					InputStreamReader reader = new InputStreamReader(is, Charset.forName("UTF-8"));
					messageBundle = new PropertyResourceBundle(reader);
				}
				catch (Exception e2) {
					log.error(e2);
				}
			}
			
			bean = new OOoBeanProxy(classLoader);
			
			bean.setLayout(new BorderLayout());
		
			log.info("A conexao com o LibreOffice foi carregada com sucesso!");
		}
		catch (Exception e) {
			bean = null;
			log.error("Erro ao iniciar a conexao com o LibreOffice, mensagem interna: " + (e.getMessage() != null ? e.getMessage() : e.getClass().getSimpleName()));
		}
	}

	public void open(File arquivo) throws Exception {
		
		try {
			FileInputStream fis = new FileInputStream(arquivo);			
			bean.loadFromStream(fis);       
            bean.aquireSystemWindow();
		}
		catch (Exception e)	{
			log.error("JLibreOffice.open: " + e.getMessage());
			e.printStackTrace();			
			throw new Exception(e.getMessage());
		}
	}
	
	public void open(String url) throws Exception {

		try {
            bean.loadFromURL(url);         
            bean.aquireSystemWindow();
		}
		catch (Exception e)	{
			log.error("JLibreOffice.open: " + e.getMessage());
			e.printStackTrace();			
			throw new Exception(e.getMessage());
		}
	}

	public String exportToPdf() throws Exception {
		return bean.exportToPdf();
	}
	
	public void execute(String cmd) throws Exception {
		bean.execute(cmd, null);
	}
	
	public void execute(String cmd, Object[] propertyValues) throws Exception {
		bean.execute(cmd, propertyValues);
	}
	
	public OOoBeanProxy getBean() {
		return bean;
	}
	
	public void closeConnection() {
		if (bean != null) {
			if (bean.isOOoConnected()) {
				bean.stopOOoConnection();
				log.info("A conexao com o LibreOffice foi fechada com sucesso!");
			}
			bean.setVisible(false);
		}
	}
	
	public boolean isConnected() {
		if (isInitialized() && bean.isOOoConnected()) {
			return true;
		}
		else {
			return false;
		}
	}
	
	public boolean isEditingDocument() {		
		try {
			return bean.getDocument() != null;
		} 
		catch (Exception e) {
			return false;
		}
	}
	
	public boolean isInitialized() {
		return bean != null;
	}
	
	public void newWriterDocument() throws Exception {
		open(JLibreOffice.NEW_WRITER);
	}

	public void save() throws Exception {
		
		if (!isConnected()) {
			return;
		}
		execute(JLibreOffice.UNO_SAVE, null);	
	}

	public void saveAs() throws Exception {
		
		if (!isConnected()) {
			return;
		}
		execute(JLibreOffice.UNO_SAVE_AS, null);		
	}
	
	public Container getContainer() {
		return bean.getContainer();
	}

	public void stopOOoConnection() {
		
		if (!isConnected()) {
			return;
		}
		bean.stopOOoConnection();
	}
	
	

	public void setMenuBarVisible(boolean value) {
		bean.setMenuBarVisible(value);
	}

	public void setStandardBarVisible(boolean value) {
		bean.setStandardBarVisible(value);
	}

	public void setToolBarVisible(boolean value) {
		bean.setToolBarVisible(value);
	}

	public ResourceBundle getMessageBundle() {
		return messageBundle;
	}

	public void clear() {
		bean.clear();		
	}
}