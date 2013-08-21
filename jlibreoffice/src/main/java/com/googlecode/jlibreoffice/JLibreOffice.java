package com.googlecode.jlibreoffice;

import java.awt.BorderLayout;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;

import javax.swing.JFrame;

import com.googlecode.jlibreoffice.installation.InstallationConfigs;
import com.googlecode.jlibreoffice.util.SystemUtils;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class JLibreOffice {

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
	public static final String UNO_INSERT_TIPO_FIELD =  ".uno:InsertTimeField";
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
	public static final String UNO_TWAIN_SELECT =  ".uno:TwainSelect";
	public static final String UNO_TWAIN_TRANSFER =  ".uno:TwainTransfer";
	public static final String UNO_UNDO =  ".uno:Undo";
	public static final String UNO_EXPORT_DIRECT_TO_PDF = ".uno:ExportDirectToPDF";
	
	private OOoBeanProxy o3Bean;
	private JFrame frame;

	public JLibreOffice(JFrame frame) {
		this.frame = frame;
		initialize();
	}

	public void open(String url) throws Exception	{

		try {
            //Get the office process to load the URL
            o3Bean.loadFromURL(url, null);         
            o3Bean.aquireSystemWindow();
		}
		catch (Exception e)	{
			System.err.println("JLibreOffice.open:" + e.getMessage());
			e.printStackTrace();			
			throw new Exception(e.getMessage());
		}
	}

	public String exportToPdf() throws Exception {
		
		String urlFile = o3Bean.getDocument().getURL();
		String urlPdf = urlFile + ".pdf";

		PropertyValue[] propertyValue = new PropertyValue[3];
		propertyValue[0] = new PropertyValue();
		propertyValue[0].Name = "URL";
		propertyValue[0].Value = urlPdf;
		propertyValue[1] = new PropertyValue();
		propertyValue[1].Name = "FilterName";
		propertyValue[1].Value = "writer_pdf_Export";
		propertyValue[2] = new PropertyValue();
		propertyValue[2].Name = "SelectionOnly";
		propertyValue[2].Value = false;

		execute(JLibreOffice.UNO_EXPORT_DIRECT_TO_PDF, propertyValue);
			
		String path = urlPdf.substring(8, urlPdf.length());
		
		try {
			return URLDecoder.decode(path, "utf-8");
		} 
		catch (UnsupportedEncodingException e) {
			return path;
		}
	}
	
	@SuppressWarnings("deprecation")
	public void execute(String cmd, PropertyValue[] propertyValue) throws Exception {
		
		try {
			XComponentContext xCc = o3Bean.getOOoConnection().getComponentContext();
			XFrame xFrame = o3Bean.getDocument().getCurrentController().getFrame();
			Object dispatchHelperObject = xCc.getServiceManager().createInstanceWithContext("com.sun.star.frame.DispatchHelper", xCc);
			XDispatchHelper xDh = (XDispatchHelper) UnoRuntime.queryInterface(XDispatchHelper.class, dispatchHelperObject);
			XDispatchProvider xDispatchProvider = (XDispatchProvider) UnoRuntime.queryInterface(XDispatchProvider.class, xFrame);
			XWindow xWindow = xFrame.getComponentWindow();

			xWindow.setFocus();
			xDh.executeDispatch(xDispatchProvider, cmd, "", 0, propertyValue);
		}
		catch (NoConnectionException e) {
			throw new Exception(e.getMessage());			
		}
	}
	
	@SuppressWarnings("deprecation")
	public void closeConnection() {
		
		// Verifica se ela foi aberta
		if (o3Bean != null) {
			if (o3Bean.isOOoConnected()) {
				o3Bean.stopOOoConnection();
			}
			o3Bean.setVisible(false);
		}
	}
	
	public OOoBeanProxy getO3Bean() {
		return o3Bean;
	}
	
	@SuppressWarnings("deprecation")
	private void initialize() {

		System.out.println("TESTE - JAVA_UNO : "  + System.getenv("JAVA_UNO"));
		
		try {
			InstallationConfigs.iniciar();

			// !!! Importante - Faz nao aparecer tela de restauracao de arquivos
			System.setProperty("com.sun.star.officebean.Options", "--norestore");
			
			// !!! Importante - Corrige o bug do linux que nao deixa digitar
			if (SystemUtils.isOsLinux()) {
				System.out.println("\nSISTEMA OPERACIONAL: \n - Linux");
				//System.setProperty("sun.awt.xembedserver", "true");
			}
			else {
				System.out.println("\nSistema Operacional: \n - Windows");
			}         	
			
			o3Bean = new OOoBeanProxy();
			
			//o3Bean.setLayout(new BorderLayout());
			//o3Bean.setMenuBarVisible(false);
			//o3Bean.setStandardBarVisible(false);
			//o3Bean.setToolBarVisible(false);

			frame.add(o3Bean, BorderLayout.CENTER);
		
			System.out.println("\n------------------------------------------------------");
			System.out.println(" A conexao com o Open Office foi carregada com sucesso!");
			System.out.println("------------------------------------------------------");
		}
		catch (Exception e) {
			o3Bean = null;
			System.out.println("\n------------------------------------------------------");
			System.out.println(" Erro ao iniciar a conexão com o Open Office !");
			System.out.println("  - " + (e.getMessage() != null ? e.getMessage() : e.getClass().getSimpleName()));
			System.out.println("------------------------------------------------------");
			e.printStackTrace();
		}
	}
	
	@SuppressWarnings("deprecation")
	public boolean isConnected() {
		if (o3Bean != null && o3Bean.isOOoConnected()) {
			return true;
		}
		else {
			return false;
		}
	}
	
	public boolean isEditingDocument() {		
		try {
			return o3Bean.getDocument() != null;
		} 
		catch (NoConnectionException e) {
			return false;
		}
	}
	
	public boolean isInitialized() {
		return o3Bean != null;
	}
	
	public void newWriterDocument() throws Exception {
		open(JLibreOffice.NEW_WRITER);
	}

	public void save() throws Exception {
		
		// Verifica se tem alguma conexao
		if (!isConnected()) {
			return;
		}
		execute(JLibreOffice.UNO_SAVE, null);	
	}

	public void saveAs() throws Exception {
		
		// Verifica se tem alguma conexao
		if (!isConnected()) {
			return;
		}
		execute(JLibreOffice.UNO_SAVE_AS, null);		
	}
}
