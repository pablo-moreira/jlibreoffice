package com.googlecode.jlibreoffice.installation;

import java.io.File;
import java.net.URL;
import java.net.URLClassLoader;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import com.sun.star.lib.loader.InstallationFinder;

public class InstallationConfigs {

	private static InstallationConfigs instance;
	
	private static final String OS_LINUX = "Linux";	
	private static final String OS_WINDOWS = "Windows";
	private static final String PROPERTY_OS_NAME = "os.name";
			
	public static String getEnvVar( String envVar ) {        

        String path = null;
        
        try {
            path = System.getenv(envVar);
        } 
        catch ( SecurityException e ) {} 
        catch ( java.lang.Error err ) {}
        return path;
    }
		
	public static InstallationConfigs getInstance() {
		return instance;
	}	
	
	private static String getOsName() {
		return System.getProperty(PROPERTY_OS_NAME);
	}
	
	public static String getTmpDirPath() {		
		return System.getProperty("java.io.tmpdir");
	}
	    
    public static void iniciar() throws Exception {	

		// Verifica se consegue encontrar o caminho de instalacao
		String unoPath = InstallationFinder.getPath();
				
		// Verifica se encontrou o uno_path
		if (unoPath == null || unoPath.equals("")) {
			throw new Exception("O caminho para o diretório de instalação do LibreOffice não foi encontrado!");
		}
		
		System.out.println("\nUNO_PATH ENCONTRADO:\n - " + unoPath);
			
		instance = new InstallationConfigs(unoPath);
	}	
	
	public static boolean isOsLinux() {
		return getOsName().startsWith(OS_LINUX);
	}

	public static boolean isOsWindows() {
		return getOsName().startsWith(OS_WINDOWS);
	}

	private URLClassLoader classLoader;
	private String unoPath; 

	private InstallationConfigs(String unoPath) throws Exception {
		
		this.unoPath = unoPath;
		
		String unoPathRoot = new File(unoPath).getParent();
		
		// Verifica qual e o SO
		if (isOsWindows()) {

			List<LibraryPath> libs = new ArrayList<LibraryPath>(); 
			
			libs.add(new LibraryPath("msvcr70.dll", false));
			libs.add(new LibraryPath("msvcr71.dll", false));
			libs.add(new LibraryPath("uwinapi.dll", false));
			libs.add(new LibraryPath("jawt.dll", false));
			libs.add(new LibraryPath("officebean.dll", true));
			libs.add(new LibraryPath("sal3.dll", true));
			libs.add(new LibraryPath("jpipe.dll", true));
			
			List<File> dirs = new ArrayList<File>();
			
			dirs.add(new File(unoPathRoot + "\\program"));
			dirs.add(new File(unoPathRoot + "\\program\\classes"));
			dirs.add(new File(unoPathRoot + "\\URE\\bin"));
			dirs.add(new File(unoPathRoot + "\\Basis\\program"));			
			dirs.add(new File(unoPathRoot + "\\URE\\bin"));
			dirs.add(new File(unoPathRoot + "\\Basis\\program"));
			
			initializeClassloader(libs, dirs);
		}
		else if (isOsLinux()) {
						
			List<LibraryPath> libs = new ArrayList<LibraryPath>(); 
			
			libs.add(new LibraryPath("libofficebean.so", true));
			libs.add(new LibraryPath("libjpipe.so", true));
			libs.add(new LibraryPath("libjuh.so", true));

			List<File> dirs = new ArrayList<File>();

			dirs.add(new File(unoPathRoot + "/basis-link/program/"));
			dirs.add(new File(unoPathRoot + "/program/"));
			dirs.add(new File("/usr/lib/ure/lib/"));
			
			initializeClassloader(libs, dirs);
		}
		else {
			throw new Exception ("O JLibreOffice somente esta preparado para funcionar com os Sistemas Operacionais Windows e Linux Ubuntu!");
		}
	}
	
	private void initializeClassloader(List<LibraryPath> libs, List<File> dirs) throws Exception {

		Set<URL> urls = new HashSet<URL>();
		
		for (LibraryPath lib : libs) {

			URL found = null;
			
			for (File dir : dirs) {
				if (new File(dir, lib.getLibraryName()).exists()) {
					found = dir.toURI().toURL();
					break;
				}
			}
			
			if (found != null) {
				urls.add(found);
			}
			else if (lib.isRequired()) { 
				throw new Exception(MessageFormat.format("A biblioteca {0} não foi encontrada!", lib.getLibraryName()));
			}
		}

		this.classLoader = new URLClassLoader(urls.toArray(new URL[]{}), InstallationConfigs.class.getClassLoader());		
	}

	public URLClassLoader getClassLoader() {
		return classLoader;	
	}

	public String getUnoPath() {
    	return unoPath;
    }
}