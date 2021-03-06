package com.googlecode.jlibreoffice.installation;

import java.io.File;
import java.net.URL;
import java.net.URLClassLoader;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.log4j.Logger;

import com.googlecode.jlibreoffice.util.CustomURLClassLoader;
import com.googlecode.jlibreoffice.util.SystemUtils;
import com.sun.star.lib.loader.InstallationFinder;

public class JLibreOfficeInstallation {

	private static JLibreOfficeInstallation instance;
	private static Logger log = Logger.getLogger(JLibreOfficeInstallation.class);
	
	private CustomURLClassLoader classLoader;
	private String unoPath;
	
	public static JLibreOfficeInstallation getInstance() {
		return instance;
	}	
			    
    public static void iniciar() throws Exception {	

    	if (instance == null) {
    		
			// Verifica se consegue encontrar o caminho de instalacao
			String unoPath = InstallationFinder.getPath();
											
			// Verifica se encontrou o uno_path
			if (unoPath == null || unoPath.equals("")) {
				String msg = "O diretório de instalação do LibreOffice não foi encontrado!";
				throw new Exception(msg);
			}

			log.info(InstallationFinder.UNO_PATH + " = " + unoPath);

			SystemUtils.putEnvVar(InstallationFinder.UNO_PATH, unoPath);
			
			if (SystemUtils.isOsLinux()) {
				System.setProperty("sun.awt.xembedserver", "true");
			}

			instance = new JLibreOfficeInstallation(unoPath);
    	}
	}	

	private JLibreOfficeInstallation(String unoPath) throws Exception {
		
		this.unoPath = unoPath;
		
		String unoPathRoot = new File(unoPath).getParent();
		
		List<DependencyPath> libs = new ArrayList<DependencyPath>();
		
		libs.add(new DependencyPath("officebean.jar", true));
		libs.add(new DependencyPath("unoil.jar", true));
		libs.add(new DependencyPath("jurt.jar", true));
		libs.add(new DependencyPath("ridl.jar", true));
		libs.add(new DependencyPath("java_uno.jar", true));
		libs.add(new DependencyPath("juh.jar", true));
		
		// Verifica qual e o SO
		if (SystemUtils.isOsWindows()) {
			
			libs.add(new DependencyPath("msvcr70.dll", false));
			libs.add(new DependencyPath("msvcr71.dll", false));
			libs.add(new DependencyPath("uwinapi.dll", false));
			libs.add(new DependencyPath("jawt.dll", false));
			libs.add(new DependencyPath("officebean.dll", true));
			libs.add(new DependencyPath("sal3.dll", true));
			libs.add(new DependencyPath("jpipe.dll", true));
			
			List<File> dirs = new ArrayList<File>();
			
			dirs.add(new File(unoPathRoot + "\\program"));
			dirs.add(new File(unoPathRoot + "\\program\\classes"));
			dirs.add(new File(unoPathRoot + "\\Basis\\program"));			
			dirs.add(new File(unoPathRoot + "\\Basis\\program\\classes"));
			dirs.add(new File(unoPathRoot + "\\URE\\bin"));
			dirs.add(new File(unoPathRoot + "\\URE\\java\\"));
			
			
			initializeClassloader(libs, dirs);
		}
		else if (SystemUtils.isOsLinux()) {

			libs.add(new DependencyPath("libofficebean.so", true));
			libs.add(new DependencyPath("libjpipe.so", true));
			libs.add(new DependencyPath("libjuh.so", true));

			List<File> dirs = new ArrayList<File>();

			dirs.add(new File(unoPathRoot + "/program/"));
			dirs.add(new File(unoPathRoot + "/program/classes/"));
			dirs.add(new File(unoPathRoot + "/ure-link/share/java/"));
			dirs.add(new File("/usr/lib/ure/lib/"));
			
			initializeClassloader(libs, dirs);
		}
		else {
			throw new Exception ("O JLibreOffice somente esta preparado para funcionar com os Sistemas Operacionais Windows e Linux Ubuntu!");
		}
	}
	
	private void initializeClassloader(List<DependencyPath> dependencies, List<File> dirs) throws Exception {

		Set<URL> jars = new HashSet<URL>();
		Set<File> libs = new HashSet<File>();
		
		for (DependencyPath dependency : dependencies) {

			File found = null;
			
			for (File dir : dirs) {
				File file = new File(dir, dependency.getDependencyName());
				if (file.exists()) {
					found = file;
					log.debug(MessageFormat.format("A dependência ({0}) foi encontrada em: {1}", dependency.getDependencyName(), file.toString()));
					break;
				}
			}
			
			if (found != null) {
				if (dependency.isLib()) {
					libs.add(found.getParentFile());
				}
				else {
					jars.add(found.toURI().toURL());
				}
			}
			else if (dependency.isRequired()) { 
				throw new Exception(MessageFormat.format("A dependência {0} não foi encontrada!", dependency.getDependencyName()));
			}
		}

		this.classLoader = new CustomURLClassLoader(jars.toArray(new URL[]{}), JLibreOfficeInstallation.class.getClassLoader());
		this.classLoader.setResourcePaths(libs);
	}

	public URLClassLoader getClassLoader() {
		return classLoader;	
	}

	public String getUnoPath() {
    	return unoPath;
    }
	
	public File getUnoPathFile() {
		return new File(getUnoPath());
	}
}