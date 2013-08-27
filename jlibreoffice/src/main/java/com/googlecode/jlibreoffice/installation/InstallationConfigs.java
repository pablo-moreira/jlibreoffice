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

public class InstallationConfigs {

	private static InstallationConfigs instance;
	private static Logger log = Logger.getLogger(InstallationConfigs.class);

	private CustomURLClassLoader classLoader;
	private String unoPath;
	
	public static InstallationConfigs getInstance() {
		return instance;
	}	
			    
    public static void iniciar() throws Exception {	

    	if (instance == null) {
    		
			// Verifica se consegue encontrar o caminho de instalacao
			String unoPath = InstallationFinder.getPath();
											
			// Verifica se encontrou o uno_path
			if (unoPath == null || unoPath.equals("")) {
				String msg = "O diretório de instalação do LibreOffice não foi encontrado!";
				log.error(msg);
				throw new Exception(msg);
			}

			log.info(InstallationFinder.UNO_PATH + " = " + unoPath);

			SystemUtils.putEnvVar(InstallationFinder.UNO_PATH, unoPath);

			instance = new InstallationConfigs(unoPath);	
    	}
	}	

	private InstallationConfigs(String unoPath) throws Exception {
		
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

			dirs.add(new File(unoPathRoot + "/basis-link/program/"));
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

		Set<URL> urls = new HashSet<URL>();
		
		for (DependencyPath dependency : dependencies) {

			URL found = null;
			
			for (File dir : dirs) {
				File file = new File(dir, dependency.getDependencyName());
				if (file.exists()) {
					found = file.toURI().toURL();
					break;
				}
			}
			
			if (found != null) {
				urls.add(found);
			}
			else if (dependency.isRequired()) { 
				throw new Exception(MessageFormat.format("A biblioteca {0} não foi encontrada!", dependency.getDependencyName()));
			}
		}

		this.classLoader = new CustomURLClassLoader(urls.toArray(new URL[]{}), InstallationConfigs.class.getClassLoader());
		this.classLoader.addResourcePath(new File(getUnoPath()).toURI().toURL());
	}

	public URLClassLoader getClassLoader() {
		return classLoader;	
	}

	public String getUnoPath() {
    	return unoPath;
    }
}