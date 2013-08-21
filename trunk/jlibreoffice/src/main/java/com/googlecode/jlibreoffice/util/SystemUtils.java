package com.googlecode.jlibreoffice.util;

import java.lang.reflect.Field;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class SystemUtils {
	
	private static final String OS_LINUX = "Linux";	
	private static final String OS_WINDOWS = "Windows";
	private static final String PROPERTY_OS_NAME = "os.name";
	
	private static String getOsName() {
		return System.getProperty(PROPERTY_OS_NAME);
	}
	
	public static String getTmpDirPath() {		
		return System.getProperty("java.io.tmpdir");
	}
		
	public static boolean isOsLinux() {
		return getOsName().startsWith(OS_LINUX);
	}

	public static boolean isOsWindows() {
		return getOsName().startsWith(OS_WINDOWS);
	}
	
	public static String getEnvVar( String envVar ) {        

        String path = null;
        
        try {
            path = System.getenv(envVar);
        } 
        catch ( SecurityException e ) {} 
        catch ( java.lang.Error err ) {}
        return path;
    }
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void putEnvVar(String key, String value) {
					
		HashMap<String,String> newenv = new HashMap<String,String>();
		newenv.putAll(System.getenv());
		newenv.put(key, value);

		try {
			Class<?> processEnvironmentClass = Class.forName("java.lang.ProcessEnvironment");
			Field theEnvironmentField = processEnvironmentClass.getDeclaredField("theEnvironment");
			theEnvironmentField.setAccessible(true);
			Map<String, String> env = (Map<String, String>) theEnvironmentField.get(null);
			env.putAll(newenv);
			Field theCaseInsensitiveEnvironmentField = processEnvironmentClass.getDeclaredField("theCaseInsensitiveEnvironment");
			theCaseInsensitiveEnvironmentField.setAccessible(true);
			Map<String, String> cienv = (Map<String, String>) theCaseInsensitiveEnvironmentField.get(null);
			cienv.putAll(newenv);
		}
		catch (NoSuchFieldException e) {
			try {
				Class[] classes = Collections.class.getDeclaredClasses();
				Map<String, String> env = System.getenv();
				for (Class cl : classes) {
					if ("java.util.Collections$UnmodifiableMap".equals(cl
							.getName())) {
						Field field = cl.getDeclaredField("m");
						field.setAccessible(true);
						Object obj = field.get(env);
						Map<String, String> map = (Map<String, String>) obj;
						map.clear();
						map.putAll(newenv);
					}
				}
			} 
			catch (Exception e2) {
				e2.printStackTrace();
			}
		} 
		catch (Exception e1) {
			e1.printStackTrace();
		}
	}
}