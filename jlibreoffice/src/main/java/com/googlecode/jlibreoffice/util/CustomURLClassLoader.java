package com.googlecode.jlibreoffice.util;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.Set;

public class CustomURLClassLoader extends URLClassLoader {
	
    private Set<File> resourcePaths;

    public CustomURLClassLoader( URL[] urls, ClassLoader urlClassLoader) {
        super( urls , urlClassLoader );
    }

    protected Class<?> findClass( String name ) throws ClassNotFoundException {
        // This is only called via this.loadClass -> super.loadClass ->
        // this.findClass, after this.loadClass has already called
        // super.findClass, so no need to call super.findClass again:
        throw new ClassNotFoundException( name );
//        return super.findClass(name);
    }

    protected Class<?> loadClass( String name, boolean resolve ) throws ClassNotFoundException
    {
        Class<?> c = findLoadedClass( name );
        if ( c == null ) {
            try {
                c = super.findClass( name );
            } catch ( ClassNotFoundException e ) {
                return super.loadClass( name, resolve );
            } catch ( SecurityException e ) {
                // A SecurityException "Prohibited package name: java.lang"
                // may occur when the user added the JVM's rt.jar to the
                // java.class.path:
                return super.loadClass( name, resolve );
            }
        }
        if ( resolve ) {
            resolveClass( c );
        }
        return c;
    }

    public void setResourcePaths(Set<File> dirs) {        
        resourcePaths = dirs;
    }

    public URL getResource(String name) {
        
    	if (resourcePaths == null) {
    		return null;
    	}

        URL result = super.getResource(name);
        
        if (result != null) {
            return result;
        }

        for (File dir : resourcePaths) {
            
			File lib = new File(dir, name);
            
			try {
				if (lib.exists()) {
					return new URL(lib.toURI().toASCIIString());                
				}
			}
        	catch (MalformedURLException e1) {
                System.err.println("malformed url: " + e1.getMessage());
                continue;
            }            
        }
        return null;
    }
}