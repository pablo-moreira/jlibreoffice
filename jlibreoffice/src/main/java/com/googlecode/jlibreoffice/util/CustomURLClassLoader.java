package com.googlecode.jlibreoffice.util;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.ArrayList;

public class CustomURLClassLoader extends URLClassLoader {

    private ArrayList<URL> resourcePaths;

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

    public void addResourcePath(URL rurl) {
        if (resourcePaths == null) resourcePaths = new ArrayList<URL>();
        resourcePaths.add(rurl);
    }

    public URL getResource(String name) {
        if (resourcePaths == null) return null;

        URL result = super.getResource(name);
        if (result != null) {
            return result;
        }

        for (URL u : resourcePaths) {
            if (u.getProtocol().startsWith("file")){
                try {
                    File f1 = new File(u.getPath());
                    File f2 = new File(f1, name);
                    if (f2.exists()) {
                        return new URL(f2.toURI().toASCIIString());
                    }
                } catch (MalformedURLException e1) {
                    System.err.println("malformed url: "+e1.getMessage());
                    continue;
                }
            }
        }
        return null;
    }

}