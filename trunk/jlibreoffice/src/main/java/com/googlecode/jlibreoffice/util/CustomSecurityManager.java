package com.googlecode.jlibreoffice.util;

import java.security.Permission;

public class CustomSecurityManager extends SecurityManager {

	@Override
    public void checkPermission(Permission perm) {    	
        check(perm, null);
    } 

    @Override
    public void checkPermission(Permission perm, Object context) {
        check(perm, context);
    }

    private void check(Permission perm, Object context) {    			
    	return;
    }
}