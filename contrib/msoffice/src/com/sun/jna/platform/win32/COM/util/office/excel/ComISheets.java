package com.sun.jna.platform.win32.COM.util.office.excel;


import com.sun.jna.platform.win32.COM.util.IConnectionPoint;
import com.sun.jna.platform.win32.COM.util.IUnknown;
import com.sun.jna.platform.win32.COM.util.annotation.ComInterface;
import com.sun.jna.platform.win32.COM.util.annotation.ComMethod;
import com.sun.jna.platform.win32.COM.util.annotation.ComProperty;

@ComInterface(iid="{000208D7-0000-0000-C000-000000000046}")
public interface ComISheets extends IUnknown, IConnectionPoint {



    @ComProperty
    ComIApplication Application();

    @ComProperty
    int getCount();

    @ComProperty
    ComIWorksheet getItem(int object);

    @ComProperty
    boolean Visible();




}
