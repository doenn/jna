/* Copyright (c) 2019 Lewis Sall, All Rights Reserved
 * The contents of this file is dual-licensed under 2
 * alternative Open Source/Free licenses: LGPL 2.1 or later and
 * Apache License 2.0. (starting with JNA version 4.0.0).
 *
 * You can freely decide which license you want to apply to
 * the project.
 *
 * You may obtain a copy of the LGPL License at:
 *
 * http://www.gnu.org/licenses/licenses.html
 *
 * A copy is also included in the downloadable source code package
 * containing JNA, in file "LGPL2.1".
 *
 * You may obtain a copy of the Apache License at:
 *
 * http://www.apache.org/licenses/
 *
 * A copy is also included in the downloadable source code package
 * containing JNA, in file "AL2.0".
 */

package com.sun.jna;

import junit.framework.TestCase;
import com.sun.jna.platform.win32.COM.util.*;
import com.sun.jna.platform.win32.COM.*;
import com.sun.jna.Pointer;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.Helper;
import java.io.File;

import com.sun.jna.platform.win32.COM.util.AbstractComEventCallbackListener;
import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.COM.util.IComEventCallbackCookie;
import com.sun.jna.platform.win32.COM.util.annotation.ComInterface;
import com.sun.jna.platform.win32.COM.util.office.excel.ComExcel_Application;
import com.sun.jna.platform.win32.COM.util.office.excel.ComIAppEvents;
import com.sun.jna.platform.win32.COM.util.office.excel.ComIApplication;
import com.sun.jna.platform.win32.COM.util.office.excel.ComIRange;
import com.sun.jna.platform.win32.COM.util.office.excel.ComISheets;
import com.sun.jna.platform.win32.COM.util.office.excel.ComIWorkbook;
import com.sun.jna.platform.win32.COM.util.office.excel.ComIWorksheet;
import com.sun.jna.platform.win32.Ole32;
import java.io.IOException;

public class ExcelTest extends TestCase {

public void testExcel(){

    Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
    try {

        String filename = "C:\\Users\\lsall\\Desktop\\TestSpreadsheet1.xlsx";

        File demoDocument = new File(filename);
        ComIApplication msExcel = null;
        Factory factory = new Factory();
        ComExcel_Application excelObject = factory.createObject(ComExcel_Application.class);
        msExcel = excelObject.queryInterface(ComIApplication.class);

        //System.out.println("MSExcel version: " + msExcel.getVersion());
        System.out.println("Desired filename: " + filename);
        System.out.println("Attempting to open Excel..." + msExcel.getVersion());

         //   msExcel.setVisible(true);

        msExcel.setVisible(true);
        Helper.sleep(5);

        // demoDocument = Helper.createNotExistingFile("jnatest", ".xls");

        //   ComIWorkbook workbook = msExcel.getWorkbooks().Open(demoDocument.getAbsolutePath());
        // msExcel.getActiveSheet().getRange("A1").setValue("Hello from JNA!");
       ComIWorkbook workbook = msExcel.getWorkbooks().Open(demoDocument.getAbsolutePath());

        ComISheets sheets = workbook.getSheets();
        ComIWorksheet worksheet = sheets.getItem(2);
        worksheet.Select();

        // sheets.Select(worksheet);
        int index = 0;
        int count = 49;
        int base = 2;
        while (index < count) {
            System.out.println("Vals:" + worksheet.getRange("A"+base+index).getValue() + " "+"Vals:" + worksheet.getRange("B"+base+index).getValue() + " "+"Vals:" + worksheet.getRange("C"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("D"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("E"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("F"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("G"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("H"+base+index).getValue()+ " "+"Val:" + worksheet.getRange("I"+base+index).getValue());
            Helper.sleep(2);
            ++index;
        }
        // worksheet.getRange("A2").getValue();


        // .setValue("Heljdla JNA!");

    } finally {
        Ole32.INSTANCE.CoUninitialize();
    }
}


    public static void main(String[] args) {
        junit.textui.TestRunner.run(ExcelTest.class);
    }
}
