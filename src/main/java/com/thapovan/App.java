package com.thapovan;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {

    }

    public static void createDocs() throws Exception{
        createPasswordProtectedDoc("C:\\Users\\hnarayanan\\Desktop\\Src.docx");
        createPasswordProtectedExcel("C:\\Users\\hnarayanan\\Desktop\\Src.xls");
    }

    public static void createPasswordProtectedDoc(String filePath) throws Exception{
        //Blank Document
        XWPFDocument document = new XWPFDocument();
    //    document.enforceReadonlyProtection("aa",null);
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream( new File(filePath));
        document.write(out);
        out.close();
        System.out.println(String.format("Password document created in %s successully",filePath));
    }

    public static void createPasswordProtectedExcel(String filePath) throws Exception{

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = null;
       // sheet = workbook.createSheet("All Transaction Report");
        FileOutputStream out = new FileOutputStream( new File(filePath));
        sheet.protectSheet("aaaa");
        workbook.write(out);
        workbook.close();

        System.out.println(String.format("Password excel created in %s successully",filePath));
    }


}
