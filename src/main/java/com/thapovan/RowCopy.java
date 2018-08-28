package com.thapovan;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class RowCopy {

    public static void main(String[] args) throws Exception {
        String excelFilePath = "C:\\Users\\hnarayanan\\Desktop\\Src\\Src.xls";

        String destFilePath = "C:\\Users\\hnarayanan\\Desktop\\Dest\\Dest.xls";

        String srcXlsx = "C:\\Users\\hnarayanan\\Desktop\\Src\\Src.xlsx";

        String destXslx = "C:\\Users\\hnarayanan\\Desktop\\Dest\\Dest.xlsx";

        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(excelFilePath));

        XSSFWorkbook inputWorkbook=new XSSFWorkbook(new FileInputStream(srcXlsx));

        HSSFSheet sheet = workbook.getSheet("A");
        copyRow(workbook, sheet, 0, 0);

        XSSFSheet sheetXlsx = inputWorkbook.getSheet("A");

        copyRowXlsx(inputWorkbook,sheetXlsx,0,0);

        FileOutputStream outStream = new FileOutputStream(destXslx);
        inputWorkbook.write(outStream);
        outStream.close();

        FileOutputStream out = new FileOutputStream(destFilePath);
        workbook.write(out);
        out.close();
    }




    private static void copyRowXlsx(XSSFWorkbook workbook, XSSFSheet worksheet, int srcRowNum, int destRowNum) {
        worksheet.protectSheet("pass");

        int lastRowNum = workbook.getSheet("A").getLastRowNum();
        int minRowNum = workbook.getSheet("A").getFirstRowNum();
        int l =0;

        System.out.println("lastRowNum: "+lastRowNum);
        System.out.println("minRowNum: "+minRowNum);

        while(l<lastRowNum) {
            XSSFRow srcRow = workbook.getSheet("A").getRow(l);



            int lastCellNum = srcRow.getLastCellNum();

            // If the row exist in destination, push down all rows by 1 else create a new row

            XSSFRow newRow = worksheet.createRow(destRowNum);


            for (int j = 0; j < lastCellNum; j++) {
                System.out.println("Inside the loop...");
                // Grab a copy of the old/new cell
                XSSFCell oldCell = srcRow.getCell(j);
                XSSFCell newCell = newRow.createCell(j);

                // If the old cell is null jump to next cell
                if (oldCell == null) {
                    newCell = null;
                    continue;
                }

                // Copy style from old cell and apply to new cell
                XSSFCellStyle newCellStyle = workbook.createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

                newCell.setCellStyle(newCellStyle);

                // If there is a cell comment, copy
                if (oldCell.getCellComment() != null) {
                    newCell.setCellComment(oldCell.getCellComment());
                }

                // If there is a cell hyperlink, copy
                if (oldCell.getHyperlink() != null) {
                    newCell.setHyperlink(oldCell.getHyperlink());
                }

                // Set the cell data type
                newCell.setCellType(oldCell.getCellType());

                // Set the cell data value
                switch (oldCell.getCellType()) {
                    case Cell.CELL_TYPE_BLANK:
                        newCell.setCellValue(oldCell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        newCell.setCellValue(oldCell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        newCell.setCellErrorValue(oldCell.getErrorCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        newCell.setCellFormula(oldCell.getCellFormula());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        newCell.setCellValue(oldCell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                        newCell.setCellValue(oldCell.getRichStringCellValue());
                        break;
                }
            }
            l++;
        }
    }

    private static void copyRow(HSSFWorkbook workbook, HSSFSheet worksheet, int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        worksheet.protectSheet("pass");
        HSSFRow sourceRow = worksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        HSSFRow newRow = worksheet.createRow(destinationRowNum);

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            HSSFCell oldCell = sourceRow.getCell(i);
            HSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            HSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }
    }
}