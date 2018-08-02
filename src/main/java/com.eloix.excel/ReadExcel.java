package com.eloix.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("excel_poi_test1.xls"));
            HSSFSheet sheet = workbook.getSheetAt(0);
            HSSFRow row = sheet.getRow(0);
            if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                System.out.println(row.getCell(0).getStringCellValue());
            }
            if (row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                System.out.println(row.getCell(1).getDateCellValue());
            }
        }catch (FileNotFoundException fnf_ex){
            System.out.println("ELOIX: Couldn't write to file! Something went wrong!");
            fnf_ex.printStackTrace();
        }catch (IOException io_ex){
            System.out.println("ELOIX: Couldn't close the file! Something went wrong!");
            io_ex.printStackTrace();
        }
    }
}
