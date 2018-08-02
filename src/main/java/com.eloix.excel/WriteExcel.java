package com.eloix.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    public static void main(String[] args){
        HSSFWorkbook workbook = new HSSFWorkbook();
        //XSSFWorkbook workbook = new XSSFWorkbook();

        HSSFSheet sheet1 = workbook.createSheet("FirstSheet");
        HSSFRow row1 = sheet1.createRow(0);
        HSSFCell cell1 = row1.createCell(0);
        cell1.setCellValue("1st cell");


        try {
            workbook.write(new FileOutputStream("excel_poi_test1.xls"));
            workbook.close();
        }catch (FileNotFoundException fnf_ex){
            System.out.println("ELOIX: Couldn't write to file! Something went wrong!");
            fnf_ex.printStackTrace();
        }catch (IOException io_ex){
            System.out.println("ELOIX: Couldn't close the file! Something went wrong!");
            io_ex.printStackTrace();
        }
    }
}
