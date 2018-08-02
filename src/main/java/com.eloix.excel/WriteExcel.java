package com.eloix.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class WriteExcel {
    public static void main(String[] args){
        HSSFWorkbook workbook = new HSSFWorkbook();
        //XSSFWorkbook workbook = new XSSFWorkbook();

        HSSFSheet sheet1 = workbook.createSheet("FirstSheet");
        HSSFRow row = sheet1.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("1st cell");

        cell = row.createCell(1);
        DataFormat format = workbook.createDataFormat();
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        cell.setCellStyle(dataStyle);
        cell.setCellValue(new Date());

        row.createCell(2).setCellValue("3rd cell");

        sheet1.autoSizeColumn(1);

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
