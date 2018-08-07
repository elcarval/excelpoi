package com.eloix.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CAModernizerInputFile {

    public static final String file_name = "CA_Modernizer_Input.xlsx";

    public static void main(String[] args){
        createFile();
    }

    public static void createFile(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {

            XSSFSheet sheet;
            XSSFRow row;
            XSSFCell cell;

            // Hardware Configuration Sheet
            sheet = workbook.createSheet("HW_Config");
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("mrbtsId");
            cell = row.createCell(1);
            cell.setCellValue("siteCode");
            cell = row.createCell(2);
            cell.setCellValue("configName");

            formatCells(sheet, 0, 3);

            // Cells info Sheet
            sheet = workbook.createSheet("Cells");
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("mrbtsId");
            cell = row.createCell(1);
            cell.setCellValue("siteCode");
            cell = row.createCell(2);
            cell.setCellValue("cellId");

            formatCells(sheet, 0, 3);

            // Transport Sheet
            sheet = workbook.createSheet("Transport");
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("mrbtsId");
            cell = row.createCell(1);
            cell.setCellValue("siteCode");
            cell = row.createCell(2);
            cell.setCellValue("tceIpAddress");

            formatCells(sheet, 0, 3);

            String file_path = "D:\\OneDrive - Nokia\\MTEL_BG\\CA\\CA_Modernizer_v2+\\";

            FileOutputStream fos = new FileOutputStream(file_path + file_name);
            workbook.write(fos);
            fos.close();
        }catch (FileNotFoundException fnf_ex){
            System.out.println("ELOIX: Couldn't write to file! Something went wrong!");
            fnf_ex.printStackTrace();
        }catch (IOException io_ex){
            System.out.println("ELOIX: Couldn't close the file! Something went wrong!");
            io_ex.printStackTrace();
        }
    }

    public static void formatCells(XSSFSheet sheet, int row, int nCells){
        XSSFFont font = sheet.getWorkbook().createFont();
        font.setBold(true);

        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFillForegroundColor(IndexedColors.GOLD.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(font);

        for(int i = row; i < nCells; ++i){
            sheet.getRow(0).getCell(i).setCellStyle(style);
            //sheet.autoSizeColumn(i);
            //sheet.setColumnWidth(i, 12 * 256); // 11.29 in excel
            sheet.setColumnWidth(i, 3265); // 12 in excel -> 12*256/11.29*12
        }
    }
}
