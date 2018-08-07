package com.eloix.excel;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadTRM {

    private static final String tceIpAddress_header_cell = "CD5";

    public static void main(String[] args) {
        String input_file_path = "D:\\OneDrive - Nokia\\MTEL_BG\\DF\\w33_18\\TRM_w33_v2.xlsm";
        String sheet_name = "Transport";
        String mrbtsId_header_cell = "A5";


        String siteCode = "SHU0020";

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(input_file_path));
            XSSFSheet sheet = workbook.getSheet(sheet_name);
            //int last_row = sheet.getLastRowNum();
            //XSSFRow row = sheet.getRow(4); // line 5
            //int last_column = row.getLastCellNum();
            //System.out.println("last_row: " + last_row + "; last_column: " + last_column); // DEBUG

            //CellReference cr  = new CellReference(tceIpAddress_header_cell);
            //System.out.println(cr); // DEBUG
            //System.out.println(sheet.getRow(cr.getRow()+120).getCell(cr.getCol()).getStringCellValue()); // DEBUG

            String tceIpAddress = getTceIpAddress(siteCode, 5, sheet);
            if(tceIpAddress == null){
                System.out.println("Site (" + siteCode + ") doesn't exist in the TRM file or there's no tceIpAddress defined for it!");
            }
            else{
                System.out.println("The tceIpAddress for site (" + siteCode + ") is: " + tceIpAddress);
            }

        }catch (FileNotFoundException fnf_ex){
            System.out.println("ELOIX: Couldn't open or write to file! Something went wrong!");
            fnf_ex.printStackTrace();
        }catch (IOException io_ex){
            System.out.println("ELOIX: Couldn't close the file! Something went wrong!");
            io_ex.printStackTrace();
        }
    }

    private static String getTceIpAddress(String site_code, int start_row, XSSFSheet sheet){
        int last_row = sheet.getLastRowNum();
        CellReference cr = new CellReference(tceIpAddress_header_cell);
        for(int i = start_row; i < last_row; ++i){
            XSSFCell cell = sheet.getRow(i).getCell(0);
            if(cell.getStringCellValue().equals(site_code)){
                System.out.println("Found at row: " + i); // DEBUG
                return sheet.getRow(i).getCell(cr.getCol()).getStringCellValue();
            }
        }
        return null;
    }
}
