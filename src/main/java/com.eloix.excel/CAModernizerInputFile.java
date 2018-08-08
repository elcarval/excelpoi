package com.eloix.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CAModernizerInputFile {

    private static final String file_name = "CA_Modernizer_Input.xlsx";

    public static void main(String[] args){
        createFile();
    }

    private static void createFile(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {

            XSSFSheet sheet;
            XSSFRow row;
            XSSFCell cell;

            // Hardware Configuration Sheet
            sheet = workbook.createSheet("HW_Config");
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("mrbtsId");           // Mandatory
            cell = row.createCell(1);
            cell.setCellValue("siteCode");          // Optional
            cell = row.createCell(2);
            cell.setCellValue("configName");        // Mandatory

            formatCells(sheet, 0, 3);

            // Cells info Sheet
            //sheet = workbook.createSheet("Cells");
            //row = sheet.createRow(0);
            //cell = row.createCell(0);
            //cell.setCellValue("mrbtsId");
            //cell = row.createCell(1);
            //cell.setCellValue("siteCode");
            //cell = row.createCell(2);
            //cell.setCellValue("cellName");

            //formatCells(sheet, 0, 3);

            // Transport Sheet
            //sheet = workbook.createSheet("Transport");
            //row = sheet.createRow(0);
            //cell = row.createCell(0);
            //cell.setCellValue("mrbtsId");
            //cell = row.createCell(1);
            //cell.setCellValue("siteCode");
            //cell = row.createCell(2);
            //cell.setCellValue("tceIpAddress");

            //formatCells(sheet, 0, 3);

            // Input Sheet
            String[] parameters = {
                    "mrbtsId",
                    // cell params
                    "lcrId",                    // @Radio (CellPar)!C | @NOKLTE:LNCEL
                    "cellName",                 // @Radio (CellPar)!D | @NOKLTE:LNCEL
                    "pMax",                     // @Radio (CellPar)!F | @NOKLTE:LNCEL
                    "tac",                      // @Radio (CellPar)!J | @NOKLTE:LNCEL
                    "chBW",                     // @Radio (CellPar)!K | @NOKLTE:LNCEL_FDD (2x)
                    "earfcnDL",                 // @Radio (CellPar)!L | @NOKLTE:LNCEL_FDD
                    "earfcnUL",                 // @Radio (CellPar)!M | @NOKLTE:LNCEL_FDD
                    "pci",                      // @Radio (CellPar)!N | @NOKLTE:LNCEL
                    "rsi",                      // @Radio (CellPar)!O | @NOKLTE:LNCEL_FDD
                    // antenna params
                    "totalLoss",                // @LTE Antenna!V | @com.nokia.srbts.eqm:ANTL
                    "antennaRoundTripDelay",    // @LTE Antenna!W | @com.nokia.srbts.eqm:ANTL
                    // transport param(s)
                    "tceIpAddress"              // @Transport!CD | @NOKLTE:MTRACE
            };

            sheet = workbook.createSheet("Input");
            row = sheet.createRow(0);

            for(int i = 0; i < parameters.length; ++i){
                cell = row.createCell(i);
                cell.setCellValue(parameters[i]);
            }

            formatCells(sheet, 0, parameters.length);

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

    private static void formatCells(XSSFSheet sheet, int row, int nCells){
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
