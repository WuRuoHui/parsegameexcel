package top.wu.parsegameexcel.parseexcel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.text.DecimalFormat;

public class ExcelParse {

    //剑玲珑
    //0：累计充值,2：新增认证,1：单日充值
    public XSSFWorkbook jjlExcelParse(XSSFSheet sheet ,Integer type) {
        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);
            if (rowData.getCell(0) == null || rowData.getCell(0).getCellType() == CellType.BLANK) continue;
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
//            System.out.println(rowData == null);
            XSSFCell cell = rowData.getCell(2);
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            if ( type == 0) {
                TotalMoneyGift totalMoneyGift = new TotalMoneyGift();
                totalMoneyGift.parseSingleTotalGift(rowData,outputSheet,outputWorkbook,district,roleId,true);
            }
            if (type == 2) {
                TotalMoneyGift totalMoneyGift = new TotalMoneyGift();
                totalMoneyGift.parseSingleTotalGift(rowData,outputSheet,outputWorkbook,district,roleId,false);
            }
            if (type == 1) {
                SingleDayGift singleDayGift = new SingleDayGift();
                singleDayGift.parseSingleDayGift(sheet,rowData,outputSheet,outputWorkbook,district,roleId);
            }

        }

        System.out.print("共计"+outputSheet.getLastRowNum()+"行，");
        return outputWorkbook;
    }

    //王者射击
    //0：累计充值,2：新增认证,1：单日充值
    public XSSFWorkbook wzsjExcelParse(XSSFSheet sheet ,Integer type) {
        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);
            if (rowData.getCell(0) == null || rowData.getCell(0).getCellType() == CellType.BLANK) continue;
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
//            System.out.println(rowData == null);
            XSSFCell cell = rowData.getCell(2);
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            if ( type == 0) {
                TotalMoneyGift totalMoneyGift = new TotalMoneyGift();
                totalMoneyGift.parseWZSJSingleTotalGift(rowData,outputSheet,outputWorkbook,district,roleId,true);
            }
            if (type == 1) {
                SingleDayGift singleDayGift = new SingleDayGift();
                singleDayGift.parseWZSJSingleDayGift(sheet,rowData,outputSheet,outputWorkbook,district,roleId);
            }

        }

        System.out.print("共计"+outputSheet.getLastRowNum()+"行，");
        return outputWorkbook;
    }

    //狩猎使命
    //0：累计充值,2：新增认证,1：单日充值
    public XSSFWorkbook slsmExcelParse(XSSFSheet sheet ,Integer type) {
        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);
            if (rowData.getCell(0) == null || rowData.getCell(0).getCellType() == CellType.BLANK) continue;
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
//            System.out.println(rowData == null);
            XSSFCell cell = rowData.getCell(2);
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            if ( type == 0) {
                TotalMoneyGift totalMoneyGift = new TotalMoneyGift();
                totalMoneyGift.parseWZSJSingleTotalGift(rowData,outputSheet,outputWorkbook,district,roleId,true);
            }
            if (type == 1) {
                SingleDayGift singleDayGift = new SingleDayGift();
                singleDayGift.parseSLSMSingleDayGift(sheet,rowData,outputSheet,outputWorkbook,district,roleId);
            }

        }

        System.out.print("共计"+outputSheet.getLastRowNum()+"行，");
        return outputWorkbook;
    }

}
