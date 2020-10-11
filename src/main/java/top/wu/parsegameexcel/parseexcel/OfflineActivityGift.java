package top.wu.parsegameexcel.parseexcel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.wu.parsegameexcel.utils.ExcelUtils;
import top.wu.parsegameexcel.utils.PropUtils;

import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class OfflineActivityGift {

    public XSSFWorkbook parseOfflineActivityGift(XSSFSheet sheet, String newExcelName) {

        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);
            if (rowData.getCell(0) == null) continue;
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            String money = (int)rowData.getCell(3).getNumericCellValue()+"";
            //充值档位
//            List<String> subTotalKeys = PropUtils.getSubTotalKeys((int) rowData.getCell(3).getNumericCellValue() + "", isSingleKey);

            ExcelUtils.setExcelHeader(outputSheet);

//            for (String money : subTotalKeys) {
            List appendRow = ExcelUtils.setOfflineActivityGiftRow(district, roleId, money);
            ExcelUtils.appendListWithHeader(outputSheet, appendRow, outputWorkbook);
//            }

        }

        System.out.print("共计" + outputSheet.getLastRowNum() + "行，");
        return outputWorkbook;
    }

    public XSSFWorkbook parseAutumnOfflineActivityGift(XSSFSheet sheet, String sheetName) {
        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);
            if (rowData.getCell(0) == null) continue;
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            //等级
            String level = (int)rowData.getCell(3).getNumericCellValue()+"";
            ExcelUtils.setExcelHeader(outputSheet);

            for (int i = 4;i <= lastCellNum; i++) {
                if (rowData.getCell(i) == null || rowData.getCell(i).getCellType() == CellType.BLANK) continue;
                int money = (int) rowData.getCell(i).getNumericCellValue();
                String day = sheet.getRow(0).getCell(i).getStringCellValue();
                List appendRow = ExcelUtils.setOfflineActivityGiftRow(district, roleId, level, money ,day);
                ExcelUtils.appendListWithHeader(outputSheet, appendRow, outputWorkbook);

            }

            //充值档位
//            List<String> subTotalKeys = PropUtils.getSubTotalKeys((int) rowData.getCell(3).getNumericCellValue() + "", isSingleKey);


//            for (String money : subTotalKeys) {
//            }

        }

        System.out.print("共计" + outputSheet.getLastRowNum() + "行，");
        return outputWorkbook;
    }
}
