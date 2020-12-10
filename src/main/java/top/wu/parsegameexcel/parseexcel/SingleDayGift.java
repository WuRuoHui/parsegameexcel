package top.wu.parsegameexcel.parseexcel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.wu.parsegameexcel.utils.ExcelUtils;

import java.text.DecimalFormat;
import java.util.List;

public class SingleDayGift {

    public XSSFWorkbook parseSingleDayGift(XSSFSheet sheet, String newExcelName) {

        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
        XSSFSheet outputSheet = outputWorkbook.createSheet();

        //获取单个sheet最大行数
        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = 1; rowNum < lastRowNum; rowNum++) {
            XSSFRow rowData = sheet.getRow(rowNum);   //获取某行
            if (rowData.getCell(0) == null || rowData.getCell(0).getCellType() == CellType.BLANK) continue;

            //获取单行最大列数
            short lastCellNum = rowData.getLastCellNum();

            //区服
            String district = rowData.getCell(1).getStringCellValue();
            //角色ID
            //对角色ID进行格式化，防止因角色ID过长而导致填充时使用科学计数法
            Double numericCellValue = rowData.getCell(2).getNumericCellValue();
            DecimalFormat df = new DecimalFormat("0");//设置两位小数
            String roleId = df.format(numericCellValue);
            //开服天数
            int day = (int) rowData.getCell(3).getNumericCellValue();
            //充值档位
            ExcelUtils.setExcelHeader(outputSheet);
            for (int i = 4; i < lastCellNum; i++) {
                //当单元格为空时跳过
                if (rowData.getCell(i) == null || rowData.getCell(i).getCellType() == CellType.BLANK) continue;
                String payDay = sheet.getRow(0).getCell(i).getStringCellValue();
                int money = (int) rowData.getCell(i).getNumericCellValue();
                List<String> singleGiftRow = ExcelUtils.setSingleGiftRow(district, roleId, day, payDay, String.valueOf(money));
                ExcelUtils.appendListWithHeader(outputSheet, singleGiftRow, outputWorkbook);
            }

        }
        System.out.print("共计" + outputSheet.getLastRowNum() + "行，");
        return outputWorkbook;
    }

    public void parseSingleDayGift(XSSFSheet sheet, XSSFRow rowData, XSSFSheet outputSheet, XSSFWorkbook outputWorkbook,
                                           String district, String roleId) {
        //开服天数
        int day = (int) rowData.getCell(3).getNumericCellValue();
        short lastCellNum = rowData.getLastCellNum();
        //充值档位
        ExcelUtils.setExcelHeader(outputSheet);
        for (int i = 4; i < lastCellNum; i++) {
            //当单元格为空时跳过
            if (rowData.getCell(i) == null || rowData.getCell(i).getCellType() == CellType.BLANK) continue;
            String payDay = sheet.getRow(0).getCell(i).getStringCellValue();
            int money = (int) rowData.getCell(i).getNumericCellValue();
//            List<String> singleGiftRow = ExcelUtils.setSingleGiftRow(district, roleId, day, payDay, String.valueOf(money));
            List<String> singleGiftRow = ExcelUtils.setSingleGiftRowNew(district, roleId, day, payDay, String.valueOf(money));
            ExcelUtils.appendListWithHeader(outputSheet, singleGiftRow, outputWorkbook);
        }
    }
    }

