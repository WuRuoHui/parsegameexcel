package top.wu.parsegameexcel.parseexcel;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.wu.parsegameexcel.utils.ExcelUtils;
import top.wu.parsegameexcel.utils.PropUtils;

import java.text.DecimalFormat;
import java.util.List;

public class TotalMoneyGift {

    public XSSFWorkbook parseSingleTotalGift(XSSFSheet sheet, String newExcelName, Boolean isSingleKey) {

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
            //充值档位
            System.out.println(roleId);
            System.out.println(numericCellValue);
            List<String> subTotalKeys = PropUtils.getSubTotalKeys((int) rowData.getCell(3).getNumericCellValue() + "", isSingleKey);
            ExcelUtils.setExcelHeader(outputSheet);

            for (String money : subTotalKeys) {
                List appendRow = ExcelUtils.setTotalMoneyGiftRow(district, roleId, money);
                ExcelUtils.appendListWithHeader(outputSheet, appendRow, outputWorkbook);
            }

        }

        System.out.print("共计" + outputSheet.getLastRowNum() + "行，");
        return outputWorkbook;
    }

    public void parseSingleTotalGift(XSSFRow rowData, XSSFSheet outputSheet, XSSFWorkbook outputWorkbook,
                                             String district, String roleId, Boolean isSingleKey) {
        List<String> subTotalKeys = PropUtils.getSubTotalKeys((int) rowData.getCell(3).getNumericCellValue() + "", isSingleKey);
        ExcelUtils.setExcelHeader(outputSheet);

        for (String money : subTotalKeys) {
            List appendRow = ExcelUtils.setTotalMoneyGiftRow(district, roleId, money);
            ExcelUtils.appendListWithHeader(outputSheet, appendRow, outputWorkbook);
        }
    }



    //王者射击累计充值档位数组获取   isSingleKey用来处理剑玲珑新增认证
    public void parseWZSJSingleTotalGift(XSSFRow rowData, XSSFSheet outputSheet, XSSFWorkbook outputWorkbook,
                                     String district, String roleId, Boolean isSingleKey) {
        List<String> subTotalKeys = PropUtils.getWZSJSubTotalKeys((int) rowData.getCell(3).getNumericCellValue() + "", isSingleKey);
        ExcelUtils.setExcelHeader(outputSheet);

        for (String money : subTotalKeys) {
            List appendRow = ExcelUtils.setWZSJTotalMoneyGiftRow(district, roleId, money);
            ExcelUtils.appendListWithHeader(outputSheet, appendRow, outputWorkbook);
        }
    }

}
