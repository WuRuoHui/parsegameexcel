package top.wu.parsegameexcel.parseexcel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class MergeExcel {
    public XSSFSheet mergeGameExcel(XSSFSheet mergeSheet, XSSFSheet oldSheet, int isFirst) {
        int mergeSheetLastRowNum = mergeSheet.getLastRowNum();
        System.out.println("mergeSheetLastRowNum:" + mergeSheetLastRowNum);
        int oldSheetLastRow = oldSheet.getLastRowNum();

        for (int i = (isFirst == 0 ? 0 : 1); i <= oldSheetLastRow; i++) {
//            XSSFRow row = mergeSheet.createRow(mergeSheetLastRowNum == 0 ? mergeSheetLastRowNum : mergeSheetLastRowNum + i);
            XSSFRow row = mergeSheet.createRow(mergeSheetLastRowNum+i);
            XSSFRow oldSheetRow = oldSheet.getRow(i);
            short lastCellNum = oldSheetRow.getLastCellNum();

            for (short k = 0; k <= lastCellNum; k++) {
                if (oldSheetRow.getCell(k) == null || oldSheetRow.getCell(k).getCellType() == CellType.BLANK) continue;
                copyCellDate(row.createCell(k), oldSheetRow.getCell(k));
            }
//            mergeSheetLastRowNum++;
        }
        return mergeSheet;
    }

    private XSSFCell copyCellDate(XSSFCell newCell, XSSFCell oldCell) {
        if (oldCell.getCellType() == CellType.STRING) {
            newCell.setCellValue(oldCell.getStringCellValue());
        }
        if (oldCell.getCellType() == CellType.NUMERIC) {
            newCell.setCellValue(oldCell.getNumericCellValue());
        }
        return newCell;
    }

}
