package top.wu.parsegameexcel.service.impl;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import top.wu.parsegameexcel.GiftType;
import top.wu.parsegameexcel.parseexcel.ExcelParse;
import top.wu.parsegameexcel.parseexcel.OfflineActivityGift;
import top.wu.parsegameexcel.parseexcel.SingleDayGift;
import top.wu.parsegameexcel.parseexcel.TotalMoneyGift;
import top.wu.parsegameexcel.service.ParseJLLService;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

@Service
public class ParseJLLServiceImpl implements ParseJLLService {

    @Override
    public Map<String, XSSFWorkbook> ParseJLLExcel(MultipartFile file) {
        Map<String, XSSFWorkbook> map = new HashMap<>();
        XSSFWorkbook xssfWorkbook = null;
        InputStream inputStream = null;
        try {
            inputStream = file.getInputStream();
            xssfWorkbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
//            XSSFSheet outputSheet = outputWorkbook.createSheet();
        //获取sheet的个数
        int numberOfSheets = xssfWorkbook.getNumberOfSheets();
        //获取指定的sheet
        for (int sheetId = 0; sheetId < numberOfSheets; sheetId++) {
            String sheetName = xssfWorkbook.getSheetName(sheetId);
            XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
            System.out.println(sheetName);
            if (sheetName.contains("新增认证")) {
                ExcelParse excelParse = new ExcelParse();
                XSSFWorkbook newTotalMoneyWorkBook = excelParse.jjlExcelParse(sheet, GiftType.NEW_AUTH_GIFT.getType());
                map.put(sheetName, newTotalMoneyWorkBook);
                System.out.println("新增认证奖励文档解析完毕");
            }
            if (sheetName.contains("累充额度")) {
                ExcelParse excelParse = new ExcelParse();
                XSSFWorkbook oldTotalMoneyWorkbook = excelParse.jjlExcelParse(sheet, GiftType.TOTAL_MONEY_GIFT.getType());
                map.put(sheetName, oldTotalMoneyWorkbook);
                System.out.println("累计充值文档解析完毕");
            }
            if (sheetName.contains("每日充值")) {
                ExcelParse excelParse = new ExcelParse();
                XSSFWorkbook singleDayGiftWorkBook = excelParse.jjlExcelParse(sheet, GiftType.SINGLE_DAY_GIFT.getType());
//                SingleDayGift singleDayGift = new SingleDayGift();
//                XSSFWorkbook singleDayGiftWorkBook = singleDayGift.parseSingleDayGift(sheet, sheetName);
                map.put(sheetName, singleDayGiftWorkBook);
                System.out.println("每日充值奖励文档完毕解析");
            }
            if (sheetName.contains("线下充值")) {
                OfflineActivityGift offlineActivityGift = new OfflineActivityGift();
//                XSSFWorkbook offlineActivityGiftWorkBook = offlineActivityGift.parseOfflineActivityGift(sheet, sheetName);
                XSSFWorkbook offlineActivityGiftWorkBook = offlineActivityGift.parseAutumnOfflineActivityGift(sheet, sheetName);
                map.put(sheetName, offlineActivityGiftWorkBook);
                System.out.println("线下活动奖励文档完毕解析");
            }
        }
        try {
            inputStream.close();
            xssfWorkbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return map;
    }
}
