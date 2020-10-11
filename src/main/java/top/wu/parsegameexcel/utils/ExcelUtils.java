package top.wu.parsegameexcel.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

import java.util.*;

public class ExcelUtils {

    //表头参数
    public static String[] excelHeader = {"平台", "区服", "角色ID", "申请原因", "是否为内部使用（1为是，0为否）", "邮件标题", "邮件内容", "道具id-1",
            "道具数量", "道具id-2", "道具数量", "道具id-3", "道具数量", "道具id-4", "道具数量", "道具id-5", "道具数量"};

    //设置表头
    public static XSSFSheet setExcelHeader(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(0);
        List<String> headers = Arrays.asList(excelHeader);
        for (int i = 0; i < headers.size(); i++) {
            row.createCell(i).setCellValue(headers.get(i));
        }
        return sheet;
    }

    //设填充一行Excel数据的List（累计充值）
    public static List<String> setTotalMoneyGiftRow(String district, String roleId, String money) {
        money = PropUtils.getSingleSubTotalIds(money);
        String reason = "累充" + money + "元";
        String emailTitle = "累充" + money + "元";
        String emailContent = "亲爱的仙友，累充" + money + "礼包已经发放，请留意查收哦~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);
        list.add(PropUtils.getTotalMoneyId(money));
        list.add("1");
        return list;
    }

    //设填充一行Excel数据的List（单日充值）
    public static List<String> setSingleGiftRow(String district, String roleId, int day, String payDay, String money) {
        List<String> subSingleKeys = PropUtils.getSubSingleKeys(money);
        String reason = payDay + "充值" + subSingleKeys.get(subSingleKeys.size() - 1) + "元";
        String emailTitle = reason;
        String emailContent = "亲爱的仙友，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);
        if (day <= 30) {
            for (String key : subSingleKeys) {
                list.add(PropUtils.getSingleIdLt30(key));
                list.add("1");
            }
        } else {
            for (String key : subSingleKeys) {
                list.add(PropUtils.getSingleIdMt30(key));
                list.add("1");
            }
        }
        return list;
    }

    //存储部分数据在list中（区服、角色ID、申请原因、邮件标题、邮件内容）
    public static List<String> setSubRow(String district, String roleId, String reason, String emailTitle, String emailContent) {
        List<String> list = new ArrayList<String>();
        list.add("mix_cn");
        list.add(district);
        list.add(roleId);
        list.add(reason);
        list.add("0");
        list.add(emailTitle);
        list.add(emailContent);
        return list;
    }

    //添加一行数据（一行数据存储在list中）
    public static XSSFSheet appendListWithHeader(XSSFSheet sheet, List arrayList, XSSFWorkbook workbook) {
        int lastRowNum = sheet.getLastRowNum();
        XSSFRow newRow = sheet.createRow(lastRowNum + 1);
        for (int i = 0; i < arrayList.size(); i++) {
            XSSFCell newCell = newRow.createCell(i);
            newCell.setCellValue(String.valueOf(arrayList.get(i)));

        }
        return sheet;
    }

    public static List<String> setOfflineActivityGiftRow(String district, String roleId, String money) {
        List giftList = new ArrayList();
        if (Integer.valueOf(money) < 1000) {
            //769	3	24401	1	3000124	1	3000235	8
            money = "500";
            giftList.add("769");
            giftList.add("3");
            giftList.add("24401");
            giftList.add("1");
            giftList.add("3000124");
            giftList.add("1");
            giftList.add("3000235");
            giftList.add("8");
        } else if (Integer.valueOf(money) < 1500) {
            //100绑元	6	占星石·朱雀	20	打磨石礼包	5	图鉴·委蛇（紫）	1	3阶橙色2星翅膀选择礼包	1	寻宝令选择宝箱	12
            //769	6	1800003	20	3100018	5	24401	1	3000119	1	3000235	12
            money = "1000";
            String[] ids = {"769", "6", "1800003", "20", "3100018", "5", "24401", "1", "3000119", "1", "3000235", "12"};
            giftList.addAll(Arrays.asList(ids));
        } else if (Integer.valueOf(money) < 3000) {
            //769	8	2110024	1	24501	1	3000114	1	3000235	16
            String[] ids = {"769", "8", "2110024", "1", "24501", "1", "3000114", "1", "3000235", "16"};
            money = "1500";
            giftList.addAll(Arrays.asList(ids));
        } else if (Integer.valueOf(money) < 5000) {
            //769	12	3000219	1	1901350	1	24501	1	3000131	1	1403012	1	3000235	20
            String[] ids = {"769", "12", "3000219", "1", "1901350", "1", "24501", "1", "3000131", "1", "1403012", "1", "3000235", "20"};
            money = "3000";
            giftList.addAll(Arrays.asList(ids));
        } else {
            //769	16	3000219	1	1901350	1	3100074	1	3100228	1	24601	1	3000131	1	1403019	1
            String[] ids = {"769", "16", "3000219", "1", "1901350", "1", "3100074", "1", "3100228", "1", "24601", "1", "3000131", "1", "1403019", "1"};
            money = "5000";
            giftList.addAll(Arrays.asList(ids));
        }
        String reason = "线下活动" + money + "档位";
        String emailTitle = "夏季福利活动" + money + "档位";
        String emailContent = "亲爱的仙友，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);

        list.addAll(giftList);
        return list;
    }

    public static List<String> setOfflineActivityGiftRow(String district, String roleId, String level, int money, String day) {

        List giftList = new ArrayList();
        int intMoney = Integer.valueOf(money);
        if (Integer.valueOf(level) < 500) {
            if (intMoney < 1000) {
                //绑元*600、随机橙品图鉴礼包*1、寻宝令宝箱*8
                String ids[] = {"769", "6", "3520100", "1", "3000509", "8"};
                giftList.addAll(Arrays.asList(ids));
            } else if (intMoney < 1500) {
                //绑元*800、饰品合成石*1、寻宝令宝箱*12
                String ids[] = {"769", "8", "2110023", "1", "3000509", "12"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 3000) {
                //绑元*1200、装备幻色石*1、寻宝令宝箱*16
                String ids[] = {"769", "12", "2110021", "1", "3000509", "16"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 5000) {
                //绑元*2400、神石选择礼包*1、寻宝令宝箱*20
                String ids[] = {"769", "24", "3000219", "1", "3000509", "20"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 10000) {
                //绑元*4000、神石选择礼包*1、防具圣装自选礼盒*1、寻宝令宝箱*24
                String ids[] = {"769", "40", "3000219", "1", "1403012", "1", "3000509", "24"};
                giftList.addAll(Arrays.asList(ids));

            } else {
                //绑元*8000、神石选择礼包*1、圣装自选礼盒*1、寻宝令宝箱*30
                String ids[] = {"769", "80", "3000219", "1", "1403019", "1", "3000509", "30"};
                giftList.addAll(Arrays.asList(ids));

            }
        } else {
            if (intMoney < 1000) {
                //绑元*800、随机橙品图鉴礼包*1、寻宝令宝箱*8
                String ids[] = {"769", "8", "3520100", "1", "3000509", "8"};
                giftList.addAll(Arrays.asList(ids));
            } else if (intMoney < 1500) {
                //绑元*1200、饰品合成石*1、寻宝令宝箱*12
                String ids[] = {"769", "12", "2110023", "1", "3000509", "12"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 3000) {
                //绑元*1500、龙珠幻色石*1、寻宝令宝箱*16
                String ids[] = {"769", "15", "2110024", "1", "3000509", "16"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 5000) {
                //绑元*3000、神石选择礼包*1、寻宝令宝箱*20
                String ids[] = {"769", "30", "3000219", "1", "3000509", "20"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 10000) {
                //绑元*5000、神石选择礼包*1、防具圣装自选礼盒*1、寻宝令宝箱*24
                String ids[] = {"769", "50", "3000219", "1", "1403012", "1", "3000509", "24"};
                giftList.addAll(Arrays.asList(ids));

            } else {
                //绑元*10000、神石选择礼包*1、圣装自选礼盒*1、寻宝令宝箱*30
                String ids[] = {"769", "100", "3000219", "1", "1403019", "1", "3000509", "30"};
                giftList.addAll(Arrays.asList(ids));

            }
        }

        int giftLevels[] = {500, 1000, 1500, 3000, 5000, 10000};

        Integer lastMoney = 0;

        for (int i = 0; i < giftLevels.length; i++) {
            if (giftLevels[i] > money) {
                lastMoney = giftLevels[i - 1];
                break;
            }
        }

        System.out.println(lastMoney);

        String reason = day + "线下活动" + lastMoney + "档位";
        String emailTitle = day + "秋季福利活动" + lastMoney + "档位";
        String emailContent = "亲爱的仙友，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);

        list.addAll(giftList);
        return list;
    }


}
