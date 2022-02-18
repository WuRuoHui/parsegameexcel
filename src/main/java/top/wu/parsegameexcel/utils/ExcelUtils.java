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

    //狩猎使命
    public static List<String> setSLSMTotalMoneyGiftRow(String district, String roleId, String money) {
        money = PropUtils.getWZSJSingleSubTotalIds(money);
        String reason = "累充" + money + "元";
        String emailTitle = "累充" + money + "元";
        String emailContent = "亲爱的玩家，累充" + money + "礼包已经发放，请留意查收哦~";
        List<String> list = setSLSMSubRow(district, roleId, reason, emailTitle, emailContent);
        list.add(PropUtils.getWZSJTotalMoneyId(money));
        list.add("1");
        return list;
    }

    //王者射击
    //设填充一行Excel数据的List（累计充值）
    public static List<String> setWZSJTotalMoneyGiftRow(String district, String roleId, String money) {
        money = PropUtils.getWZSJSingleSubTotalIds(money);
        String reason = "累充" + money + "元";
        String emailTitle = "累充" + money + "元";
        String emailContent = "亲爱的玩家，累充" + money + "礼包已经发放，请留意查收哦~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);
        list.add(PropUtils.getWZSJTotalMoneyId(money));
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

    //设填充一行Excel数据的List（单日充值 12.10版本）
    public static List<String> setSingleGiftRowNew(String district, String roleId, int day, String payDay, String money) {
        String key = PropUtils.getSubSingleKeyNew(money);
        String reason = payDay + "充值" + key + "元";
        String emailTitle = reason;
        String emailContent = "亲爱的仙友，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);
        if (day <= 7) {
            list.add(PropUtils.getSingleIdLt7(key));
            list.add("1");
        } else {
            list.add(PropUtils.getSingleIdMt7(key));
            list.add("1");
        }
        return list;
    }

    //王者射击 设填充一行Excel数据的List（单日充值 12.10版本）
    public static List<String> setWZSJSingleGiftRowNew(String district, String roleId, String payDay, String money) {
        String key = PropUtils.getWZSJSubSingleKeyNew(money);
        String reason = payDay + "充值" + key + "元";
        String emailTitle = reason;
        String emailContent = "亲爱的玩家，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);

        list.add(PropUtils.getWZSJSingleId(key));
        list.add("1");
       /* if (day <= 7) {
            list.add(PropUtils.getSingleIdLt7(key));
            list.add("1");
        } else {
            list.add(PropUtils.getSingleIdMt7(key));
            list.add("1");
        }*/
        return list;
    }

    //狩猎使命 设填充一行Excel数据的List（单日充值 12.10版本）
    public static List<String> setSLSMSingleGiftRowNew(String district, String roleId, String payDay, String money) {
        String key = PropUtils.getSLSMSubSingleKeyNew(money);
        String reason = payDay + "充值" + key + "元";
        String emailTitle = reason;
        String emailContent = "亲爱的玩家，您的充值奖励已发放，请留意查收~";
        List<String> list = setSLSMSubRow(district, roleId, reason, emailTitle, emailContent);

        switch (key) {
            case "1000":
                list.add("10009051");
                list.add("3");
                break;
            case "2000":
                //"10009051","3","10009021","1","10009011","1"
                list.addAll(Arrays.asList(new String[] {"10009051","3","10009021","1","10009011","1"}));
                break;
            case "3000":

                list.addAll(Arrays.asList(new String[] {"10009051","3","10009021","1","10009011","1","10009081","1"}));
                /*list.add("10009021");
                list.add("1");
                list.add("10009021");
                list.add("1");
                list.add("10009011");
                list.add("1")*/
                break;
            case "5000":
                //"10009011","1","10009081","1","10009061","1","10009031","1","10009041","10"
                list.addAll(Arrays.asList(new String[] {"10009011","1","10009081","1","10009061","1","10009031","1","10009041","10"}));
                break;
            case "7000":
                //"10009011","2","10009081","2","10009061","2","10009031","1","10009041","10"
                list.addAll(Arrays.asList(new String[] {"10009011","2","10009081","2","10009061","2","10009031","1","10009041","10"}));
                break;
            case "10000":
               // "10009011","2","10009081","3","10009061","2","10009031","2","10009041","10","10009071","1"
                list.addAll(Arrays.asList(new String[] {"10009011","2","10009081","3","10009061","2","10009031","2","10009041","10","10009071","1"}));
                break;
            case "15000":
                //"10009011","4","10009081","4","10009061","2","10009031","4","10009041","20","10009071","2"
                list.addAll(Arrays.asList(new String[] {"10009011","4","10009081","4","10009061","2","10009031","4","10009041","20","10009071","2"}));
                break;
            case "20000":
                //"10009011","6","10009081","5","10009061","3","10009031","5","10009041","30","10009071","2","10000076","3"
                list.addAll(Arrays.asList(new String[] {"10009011","6","10009081","5","10009061","3","10009031","5","10009041","30","10009071","2","10000076","3"}));
                break;

            case "30000":
                //"10009011","8","10009081","6","10009061","3","10009031","6","10009041","40","10009071","3","10000044","1"
                list.addAll(Arrays.asList(new String[] {"10009011","8","10009081","6","10009061","3","10009031","6","10009041","40","10009071","3","10000044","1"}));
                break;

            case "50000":
//                "10009011","12","10009081","9","10009061","5","10009031","12","10009041","50","10009071","5","10000044","2"
                list.addAll(Arrays.asList(new String[] {"10009011","12","10009081","9","10009061","5","10009031","12","10009041","50","10009071","5","10000044","2"}));
                break;

        }

/*        list.add(PropUtils.getWZSJSingleId(key));
        list.add("1");*/
       /* if (day <= 7) {
            list.add(PropUtils.getSingleIdLt7(key));
            list.add("1");
        } else {
            list.add(PropUtils.getSingleIdMt7(key));
            list.add("1");
        }*/
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

    //存储部分数据在list中（区服、角色ID、申请原因、邮件标题、邮件内容）
    public static List<String> setSLSMSubRow(String district, String roleId, String reason, String emailTitle, String emailContent) {
        List<String> list = new ArrayList<String>();
        list.add("mix_mob4ios");
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

        /*List giftList = new ArrayList();
        int intMoney = Integer.valueOf(money);
        if (Integer.valueOf(level) < 650) {
            if (intMoney < 3000) {
                //1000	绑元*500、巢穴怒气重置卡（ID：3120049）*1、魔王挑战礼盒*3
                String ids[] = {"769", "5", "3120049", "1", "3000202", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else if (intMoney < 8000) {
                //3000	绑元*1500、祝余年套装（ID：1011029，自动化识别性别）*1、巢穴等级限制解除卡（ID：3120047）*1、寻宝令宝箱（ID：3000521）*20
                String ids[] = {"769", "15", "1011029", "1", "3120047", "1", "3000521", "20"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 15000) {
                //8000	绑元*5000、祝余年·阵（ID：2005254）*1、岁岁年年-气泡*1、寻宝令宝箱（ID：3000521）*30、刷新卡选择礼包*2
                String ids[] = {"769", "50", "2005254", "1", "1021029", "1", "3000521", "30", "2210", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 30000) {
                //15000	绑元*8000、祝余年·宠（ID：3414）*1、祈丰年-相框*1、寻宝令宝箱（ID：3000521）*50、刷新卡选择礼包*3
                String ids[] = {"769", "80", "3414", "1", "1020035", "1", "3000521", "50", "2210", "3"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 50000) {
                //30000	绑元*12000、祝余年·魂（ID：2005148）*1、神石选择礼包*1、防具圣装选择礼盒*1、高级红品装备宝箱*2（650级可更换神武打造符*2）
                String ids[] = {"769", "120", "2005148", "1", "3000219", "1", "1403012", "1", "3100034", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 80000){
                //50000	绑元*20000、祝余年·翼（ID：1050428）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "200", "1050428", "1", "3000219", "2", "1403019", "1", "3100034", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else {
                //80000	绑元*30000、堕火凤凰（ID：2112025）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "300", "2112025", "1", "3000219", "2", "1403019", "1", "3100034", "3"};
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

            }*/
        List giftList = new ArrayList();
        int intMoney = Integer.valueOf(money);
        if (Integer.valueOf(level) < 650) {
            if (intMoney < 3000) {
                //1000	绑元*500、巢穴怒气重置卡（ID：3120049）*1、魔王挑战礼盒*3
                String ids[] = {"769", "5", "3120049", "1", "3000202", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else if (intMoney < 8000) {
                //3000	绑元*1500、祝余年套装（ID：1011029，自动化识别性别）*1、巢穴等级限制解除卡（ID：3120047）*1、寻宝令宝箱（ID：3000521）*20
                String ids[] = {"769", "15", "1011029", "1", "3120047", "1", "3000521", "20"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 15000) {
                //8000	绑元*5000、祝余年·阵（ID：2005254）*1、岁岁年年-气泡*1、寻宝令宝箱（ID：3000521）*30、刷新卡选择礼包*2
                String ids[] = {"769", "50", "2005254", "1", "1021029", "1", "3000521", "30", "2218", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 30000) {
                //15000	绑元*8000、祝余年·宠（ID：3414）*1、祈丰年-相框*1、寻宝令宝箱（ID：3000521）*50、刷新卡选择礼包*3
                String ids[] = {"769", "80", "3414", "1", "1020035", "1", "3000521", "50", "2218", "3"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 50000) {
                //30000	绑元*12000、祝余年·魂（ID：2005148）*1、神石选择礼包*1、防具圣装选择礼盒*1、高级红品装备宝箱*2（650级可更换神武打造符*2）
                String ids[] = {"769", "120", "2005148", "1", "3000219", "1", "1403012", "1", "3100034", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 80000){
                //50000	绑元*20000、祝余年·翼（ID：1050428）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "200", "1050428", "1", "3000219", "2", "1403019", "1", "3100034", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else {
                //80000	绑元*30000、堕火凤凰（ID：2112025）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "300", "2112025", "1", "3000219", "2", "1403019", "1", "3100034", "3"};
                giftList.addAll(Arrays.asList(ids));
            }
        } else {
            if (intMoney < 3000) {
                //1000	绑元*500、巢穴怒气重置卡（ID：3120049）*1、魔王挑战礼盒*3
                String ids[] = {"769", "5", "3120049", "1", "3000202", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else if (intMoney < 8000) {
                //3000	绑元*1500、祝余年套装（ID：1011029，自动化识别性别）*1、巢穴等级限制解除卡（ID：3120047）*1、寻宝令宝箱（ID：3000521）*20
                String ids[] = {"769", "15", "1011029", "1", "3120047", "1", "3000521", "20"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 15000) {
                //8000	绑元*5000、祝余年·阵（ID：2005254）*1、岁岁年年-气泡*1、寻宝令宝箱（ID：3000521）*30、刷新卡选择礼包*2
                String ids[] = {"769", "50", "2005254", "1", "1021029", "1", "3000521", "30", "2218", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 30000) {
                //15000	绑元*8000、祝余年·宠（ID：3414）*1、祈丰年-相框*1、寻宝令宝箱（ID：3000521）*50、刷新卡选择礼包*3
                String ids[] = {"769", "80", "3414", "1", "1020035", "1", "3000521", "50", "2218", "3"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 50000) {
                //30000	绑元*12000、祝余年·魂（ID：2005148）*1、神石选择礼包*1、防具圣装选择礼盒*1、高级红品装备宝箱*2（650级可更换神武打造符*2）
                String ids[] = {"769", "120", "2005148", "1", "3000219", "1", "1403012", "1", "3900012", "2"};
                giftList.addAll(Arrays.asList(ids));

            } else if (intMoney < 80000){
                //50000	绑元*20000、祝余年·翼（ID：1050428）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "200", "1050428", "1", "3000219", "2", "1403019", "1", "3900012", "3"};
                giftList.addAll(Arrays.asList(ids));
            } else {
                //80000	绑元*30000、堕火凤凰（ID：2112025）、神石选择礼包*2、圣装选择礼盒*1、高级红品装备宝箱*3（650级可更换神武打造符*3）
                String ids[] = {"769", "300", "2112025", "1", "3000219", "2", "1403019", "1", "3900012", "3"};
                giftList.addAll(Arrays.asList(ids));
            }
        }

        int giftLevels[] = {1000, 3000, 8000, 15000, 30000, 50000,80000};

        Integer lastMoney = 0;

        for (int i = 0; i < giftLevels.length; i++) {
            if (giftLevels[i] > money) {
                lastMoney = giftLevels[i - 1];
                break;
            }
            lastMoney = 80000;
        }


        System.out.println(lastMoney);

        String reason = /*day + */"春节线下活动" + lastMoney + "档位";
        String emailTitle = /*day + */"春节福利活动" + lastMoney + "档位";
        String emailContent = "亲爱的仙友，您的充值奖励已发放，请留意查收~";
        List<String> list = setSubRow(district, roleId, reason, emailTitle, emailContent);

        list.addAll(giftList);
        return list;
    }


}
