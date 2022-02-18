package top.wu.parsegameexcel.utils;


import top.wu.parsegameexcel.comparator.MyComparator;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class PropUtils {

    private final static String TOTAL_GIFT_NAME = "static/jll_total_money_gift.properties";

    //王者射击累计充值档位与ID对应
    private final static String WZSJ_TOTAL_GIFT_NAME = "static/wzsj_total_money_gift.properties";


    private final static String SINGLE_GIFT_LT_30 = "static/jll_single_gift_lt_30.properties";
    private final static String SINGLE_GIFT_MT_30 = "static/jll_single_gift_mt_30.properties";

    private final static String SINGLE_GIFT_LT_7 = "static/jll_single_gift_lt_7.properties";
    private final static String SINGLE_GIFT_MT_7 = "static/jll_single_gift_mt_7.properties";

    private final static String WZSJ_SINGLE_GIFT = "static/wzsj_single_gift.properties";




    private static Properties properties = new Properties();


    //获得所有累计充值档位的ID List
    public static ArrayList<String> getTotalMoneyGift() {
        return getAllPropIds(TOTAL_GIFT_NAME);
    }

    //获得王者射击所有累计充值档位的ID List
    public static ArrayList<String> getWZSJTotalMoneyGift() {
        return getAllPropIds(WZSJ_TOTAL_GIFT_NAME);
    }



    //根据键获得累计充值档位的ID
    public static String getTotalMoneyId(String key) {
        return getPropId(TOTAL_GIFT_NAME, key);
    }

    //根据键获得王者射击累计充值档位的ID
    public static String getWZSJTotalMoneyId(String key) {
        return getPropId(WZSJ_TOTAL_GIFT_NAME, key);
    }

    //获取所有开服天数小于或等于30天的单日充值档位ID List
    public static ArrayList<String> getSingleGiftLt30() {
        return getAllPropIds(SINGLE_GIFT_LT_30);
    }

    //根据键获得当日充值档位的ID（小于或等于30天）
    public static String getSingleIdLt30(String key) {
        return getPropId(SINGLE_GIFT_LT_30, key);
    }

    //获取所有开服小于或等于7天的档位充值档位ID List
    public static ArrayList<String> getSingleGiftLt7() {
        return getAllPropIds(SINGLE_GIFT_LT_7);
    }

    //获取王者射击档位充值档位ID List
    public static ArrayList<String> getSWZSJingleGift() {
        return getAllPropIds(WZSJ_SINGLE_GIFT);
    }

    //根据键获得当日充值档位的ID（小于或等于7天）
    public static String getSingleIdLt7(String key) {
        return getPropId(SINGLE_GIFT_LT_7, key);
    }

    public static String getWZSJSingleId(String key) {
        return getPropId(WZSJ_SINGLE_GIFT, key);
    }

    //获取所有开服天数大于30天的单日充值档位ID List
    public static ArrayList<String> getSingleGiftMt30() {
        return getAllPropIds(SINGLE_GIFT_MT_30);
    }

    //根据键获得当日充值档位的ID（大于30天）
    public static String getSingleIdMt30(String key) {
        return getPropId(SINGLE_GIFT_MT_30, key);
    }

    //获取所有开服天数大于7天的单日充值档位ID List
    public static ArrayList<String> getSingleGiftMt7() {
        return getAllPropIds(SINGLE_GIFT_MT_7);
    }

    //根据键获得当日充值档位的ID（大于7天）
    public static String getSingleIdMt7(String key) {
        return getPropId(SINGLE_GIFT_MT_7, key);
    }

    //根据配置文件名称获取所有的键
    private static ArrayList<String> getAllPropIds(String propFileName) {
        InputStream is = PropUtils.class.getClassLoader().getResourceAsStream(propFileName);
        try {
            properties.load(is);
        } catch (IOException e) {
            e.printStackTrace();
        }

        //将键的Set转换为List再进行排序
        Set<Object> set = properties.keySet();

        ArrayList<String> ids = new ArrayList(set);

        Collections.sort(ids, new MyComparator());
        try {
            properties.clear();
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return ids;
    }

    //根据配置文件名称和键获取值（即获取档位道具ID）
    private static String getPropId(String propFileName, String key) {
        InputStream is = PropUtils.class.getClassLoader().getResourceAsStream(propFileName);
        try {
            properties.load(is);
        } catch (IOException e) {
            e.printStackTrace();
        }

        String id = properties.getProperty(key);

        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        properties.clear();

        return id;
    }


    //获取累计充值奖励的key值（即档位的列表）dist：累计充值金额  isSingleKey：是否返回单个元素的列表（当为累计充值时返回单个元素）
    public static List<String> getSubTotalKeys(String dist, Boolean isSingleKey) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }

        ArrayList<String> totalMoneyGiftIds = getTotalMoneyGift();
        int index = 0;

        if (isSingleKey) {
            if (Integer.valueOf(dist) >= 1000000)
                return totalMoneyGiftIds.subList(totalMoneyGiftIds.size() - 1, totalMoneyGiftIds.size());

            for (String id : totalMoneyGiftIds) {
                if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                    index = totalMoneyGiftIds.indexOf(id);
                    break;
                }
            }
            return totalMoneyGiftIds.subList(index - 1, index);
        }

        if (Integer.valueOf(dist) >= 1000000) return totalMoneyGiftIds;
        for (String id : totalMoneyGiftIds) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = totalMoneyGiftIds.indexOf(id);
                break;
            }
        }
        return totalMoneyGiftIds.subList(0, index);
    }

    //获取王者射击累计充值奖励的key值（即档位的列表）dist：累计充值金额  isSingleKey：是否返回单个元素的列表（当为累计充值时返回单个元素）
    public static List<String> getWZSJSubTotalKeys(String dist, Boolean isSingleKey) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }

        ArrayList<String> totalMoneyGiftIds = getWZSJTotalMoneyGift();
        System.out.println("id"+totalMoneyGiftIds);
        int index = 0;

        if (isSingleKey) {
            if (Integer.valueOf(dist) >= 300000)
                return totalMoneyGiftIds.subList(totalMoneyGiftIds.size() - 1, totalMoneyGiftIds.size());

            for (String id : totalMoneyGiftIds) {
                if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                    index = totalMoneyGiftIds.indexOf(id);
                    break;
                }
            }
            return totalMoneyGiftIds.subList(index - 1, index);
        }

        if (Integer.valueOf(dist) >= 300000) return totalMoneyGiftIds;
        for (String id : totalMoneyGiftIds) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = totalMoneyGiftIds.indexOf(id);
                break;
            }
        }
        return totalMoneyGiftIds.subList(0, index);
    }

    public static String getSingleSubTotalIds(String dist) {

        ArrayList<String> totalMoneyGiftIds = getTotalMoneyGift();
        int index = 0;
        if (Integer.valueOf(dist) >= 1000000) return totalMoneyGiftIds.get(totalMoneyGiftIds.size() - 1);

        for (String id : totalMoneyGiftIds) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = totalMoneyGiftIds.indexOf(id);
                break;
            }
        }
        return totalMoneyGiftIds.get(index - 1);
    }

    //获取王者射击单日充值档位
    public static String getWZSJSingleSubTotalIds(String dist) {

        //狩猎单日充档位
//        String[] singleDayLevel = {"1000","2000","3000","5000","7000","10000","15000","20000","30000","50000"};
        ArrayList<String> totalMoneyGiftIds = getWZSJTotalMoneyGift();
//        List<String> totalMoneyGiftIds = Arrays.asList(singleDayLevel);

        int index = 0;
        if (Integer.valueOf(dist) >= 30000) return totalMoneyGiftIds.get(totalMoneyGiftIds.size() - 1);

        for (String id : totalMoneyGiftIds) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = totalMoneyGiftIds.indexOf(id);
                break;
            }
        }
        return totalMoneyGiftIds.get(index - 1);
    }

    //获取狩猎使命单日充值档位
    public static String getSLSMSingleSubTotalIds(String dist) {

        //狩猎单日充档位
        String[] singleDayLevel = {"1000","2000","3000","5000","7000","10000","15000","20000","30000","50000"};
                List<String> totalMoneyGiftIds = Arrays.asList(singleDayLevel);

//        ArrayList<String> totalMoneyGiftIds = getWZSJTotalMoneyGift();
        int index = 0;
        if (Integer.valueOf(dist) >= 500000) return totalMoneyGiftIds.get(totalMoneyGiftIds.size() - 1);

        for (String id : totalMoneyGiftIds) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = totalMoneyGiftIds.indexOf(id);
                break;
            }
        }
        return totalMoneyGiftIds.get(index - 1);
    }

    public static List<String> getSubSingleKeys(String dist) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }
        ArrayList<String> singleGiftKeys = getSingleGiftLt30();
        int index = 0;

        if (Integer.valueOf(dist) >= 50000) return singleGiftKeys;
        for (String id : singleGiftKeys) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = singleGiftKeys.indexOf(id);
                break;
            }
        }
        return singleGiftKeys.subList(0, index);
    }

    //12.10版本
    public static String getSubSingleKeyNew(String dist) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }
        ArrayList<String> singleGiftKeys = getSingleGiftLt7();
        int index = 0;

        if (Integer.valueOf(dist) >= 50000) return singleGiftKeys.get(singleGiftKeys.size()-1);
        for (String id : singleGiftKeys) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = singleGiftKeys.indexOf(id);
                break;
            }
        }
        System.out.println(singleGiftKeys.get(index - 1));
        return singleGiftKeys.get(index-1);
    }

    public static String getWZSJSubSingleKeyNew(String dist) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }
        ArrayList<String> singleGiftKeys = getSWZSJingleGift();
        int index = 0;

        if (Integer.valueOf(dist) >= 30000) return singleGiftKeys.get(singleGiftKeys.size()-1);
        for (String id : singleGiftKeys) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = singleGiftKeys.indexOf(id);
                break;
            }
        }
        System.out.println(singleGiftKeys.get(index - 1));
        return singleGiftKeys.get(index-1);
    }

    //狩猎使命
    public static String getSLSMSubSingleKeyNew(String dist) {
        if (Integer.valueOf(dist) < 1000) {
            System.out.println("warning：数值错误，请检查");
            System.exit(0);
        }

        String[] singleDayLevel = {"1000","2000","3000","5000","7000","10000","15000","20000","30000","50000"};


        List<String> singleGiftKeys = Arrays.asList(singleDayLevel);
        int index = 0;

        if (Integer.valueOf(dist) >= 50000) return singleGiftKeys.get(singleGiftKeys.size()-1);
        for (String id : singleGiftKeys) {
            if (Integer.valueOf(dist) < Integer.valueOf(id)) {
                index = singleGiftKeys.indexOf(id);
                break;
            }
        }
        System.out.println(singleGiftKeys.get(index - 1));
        return singleGiftKeys.get(index-1);
    }
}
