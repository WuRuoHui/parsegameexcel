package top.wu.parsegameexcel.comparator;

import java.util.Comparator;

public class MyComparator implements Comparator<String> {

    public int compare(String o1, String o2) {
        return Integer.valueOf(o1) - Integer.valueOf(o2);
    }
}
