package top.wu.parsegameexcel;

public enum GiftType {
    TOTAL_MONEY_GIFT(0),
    SINGLE_DAY_GIFT(1),
    NEW_AUTH_GIFT(2);

    private Integer type;

    GiftType(Integer type) {
        this.type = type;
    }

    public Integer getType() {
        return type;
    }
}
