package top.wu.parsegameexcel;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ParsegameexcelApplicationTests {

    @Test
    void contextLoads() {
        ParsegameexcelApplicationTests.class.getClassLoader().getResourceAsStream("static/jjl_total_money_gift.properties");
    }

}
