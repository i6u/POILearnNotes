import org.apache.commons.beanutils.BeanUtils;
import org.junit.Test;
import zyr.learn.poi.model.User;
import zyr.learn.poi.util.ExcelTemplate;
import zyr.learn.poi.util.ExcelUtil;

import java.util.*;

/**
 * Created by zhouweitao on 2016/12/4.
 *
 */
public class ExcelUtilTest {

    @Test
    public void test01() throws Exception {
        ExcelTemplate excelTemplate = ExcelTemplate.newInstance();

        ExcelTemplate et = excelTemplate.readTemplateByPath("/Users/zhouweitao/Desktop/temp/b.xlsx");
        et.createNewRow();
        et.createCell("111");
        et.createCell("a");
        et.createCell("A");
        et.createCell("001");
        et.createNewRow();
        et.createCell("222");
        et.createCell("b");
        et.createCell("B");
        et.createCell("002");
        et.createNewRow();
        et.createCell("333");
        et.createCell("c");
        et.createCell("C");
        et.createCell("003");
        et.createNewRow();
        et.createCell("444");
        et.createCell("d");
        et.createCell("D");
        et.createCell("004");
        et.createNewRow();
        et.createCell(555);
        et.createCell(false);
        et.createCell(Calendar.DAY_OF_YEAR);
        et.createCell(12.14d);
        et.createNewRow();
        et.createCell(666);
        et.createCell(24.53d);
        et.createCell(Calendar.WEEK_OF_YEAR);
        et.createCell(006);
        Map<String, String> info = new HashMap<>();
        info.put("title", "测试学生信息表");
        info.put("date", new Date().toString());
        info.put("dep", "天下为公");

        et.replaceFinalData(info);
        et.insertNO();

        et.writeToFIle("/Users/zhouweitao/Desktop/temp/a1.xlsx");

    }


    @Test
    public void test02() throws Exception {

        List<User> users = new ArrayList<User>();
        users.add(new User(1,"悟空","a",22));
        users.add(new User(2,"八戒","b",23));
        users.add(new User(3,"沙僧","c",18));
        users.add(new User(4,"和尚","d",27));
        users.add(new User(5,"如来","e",31));

        Map<String, String> parm = new HashMap<>();
        parm.put("title", "西游记");
        parm.put("date", "1000");
        parm.put("dep", "最强四人组");

        ExcelUtil eu = ExcelUtil.newInstance();
//        eu.exportObj2ExcelByTemplate(parm,"/Users/zhouweitao/Desktop/temp/c.xlsx","/Users/zhouweitao/Desktop/temp/c111.xlsx",users, User.class,false);
//        eu.exportObj2Excel("/Users/zhouweitao/Desktop/temp/aaa.xlsx",users,User.class,true);
        eu.exportObj2Excel("/Users/zhouweitao/Desktop/temp/aaa.xls",users,User.class,false);
    }


    @Test
    public void test03() throws Exception {
        ExcelUtil eu = ExcelUtil.newInstance();
        List<Object> objects = eu.readExcel2ObjsByFilePath("/Users/zhouweitao/Desktop/temp/aaa.xls",User.class);
        for (Object o : objects) {
            System.out.println(o);
        }

    }

    @Test
    public void testRe() throws Exception {
        System.out.println(ExcelUtilTest.class.getResourceAsStream(""));

    }

    @Test
    public void beanUtilsTest() throws Exception {
        Class clazz = User.class;
        Object obj = clazz.newInstance();
        BeanUtils.copyProperty(obj, "username", "张三");
        String str = BeanUtils.getProperty(obj, "username");
        System.out.println(str);

    }
}
