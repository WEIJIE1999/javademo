package com.wj.boot;

import cn.hutool.core.date.DateUtil;
import com.wj.bean.User;
import com.wj.utils.ExportUtils;
import org.apache.poi.ss.formula.functions.T;

import java.util.*;

/**
 * @author weijie
 * @create 2021-01-07 下午 23:33
 */
public class Main {
    public static void main(String[] args) {

        List<User> list = new ArrayList<>();
        list.add(new User("zhangsan", "1231", new Date()));
        list.add(new User("zhangsan1", "1232", new Date()));
        list.add(new User("zhangsan2", "1233", new Date()));
        list.add(new User("zhangsan3", "1234", new Date()));
        list.add(new User("zhangsan4", "1235", new Date()));
        list.add(new User("zhangsan5", "1236", DateUtil.date(new Date())));

        Map map = new HashMap<String, String>(8);
        map.put("name", "姓名");
        map.put("age", "年龄");
        map.put("birthDay", "生日");

        ExportUtils.export("D:\\Users",list,map,2,"申请人员信息","申请学院");
    }
}
