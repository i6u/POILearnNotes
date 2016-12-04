package zyr.learn.poi.model;

import zyr.learn.poi.util.ExcelResources;

/**
 * Created by zhouweitao on 2016/12/4.
 */
public class User {
    private int uid;
    private String username;
    private String nickname;
    private int age;


    public User() {
    }

    public User(int uid, String username, String nickname, int age) {
        this.uid = uid;
        this.username = username;
        this.nickname = nickname;
        this.age = age;
    }

    @ExcelResources(title = "用户标识",order = 1)
    public int getUid() {
        return uid;
    }

    public void setUid(int uid) {
        this.uid = uid;
    }

    @ExcelResources(title = "用户名称",order = 2)
    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    @ExcelResources(title = "用户昵称",order = 3)
    public String getNickname() {
        return nickname;
    }

    public void setNickname(String nickname) {
        this.nickname = nickname;
    }

    @ExcelResources(title = "用户年龄",order = 4)
    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return "User{" +
                "uid=" + uid +
                ", username='" + username + '\'' +
                ", nickname='" + nickname + '\'' +
                ", age=" + age +
                '}';
    }
}
