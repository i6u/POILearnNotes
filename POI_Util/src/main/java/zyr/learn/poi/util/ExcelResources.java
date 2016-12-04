package zyr.learn.poi.util;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * Created by zhouweitao on 2016/12/4.
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelResources {
    String title();
    int order() default 99999;
}
