package pers.chao.document.helper.annontation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author: WangYichao
 * @description: Excel导出列注解
 * @date: 2018/8/3 22:03d
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {

    /**
     * 列名
     * @return
     */
    String value();

    /**
     * 列顺序
     * @return
     */
    int order() default Integer.MIN_VALUE;

}
