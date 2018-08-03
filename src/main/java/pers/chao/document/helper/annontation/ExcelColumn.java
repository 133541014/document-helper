package pers.chao.document.helper.annontation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author: WangYichao
 * @description: Excel������ע��
 * @date: 2018/8/3 22:03d
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {

    /**
     * ����
     * @return
     */
    String value();

    /**
     * ��˳��
     * @return
     */
    int order() default Integer.MIN_VALUE;

}
