package com.coderlucifar.excel;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExplicitConstraint {
    // 定义固定下拉内容
    String[] source() default {};
    // 定义动态下拉内容，
    Class[] sourceClass() default {};
    // 列标号必须和字段下标一致
    int indexNum() default 0;
    // 字典type
    String type() default "";
}
