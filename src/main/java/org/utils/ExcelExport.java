package org.utils;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(ExcelExports.class)
@Documented
public @interface ExcelExport {


    String title();

    CellType cellType();

    String[] scene() default {""};

    int order();

    String serializeBeanName() default "";

    Class serializeBeanClass() default void.class;


}
