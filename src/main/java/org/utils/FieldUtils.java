package org.utils;

import org.apache.commons.lang3.Validate;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class FieldUtils extends org.apache.commons.lang3.reflect.FieldUtils {

    public static List<Field> getFieldsListWithAnnotation(final Class<?> cls, final Class<? extends Annotation> annotationCls) {
        Validate.notNull(annotationCls, "annotationCls");
        final List<Field> allFields = getAllFieldsList(cls);
        final List<Field> annotatedFields = new ArrayList<>();
        for (final Field field : allFields) {
            if (field.getAnnotationsByType(annotationCls).length != 0) {
                annotatedFields.add(field);
            }
        }
        return annotatedFields;
    }

}
