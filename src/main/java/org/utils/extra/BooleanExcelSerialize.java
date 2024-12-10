package org.utils.extra;

import org.utils.AbstractExcelSerialize;

public class BooleanExcelSerialize extends AbstractExcelSerialize {


    @Override
    public Object getExcelObject(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof Boolean) {
            return ((Boolean) value) ? "是" : "否";
        } else if (value instanceof String) {
            String string = (String) value;
            if ("1".equals(string) || "true".equalsIgnoreCase(string) || "Y".equalsIgnoreCase(string)) {
                return "是";
            } else if ("0".equals(string) || "false".equalsIgnoreCase(string) || "N".equalsIgnoreCase(string)) {
                return "否";
            }
        }
        return value;
    }
}
