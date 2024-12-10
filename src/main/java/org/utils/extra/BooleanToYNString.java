package org.utils.extra;

import org.utils.AbstractExcelSerialize;

/**
 * @author Jackson
 * @date 2024/11/29
 * @description
 */
public class BooleanToYNString extends AbstractExcelSerialize {
    @Override
    public Object getExcelObject(Object value) {
        if (!(value instanceof Boolean)) {
            return value;
        }
        Boolean bool = (Boolean) value;
        if (bool) {
            return "Y";
        } else {
            return "N";
        }
    }
}
