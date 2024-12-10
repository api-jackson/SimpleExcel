package org.utils.extra;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.Date;

public class DateTimeUtils {

    public static String PATTERN_DATETIME_MINI = "yyyyMMddHHmmss";

    public static LocalDateTime getMaxTimeOfDay(LocalDateTime localDateTime) {
        if (localDateTime != null) {
            LocalDateTime maxLocalDateTime = localDateTime.withHour(23).withMinute(59).withSecond(59);
            return maxLocalDateTime;
        }
        return null;
    }

    public static LocalDateTime getMinTimeOfDay(LocalDateTime localDateTime) {
        if (localDateTime != null) {
            LocalDateTime minLocalDateTime = localDateTime.withHour(0).withMinute(0).withSecond(0);
            return minLocalDateTime;
        }
        return null;
    }

    /**
     * 日期转字符串
     *
     * @param date   日期对象
     * @param format 日期格式
     * @return 日期字符串
     */
    public static String date2String(Date date, String format) {
        if (date == null) {
            return null;
        }
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }

}
