package org.utils.test.vo;

import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellType;
import org.utils.ExcelExport;
import org.utils.ExcelImport;
import org.utils.extra.BooleanToYNString;

import java.time.LocalDateTime;

/**
 * @author Jackson
 * @date 2024/11/29
 * @description
 */
@Data
public class SampleVO {


    @ExcelExport(title = "测试字符串", cellType = CellType.STRING, order = 1)
    @ExcelImport(order = 1)
    public String sampleString;

    @ExcelExport(title = "测试日期", cellType = CellType.STRING, order = 2)
    @ExcelImport(order = 2)
    public LocalDateTime sampleDate;

    @ExcelExport(title = "测试数字", cellType = CellType.NUMERIC, order = 3)
    @ExcelImport(order = 3)
    public Double sampleDouble;

    @ExcelExport(title = "测试布尔", cellType = CellType.STRING, order = 4, serializeBeanClass = BooleanToYNString.class)
    @ExcelImport(order = 4)
    public Boolean sampleBoolean;

}
