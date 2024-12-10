package org.utils.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.utils.ExcelUtils;
import org.utils.extra.WorkbookHandler;
import org.utils.test.configuration.ExcelToolsConfiguration;
import org.utils.test.vo.SampleVO;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.PriorityQueue;

/**
 * @author Jackson
 * @date 2024/11/30
 * @description
 */

public class ExportTest {

    @Test
    public void test() throws IOException {
        List<SampleVO> sampleVOList = getSampleVOList();
        AnnotationConfigApplicationContext context = new AnnotationConfigApplicationContext();
        context.register(ExcelToolsConfiguration.class);
        context.scan("org.utils");
        context.refresh();
        Workbook workbook = new XSSFWorkbook();
        ExcelUtils.exportSheet(workbook, sampleVOList, SampleVO.class, "测试");
        WorkbookHandler.saveExcelToFile(workbook, "test.xlsx");
        context.close();


    }

    private List<SampleVO> getSampleVOList() {
        List<SampleVO> sampleVOList = new ArrayList<>();

        SampleVO.SampleVOBuilder test1 = SampleVO.builder().sampleString("test1").sampleDouble(1.1).sampleDate(LocalDateTime.now()).sampleBoolean(true);
        sampleVOList.add(test1.build());

        SampleVO.SampleVOBuilder test2 = SampleVO.builder().sampleString("test2").sampleDouble(2.2).sampleDate(LocalDateTime.now().minusDays(1)).sampleBoolean(false);
        sampleVOList.add(test2.build());

        return sampleVOList;

    }

}
