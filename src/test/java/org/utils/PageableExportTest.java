package org.utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.utils.extra.WorkbookHandler;
import org.utils.test.configuration.ExcelToolsConfiguration;
import org.utils.test.vo.SampleVO;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class PageableExportTest {




    @Test
    public void test() throws IOException {
        // 导出测试
        List<SampleVO> sampleVOList = getSampleVOList();
        AnnotationConfigApplicationContext context = new AnnotationConfigApplicationContext();
        context.register(ExcelToolsConfiguration.class);
        context.scan("org.utils");
        context.refresh();
        Workbook exportWorkbook = new XSSFWorkbook();
        PageableExcelExportUtils.exportSheet(exportWorkbook, SampleVO.class, "测试", "", (pageNum, pageSize) -> {
            if (pageNum < 10) {
                return Arrays.asList(sampleVOList.get(pageNum % sampleVOList.size()));
            }
            return Collections.EMPTY_LIST;
        });

        String filePath = WorkbookHandler.saveExcelToFile(exportWorkbook, "test.xlsx");

        // 导入测试
//        Workbook importWorkbook = ExcelImportUtils.determineWorkbook(new File(filePath));
//        List<SampleVO> importVOList = ExcelImportUtils.getObjectListFromExcel(importWorkbook, SampleVO.class);


        context.close();


    }


    private List<SampleVO> getSampleVOList() {
        List<SampleVO> sampleVOList = new ArrayList<>();

        SampleVO test1 = new SampleVO();
        test1.setSampleString("test1");
        test1.setSampleDouble(1.1);
        test1.setSampleDate(LocalDateTime.now());
        test1.setSampleBoolean(true);
        sampleVOList.add(test1);

        SampleVO test2 = new SampleVO();
        test2.setSampleString("test2");
        test2.setSampleDouble(2.2);
        test2.setSampleDate(LocalDateTime.now().minusDays(1));
        test2.setSampleBoolean(false);
        sampleVOList.add(test2);

        return sampleVOList;

    }

}
