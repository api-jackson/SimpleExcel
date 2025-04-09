package org.utils;

import com.google.common.collect.BiMap;
import com.google.common.collect.HashBiMap;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.CollectionUtils;
import org.utils.pageable.PageableData;

import java.lang.reflect.Field;
import java.util.*;

@Slf4j
public class PageableExcelExportUtils {

    private static final Integer DEFAULT_PAGE_SIZE = 100;

    public static <T> void exportSheet(Workbook workbook, Class<T> klass, String sheetName, String scene, PageableData pageableData) {
        try {
            List<Field> exportFieldList = FieldUtils.getFieldsListWithAnnotation(klass, ExcelExport.class);
            Map<ExcelExport, Field> excelExportFieldMap = filterScene(exportFieldList, scene);
            BiMap<Field, ExcelExport> excelExportFieldBiMap = HashBiMap.create(excelExportFieldMap).inverse();
            Set<Map.Entry<ExcelExport, Field>> excelExportSet = excelExportFieldMap.entrySet();

            Map<String, List> exportListMap = generateExportList(excelExportSet);
            pageableExportExcel(exportListMap, workbook, sheetName, pageableData, excelExportFieldBiMap);

        } catch (IllegalAccessException e) {
            log.error("导出"+sheetName+"异常", e);
        }
    }

    /**
     * 取出各种list，进行导出处理
     * @param exportListMap
     * @param workbook
     * @param sheetName
     */
    private static <T> void pageableExportExcel(Map<String, List> exportListMap, Workbook workbook, String sheetName,
                                           PageableData<T> pageableData, BiMap<Field, ExcelExport> excelExportFieldMap) throws IllegalAccessException {
        List<String> titleList = exportListMap.get("titleList");
        List<PropertyNameType> propertyNameTypeList = exportListMap.get("propertyList");

        Sheet sheet = workbook.createSheet(sheetName);
        Row titleRow = sheet.createRow(0);
        buildTitle(titleRow, titleList);
        int rownum = 1;
        int pagenum = 0;
        List<T> data = pageableData.getData(pagenum++, DEFAULT_PAGE_SIZE);
        while (!CollectionUtils.isEmpty(data)) {

            List<Map<String, Object>> rowObjectList = getMapObjectList(data, excelExportFieldMap) ;

            for (int i=0; i<data.size(); i++) {  // 行处理
                Row row = sheet.createRow(rownum++);
                Map<String, Object> rowObject = rowObjectList.get(i);
                for (int column=0; column < propertyNameTypeList.size(); column++) {  // 行中每个单元格处理
                    PropertyNameType propertyNameType = propertyNameTypeList.get(column);
                    Cell cell = row.createCell(column, propertyNameType.getPropertyType());
                    Object exportObject = rowObject.get(propertyNameType.getPropertyName());
                    if (rowObject.get(propertyNameType.getPropertyName()) == null) {  // 对象为空，跳过
                        continue;
                    }
                    cell.setCellValue(exportObject.toString());
                }
            }
            data = pageableData.getData(pagenum++, DEFAULT_PAGE_SIZE);
        }
    }

    /**
     * 生成标题行titleList，Excel属性propertyList
     * @param excelExportSet 有序set，根据ExcelExport注解的order属性进行排序
     * @return
     * @param <T>
     * @throws IllegalAccessException
     */
    private static <T> Map<String, List> generateExportList(Set<Map.Entry<ExcelExport, Field>> excelExportSet) throws IllegalAccessException {



        List<String> titleList = new ArrayList<>();
        List<Field> fieldList = new ArrayList<>();
        List<PropertyNameType> propertyNameTypeList = new ArrayList<>();

        buildPropertyTypeAndTitleList(excelExportSet, fieldList, titleList, propertyNameTypeList);

        Map<String, List> exportExcelListMap = new HashMap<>();
        exportExcelListMap.put("titleList", titleList);
        exportExcelListMap.put("propertyList", propertyNameTypeList);
        return exportExcelListMap;

    }

    /**
     * 过滤出指定场景的ExcelExport注解
     * @param fieldList
     * @param sceneParam
     * @return
     */
    private static Map<ExcelExport, Field> filterScene(List<Field> fieldList, String sceneParam) {
        Map<ExcelExport, Field> sceneFieldMap = new TreeMap<>(new Comparator<ExcelExport>() {
            @Override
            public int compare(ExcelExport o1, ExcelExport o2) {
                return o1.order()- o2.order();
            }
        });

        for (Field field : fieldList) {  // Field Loop
            ExcelExport[] excelExportArray = field.getAnnotationsByType(ExcelExport.class);
            excelExportAnnotationLoop:
            for (ExcelExport excelExport : excelExportArray) {  // Annotation ExcelExport Loop
                for (String sceneInField : excelExport.scene()) {  //  ExcelExport.scene Array Loop
                    if (sceneInField.equals(sceneParam)) {
                        sceneFieldMap.put(excelExport, field);
                        break excelExportAnnotationLoop;
                    }
                }
            }
        }
        return sceneFieldMap;
    }

    private static List<PropertyNameType> buildPropertyTypeAndTitleList(Set<Map.Entry<ExcelExport, Field>> excelExportSet, List<Field> fieldList,
                                                                        List<String> titleList, List<PropertyNameType> propertyNameTypeList) {

        for (Map.Entry<ExcelExport, Field> entry : excelExportSet) {
            ExcelExport excelExport = entry.getKey();
            Field field = entry.getValue();
            fieldList.add(field);
            titleList.add(excelExport.title());
            propertyNameTypeList.add(new PropertyNameType(field.getName(), excelExport.cellType()));
        }
        return propertyNameTypeList;
    }

    /**
     * 构建导出Excel标题行
     * @param row
     * @param titleList
     */
    private static void buildTitle(Row row, List<String> titleList) {
        for (int i=0; i<titleList.size(); i++) {
            String title = titleList.get(i);
            Cell cell = row.createCell(i, CellType.STRING);
            cell.setCellValue(title);
        }
    }



    private static <T> List<Map<String, Object>> getMapObjectList(List<T> queryList, BiMap<Field, ExcelExport> excelExportFieldBiMap) throws IllegalAccessException {
        Set<Field> fields = excelExportFieldBiMap.keySet();
        List<Field> fieldList = new ArrayList<>(fields);
        List<Map<String, Object>> exportList = new ArrayList<>();
        for (T t : queryList) {
            Map<String, Object> objectMap = new HashMap<>();
            for (Field field : fieldList) {
                field.setAccessible(true);
                Object object = field.get(t);
                object = ExcelExportUtils.processBySerialize(object, excelExportFieldBiMap.get(field));
                objectMap.put(field.getName(), object);
            }
            exportList.add(objectMap);
        }
        return exportList;
    }

}
