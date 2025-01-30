package org.utils;

import com.google.common.collect.BiMap;
import com.google.common.collect.HashBiMap;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author Jackson
 * @date 2024/12/28
 * @description
 */
@Slf4j
public class ExcelExportUtils {

    /******************** 导出处理方法  start ********************/

    private static void buildTitle(Row row, List<String> titleList) {
        for (int i=0; i<titleList.size(); i++) {
            String title = titleList.get(i);
            Cell cell = row.createCell(i, CellType.STRING);
            cell.setCellValue(title);
        }
    }

    private static void getExcelBytes(List<PropertyNameType> propertyTypeMapList, List<Map<String, Object>> exportObjectMapList,
                                      List<String> titleList, Workbook workbook, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row titleRow = sheet.createRow(0);
        buildTitle(titleRow, titleList);
        for (int rownum = 0; rownum<exportObjectMapList.size(); rownum++) {  // 行处理
            Row row = sheet.createRow(rownum+1);
            Map<String, Object> rowObject = exportObjectMapList.get(rownum);
            for (int column=0; column < propertyTypeMapList.size(); column++) {  // 行中每个单元格处理
                PropertyNameType propertyNameType = propertyTypeMapList.get(column);
                Cell cell = row.createCell(column, propertyNameType.getPropertyType());
                Object exportObject = rowObject.get(propertyNameType.getPropertyName());
                if (rowObject.get(propertyNameType.getPropertyName()) == null) {  // 对象为空，跳过
                    continue;
                }
                cell.setCellValue(exportObject.toString());
            }
        }
    }



//    public static void download(HttpServletResponse response, Workbook workbook, String filename) throws IOException {
//        int surfixIndex = filename.lastIndexOf(".");
//        String dateString = DateTimeUtils.date2String(new Date(), DateTimeUtils.PATTERN_DATETIME_MINI);
//        String nowFilename = new String(filename.substring(0, surfixIndex)+"_"+dateString+"."+filename.substring(surfixIndex+1));
//        File file = new File(nowFilename);
//        file.createNewFile();
//        try (FileOutputStream fos = new FileOutputStream(file)) {
//            workbook.write(fos);
//            fos.flush();
//        } catch (Exception e) {
//            throw e;
//        }
//
//        FileInputStream inputStream = new FileInputStream(file);
//        try (OutputStream outputStream = new BufferedOutputStream(response.getOutputStream())) {
//            byte[] bytes = StreamUtils.copyToByteArray(inputStream);
//            String exportFilename = new String(nowFilename.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
//            response.addHeader(HttpHeaders.CONTENT_DISPOSITION,
//                    "attachment;filename=" + exportFilename);
//            response.addHeader("Content-length", String.valueOf(bytes.length));
//            response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
//            outputStream.write(bytes);
//            outputStream.flush();
//        } catch (Exception e) {
//            throw e;
//        } finally {
//            if (inputStream != null) {
//                inputStream.close();
//            }
//        }
//        file.delete();
//        response.getOutputStream();
//        outputStream.write(bytes);
//        response.flushBuffer();
//    }

    private static <T> Map<String, List> generateExportList(List<T> queryList, Class<T> klass, String scene) throws IllegalAccessException {
        List<Field> exportFieldList = FieldUtils.getFieldsListWithAnnotation(klass, ExcelExport.class);
        Map<ExcelExport, Field> excelExportFieldMap = filterScene(exportFieldList, scene);
        BiMap<Field, ExcelExport> excelExportFieldBiMap = HashBiMap.create(excelExportFieldMap).inverse();
        Set<Map.Entry<ExcelExport, Field>> excelExportSet = excelExportFieldMap.entrySet();

        List<String> titleList = new ArrayList<>();
        List<PropertyNameType> propertyNameTypeList = new ArrayList<>();
        List<Map<String, Object>> exportList = new ArrayList<>();
        List<Field> fieldList = new ArrayList<>();

        for (Map.Entry<ExcelExport, Field> entry : excelExportSet) {
            ExcelExport excelExport = entry.getKey();
            Field field = entry.getValue();
            fieldList.add(field);
            titleList.add(excelExport.title());
            propertyNameTypeList.add(new PropertyNameType(field.getName(), excelExport.cellType()));
        }


        for (T t : queryList) {
            Map<String, Object> objectMap = new HashMap<>();
            for (Field field : fieldList) {
                field.setAccessible(true);
                Object object = field.get(t);
                object = processBySerialize(object, excelExportFieldBiMap.get(field));
                objectMap.put(field.getName(), object);
            }
            exportList.add(objectMap);
        }

        Map<String, List> exportExcelListMap = new HashMap<>();
        exportExcelListMap.put("titleList", titleList);
        exportExcelListMap.put("propertyList", propertyNameTypeList);
        exportExcelListMap.put("objectList", exportList);
        return exportExcelListMap;

    }

    private static Object processBySerialize(Object object, ExcelExport excelExport) {

        object = handleDateFormatter(object, excelExport);

        final boolean emptyBeanName = StringUtils.isBlank(excelExport.serializeBeanName());
        final boolean emptyBeanClass = excelExport.serializeBeanName() == null || excelExport.serializeBeanClass().equals(void.class);
        if (emptyBeanName && emptyBeanClass) {  // beanName 且 beanClass 都为空
            return object;
        }
        if (!emptyBeanName) {  // beanName 不为空，优先取beanName
            return processBySerializeBeanName(object, excelExport);
        } else {  // beanClass 不为空
            return processBySerializeBeanClass(object, excelExport);
        }
    }

    private static Object processBySerializeBeanName(Object object, ExcelExport excelExport) {
        final AbstractExcelSerialize serializeBean = (AbstractExcelSerialize) SpringBeanUtils.getBean(excelExport.serializeBeanName());
        return serializeBean.getExcelObject(object);
    }

    private static Object processBySerializeBeanClass(Object object, ExcelExport excelExport) {
        final AbstractExcelSerialize serializeBean = (AbstractExcelSerialize) SpringBeanUtils.getBean(excelExport.serializeBeanClass());
        return serializeBean.getExcelObject(object);
    }

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


    private static void processExcelSheetByListMap(Map<String, List> exportListMap, Workbook workbook, String sheetName) {
        List<String> titleList = exportListMap.get("titleList");
        List<PropertyNameType> propertyNameTypeList = exportListMap.get("propertyList");
        List<Map<String, Object>> objectList = exportListMap.get("objectList");
        ExcelExportUtils.getExcelBytes(propertyNameTypeList, objectList, titleList, workbook, sheetName);
    }

    public static <T> void exportSheet(Workbook workbook, List<T> dataList, Class<T> klass, String sheetName, String scene) {
        try {
            Map<String, List> exportListMap = ExcelExportUtils.generateExportList(dataList, klass, scene);
            ExcelExportUtils.processExcelSheetByListMap(exportListMap, workbook, sheetName);
        } catch (IllegalAccessException e) {
            log.error("导出"+sheetName+"异常", e);
        }
    }

    public static <T> void exportSheet(Workbook workbook, List<T> dataList, Class<T> klass, String sheetName) {
        exportSheet(workbook, dataList, klass, sheetName, "");
    }


    private static Object handleDateFormatter(Object object, ExcelExport excelExport) {
        if (object instanceof Date) {
            Date date = (Date) object;
            return DateFormatUtils.format(date, StringUtils.defaultIfEmpty(excelExport.dateFormat(), "yyyy-MM-dd HH:mm:ss"));
        } else if (object instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) object;
            return DateTimeFormatter.ofPattern(StringUtils.defaultIfEmpty(excelExport.dateFormat(), "yyyy-MM-dd HH:mm:ss")).format(localDateTime);
        } else if (object instanceof LocalDate) {
            LocalDate localDate = (LocalDate) object;
            return DateTimeFormatter.ofPattern(StringUtils.defaultIfEmpty(excelExport.dateFormat(), "yyyy-MM-dd")).format(localDate);
        } else if (object instanceof LocalTime) {
            LocalTime localTime = (LocalTime) object;
            return DateTimeFormatter.ofPattern(StringUtils.defaultIfEmpty(excelExport.dateFormat(), "HH:mm:ss")).format(localTime);
        } else {
            return object;
        }
    }

    /******************** 导出处理方法  end ********************/


}
