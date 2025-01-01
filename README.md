# SimpleExcel


## 导出Excel工具类 ExcelExportUtils

### 相关类说明：
  - 导出工具类：`org.utils.ExcelExportUtils`
  - 导出字段注解：`org.utils.ExcelExport`
  - 导出字段注解集合，用于在同一个字段上标注多个ExcelExport：`org.utils.ExcelExports`

### 可以使用的方法：
1. 根据顺序导出字段数据，可设置导出字段的标题
2. 根据特定场景导出字段的值
3. 根据字段值转换导出的值（内置 Boolean -> 是否，0/1 -> 是否）

### `org.utils.ExcelExport` 相关类参数说明


  - `title(String)`: String，生成 Excel 表格的标题

  - `cellType(CellType)`: CellType枚举类，该列数据填充的数据类型，当前有 _NONE，BLANK，BOOLEAN（常用），ERROR，FORMULA，NUMERIC（常用），STRING（常用，日期等）

  - `order(int)`: int，该字段导出列的排序

  - `scene(String[])`: String数组，场景编码数组，在字段上标识该列。调用方通过传入不同的场景编码，来决定该注解是否生效
    - 使用方法：
      > 
      > 假设字段定义如下：
      > ```java
      > @ExcelExport(title="field1-noscene", cellType=CellType.STRING, order=1)
      > private String field1;
      > 
      > @ExcelExport(title="field2-scene1", cellType=CellType.STRING, order=1, scene={"scene1"})
      > private String field2;
      > 
      > @ExcelExport(title="field3-noscene", cellType=CellType.STRING, order=2)
      > @ExcelExport(title="field3-scene1", cellType=CellType.STRING, order=2, scene={"scene1"})
      > private String field3;
      > ```
      > 
      > ---
      >
      > 调用方1
      >
      > ```java  
      > ExcelExportUtils.exportSheet(exportWorkbook, sampleVOList, SampleVO.class, "测试", "");
      > ```
      >
      > 导出结果：
      > | field1-noscene | field3-noscene |
      >
      > ---
      >
      > 调用方2
      >
      > ```java  
      > ExcelExportUtils.exportSheet(exportWorkbook, sampleVOList, SampleVO.class, "测试", "scene1");
      > ```
      >
      > 导出结果：
      > | field2-scene1 | field3-scene1 |

  - `serializeBeanName(String)`: 使用该Bean对此字段进行序列化，优先取beanName

  - `serializeBeanClass(Class)`: 使用该Bean对此字段进行序列化，当 beanName 和 beanClass 同时存在时，优先取beanName
    - 使用方法：
      > 
      > 假设字段定义如下：
      > ```java
      > @ExcelExport(title="booleanfield1", cellType=CellType.STRING, order=1, serializeBeanClass=BooleanToYNString.class)
      > private Boolean booleanfield1;
      >
      > @ExcelExport(title="booleanfield2", cellType=CellType.STRING, order=2, serializeBeanClass=BooleanExcelSerialize.class
      > private Boolean booleanfield2;
      >
      > ```
      >
      > 导出结果：
      >
      > | booleanfield1 | booleanfield2 |
      >
      > | Y | 是 |
      >
      > | N | 否 |





## 导入Excel工具类 ExcelImportUtils

### 相关类说明：
  - 导出工具类：`org.utils.ExcelImportUtils`
  - 导出字段注解：`org.utils.ExcelImport`
  - 导出字段注解集合，用于在同一个字段上标注多个ExcelExport：`org.utils.ExcelImports`

### 可以使用的方法：
1. 根据顺序，从Excel中取字段导入到list中
2. 根据场景选择导入字段

### 相关参数说明
  - `order(int)`: int，导入Excel中该列的位置排序

  - `scene(String[])`: String数组，场景编码数组，在字段上标识该列。调用方通过传入不同的场景编码，来决定该注解是否生效
    - 使用方法：
      > 
      > 假设字段定义如下：
      > ```java
      > @ExcelImport(order=1)
      > private String field1;
      > 
      > @ExcelImport(order=1, scene={"import-scene1"})
      > private String field2;
      > 
      > @ExcelImport(order=2)
      > @ExcelImport(order=2, scene={"import-scene1"})
      > private String field3;
      > ```
      > 
      > ---
      > 导入Excel 1: 
      > 
      > | field1-noscene | field3-noscene |
      >
      > 调用方1
      >
      > ```java  
      > ExcelImportUtils.getObjectListFromExcel(importWorkbook, SampleVO.class, "");
      > ```
      >
      >
      > ---
      >
      >
      > 导入Excel 2：
      >
      > | field2-scene1 | field3-scene1 |
      >
      > 调用方2
      >
      > ```java  
      > ExcelExportUtils.exportSheet(importWorkbook, SampleVO.class, "import-scene1");
      > ```

