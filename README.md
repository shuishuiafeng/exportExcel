# exportExcel
Java + Spring Boot + POI + 反射机制实现的excel导出

## 先在pom.xml中增加poi的依赖(版本根据情况去设置即可)
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>3.17</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.17</version>
</dependency>

## 简单代码文件类初步讲解一下
### 一、注解类
（1） 注解类ExcelField，这个注解类将被注解到要生成excel的实体类的字段上，比如生成订单实体对应的excel文件，需要导出其中的字段orderCode（订单码）、orderName(订单名称），那么这个ExcelField注解就是放到这两个字段上添加；
### 二、工具类
（1）首先就是进行导出操作的工具类ExportExcelUtil了，实现产生表头、表体内容和表格输出等操作功能；
（2）ReflectionUtil反射工具类，辅助于上一个ExportExcelUtil的，用户获取注解对应的字段的方法进行执行等等；
（3）ExportEnumUtils是配合ExcelField中的枚举字段的导出显示；
（4）Money配合ExportExcelUtil针对金钱的显示；
（5）最后是没有测试的ImportExcelUtil这个导入excel的工具；
### 三、实际使用相关类
（1）用于导出的实体类Organization,在这个类的字段上添加注解ExcelField决定此字段是否导出；
（2）在ExcelController中调用工具类的导出功能测试实验；
