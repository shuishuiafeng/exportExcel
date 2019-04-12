package com.xiaofeng.pro.common.utils;

import com.xiaofeng.pro.common.annotation.ExcelField;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URLEncoder;
import java.util.*;

/**
 * @Auther: xiaofeng
 * @Date: 2019/3/2 13:31
 * @Description:
 */
@Slf4j
public class ExportExcelUtil {
    private static final String DEFAULT_URL_ENCODING = "UTF-8";

    /**
     * 注解列表,可能一个字段上包含多个注解（List中元素的类型Object[]{ ExcelField, Field/Method }）
     */
    List<Object[]> annotationlist = new ArrayList<>();

    /**
     * 工作簿对象
     */
    private XSSFWorkbook xssfWorkbook;

    /**
     * sheet对象
     */
    private XSSFSheet xssfSheet;

    /**
     * 样式列表
     */
    private Map<String, CellStyle> styles;

    /**
     * 当前行号，不给设置值，默认也是0
     */
    private int rownum = 0;

    /**
     * 构造函数
     *
     * @param title 表格标题，传“空值”，表示无标题
     * @param cls   实体对象，通过annotation.ExportField获取标题
     */
    public ExportExcelUtil(String title, Class<?> cls) {
        this(title, cls, 1);
    }

    /**
     * 构造函数
     *
     * @param title  表格标题，传“空值”，表示无标题
     * @param cls    实体对象类型
     * @param type   导出类型（1:导出数据；2：导出模板）
     * @param groups 导入导出分组（设定了此groups中group的字段才会导出
     *               eg:某个字段上的注解为@ExcelField(groups=[2,3]),此构造方法的groups为3,4，则此字段因为包含了3所以会导出此字段）
     */
    public ExportExcelUtil(String title, Class<?> cls, int type, int... groups){
        /**
         * 理一下开发思路
         * 1. 根据cls找到当前类上的ExcelField注解
         * 2. 根据注解可以找到对应字段上的title字段标题、sort排序、左右对齐等信息
         * 3. 然后创建excel表格，带上title的信息作为头，带上字段title作为列头
         * 还需要注意不仅仅在字段上有，在method上也会有
         */
        /** 获取带有注解的字段 */
        Field[] clsFields = cls.getDeclaredFields();
        for (Field field:
                clsFields) {
            // 获取针对于ExcelField的注解信息
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if(excelField != null && (excelField.type() == 0 || excelField.type() == type)){
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : excelField.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationlist.add(new Object[]{excelField, field}); // 仅添加符合group要求的字段到注解列表中
                                break;
                            }
                        }
                    }
                } else {
                    annotationlist.add(new Object[]{excelField, field}); // 添加所有加了注解的字段到注解列表中
                }
            }
        }

        /** 获取带有注解的方法 */
        Method[] ms = cls.getDeclaredMethods();
        for (Method m:
                ms) {
            ExcelField excelField = m.getAnnotation(ExcelField.class);
            if(excelField != null && (excelField.type() == 0 || excelField.type() == type)){
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : excelField.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationlist.add(new Object[]{excelField, m});
                                break;
                            }
                        }
                    }
                } else {
                    annotationlist.add(new Object[]{excelField, m});
                }
            }
        }

        /** 字段排序 , 按照ExcelField的sort()值排序*/
        annotationlist.sort(Comparator.comparing(o -> ((ExcelField) o[0]).sort()));

        /** 初始化表头信息 */
        List<String> headerList = new ArrayList<>();
        for (Object[] ob : annotationlist) {
            String t = ((ExcelField) ob[0]).title(); // 获取注解中的title标题
            if(type == 1){
                // 有可能title上有**注释，要除去
                String[] ss = StringUtils.split(t, "**",2);
                if(ss.length == 2){
                    t = ss[0];
                }
            }
            headerList.add(t);
        }
        initialize(title, headerList);
    }

    /**
     * 初始化表信息
     * @param title 表格标题头，就是表头标题
     * @param excelFields  列头标题列表
     */
    private void initialize(String title, List<String> excelFields){
        // 创建workbook
        this.xssfWorkbook = new XSSFWorkbook();
        // 创建sheet
        this.xssfSheet = this.xssfWorkbook.createSheet();
        // 创建样式
        this.styles = createStyles(xssfWorkbook);
        //sheet顶部头信息
        if (StringUtils.isNotBlank(title)) {
            XSSFRow titleRow = this.xssfSheet.createRow(rownum++); // 创建表头行
            titleRow.setHeightInPoints(30);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get("title"));
            titleCell.setCellValue(title);
            if(excelFields.size() > 1){ // 当列多余1列的情况则表头需要 合并单元格
                this.xssfSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),titleRow.getRowNum(),0,excelFields.size()-1));
            }
        }

        if (excelFields == null) {
            throw new RuntimeException("headerList not null!");
        }
        /** 添加列头行 */
        XSSFRow headerRow = this.xssfSheet.createRow(rownum++);
        headerRow.setHeightInPoints(16);

        for(int i = 0 ; i < excelFields.size() ; i++){
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellStyle(styles.get("header"));
            String[] ss = StringUtils.split(excelFields.get(i), "**", 2);
            if(ss.length == 2){
                cell.setCellValue(ss[0]);
                // 注释批注 : XSSFClientAnchor画斜线
                Comment comment = this.xssfSheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            }else{
                cell.setCellValue(excelFields.get(i));
            }
            xssfSheet.autoSizeColumn(i);
        }

        // 将每一列的宽度设置成其两倍长度，如果两倍长度小于3000则设为3000
        for (int i = 0; i < excelFields.size(); i++) {
            int colWidth = xssfSheet.getColumnWidth(i) * 2;
            xssfSheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);
        }
    }

    /**
     * 设置excel的数据列表
     * @param datalist 传入数据列表
     * @param <E> 实体类的类型
     * @return
     * 梳理一下此方法的功能
     * 当搜索的时候要判断对象的某个字段和注解、和字段对应上,理一下流程
     * 1. 遍历所有的datalist，获取其中的一个E，也就是要导出的实体对象
     * 2. 将实体对象中的字段与annotationlist中的field对应起来，通过反射调用对象中的字段
     * 3. 获取当前这个field中的那个excelfield所在的列，不用麻烦了，而且我也没找到根据名称获取当前column的，反正annotationlist是按照顺序设置的表头，所以这里就顺序获取就可以了
     * 4. 遍历一条数据，包含多个字段，每个都要获取出来，相当于对一条数据进行整个annotation的遍历，获取到想要的值
     */
    public <E> ExportExcelUtil setExcelData(List<E> datalist){
        for (E data:
             datalist) {
            XSSFRow row = this.addRow();
            StringBuilder sb = new StringBuilder();
            int column = 0;
            for (Object[] objArr:
                 annotationlist) {
                ExcelField ef = (ExcelField) objArr[0];
                Object val = null;
                /** 获取实体中字段对应的值
                 * 1. 先从注解ExcelField的value()中获取
                 * 2. 如果没设置则从注解字段的名称调用其get方法（方法则直接调用对应本身方法）
                 * */
                try {
                    if (StringUtils.isNotBlank(ef.value())) {
                        val = ReflectionUtil.invokeGetter(data, ef.value());
                    } else {
                        if (objArr[1] instanceof Field) {
                            val = ReflectionUtil.invokeGetter(data, ((Field) objArr[1]).getName());
                        } else if (objArr[1] instanceof Method) {
                            val = ReflectionUtil.invokeMethod(data, ((Method) objArr[1]).getName(), new Class[]{}, new Object[]{});
                        }
                    }
                    // 如果设置了这个注解字段为枚举类，则获取枚举的标签名称
                    if (Class.class != ef.enumType()) {
                        val = ExportEnumUtils.getEnumName(val == null ? "" : val.toString(), ef.enumType(), "");
                    }
                } catch (Exception ex) {
                    // Failure to ignore
                    log.info(ex.toString());
                    val = "";
                }
                // 添加单元格
                this.addCell(row, column++, val, ef.align(), ef.fieldType(), ef.format());
                sb.append(val).append(", ");
            }
            log.debug("Write success: [" + row.getRowNum() + "] " + sb.toString());
        }
        return this;
    }

    /**
     * 将表格输出到客户端
     * @param response
     * @param fileName
     * @return
     */
    public ExportExcelUtil write(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("application/vnd.ms-excel; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename*=utf-8'zh_cn'" + URLEncoder.encode(fileName, DEFAULT_URL_ENCODING));
//        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, DEFAULT_URL_ENCODING));
        write(response.getOutputStream());
        return this;
    }

    /**
     * 输出数据流
     *
     * @param os 输出数据流
     */
    public ExportExcelUtil write(OutputStream os) throws IOException {
        xssfWorkbook.write(os);
        os.flush();
        os.close();
        return this;
    }

    /**
     * 清理临时文件
     */
    public void dispose() {
        //wb.dispose();
    }
    /**
     * 添加一个单元格
     *
     * @param row       添加的行
     * @param column    添加列号
     * @param val       添加值
     * @param align     对齐方式（1：靠左；2：居中；3：靠右）
     * @param fieldType 注解的类型 主要是针对时间、价钱的处理
     * @param format    注解的格式
     * @return 单元格对象
     */
    private Cell addCell(XSSFRow row, int column, Object val, int align, Class<?> fieldType, String format){
        Cell cell = row.createCell(column);
        String cellFormatString = "@";
        try {
            if (val == null) {
                cell.setCellValue("");
            } else if (fieldType == Date.class) {
                if ((Long) val == 0L) {
                    cell.setCellValue("");
                } else {
                    Date temVal = new Date((Long) val);
                    cell.setCellValue(temVal);
                }

                cellFormatString = StringUtils.defaultIfEmpty(format, "yyyy-MM-dd");
            } else if (fieldType == Money.class) {
                BigDecimal hundred = new BigDecimal(100);
                BigDecimal result;
                if (val instanceof Long) {
                    result = new BigDecimal((Long) val).divide(hundred, 2, RoundingMode.HALF_UP);
                } else if (val instanceof BigDecimal) {
                    result = ((BigDecimal) val).divide(hundred, 2, RoundingMode.HALF_UP);
                } else {
                    result = new BigDecimal((Integer) val).divide(hundred, 2, RoundingMode.HALF_UP);
                }
                cell.setCellValue(result.toString());
                /*cellFormatString = StringUtils.defaultIfEmpty(format, "0.00");*/
            } else if (fieldType != Class.class) {
                cell.setCellValue((String) fieldType.getMethod("setValue", Object.class).invoke(null, val));
            } else {
                if (val instanceof String) {
                    cell.setCellValue((String) val);
                } else if (val instanceof Integer) {
                    cell.setCellValue((Integer) val);
                    cellFormatString = "0";
                } else if (val instanceof Long) {
                    cell.setCellValue((Long) val);
                    cellFormatString = "0";
                } else if (val instanceof Double) {
                    cell.setCellValue((Double) val);
                    cellFormatString = "0.00";
                } else if (val instanceof Float) {
                    cell.setCellValue((Float) val);
                    cellFormatString = "0.00";
                } else if (val instanceof Date) {
                    try {
                        cell.setCellValue((Date) val);
                    } catch (Exception e) {
                        Date temVal = new Date((Long) val);
                        cell.setCellValue(temVal);
                    }

                    cellFormatString = StringUtils.defaultIfEmpty(format, "yyyy-MM-dd");
                } else {
                    cell.setCellValue(
                            (String) Class.forName(
                                    this.getClass().getName().replaceAll(
                                            this.getClass().getSimpleName(),
                                            "fieldtype." + val.getClass().getSimpleName() + "Type")
                            ).getMethod("setValue", Object.class).invoke(null, val));
                }
            }
            CellStyle style;
            if (val != null) {
                style = styles.get("data_column_val_" + column);
                if (style == null) {
                    style = xssfWorkbook.createCellStyle();
                    style.cloneStyleFrom(styles.get("data" + (align >= 1 && align <= 3 ? align : "")));
                    style.setDataFormat(xssfWorkbook.createDataFormat().getFormat(cellFormatString));
                    styles.put("data_column_val_" + column, style);
                }
            } else {
                style = styles.get("data_column_" + column);
                if (style == null) {
                    style = xssfWorkbook.createCellStyle();
                    style.cloneStyleFrom(styles.get("data" + (align >= 1 && align <= 3 ? align : "")));
                    style.setDataFormat(xssfWorkbook.createDataFormat().getFormat(cellFormatString));
                    styles.put("data_column_" + column, style);
                }
            }
            cell.setCellStyle(style);
        } catch (Exception ex) {
            log.error("Set cell value [" + row.getRowNum() + "," + column + "] error: " + ex.toString());
            cell.setCellValue(val.toString());
        }
        return cell;
    }

    private XSSFRow addRow(){
        return this.xssfSheet.createRow(rownum++);
    }

    /**
     * 创建表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        /** 创建表头标题样式 */
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
        style.setFont(titleFont);
        styles.put("title", style);

        /** 创建数据单元格样式 */
        style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        /** 设置左对齐单元格样式 */
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.LEFT);
        styles.put("data1", style);

        /** 设置居中对齐单元格样式 */
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("data2", style);

        /** 设置右对齐单元格样式 */
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("data3", style);

        /** 设置列头单元格样式 */
        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        return styles;
    }
}
