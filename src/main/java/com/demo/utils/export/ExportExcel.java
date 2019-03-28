package com.demo.utils.export;


import com.demo.utils.ConstantUtil;
import com.demo.utils.Reflections;
import com.demo.utils.export.annotation.ExcelField;
import com.google.common.collect.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.util.*;
import java.util.zip.ZipOutputStream;

/**
 * 导出Excel文件（导出“XLSX”格式，支持大数据量导出   @see org.apache.poi.ss.SpreadsheetVersion）
 */
public class ExportExcel {

    private static Logger log = LoggerFactory.getLogger(ExportExcel.class);

    /**
     * 工作薄对象
     */
    private SXSSFWorkbook wb;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 样式列表
     */
    private Map<String, CellStyle> styles;

    /**
     * 当前行号
     */
    private int rownum;

    /**
     * 数组集合
     */
    private List<String[]> dataStringList;

    /**
     * 数组排序
     */
    private List<Integer> dataSort;

    /**
     * 注解列表（Object[]{ ExcelField, Field/Method }）
     * 所属对象的每一列注释列表的属性
     */
    List<Object[]> annotationList = Lists.newArrayList();

    public List<String> headerList = Lists.newArrayList();


    public List<Integer> reds = Lists.newArrayList();

    public ExportExcel() {
    }

    /**
     * 构造函数
     *
     * @param title 表格标题，传“空值”，表示无标题
     * @param cls   实体对象，通过annotation.ExportField获取标题
     */
    public ExportExcel(String title, String anno, Class<?> cls) {
        this(title, cls, 1, anno);
    }

    /**
     * 没有校区，使用默认
     *
     * @param title
     * @param cls
     * @param isHide
     * @param type
     * @param anno
     * @param groups
     */
    public ExportExcel(String title, Class<?> cls, Integer isHide, int type, String anno, int... groups) {
        this(title, cls, isHide, null, type, anno, groups);
    }

    /**
     * 构造函数
     *
     * @param title  表格标题，传“空值”，表示无标题
     * @param cls    实体对象，通过annotation.ExportField获取标题
     * @param type   导出类型（1:导出数据；2：导出模板）
     * @param groups 导入分组
     */
    public ExportExcel(String title, Class<?> cls, Integer isHide, Integer isMultiCampus, int type, String anno, int... groups) {
        if (isHide == null) {
            isHide = 0;
        }
        if (isMultiCampus == null) {
            isMultiCampus = 0;
        }
        // Get annotation field
        Field[] fs = cls.getDeclaredFields();
        for (Field f : fs) {
            ExcelField ef = f.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == 0 || ef.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : ef.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationList.add(new Object[]{ef, f});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{ef, f});
                }
            }
        }
        // Get annotation method
        Method[] ms = cls.getDeclaredMethods();
        for (Method m : ms) {
            ExcelField ef = m.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == 0 || ef.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : ef.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationList.add(new Object[]{ef, m});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{ef, m});
                }
            }
        }
        // Field sorting
        Collections.sort(annotationList, new Comparator<Object[]>() {
            @Override
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(
                        new Integer(((ExcelField) o2[0]).sort()));
            }
        });
        // Initialize
        List<String> headerList = Lists.newArrayList();

        List<Integer> redList = Lists.newArrayList();
        for (Object[] os : annotationList) {
            ExcelField field = (ExcelField) os[0];
            String t = field.title();
            // 如果是导出，则去掉注释
            if (type == 1) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            if (isHide == 0 && field.hide() == 1) {
                continue;
            }
            if (isMultiCampus == 0 && field.isMultiCampus() == 1) {
                continue;
            }
            redList.add(field.isnull());
            headerList.add(t);

        }
        initialize(title, headerList, anno, redList, type, null);
    }

    public ExportExcel(String title, Class<?> cls, int type, String anno, int... groups) {
        this(title, cls, 0, type, anno, groups);
    }

    /**
     * 重载修改下载列表调用方法，添加hide
     * isMultiCampus的标识不好使，下拉框的数据验证使用下标的方式对应，如果不现实，和sort对应的下标对应不上，太麻烦了，现在是不好使的
     */
    public ExportExcel(int type, Class<?> cls, int isHide, int isMultiCampus, int... groups) {
        // Get annotation field
        Field[] fs = cls.getDeclaredFields();
        for (Field f : fs) {
            ExcelField ef = f.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == 0 || ef.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : ef.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationList.add(new Object[]{ef, f});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{ef, f});
                }
            }
        }
        // Get annotation method
        Method[] ms = cls.getDeclaredMethods();
        for (Method m : ms) {
            ExcelField ef = m.getAnnotation(ExcelField.class);
            if (ef != null && (ef.type() == 0 || ef.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : ef.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationList.add(new Object[]{ef, m});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{ef, m});
                }
            }
        }
        // Field sorting
        Collections.sort(annotationList, new Comparator<Object[]>() {

            @Override
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(
                        new Integer(((ExcelField) o2[0]).sort()));
            }
        });
        // Initialize
        headerList = Lists.newArrayList();
        reds = Lists.newArrayList();
        for (Object[] os : annotationList) {
            String t = ((ExcelField) os[0]).title();
            int red = ((ExcelField) os[0]).isnull();
            int hide = ((ExcelField) os[0]).hide();
            int multi = ((ExcelField) os[0]).isMultiCampus();
            // 如果是导出，则去掉注释
            if (type == 1) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            if (isHide == 0 && hide == 1) {
                continue;
            }
            if (isMultiCampus == 0 && multi == 1) {
                continue;
            }
            reds.add(red);
            headerList.add(t);
        }
    }

    /**
     * 修改下拉列表调用方法
     *
     * @param cls
     * @param type
     * @param groups
     */
    public ExportExcel(Class<?> cls, int type, int... groups) {
        this(type, cls, 0, 0, groups);
    }

    public ExportExcel(int type, Class<?> cls, String isHide, int... groups) {
        this(type, cls, Integer.parseInt(isHide), 0, groups);
    }


    /**
     * 构造函数
     *
     * @param title   表格标题，传“空值”，表示无标题
     * @param headers 表头数组
     */
    public ExportExcel(String title, String[] headers, String anno) {
        initialize(title, Lists.newArrayList(headers), anno, null, null, null);
    }

    /**
     * 构造函数
     *
     * @param title      表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    public ExportExcel(String title, List<String> headerList, String anno, int type) {
        initialize(title, headerList, anno, null, type, null);
    }

    /**
     * 初始化函数
     *
     * @param title      表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     *                   注释anno
     *                   reds:
     */
    public void initialize(String title, List<String> headerList, String anno, List<Integer> reds, Integer type, String name) {
        this.wb = new SXSSFWorkbook(500);
        if (name == null) {
            name = "Export";
        }
        this.sheet = wb.createSheet(name);
        this.styles = createStyles(wb);
        // Create title
        if (StringUtils.isNotBlank(title)) {
            Row titleRow = sheet.createRow(rownum++);
            titleRow.setHeightInPoints(30);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get("title"));
            titleCell.setCellValue(title);
            if (headerList.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                        titleRow.getRowNum(), titleRow.getRowNum(), headerList.size() - 1));
            }
        }


        if (type == 2) {
            Row introRow = sheet.createRow(rownum++);
            introRow.setHeightInPoints(90);
            Cell introCell = introRow.createCell(0);
            introCell.setCellStyle(styles.get("intro"));
            introCell.setCellValue(new XSSFRichTextString(anno));       //HSSFRichTextString
            sheet.addMergedRegion(new CellRangeAddress(introRow.getRowNum(),
                    introRow.getRowNum(), introRow.getRowNum() - 1, headerList.size() - 1));
        }
        // Create header
        if (headerList == null) {
            throw new RuntimeException("headerList not null!");
        }
        Row headerRow = sheet.createRow(rownum++);
        headerRow.setHeightInPoints(16);

        Integer recordTotalCell = 0;
        for (int i = 0; i < headerList.size(); i++) {

            //==========================滑腻腻的分割线以下部分是LL所写  有问题找她哦===========================
            //这里是控制表头是否合并单元格的
            //创建单元格的时候   index包含头不包含尾 （和substring很像哦）
            String oneHead = headerList.get(i);
            Integer cells = 0;
            if (oneHead.indexOf(";") > 0) {
                String[] headAndCellCount = oneHead.split(";");
                String strCell = headAndCellCount[1];
                cells = Integer.parseInt(strCell);
            }

            //是否合单元格的判断
            Cell cell = null;
            if (i > 0) {
                cell = headerRow.createCell(recordTotalCell);
            } else {
                cell = headerRow.createCell(i);
            }
            if (cells > 0) {
                for (int m = 1; m < cells; m++) {
                    recordTotalCell++;
                    headerRow.createCell(recordTotalCell);
                    if (cells - 1 == m) {
                        recordTotalCell++;
                    }
                }
            } else {
                recordTotalCell++;
            }
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            String[] ss = StringUtils.split(headerList.get(i), "**", 2);
            if (ss.length == 2) {
                if (ss[0].indexOf(";") > 0) {
                    cell.setCellValue(ss[0].split(";")[0]);
                } else {
                    cell.setCellValue(ss[0]);
                }

                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            } else {
                if (ss[0].indexOf(";") > 0) {
                    cell.setCellValue(ss[0].split(";")[0]);
//                    cell.setCellValue(new XSSFRichTextString(ss[0].split(";")[0]));
                } else {
                    cell.setCellValue(new XSSFRichTextString(headerList.get(i)));
                }
            }


            if (cells > 0) {
                CellRangeAddress region = new CellRangeAddress(headerRow.getRowNum(),
                        headerRow.getRowNum(), recordTotalCell - cells, recordTotalCell - 1);
                sheet.addMergedRegion(region);
            }


            if (reds != null && reds.size() > 0) {
                if (reds.get(i) == 0) {
                    cell.setCellStyle(styles.get("header"));
                }
                if (reds.get(i) == 1) {
                    cell.setCellStyle(styles.get("header2"));
                }
            }

//            sheet.autoSizeColumn(i);
        }
        for (int i = 0; i < headerList.size(); i++) {
            int colWidth = sheet.getColumnWidth(i) * 2;
            sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);
        }

        // 遍历所有的属性值 获得当前属性值所在列
//        如果是下拉且是导出并且不是联动效果
//        则 获取 下拉框的值其中 dataSort 是 判断是否有需要联动的顺序值
//        dataStringList 是当前顺序下的 数组值
        for (Object[] os : annotationList) {
            ExcelField field = (ExcelField) os[0];
            Integer index = field.sort() - 1;
            if (field.isDropDown() == 1 && field.isLlinkage() == 0) {
                String[] dlData = field.dropDownList();
//               this.set
                if (dataSort != null && dataSort.size() != 0) {
                    if (dataSort.contains(index)) {
                        Integer i = dataSort.indexOf(index);
                        dlData = dataStringList.get(i);
                    }
                }
                if (type == 2) {
                    sheet.addValidationData(setDataValidation(sheet, dlData, 3, 50000, index, index));
                } else {
                    sheet.addValidationData(setDataValidation(sheet, dlData, 2, 50000, index, index));
                }
            }

        }
        log.debug("Initialize success.");
    }

    public void cascade(String area, String firstList, String[] parentsArray, Map<String, String[]> childrenArrayMap, int rownum) {

//       todo: byfengzhi 2018/9/25 区别纯数字，研究以后优化
        for (int i = 0; i < parentsArray.length; i++) {
            if (ConstantUtil.firstDigit(parentsArray[i])) {
                parentsArray[i] = "_" + parentsArray[i];
            }
        }
        Map<String, String[]> _childrenArrayMap = new HashMap<>();
        List<String> keys = new ArrayList<>();
        int x = 0;
        for (String key : childrenArrayMap.keySet()) {
            if (ConstantUtil.firstDigit(key)) {
                String[] array = childrenArrayMap.get(key);
                keys.add(key);
                key = "_" + key;
                _childrenArrayMap.put(key, array);


            }
        }
        childrenArrayMap.putAll(_childrenArrayMap);
        for (String v : keys) {
            childrenArrayMap.remove(v);
        }
//        ---------------分割线---------------------
        //创建一个专门用来存放地区信息的隐藏sheet页
        //因此也不能在现实页之前创建，否则无法隐藏。
        Sheet hideSheet = wb.createSheet(area);
        //这一行作用是将此sheet隐藏，功能未完成时注释此行,可以查看隐藏sheet中信息是否正确
        wb.setSheetHidden(wb.getSheetIndex(hideSheet), true);

        int rowId = 0;
        // 设置第一行，存省的信息
        Row provinceRow = hideSheet.createRow(rowId++);
        provinceRow.createCell(0).setCellValue(firstList);
        for (int i = 0; i < parentsArray.length; i++) {
            Cell provinceCell = provinceRow.createCell(i + 1);
            provinceCell.setCellValue(parentsArray[i]);
        }
        // 将具体的数据写入到每一行中，行开头为父级区域，后面是子区域。
        for (int i = 0; i < parentsArray.length; i++) {
            String key = parentsArray[i];

            String[] son = childrenArrayMap.get(key);
            Row row = hideSheet.createRow(rowId++);
            row.createCell(0).setCellValue(key);
            if (son != null) {
                for (int j = 0; j < son.length; j++) {
                    Cell cell = row.createCell(j + 1);
                    cell.setCellValue(son[j]);
                }
            }
            // 添加名称管理器
            Integer length = 0;
            if (son != null) {
                length = son.length;
            }
            String range = getRange(1, rowId, length);
            Name name = wb.createName();
            //key不可重复

            name.setNameName(key);

            String formula = area + "!" + range;
            name.setRefersToFormula(formula);
        }

        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        // 省规则
        if (parentsArray.length != 0) {
            DataValidationConstraint provConstraint = dvHelper.createExplicitListConstraint(parentsArray);
            // 四个参数分别是：起始行、终止行、起始列、终止列
//      注意不可以大于所设置联动 即 假如联动 是第三列第四列 则 不能设置从第四列开始
//        第一位是0
            CellRangeAddressList provRangeAddressList = new CellRangeAddressList(3, 50000, rownum - 1, rownum - 1);
            DataValidation provinceDataValidation = dvHelper.createValidation(provConstraint, provRangeAddressList);
            //验证
            provinceDataValidation.createErrorBox("error", area + ":联动有问题");
            provinceDataValidation.setShowErrorBox(true);
            provinceDataValidation.setSuppressDropDownArrow(true);
            sheet.addValidationData(provinceDataValidation);
            //对前20行设置有效性
            for (int i = 3; i < 20; i++) {
                setDataValidation(doHandle(rownum), sheet, i, rownum + 1);
            }
        }
    }


    /**
     * 创建表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    public Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(titleFont);
        styles.put("title", style);

        style = wb.createCellStyle();
        Font introFont = wb.createFont();
        introFont.setFontName("Arial");
        introFont.setFontHeightInPoints((short) 10);
        introFont.setColor(IndexedColors.RED.getIndex());
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(introFont);
        style.setWrapText(true);
        styles.put("intro", style);

        style = wb.createCellStyle();
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("data1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        styles.put("data2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        styles.put("data3", style);


        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font headerFont2 = wb.createFont();
        headerFont2.setFontName("Arial");
        headerFont2.setFontHeightInPoints((short) 10);
        headerFont2.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont2.setColor(IndexedColors.WHITE.getIndex());
        headerFont2.setColor(IndexedColors.RED.index);
        style.setFont(headerFont2);
        styles.put("header2", style);

        return styles;
    }

    /**
     * 添加一行
     *
     * @return 行对象
     */
    public Row addRow() {
        return sheet.createRow(rownum++);
    }


    /**
     * 添加一个单元格
     *
     * @param row    添加的行
     * @param column 添加列号
     * @param val    添加值
     * @return 单元格对象
     */
    public Cell addCell(Row row, int column, Object val) {
        return this.addCell(row, column, val, 0, Class.class);
    }

    /**
     * 添加一个单元格
     *
     * @param row    添加的行
     * @param column 添加列号
     * @param val    添加值
     * @param align  对齐方式（1：靠左；2：居中；3：靠右）
     * @return 单元格对象
     */
    public Cell addCell(Row row, int column, Object val, int align, Class<?> fieldType) {
        Cell cell = row.createCell(column);
        CellStyle style = styles.get("data" + (align >= 1 && align <= 3 ? align : ""));
        try {
            if (val == null) {
                cell.setCellValue("");
            } else if (val instanceof String) {
                cell.setCellValue((String) val);
            } else if (val instanceof Integer) {
                cell.setCellValue((Integer) val);
            } else if (val instanceof Long) {
                cell.setCellValue((Long) val);
            } else if (val instanceof Double) {
                cell.setCellValue((Double) val);
            } else if (val instanceof Float) {
                cell.setCellValue((Float) val);
            } else if (val instanceof Date) {
                DataFormat format = wb.createDataFormat();
                style.setDataFormat(format.getFormat("yyyy-MM-dd"));
                cell.setCellValue((Date) val);
            } else {
                if (fieldType != Class.class) {
                    cell.setCellValue((String) fieldType.getMethod("setValue", Object.class).invoke(null, val));
                } else {
                    cell.setCellValue((String) Class.forName(this.getClass().getName().replaceAll(this.getClass().getSimpleName(),
                            "fieldtype." + val.getClass().getSimpleName() + "Type")).getMethod("setValue", Object.class).invoke(null, val));
                }
            }
        } catch (Exception ex) {
            log.info("Set cell value [" + row.getRowNum() + "," + column + "] error: " + ex.toString());
            cell.setCellValue(val.toString());
        }
        cell.setCellStyle(style);
        return cell;
    }

    /**
     * 添加数据（通过annotation.ExportField添加数据）
     *
     * @return list 数据列表
     */
    public <E> ExportExcel setDataList(List<E> list) {
        for (E e : list) {
            int colunm = 0;
            Row row = this.addRow();
            StringBuilder sb = new StringBuilder();
            for (Object[] os : annotationList) {
                ExcelField ef = (ExcelField) os[0];
                Object val = null;
                // Get entity value
                try {
                    if (StringUtils.isNotBlank(ef.value())) {
                        val = Reflections.invokeGetter(e, ef.value());
                    } else {
                        if (os[1] instanceof Field) {
                            val = Reflections.invokeGetter(e, ((Field) os[1]).getName());
                        } else if (os[1] instanceof Method) {
                            val = Reflections.invokeMethod(e, ((Method) os[1]).getName(), new Class[]{}, new Object[]{});
                        }
                    }
                } catch (Exception ex) {
                    // Failure to ignore
                    log.info(ex.toString());
                    val = "";
                }
                this.addCell(row, colunm++, val, ef.align(), ef.fieldType());
                sb.append(val + ", ");
            }
            log.debug("Write success: [" + row.getRowNum() + "] " + sb.toString());
        }
        return this;
    }

    /**
     * 输出数据流
     *
     * @param os 输出数据流
     */
    public ExportExcel write(OutputStream os) throws IOException {
        wb.write(os);

        return this;
    }

    /**
     * 输出数据流
     * @param os 输出压缩文件流
     * @return
     * @throws IOException
     */
    public ExportExcel write(ZipOutputStream os) throws IOException {
        wb.write(os);
        return this;
    }
    /**
     * 输出到客户端
     *
     * @param fileName 输出文件名
     */
    public ExportExcel write(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("multipart/form-data");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
        write(response.getOutputStream());
        return this;
    }
    /**
     * 输出到压缩包
     *
     *
     * @param zipOutputStream 压缩文件流
     */
    public ExportExcel writeZipFile(ZipOutputStream zipOutputStream) throws  IOException {
        this.write(zipOutputStream);
        return this;
    }

    /**
     * 输出到文件
     *
     * @param name 输出文件名
     */
    public ExportExcel writeFile(String name) throws FileNotFoundException, IOException {
        FileOutputStream os = new FileOutputStream(name);
        this.write(os);
        return this;
    }

    /**
     * 清理临时文件
     */
    public ExportExcel dispose() {
        wb.dispose();
        return this;
    }

    /**
     * 根据headerList设置数据
     *
     * @param list
     * @param headerList
     * @param <E>
     * @return
     */
    public <E> ExportExcel setDataListByHeader(List<E> list, List<String> headerList) {
        for (E e : list) {
            int colunm = 0;
            Row row = this.addRow();
            StringBuilder sb = new StringBuilder();
            Object val = null;
            Method[] method = e.getClass().getDeclaredMethods();
            for (String head : headerList) {
                for (Method m : method) {
                    if (m.getName().indexOf("get") != -1) {
                        ExcelField myAnnotation = m.getAnnotation(ExcelField.class);
                        if (myAnnotation.title().equals(head)) {
                            val = Reflections.invokeMethod(e, m.getName(), new Class[]{}, new Object[]{});
                            this.addCell(row, colunm++, val, 2, e.getClass());
                            sb.append(val + ", ");
                        }
                    }
                }
            }

            log.debug("Write success: [" + row.getRowNum() + "] " + sb.toString());
        }
        return this;
    }

    //    设置下拉列表
    private static DataValidation setDataValidation(Sheet sheet, String[] textList, int firstRow, int endRow, int firstCol, int endCol) {

        DataValidationHelper helper = sheet.getDataValidationHelper();
        //加载下拉列表内容
        DataValidationConstraint constraint = helper.createExplicitListConstraint(textList);
        //DVConstraint constraint = new DVConstraint();
        constraint.setExplicitListValues(textList);

        //设置数据有效性加载在哪个单元格上。四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);

        //数据有效性对象
        DataValidation data_validation = helper.createValidation(constraint, regions);
        //DataValidation data_validation = new DataValidation(regions, constraint);

        return data_validation;
    }


    /**
     * 设置有效性
     *
     * @param offset 主影响单元格所在列，即此单元格由哪个单元格影响联动
     * @param sheet
     * @param rowNum 行数
     * @param colNum 列数
     */
    public static void setDataValidation(String offset, Sheet sheet, int rowNum, int colNum) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidation data_validation_list;
        data_validation_list = getDataValidationByFormula(
                "INDIRECT($" + offset + (rowNum) + ")", rowNum, colNum, helper);
        sheet.addValidationData(data_validation_list);
    }

    /**
     * 加载下拉列表内容
     *
     * @param formulaString
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @param dvHelper
     * @return
     */
    private static DataValidation getDataValidationByFormula(
            String formulaString, int naturalRowIndex, int naturalColumnIndex, DataValidationHelper dvHelper) {
        // 加载下拉列表内容
        // 举例：若formulaString = "INDIRECT($A$2)" 表示规则数据会从名称管理器中获取key与单元格 A2 值相同的数据，
        //如果A2是江苏省，那么此处就是江苏省下的市信息。
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createFormulaListConstraint(formulaString);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        int firstRow = naturalRowIndex - 1;
        int lastRow = naturalRowIndex - 1;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = naturalColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow,
                lastRow, firstCol, lastCol);
        // 数据有效性对象
        // 绑定
        XSSFDataValidation data_validation_list = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, regions);
        data_validation_list.setEmptyCellAllowed(false);
        if (data_validation_list instanceof XSSFDataValidation) {
            data_validation_list.setSuppressDropDownArrow(true);
            data_validation_list.setShowErrorBox(true);
        } else {
            data_validation_list.setSuppressDropDownArrow(false);
        }
        // 设置输入信息提示信息
        data_validation_list.createPromptBox("下拉选择提示", "请使用下拉方式选择合适的值！");
        // 设置输入错误提示信息
        //data_validation_list.createErrorBox("选择错误提示", "你输入的值未在备选列表中，请下拉选择合适的值！");
        return data_validation_list;
    }

    /**
     * 计算formula
     *
     * @param offset   偏移量，如果给0，表示从A列开始，1，就是从B列
     * @param rowId    第几行
     * @param colCount 一共多少列
     * @return 如果给入参 1,1,10. 表示从B1-K1。最终返回 $B$1:$K$1
     */
    public static String getRange(int offset, int rowId, int colCount) {
        char start = (char) ('A' + offset);
        if (colCount <= 25) {
            char end = (char) (start + colCount - 1);
            return "$" + start + "$" + rowId + ":$" + end + "$" + rowId;
        } else {
            char endPrefix = 'A';
            char endSuffix = 'A';
            if ((colCount - 25) / 26 == 0 || colCount == 51) {// 26-51之间，包括边界（仅两次字母表计算）
                if ((colCount - 25) % 26 == 0) {// 边界值
                    endSuffix = (char) ('A' + 25);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                }
            } else {// 51以上
                if ((colCount - 25) % 26 == 0) {
                    endSuffix = (char) ('A' + 25);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26 - 1);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26);
                }
            }
            return "$" + start + "$" + rowId + ":$" + endPrefix + endSuffix + "$" + rowId;
        }
    }

    //设置ABC
    private static String doHandle(final int num) {
        String[] charArr = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
                "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
                "W", "X", "Y", "Z"};
        return charArr[num - 1];
    }


    //下拉列表数据
    public void setDate(List<String[]> setData) {
        dataStringList = setData;
    }

    public List<String[]> getDate(List<String[]> setData) {
        return dataStringList = setData;
    }

    //下拉列表排序
    public void setSort(List<Integer> setNeedData) {
        dataSort = setNeedData;
    }

    public List<Integer> getSort(List<Integer> setNeedData) {
        return dataSort = setNeedData;
    }


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//    关于 Excel 表格 工具类用法举例
//    a)   基本用法 第三个参数中 2 代表  有注释 1代表无注释
//        new ExportExcel("导出失败信息", InsertBatchRingView.class, 1, "", 1).setDataList(exportFile).write(response, fileName).dispose();

//    b)  多sheet 表格     wb  必须new 一个新的 SXSSFWorkbook
//    需要多少个 sheet页 就需要new ExportExcelMultiSheet  [不支持下拉列表]
//        new ExportExcelMultiSheet(wb, "网关批量绑定", 0, "网关批量绑定", StationBindView.class, 2, anno, 1).setDataList(new ArrayList(), wb);
//        new ExportExcelMultiSheet(wb, "区域列表", 1, "区域列表", AreaNameView.class, 1, "").setDataList(areaList, wb);
//        new ExportExcelMultiSheet(wb, "教室列表", 2, "教室列表", ClassroomNameView.class, 1, "").setDataList(classroomList, wb);
//        new ExportExcelMultiSheet(wb, "示例数据", 3, "网关批量绑定模板", StationBindView.class, 2, anno, 1).setDataList(list, wb).write(wb, response, fileName).dispose(wb);

//    c)  下拉列表 用法  colNumList 这个list 集合代表 代表的是列表中第几列  ccc 这个数组代表的 某一列的数值
//         要求必须: 多列同时下拉列表时 colNumList 中的顺序 必须是和 arrayList 中 selectColData 的顺序是对应的
//          在实体类中必须开启isDropDown=1 当 其等于 0时 为关闭,默认为0, dropDownList 是下拉菜单中的值
//          当在控制层中没有重新为其赋值时  在实体类中写的值就是 下拉菜单的值
//          需要给多少列设置就需要存入多少个序号及对应的 字符串数组
//        List<Integer> colNumList = new ArrayList<>();
    //      下标从0开始
//        colNumList.add(6);
//        String[] selectColData = {"1","2","3","4","5","6"};
//        List<String[]> arrayList = new ArrayList<>();
//        arrayList.add(selectColData);
//        ExportExcel exportExcel = new ExportExcel( RoleBatchView.class, 2,1);
//        exportExcel.setDate(arrayList);
//        exportExcel.setSort(colNumList);
//        exportExcel.initialize("绑定信息",exportExcel.headerList,"",exportExcel.reds,2);
//        exportExcel.setDataList(new ArrayList()).write(response, "123.xlsx").dispose();

//    d) 下拉联动用法
//    以该类举例TeachClassRoom
//    经过一系列方法获得 :
//          一个 String[] 数组,作为父影响联动那一列
//          一个Map<String, String[]>  map 作为被影响的子数据那一列
//    在定义完 exportExcel 后执行导出之前执行 cascade 方法
//    exportExcel.cascade("sheet页名称","新建sheet页第一个单元格名称",父数组,子map集合,第几列开始执行);
//    仅封装了 AB   BC  CD  相邻列做联动
//      兼容 设置重新设置下拉

//    List<TeachClassRoom> teachClassRooms = teachClassRoomService.getTeacherClassRoom(getLoginUser().getSchoolId());
//    Map<String,String> building = new HashMap<>();
//    Map<String, List<TeachClassRoom>> TeachClassRoomcollect = teachClassRooms.stream()
//            .peek(teachClassRoom -> building.put(teachClassRoom.getTeachBuilding(),teachClassRoom.getTeachBuilding()))
//            .collect(Collectors.groupingBy(TeachClassRoom::getTeachBuilding));
//    Map<String, String[]> TeachClassRoomChildrenArrayMap = new HashMap<>();
//        for (Map.Entry entry : TeachClassRoomcollect.entrySet()){
//        String key = (String) entry.getKey();
//        List<TeachClassRoom> valueList = (List<TeachClassRoom>) entry.getValue();
//        List<String> stringList = new ArrayList<>();
//        for (TeachClassRoom teachClassRoom : valueList){
//            stringList.add(teachClassRoom.getRoomName());
//        }
//        String[] array = stringList.toArray(new String[stringList.size()]);
//        TeachClassRoomChildrenArrayMap.put(key,array);
//    }
//    String[] buildingArray = building.keySet()
//            .stream()
//            .toArray(String[]::new);
//        ExportExcel exportExcel = new ExportExcel( StationBindView.class, 2,1);
//        exportExcel.initialize("网关批量绑定",exportExcel.headerList,anno,exportExcel.reds,2,"网关批量绑定");
//        exportExcel.cascade("教室","教学楼",buildingArray,TeachClassRoomChildrenArrayMap,4);
//        exportExcel.setDataList(new ArrayList()).write(response, "网关批量绑定模板.xlsx").dispose();


//    public void cascade2(String[] schoolArr ,Map<String,String[]> map) {
//
//        Workbook book = new XSSFWorkbook();
//
//        // 创建需要用户填写的数据页
//        // 设计表头
//        Sheet sheetPro = book.createSheet("校区年级班级");
//        Row row0 = sheetPro.createRow(0);
//        row0.createCell(0).setCellValue("校区");
//        row0.createCell(1).setCellValue("年级");
//        row0.createCell(2).setCellValue("班级");
//
//        //创建一个专门用来存放地区信息的隐藏sheet页
//        //因此也不能在现实页之前创建，否则无法隐藏。
//        Sheet hideSheet = book.createSheet("多校区");
//        //这一行作用是将此sheet隐藏，功能未完成时注释此行,可以查看隐藏sheet中信息是否正确
//        //book.setSheetHidden(book.getSheetIndex(hideSheet), true);
//
//        int rowId = 0;
//        // 设置第一行，存省的信息
//        Row provinceRow = hideSheet.createRow(rowId++);
//        provinceRow.createCell(0).setCellValue("校区列表");
//        for(int i = 0; i < schoolArr.length; i ++){
//            Cell provinceCell = provinceRow.createCell(i + 1);
//            provinceCell.setCellValue(schoolArr[i]);
//        }
//        // 将具体的数据写入到每一行中，行开头为父级区域，后面是子区域。
//        for(int i = 0;i < areaFatherNameArr.length;i++){
//            String key = areaFatherNameArr[i];
//            String[] son = areaMap.get(key);
//            Row row = hideSheet.createRow(rowId++);
//            row.createCell(0).setCellValue(key);
//            for(int j = 0; j < son.length; j ++){
//                Cell cell = row.createCell(j + 1);
//                cell.setCellValue(son[j]);
//            }
//
//            // 添加名称管理器
//            String range = getRange(1, rowId, son.length);
//            Name name = book.createName();
//            //key不可重复,将父区域名作为key 父区域不能为数字开头
//            name.setNameName(key);
//            String formula = "area!" + range;
//            name.setRefersToFormula(formula);
//        }
//      e)隐藏数据列
//    设置标签    @ExcelField(title = "错误信息", align = 2, sort = 3, groups = {1, 2}, isnull=1, hide = 1)
//      hide=1 这个必须得设置 默认为0
//     调用  ExportExcel(String title, Class<?> cls,  Integer _hide, int type, String anno, int... groups) 构造方法
//   例子：  new ExportExcel("导出失败信息", InsertBatchRingView.class, 1, 2, anno).setDataList(exportFile).write(response, fileName).dispose();


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
}

