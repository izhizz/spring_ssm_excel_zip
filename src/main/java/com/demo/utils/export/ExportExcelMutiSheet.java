package com.demo.utils.export;

import com.demo.utils.Reflections;
import com.demo.utils.export.annotation.ExcelField;
import com.google.common.collect.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.util.*;

/**
 * 导出Excel文件（导出“XLSX”格式，支持大数据量导出   @see org.apache.poi.ss.SpreadsheetVersion）
 */
public class ExportExcelMutiSheet {

    private static Logger log = LoggerFactory.getLogger(ExportExcelMutiSheet.class);

    /**
     * 工作薄对象
     */
//    private SXSSFWorkbook wb;

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
     * 注解列表（Object[]{ ExcelField, Field/Method }）
     */
    List<Object[]> annotationList = Lists.newArrayList();

    public Integer checkHide = 0;

    /**
     * 构造函数
     *
     * @param title 表格标题，传“空值”，表示无标题
     * @param cls   实体对象，通过annotation.ExportField获取标题
     */
    public ExportExcelMutiSheet(SXSSFWorkbook wb, String sheetName, Integer _hide, Integer sheetIndex, String title, String anno, Class<?> cls) {
        this(wb, sheetName,_hide, sheetIndex, title, cls, 1, anno);
    }

    public ExportExcelMutiSheet() { }

    /**
     * 构造函数
     *
     * @param title  表格标题，传“空值”，表示无标题
     * @param cls    实体对象，通过annotation.ExportField获取标题
     * @param type   导出类型（1:导出数据；2：导出模板）
     * @param groups 导入分组
     */
    public ExportExcelMutiSheet(SXSSFWorkbook wb, String sheetName, Integer _hide, Integer sheetIndex, String title, Class<?> cls, int type, String anno, int... groups) {
        checkHide=_hide;
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
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(
                        new Integer(((ExcelField) o2[0]).sort()));
            }

            ;
        });
        // Initialize
        List<String> headerList = Lists.newArrayList();
        List<Integer> reds = Lists.newArrayList();
        for (Object[] os : annotationList) {
            ExcelField field = (ExcelField) os[0];

            String t = field.title();
            int red = field.isnull();
            // 如果是导出，则去掉注释
            if (type == 1) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            reds.add(red);
            int hide = ((ExcelField) os[0]).hide();
            if (checkHide==0&&hide == 1) {
                continue;
            }
            //================================LL===========================================
//            int hebingcout = field.hebingcout();
//            if (hebingcout > 0) {
//                t += ";" + hebingcout;
//            }
            //==================================LL=========================================

            headerList.add(t);
        }
        initialize(wb, sheetName, sheetIndex, title, headerList, anno, reds, type);
    }

    /**
     * 构造函数
     *
     * @param title   表格标题，传“空值”，表示无标题
     * @param headers 表头数组
     */
    public ExportExcelMutiSheet(SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, String[] headers, String anno) {
        initialize(wb, sheetName, sheetIndex, title, Lists.newArrayList(headers), anno, null, null);
    }

    /**
     * 构造函数
     *
     * @param title      表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    public ExportExcelMutiSheet(SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, List<String> headerList, String anno, int type) {
        initialize(wb, sheetName, sheetIndex, title, headerList, anno, null, type);
    }


    /**
     * 初始化函数
     *
     * @param title      表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    private void initialize(SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, List<String> headerList, String anno, List<Integer> reds, Integer type) {

        this.sheet = wb.createSheet();
        wb.setSheetName(sheetIndex, sheetName);

        if (headerList.size() > 1) {
            this.styles = createStyles(wb);
        } else {
            this.styles = createSingleStyles(wb);
        }
        // Create title
        if (StringUtils.isNotBlank(title)) {
            Row titleRow = sheet.createRow(rownum++);
            titleRow.setHeightInPoints(45);
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
            introRow.setHeightInPoints(100);
            Cell introCell = introRow.createCell(0);
            styles.get("intro").setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            introCell.setCellStyle(styles.get("intro"));
            introCell.setCellValue(new XSSFRichTextString(anno));       //HSSFRichTextString
            if (headerList.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(introRow.getRowNum(),
                        introRow.getRowNum(), introRow.getRowNum() - 1, headerList.size() - 1));
            }
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
            Cell cell =  null;
            if (i>0){
                cell = headerRow.createCell(recordTotalCell);
            }else {
                cell =  headerRow.createCell(i);
            }
            if (cells>0){
                for (int m=1;m<cells;m++){
                    recordTotalCell++;
//                    System.out.println(recordTotalCell);
                    headerRow.createCell(recordTotalCell);
                    if (cells-1==m){
                        recordTotalCell++;
                    }
                }
            }else {
                recordTotalCell++;
            }

            String[] ss = StringUtils.split(headerList.get(i), "**", 2);
            if (ss.length == 2) {
                if (ss[0].indexOf(";")>0){
//                    System.out.println(ss[0].split(";")[0]);
                    cell.setCellValue(ss[0].split(";")[0]);
                } else {
//                    System.out.println(ss[0].split(";")[0]);
                    cell.setCellValue(ss[0]);
                }

                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            } else {
                if (ss[0].indexOf(";")>0){
//                    System.out.println(ss[0].split(";")[0]);
                        cell.setCellValue(ss[0].split(";")[0]);
//                    cell.setCellValue(new XSSFRichTextString(ss[0].split(";")[0]));
                }else {
//                    System.out.println(headerList.get(i));
                    cell.setCellValue(new XSSFRichTextString(headerList.get(i)));
                }
            }



            if (cells > 0) {
                CellRangeAddress region = new CellRangeAddress(headerRow.getRowNum(),
                        headerRow.getRowNum(), recordTotalCell-cells, recordTotalCell-1);
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
            sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : 5000);
        }
        log.debug("Initialize success.");
    }


    /**
     * 创建表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
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
        headerFont2.setColor(IndexedColors.GOLD.index);
        style.setFont(headerFont2);
        styles.put("header2", style);

        return styles;
    }


    /**
     * 创建单一列表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createSingleStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style = wb.createCellStyle();

        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 10);
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
        headerFont2.setColor(IndexedColors.WHITE.index);
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
    public Cell addCell(SXSSFWorkbook wb, Row row, int column, Object val) {
        return this.addCell(wb, row, column, val, 0, Class.class);
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
    public Cell addCell(SXSSFWorkbook wb, Row row, int column, Object val, int align, Class<?> fieldType) {
        Cell cell = row.createCell(column);
        CellStyle style = styles.get("data" + (align >= 1 && align <= 3 ? align : ""));
        DataFormat format = wb.createDataFormat();
        style.setDataFormat(format.getFormat("@"));
        try {
            if (val == null) {
                cell.setCellValue("");
            } else if (val instanceof String) {
                cell.setCellValue((String) val);
            } else if (val instanceof Integer) {
                cell.setCellValue((String) val);
            } else if (val instanceof Long) {
                cell.setCellValue((String) val);
            } else if (val instanceof Double) {
                cell.setCellValue((Double) val);
            } else if (val instanceof Float) {
                cell.setCellValue((Float) val);
            } else if (val instanceof Date) {
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
        cell.setCellType(Cell.CELL_TYPE_STRING);
        return cell;
    }

    /**
     * 添加数据（通过annotation.ExportField添加数据）
     *
     * @return list 数据列表
     */
    public <E> ExportExcelMutiSheet setDataList(List<E> list, SXSSFWorkbook wb) {
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
                    // If is dict, get dict label
//                    if (StringUtils.isNotBlank(ef.dictType())) {
//                        val = DictUtils.getDictLabel(val == null ? "" : val.toString(), ef.dictType(), "");
//                    }
                } catch (Exception ex) {
                    // Failure to ignore
                    log.info(ex.toString());
                    val = "";
                }
                this.addCell(wb, row, colunm++, val, ef.align(), ef.fieldType());
                sb.append(val + ", ");
            }
            if (annotationList.size() == 0) {

                Object val = null;

                Method[] method = e.getClass().getDeclaredMethods();
                List<Method> methodGet = new ArrayList<Method>();
                for (Method m : method) {

                    if (m.getName().indexOf("get") != -1) {
                        methodGet.add(m);
                    }
                }

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
    public ExportExcelMutiSheet write(SXSSFWorkbook wb, OutputStream os) throws IOException {
        wb.write(os);
        return this;
    }

    /**
     * 输出到客户端
     *
     * @param fileName 输出文件名
     */
    public ExportExcelMutiSheet write(SXSSFWorkbook wb, HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("multipart/form-data");

        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
        write(wb, response.getOutputStream());
        return this;
    }

    /**
     * 输出到客户端
     *  * 重载的原因：
     *      因为火狐浏览器自己会对url进行一次编码，所以火狐浏览器下载的文件名是乱码的。
     *      重载的函数添加request参数，对浏览器进行判断，对不同的浏览器文件名进行不同方式的编码。
     *      --liyaoheng
     *
     * @param fileName 输出文件名
     */
    public ExportExcelMutiSheet write(SXSSFWorkbook wb, HttpServletRequest request, HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("multipart/form-data");

        response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        write(wb, response.getOutputStream());
        return this;
    }

    /**
     * 输出到文件
     *
     * @param name 输出文件名
     */
    public ExportExcelMutiSheet writeFile(SXSSFWorkbook wb, String name) throws FileNotFoundException, IOException {
        FileOutputStream os = new FileOutputStream(name);
        this.write(wb, os);
        return this;
    }

    /**
     * 清理临时文件
     */
    public ExportExcelMutiSheet dispose(SXSSFWorkbook wb) {
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
    public <E> ExportExcelMutiSheet setDataListByHeader(List<E> list, List<String> headerList, SXSSFWorkbook wb) {
        try {
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
                                this.addCell(wb, row, colunm++, val, 2, e.getClass());
                                sb.append(val + ", ");
                            }
                        }
                    }
                }

                log.debug("Write success: [" + row.getRowNum() + "] " + sb.toString());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return this;
    }

    public ExportExcelMutiSheet(int width, SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, Class<?> cls, int type, String anno, int... groups) {
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
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(
                        new Integer(((ExcelField) o2[0]).sort()));
            }
        });
        // Initialize
        List<String> headerList = Lists.newArrayList();
        List<Integer> reds = Lists.newArrayList();
        for (Object[] os : annotationList) {
            ExcelField field = (ExcelField) os[0];


            String t = field.title();
            int red = field.isnull();
            // 如果是导出，则去掉注释
            if (type == 1) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            reds.add(red);
            //================================LL===========================================
//            int hebingcout = field.hebingcout();
//            if (hebingcout > 0) {
//                t += ";" + hebingcout;
//            }
            //==================================LL=========================================
            headerList.add(t);

        }
        initialize(wb, sheetName, sheetIndex, title, headerList, anno, reds, type, width);
    }


    private void initialize(SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, List<String> headerList, String anno, List<Integer> reds, Integer type, int width) {

        this.sheet = wb.createSheet();
        wb.setSheetName(sheetIndex, sheetName);

        if (headerList.size() > 1) {
            this.styles = createStyles(wb);
        } else {
            this.styles = createSingleStyles(wb);
        }
        // Create title
        if (StringUtils.isNotBlank(title)) {
            Row titleRow = sheet.createRow(rownum++);
            titleRow.setHeightInPoints(40);
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
            introRow.setHeightInPoints(100);
            Cell introCell = introRow.createCell(0);
            styles.get("intro").setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            introCell.setCellStyle(styles.get("intro"));
            introCell.setCellValue(new XSSFRichTextString(anno));       //HSSFRichTextString
            if (headerList.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(introRow.getRowNum(),
                        introRow.getRowNum(), introRow.getRowNum() - 1, headerList.size() - 1));
            }
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
            Cell cell =  null;
            if (i>0){
                cell = headerRow.createCell(recordTotalCell);
            }else {
                cell =  headerRow.createCell(i);
            }
            if (cells>0){
                for (int m=1;m<cells;m++){
                    recordTotalCell++;
//                    System.out.println(recordTotalCell);
                    headerRow.createCell(recordTotalCell);
                    if (cells-1==m){
                        recordTotalCell++;
                    }
                }

                for (int m=1;m<cells;m++){
                    headerRow.createCell(i+m+1);
                    recordTotalCell=i+m+1;
                }
                CellRangeAddress region = new CellRangeAddress(headerRow.getRowNum(),
                        headerRow.getRowNum(), i, i+recordTotalCell-1);
                sheet.addMergedRegion(region);
            }else {
                recordTotalCell++;
            }


            //=====================================滑腻腻的分割线以上部分是LL所写  有问题找她哦==========================================================
            if (reds != null && reds.size() > 0) {
                if (reds.get(i) == 0) {
                    cell.setCellStyle(styles.get("header"));
                }
                if (reds.get(i) == 1) {
                    cell.setCellStyle(styles.get("header2"));
                }
            }

            String[] ss = StringUtils.split(headerList.get(i), "**", 2);
            if (ss.length == 2) {
                if (ss[0].indexOf(";")>0){
                    cell.setCellValue(ss[0].split(";")[0]);
                }else {
                    cell.setCellValue(ss[0]);
                }

                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            } else {
                if (ss[0].indexOf(";")>0){
//                    System.out.println(ss[0]);
//                    System.out.println(ss[0].split(";")[0]);
                    cell.setCellValue(ss[0].split(";")[0]);
                }else {
//                    System.out.println(headerList.get(i));
                    cell.setCellValue(headerList.get(i));
                }
            }
            //==========================LL 有问题找她===========================
//            sheet.autoSizeColumn(i);
        }
        for (int i = 0; i < headerList.size(); i++) {
            int colWidth = sheet.getColumnWidth(i) * 2;
            sheet.setColumnWidth(i, width);
        }

        int flag=0;
        while (true){
            int colunm = 0;
            Row row = this.addRow();
            StringBuilder sb = new StringBuilder();
            for (Object[] os : annotationList) {
                ExcelField ef = (ExcelField) os[0];
                Object val = null;
                // Get entity value
//                try {
//
//                    val = Reflections.invokeGetter("", ef.value());
//
//                } catch (Exception ex) {
//                    // Failure to ignore
//                    log.info(ex.toString());
//                    val = "";
//                }
                this.addCell(wb,row, colunm++, val, ef.align(), ef.fieldType());
                sb.append(val + ", ");
            }
            flag++;
            if (flag>=300){
                break;
            }
        }
        log.debug("Initialize success.");
    }


    public ExportExcelMutiSheet(int width, SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, Class<?> cls, int type, String anno, Integer height, Short cellType, String color, int... groups) {
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
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(
                        new Integer(((ExcelField) o2[0]).sort()));
            }

            ;
        });
        // Initialize
        List<String> headerList = Lists.newArrayList();
        List<Integer> reds = Lists.newArrayList();
        for (Object[] os : annotationList) {
            ExcelField field = (ExcelField) os[0];

            String t = field.title();
            int red = field.isnull();
            // 如果是导出，则去掉注释
            if (type == 1) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            reds.add(red);
            headerList.add(t);
        }
        initialize(wb, sheetName, sheetIndex, title, headerList, anno, reds, type, width, height, color, cellType);
    }


    private void initialize(SXSSFWorkbook wb, String sheetName, Integer sheetIndex, String title, List<String> headerList, String anno, List<Integer> reds, Integer type, int width, Integer height, String color, Short cellType) {

        this.sheet = wb.createSheet();
        wb.setSheetName(sheetIndex, sheetName);

        if (headerList.size() > 1) {
            this.styles = createStyles(wb);
        } else {
            this.styles = createSingleStyles(wb);
        }
        // Create title
        if (StringUtils.isNotBlank(title)) {
            Row titleRow = sheet.createRow(rownum++);
            titleRow.setHeightInPoints(40);
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
            if (height == null) {
                height = 100;
            }
            introRow.setHeightInPoints(height);
            Cell introCell = introRow.createCell(0);
            if (cellType == null) {
                cellType = CellStyle.VERTICAL_CENTER;
            }
            styles.get("intro").setVerticalAlignment(cellType);
            if (StringUtils.isEmpty(color)) {
                color = "intro";
            }
            introCell.setCellStyle(styles.get(color));
            introCell.setCellValue(new XSSFRichTextString(anno));       //HSSFRichTextString
            if (headerList.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(introRow.getRowNum(),
                        introRow.getRowNum(), introRow.getRowNum() - 1, headerList.size() - 1));
            }
        }

        // Create header
        if (headerList == null) {
            throw new RuntimeException("headerList not null!");
        }
        Row headerRow = sheet.createRow(rownum++);
        headerRow.setHeightInPoints(16);
        for (int i = 0; i < headerList.size(); i++) {
            Cell cell = headerRow.createCell(i);
            if (reds != null && reds.size() > 0) {
                if (reds.get(i) == 0) {
                    cell.setCellStyle(styles.get("header"));
                }
                if (reds.get(i) == 1) {
                    cell.setCellStyle(styles.get("header2"));
                }
            }

            String[] ss = StringUtils.split(headerList.get(i), "**", 2);
            if (ss.length == 2) {
                cell.setCellValue(ss[0]);
                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(
                        new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                comment.setString(new XSSFRichTextString(ss[1]));
                cell.setCellComment(comment);
            } else {
                cell.setCellValue(headerList.get(i));
            }
//            sheet.autoSizeColumn(i);
        }
        for (int i = 0; i < headerList.size(); i++) {
            int colWidth = sheet.getColumnWidth(i) * 2;
            sheet.setColumnWidth(i, width);
        }
        log.debug("Initialize success.");
    }

}
