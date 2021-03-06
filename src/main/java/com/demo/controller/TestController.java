package com.demo.controller;

import com.demo.modelView.ClassView;
import com.demo.persistence.dao.LinkageTestMapper;
import com.demo.persistence.dao.TestMapper;
import com.demo.persistence.entity.LinkageTest;
import com.demo.persistence.entity.LinkageTestExample;
import com.demo.persistence.entity.Test;
import com.demo.utils.ConstantUtil;
import com.demo.utils.StreamUtils;
import com.demo.utils.ZipUtils;
import com.demo.utils.export.ExportExcel;
import com.demo.utils.export.ExportExcelMutiSheet;
import com.demo.utils.export.ImportExcel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 *
 */
@RequestMapping("/excel")
@Controller
public class TestController {
    @Autowired
    TestMapper testMapper;
    @Autowired
    LinkageTestMapper linkageTestMapper;

    /**
     * 无注释导出Excel表格
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/anno/no")
    public void annoNo(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//        type = 1 代表无注释
        new ExportExcel("无注释导出Excel", ClassView.class, 1, "", 1).setDataList(classViewList).write(response, "无注释导出Excel.xlsx").dispose();
    }

    /**
     * 有注释导出Excel表格
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/anno/yes")
    public void annoYes(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//        type = 2 代表有注释
        new ExportExcel("有注释导出Excel", ClassView.class, 2, "我是注释", 1).setDataList(classViewList).write(response, "有注释导出Excel.xlsx").dispose();
    }

    /**
     * excel 表格下拉
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/pull/down")
    public void pullDown(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//        生成下拉数组
        List<String> _className = classViewList.stream().map(ClassView::getClassName).collect(Collectors.toList());
        String[] className = _className.toArray(new String[_className.size()]);
        List<String> _nj = classViewList.stream().map(ClassView::getNj).collect(Collectors.toList());
        String[] nj = _nj.toArray(new String[_nj.size()]);
//        数组集合
        List<String[]> arrayList = new ArrayList<>();
        arrayList.add(nj);
        arrayList.add(className);
//        数据下拉位置 列 横向 0开始
        List<Integer> colNumList = new ArrayList<>();
        colNumList.add(0);
        colNumList.add(1);
        ExportExcel exportExcel = new ExportExcel(ClassView.class, 2, 1);
        exportExcel.setDate(arrayList);
        exportExcel.setSort(colNumList);
        exportExcel.initialize("表格下拉", exportExcel.headerList, "", exportExcel.reds, 2, "表格下拉");
        exportExcel.setDataList(new ArrayList()).write(response, "表格下拉.xlsx").dispose();
    }

    /**
     * excel 表格联动
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/pull/linked")
    public void pullLinked(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//      得到nj的String 一个Map
        Map<String, String> njMap = new HashMap<>();
//        根据nj分组班级map
        Map<String, List<ClassView>> njCollect = classViewList.stream()
                .peek(classView -> njMap.put(classView.getNj(), classView.getNj()))
                .collect(Collectors.groupingBy(ClassView::getNj));

//        获得班级名称的一个数组
        Map<String, String[]> classNameArrayMap = new HashMap<>();
        for (Map.Entry entry : njCollect.entrySet()) {
            String key = (String) entry.getKey();
            List<ClassView> valueList = (List<ClassView>) entry.getValue();
            List<String> stringList = new ArrayList<>();
            for (ClassView classView : valueList) {
                stringList.add(classView.getClassName());
            }
            String[] array = stringList.toArray(new String[stringList.size()]);
            classNameArrayMap.put(key, array);
        }

        String[] njArray = njMap.keySet()
                .stream()
                .toArray(String[]::new);

        ExportExcel exportExcel = new ExportExcel(ClassView.class, 2, 1);
        exportExcel.initialize("联动测试", exportExcel.headerList, "", exportExcel.reds, 2, "联动测试");
//        rownum 需要判断从哪里开始 默认开始从1
        exportExcel.cascade("班级", "年级", njArray, classNameArrayMap, 1);
        exportExcel.setDataList(new ArrayList()).write(response, "联动测试.xlsx").dispose();
    }


    /**
     * excel 写入磁盘
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/write/io")
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//        type = 1 代表无注释
        new ExportExcel("无注释导出Excel", ClassView.class, 1, "", 1).setDataList(classViewList).writeFile("d:/磁盘写入.xlsx").dispose();
    }


    /**
     * zip 写入磁盘
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/zip/io")
    public void zipFile(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
        new ExportExcel("写入磁盘导出Excel", ClassView.class, 1, "", 1).setDataList(classViewList)
                .writeFile("d:/aaa/写入磁盘导出Excel.xlsx").dispose();

        String dirzip = "E:/写入磁盘导出Excel.zip";
        FileOutputStream fos1 = new FileOutputStream(new File(dirzip));
        ZipUtils.toZip("d:/aaa", fos1, true);
    }

    /**
     * 外部资源 访问直接生成压缩包
     *
     * @param request
     * @param response
     */
    @ResponseBody
    @RequestMapping(value = "/url/download/data/zip")
    public void urlDownZip(HttpServletRequest request, HttpServletResponse response) {
        try {
//            设置响应头:设置响应流
            StreamUtils.reponseStreamUtf8("内存压缩包.zip", response);
//            外部请求资源路径支持数组，list
            String[] pictureArray = {"http://file1.ckmooc.com/1552534945951.jpg"};
//            压缩外部资源
            ZipUtils.toUrlZip(pictureArray, response.getOutputStream());
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    /**
     * 多sheet页设置
     *
     * @param request
     * @param response
     */
    @ResponseBody
    @RequestMapping(value = "/multi/sheet")
    public void mulitySheet(HttpServletRequest request, HttpServletResponse response) {
        try {
            SXSSFWorkbook wb = new SXSSFWorkbook();
            new ExportExcelMutiSheet(wb, "a", 0, "a", ClassView.class, 2, "", 1).setDataList(new ArrayList(), wb);
            new ExportExcelMutiSheet(wb, "b", 1, "b", ClassView.class, 1, "").setDataList(new ArrayList(), wb);
            new ExportExcelMutiSheet(wb, "c",2, "c", ClassView.class, 1, "").setDataList(new ArrayList(), wb);
            new ExportExcelMutiSheet(wb, "d",3, "d", ClassView.class, 2, "", 1).setDataList(new ArrayList(), wb).write(wb, response, "aa.xlsx").dispose(wb);
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }


    /**
     *
     * @param file
     * @throws IOException
     * @throws InvalidFormatException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    @ResponseBody
    @RequestMapping(value = "/in/data")
    public void importCardExcelInfo(@RequestParam(value = "file") MultipartFile file) throws IOException, InvalidFormatException, IllegalAccessException, InstantiationException {
        try {
            ImportExcel importExcel = new ImportExcel(file, 2, 0);  //从第一行开始
            List<ClassView> dataList = importExcel.getDataList(ClassView.class, 1);
            System.out.println(dataList);
        }catch (Exception e){
            e.printStackTrace();
        }
    }


    private List<ClassView> getData() {
        List<LinkageTest> linkageTests = linkageTestMapper.selectByExample(new LinkageTestExample());
        ClassView classView;
        List<ClassView> classViewList = new ArrayList<>();
        for (LinkageTest linkageTest : linkageTests) {
            classView = new ClassView();
            classView.setNj(linkageTest.getNj());
            classView.setClassName(linkageTest.getClassName());
            classViewList.add(classView);
        }
        return classViewList;
    }
}
