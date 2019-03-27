package com.demo.controller;

import com.demo.modelView.ClassView;
import com.demo.persistence.dao.LinkageTestMapper;
import com.demo.persistence.dao.TestMapper;
import com.demo.persistence.entity.LinkageTest;
import com.demo.persistence.entity.LinkageTestExample;
import com.demo.persistence.entity.Test;
import com.demo.utils.ZipUtils;
import com.demo.utils.export.ExportExcel;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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
                .peek(classView -> njMap.put(classView.getNj(),classView.getNj()))
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
     * excel 写入磁盘
     *
     * @param response
     * @throws IOException
     */
    @RequestMapping("/zip/io")
    public void zipFile(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
        new ExportExcel("无注释导出Excel", ClassView.class, 1, "", 1).setDataList(classViewList).writeFile("d:/aaa/磁盘写入.xlsx").dispose();
        String dirzip =   "E:\\aa.zip";
        FileOutputStream fos1 = new FileOutputStream(new File(dirzip));
        ZipUtils.toZip("d:/aaa", fos1, true);
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
