package com.demo.controller;

import com.demo.modelView.ClassView;
import com.demo.persistence.dao.LinkageTestMapper;
import com.demo.persistence.dao.TestMapper;
import com.demo.persistence.entity.LinkageTest;
import com.demo.persistence.entity.LinkageTestExample;
import com.demo.persistence.entity.Test;
import com.demo.utils.export.ExportExcel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
     * @param response
     * @throws IOException
     */
    @RequestMapping("/anno/yes")
    public void annoYes(HttpServletResponse response) throws IOException {
        List<ClassView> classViewList = this.getData();
//        type = 2 代表有注释
        new ExportExcel("无注释导出Excel", ClassView.class, 2, "我是注释", 1).setDataList(classViewList).write(response, "无注释导出Excel.xlsx").dispose();
    }



    private List<ClassView> getData(){
        List<LinkageTest> linkageTests = linkageTestMapper.selectByExample(new LinkageTestExample());
        ClassView classView;
        List<ClassView> classViewList = new ArrayList<>();
        for (LinkageTest linkageTest : linkageTests){
            classView = new ClassView();
            classView.setNj(linkageTest.getNj());
            classView.setClassName(linkageTest.getClassName());
            classViewList.add(classView);
        }
        return classViewList;
    }
}
