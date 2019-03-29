package com.demo.modelView;

import com.demo.utils.export.annotation.ExcelField;

public class ClassView {
    private String id;
    private String nj;
    private String className;
    @ExcelField(title = "年级", align = 2, sort = 1, groups = {1, 2}, isnull = 1)
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @ExcelField(title = "年级", align = 2, sort = 1, groups = {1, 2}, isnull = 1)
    public String getNj() {
        return nj;
    }

    public void setNj(String nj) {
        this.nj = nj;
    }
    @ExcelField(title = "班级", align = 2, sort = 2, groups = {1, 2}, isnull = 1)
    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }
}
