package com.demo.persistence.dao;

import com.demo.persistence.entity.LinkageTest;
import com.demo.persistence.entity.LinkageTestExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

public interface LinkageTestMapper {
    int deleteByExample(LinkageTestExample example);

    int deleteByPrimaryKey(Integer id);

    int insert(LinkageTest record);

    int insertSelective(LinkageTest record);

    List<LinkageTest> selectByExample(LinkageTestExample example);

    LinkageTest selectByPrimaryKey(Integer id);

    int updateByExampleSelective(@Param("record") LinkageTest record, @Param("example") LinkageTestExample example);

    int updateByExample(@Param("record") LinkageTest record, @Param("example") LinkageTestExample example);

    int updateByPrimaryKeySelective(LinkageTest record);

    int updateByPrimaryKey(LinkageTest record);
}