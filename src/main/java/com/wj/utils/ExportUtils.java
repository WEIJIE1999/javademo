package com.wj.utils;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.formula.functions.T;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author weijie
 * @create 2021-01-07 下午 22:27
 */
public class ExportUtils {

    /**
     * 导出到 浏览器
     * @param response    HttpServletResponse
     * @param rows        实体类数据集合
     * @param headerAlias 列名
     * @param lastColumn  标题合并几行
     * @param content     标题名
     * @param filename    文件名
     */
    public static void exportToWeb(HttpServletResponse response, List rows, Map<String, String> headerAlias, Integer lastColumn, String content, String filename) {
        // 通过工具类创建writer，默认创建xls格式
        ExcelWriter writer = ExcelUtil.getWriter();
        //自定义标题别名
        for (Map.Entry<String, String> entry : headerAlias.entrySet()) {
            writer.addHeaderAlias(entry.getKey(), entry.getValue());
        }

        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(lastColumn, content);

        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);

        //out为OutputStream，需要写出到的目标流
        //response为HttpServletResponse对象
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        //test.xls是弹出下载对话框的文件名，不能为中文，中文请自行编码
        //String name = StringUtils.toUtf8String(filename);
        response.setHeader("Content-Disposition", "attachment;filename=" + filename + ".xls");
        ServletOutputStream out = null;
        try {
            out = response.getOutputStream();
            writer.flush(out, true);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭writer，释放内存
            writer.close();
        }
        //此处记得关闭输出Servlet流
        IoUtil.close(out);
    }

    /**
     * 导出到指定位置
     * @param path   路径
     * @param beans    实体类数据集合
     * @param headerAlias  列名
     * @param lastColumn  标题合并几行
     * @param content     标题名
     * @param filename    文件名
     */
    public static void export(String path, List beans, Map<String, String> headerAlias, Integer lastColumn, String content, String filename) {
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(path + filename + ".xls");

        List rows = CollUtil.newArrayList(beans);

        //自定义标题别名
        for (Map.Entry<String, String> entry : headerAlias.entrySet()) {
            writer.addHeaderAlias(entry.getKey(), entry.getValue());
        }
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(lastColumn, content);
        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(rows, true);
        // 关闭writer，释放内存
        writer.close();
    }

    /**
     * 读取为Bean，Bean中的字段名为标题，字段值为标题对应的单元格值
     * @param path  文件路径
     * @param beanType  读取为的实体类
     * @return
     */
    public static List<T> inport(String path, Class<T> beanType) {
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(path));
        return reader.readAll(beanType);
    }

    /**
     * 读取为Bean，Bean中的字段名为标题，字段值为标题对应的单元格值
     * @param path  文件路径
     * @param number  第几个sheet
     * @param beanType  读取为的实体类
     * @return
     */
    public static List<T> inportForSheetNumer(String path,Integer number,Class<T> beanType) {
        //通过sheet编号获取
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(path), number);
        return reader.readAll(beanType);
    }

    /**
     * 读取为Bean，Bean中的字段名为标题，字段值为标题对应的单元格值
     * @param path  文件路径
     * @param sheetName  指定sheet的名字
     * @param beanType  读取为的实体类
     * @return
     */
    public static List<T> inportForSheetName(String path,String sheetName, Class<T> beanType) {
        //通过sheet名获取
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(path), sheetName);
        return reader.readAll(beanType);
    }

}
