package com.coderlucifar.excel;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author sunyuan
 * @version 1.0
 * @description: excel工具类
 * @date 2022/11/24 18:08
 */
public class ExcelUtils {

    public static void createExcelTemplate(HttpServletResponse httpServletResponse, String fileName, List<Class<?>> modelClassList, Map<String, Object> dataDictionaryMap, Integer headHeight, String[] sheetNames) {
        // 设置单元格风格
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = setMyCellStyle();
        try {
            //下拉列表集合 key是列号，value是下拉列表中的值
            Map<Integer, String[]> explicitListConstraintMap = new HashMap<>();
            //循环获取对应列得下拉列表信息
            Field[] declaredFields = modelClassList.get(0).getDeclaredFields();
            for (int i = 0; i < declaredFields.length; i++) {
                Field field = declaredFields[i];
                // 获取字段上的 注解， 解析下拉列表信息
                ExplicitConstraint explicitConstraint = field.getAnnotation(ExplicitConstraint.class);
                // 解析注解
                resolveExplicitConstraint(explicitListConstraintMap, explicitConstraint, dataDictionaryMap);
            }
            // 获取文件输出流
            OutputStream outputStream = getOutputStream(fileName, httpServletResponse, ExcelTypeEnum.XLSX);

            ExcelWriter excelWriter = EasyExcel.write(outputStream)
                                                .excelType(ExcelTypeEnum.XLSX)
                                                .build();
            // 写sheet1
            WriteSheet sheet1 = EasyExcel.writerSheet(0, sheetNames[0]).head(modelClassList.get(0))
                    .registerWriteHandler(new DownListCellWriteHandler(headHeight, explicitListConstraintMap))
                    .registerWriteHandler(new TemplateCellWriteHandler())
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .build();
//            excelWriter.write(data, sheet1);
            // 写sheet2
            WriteSheet sheet2 = EasyExcel.writerSheet(1, "组织机构字典").head(modelClassList.get(1))      // OrgInfoModel.class
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .registerWriteHandler(new SimpleColumnWidthStyleStrategy(40))
                    .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)33, (short)16))
                    .build();
//            excelWriter.write(orgList, sheet2);
            // sheet3
            WriteSheet sheet3 = EasyExcel.writerSheet(2, "部门字典").head(modelClassList.get(2))          // DeptInfoModel.class
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .registerWriteHandler(new SimpleColumnWidthStyleStrategy(40))
                    .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)33, (short)16))
                    .build();
//            excelWriter.write(deptList, sheet3);

            excelWriter.finish();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("系统异常！");
        }

    }

    /**
     * 解析注解内容 获取下列表信息
     * @param explicitListConstraintMap 下拉列表map （key为列号，value为下拉列表list）
     * @param explicitConstraint 下拉列表注解
     * @param dataDictionaryMap 数据字典map
     * @return
     */
    public static Map<Integer, String[]> resolveExplicitConstraint(Map<Integer, String[]> explicitListConstraintMap, ExplicitConstraint explicitConstraint, Map<String, Object> dataDictionaryMap){
        if (explicitConstraint == null) {
            return null;
        }
        // 获取'固定下拉'信息
        String[] source = explicitConstraint.source();
        if (source.length > 0) {
            // 存放对应列的下拉列表
            explicitListConstraintMap.put(explicitConstraint.indexNum(), source);
        }
        // '动态下拉'信息
        if (dataDictionaryMap != null && !dataDictionaryMap.isEmpty()){
            try {
                // 获取key对应的字典值（多个以逗号分隔）
                String dictionaryValues = (String) dataDictionaryMap.get(explicitConstraint.type());
                String[] dictionaryValueList = {};
                if (StrUtil.isNotEmpty(dictionaryValues)){
                    dictionaryValueList = dictionaryValues.split(",");
                }

                if (dictionaryValueList.length > 0){
                    explicitListConstraintMap.put(explicitConstraint.indexNum(), dictionaryValueList);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 导出文件时为Writer生成OutputStream
     * 设置http响应头参数
     */
    private static OutputStream getOutputStream(String fileName, HttpServletResponse response, ExcelTypeEnum excelTypeEnum) throws Exception {
        try {
            fileName = new String(fileName.getBytes(), StandardCharsets.ISO_8859_1);
            response.setCharacterEncoding(StandardCharsets.UTF_8.name());
            response.setContentType("application/vnd.ms-excel");
            response.addHeader("Content-Disposition", "filename=" + fileName);
            return response.getOutputStream();
        } catch (IOException e) {
            throw new Exception("系统异常");
        }
    }

    /**
     * 设置单元格风格策略
     * @return
     */
    public static HorizontalCellStyleStrategy setMyCellStyle() {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 设置表头居中对齐
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 颜色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 10);
        // 字体
        headWriteCellStyle.setWriteFont(headWriteFont);
        headWriteCellStyle.setWrapped(true);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 设置内容靠中对齐
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        return horizontalCellStyleStrategy;
    }

}
