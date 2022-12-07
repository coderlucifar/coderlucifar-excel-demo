package com.coderlucifar.excel;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * excel通用单元格格式类
 */
public class TemplateCellWriteHandler implements CellWriteHandler {
    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean isHead) {
        Workbook workbooks = writeSheetHolder.getSheet().getWorkbook();
        writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), 20 * 256);
        CellStyle cellStyle = workbooks.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//设置前景填充样式
        cellStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());//前景填充色
        Font font1 = workbooks.createFont();//设置字体
        font1.setBold(true);
        font1.setColor((short)1);
        font1.setFontHeightInPoints((short)15);
        cellStyle.setFont(font1);
        cell.setCellStyle(cellStyle);
        //其他列
        if (!isHead){
            CellStyle style = workbooks.createCellStyle();
            DataFormat dataFormat = workbooks.createDataFormat();
            style.setDataFormat(dataFormat.getFormat("@"));
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setAlignment(HorizontalAlignment.CENTER);
            cell.setCellStyle(style);
        }
        //设置日期
        if (!isHead && cell.getColumnIndex()==3 || !isHead && cell.getColumnIndex()==9|| !isHead && cell.getColumnIndex()==17){
            CellStyle style = workbooks.createCellStyle();
            DataFormat dataFormat = workbooks.createDataFormat();
            style.setDataFormat(dataFormat.getFormat("yyyy/mm"));
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setAlignment(HorizontalAlignment.CENTER);
            cell.setCellStyle(style);
        }
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }
}
