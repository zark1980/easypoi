package com.telecomjs.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by zark on 17/3/4.
 */
public class PoiExcelWriter {
    private String[] titles;
    private String sheetName;
    private OutputStream outputStream;
    private String type;
    private Workbook workbook;
    private Sheet sheet;
    private Object[] data;
    private int currentRow =0;

    public PoiExcelWriter(OutputStream os,String type){
        this(os,null,null,type);
    }
    public PoiExcelWriter(OutputStream os){
        this(os,null,null,"xls");
    }
    /**
     *
     * @param os
     * @param titles
     * @param sheetName
     */
    public PoiExcelWriter(OutputStream os , String[] titles , String sheetName,String type ){
        this.outputStream = os;
        this.titles = titles;
        this.sheetName = sheetName;
        this.type = type;
        if (os == null ){
            throw  new UnsupportedOperationException("输出流参数不能为空");
        }
    }

    /**
     * 初始化对象
     */
    public void open(){
        try {
            if (this.type.equals("xlsx"))
                workbook = new SXSSFWorkbook();
            else
                workbook = new HSSFWorkbook();
            if (sheetName != null)
                sheet = workbook.createSheet(sheetName);
            else
                sheet = workbook.createSheet();
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }

    public void writeTitle(Object[] titles){
        writeRow(titles);
    }

    public void writeRow(Object[] data){
        Row row = sheet.createRow(currentRow++);
        int columnNumber = data.length;
        for (int j=0;j<columnNumber;j++){
            Cell cell = row.createCell(j);
            cell.setCellValue((String)data[j]);
        }
    }

    public void writeAll(List<Object[]> data){
        for (Object[] arr : data){
            writeRow(arr);
        }
    }

    /**
     * 写入缓存，释放内存
     */
    public void close(){
        try {
            workbook.write(this.outputStream);
            workbook = null;
            sheet = null;
            this.outputStream.flush();
            if (workbook != null)
                workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
