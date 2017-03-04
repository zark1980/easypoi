package com.telecomjs.utils;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by zark on 17/3/4.
 */
public class PoiExcelReader implements Iterable<String[]> {
    //读取的excel内部对象
    private InputStream inputStream;
    private POIFSFileSystem fs;
    //private HSSFWorkbook wb;
    private Workbook wb;
    //private HSSFSheet sheet;
    private Sheet sheet;
    //private HSSFRow row;
    private Row row;
    private boolean hasTitle=true;
    private String[] titles;
    private String sheetName;
    //Iterator 使用的对象
    private int columnNumber = 0;
    private int rowNumber = 0;
    private int currentRowNumber = 0;


    /**
     * @param is 输入文件流
     */
    public PoiExcelReader(InputStream is){
        this(is,true,null);
    }

    public PoiExcelReader(InputStream is,String sheetName){
        this(is,true,sheetName);
    }

    /**
     *
     * @param is 输入文件流
     * @param hasTitle 是否有标题行
     */
    public PoiExcelReader(InputStream is ,boolean hasTitle ,String sheetName ){
        this.inputStream = is;
        this.hasTitle = hasTitle;
        this.sheetName = sheetName;
        if (is == null){
            throw  new UnsupportedOperationException();
        }
    }

    /**
     * 打开文件流，初始化Iterator
     */
    public void open(){
        try {

            try {
                wb = new XSSFWorkbook(this.inputStream);
            } catch (Exception ex) {
                fs = new POIFSFileSystem(this.inputStream);
                wb = new HSSFWorkbook(fs);
            }
            if (this.sheetName == null)
                sheet = wb.getSheetAt(0);
            else
                sheet = wb.getSheet(this.sheetName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void close(){
        try{
            if (wb != null)
                wb.close();
            if (fs != null)
                fs.close();
            this.inputStream = null;
            this.sheet = null;
            this.titles = null;
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel表格表头的内容
     * @param --InputStream
     * @return String 表头内容的数组
     */
    public String[] readExcelTitle() {
        if (sheet == null){
            throw  new UnsupportedOperationException("参数没有初始化，先执行open方法.");
        }
        if (titles == null) {
            row = sheet.getRow(0);
            //得到总记录数
            if (rowNumber == 0)
                rowNumber = sheet.getLastRowNum();
            // 标题总列数
            if (columnNumber == 0)
                columnNumber = row.getPhysicalNumberOfCells();
            //System.out.println("rowNum:,colNum:" + rowNumber + "," + columnNumber);
            titles = new String[columnNumber];
            for (int i = 0; i < columnNumber; i++) {
                //title[i] = getStringCellValue(row.getCell((short) i));
                titles[i] = getCellFormatValue(row.getCell((short) i));
            }
            currentRowNumber++;
        }
        return titles;
    }

    private String[] readExcelNextRow(){
        if (sheet == null){
            throw  new UnsupportedOperationException("参数没有初始化，先执行open方法.");
        }

        //得到总记录数
        if (rowNumber == 0)
            rowNumber = sheet.getLastRowNum();
        // 标题总列数
        if (columnNumber == 0)
            columnNumber = row.getPhysicalNumberOfCells();
        String[] columns= new String[columnNumber];
        //System.out.println("rowNum:,colNum:" + rowNumber + "," + columnNumber);
        // 正文内容应该从第二行开始,第一行为表头的标题
        row = sheet.getRow(currentRowNumber++);
        for (int j = 0;j < columnNumber;j++) {
            columns[j] = getCellFormatValue(row.getCell((short) j)).trim() ;
        }

        return columns;
    }
    /**
     * 读取Excel数据内容
     * @param --InputStream
     * @return Map 包含单元格数据内容的Map对象
     */
    public List<String[]> readExcelContent(InputStream is) {
        List content = new ArrayList<String[]>();
        if (sheet == null){
            throw  new UnsupportedOperationException("参数没有初始化，先执行open方法.");
        }
        // 得到总行数
        if (rowNumber == 0)
            rowNumber = sheet.getLastRowNum();
        if (columnNumber == 0) {
            row = sheet.getRow(0);
            columnNumber = row.getPhysicalNumberOfCells();
        }
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = (this.hasTitle?1:0); i <= rowNumber; i++) {
            row = sheet.getRow(i);
            int j = 0;
            String[] str=new String[columnNumber];
            while (j < columnNumber) {
                // 每个单元格的数据内容用"-"分割开，以后需要时用String类的replace()方法还原数据
                // 也可以将每个单元格的数据设置到一个javabean的属性中，此时需要新建一个javabean
                // str += getStringCellValue(row.getCell((short) j)).trim() +
                // "-";
                str[j] = getCellFormatValue(row.getCell((short) j)).trim();
                j++;
            }
            content.add(str);
        }
        return content;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    private String getStringCellValue(Cell cell) {
        String strCell = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        if (cell == null) {
            return "";
        }
        return strCell;
    }

    /**
     * 获取单元格数据内容为日期类型的数据
     *
     * @param cell
     *            Excel单元格
     * @return String 单元格数据内容
     */
    private String getDateCellValue(Cell cell) {
        String result = "";
        try {
            int cellType = cell.getCellType();
            if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
                Date date = cell.getDateCellValue();
                result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1)
                        + "-" + date.getDate();
            } else if (cellType == HSSFCell.CELL_TYPE_STRING) {
                String date = getStringCellValue(cell);
                result = date.replaceAll("[年月]", "-").replace("日", "").trim();
            } else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
                result = "";
            }
        } catch (Exception e) {
            System.out.println("日期格式不正确!");
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 根据HSSFCell类型设置数据
     * @param cell
     * @return
     */
    private String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC:
                case HSSFCell.CELL_TYPE_FORMULA: {
                    // 判断当前的cell是否为Date
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        // 如果是Date类型则，转化为Data格式

                        //方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
                        //cellvalue = cell.getDateCellValue().toLocaleString();

                        //方法2：这样子的data格式是不带带时分秒的：2011-10-12
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);

                    }
                    // 如果是纯数字
                    else {
                        // 取得当前Cell的数值
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                // 默认的Cell值
                default:
                    cellvalue = " ";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;

    }


    @Override
    public Iterator<String[]> iterator() {
        return new Iterator<String[]>() {
            @Override
            public boolean hasNext() {
                return currentRowNumber < rowNumber;
            }

            @Override
            public String[] next() {
                return  readExcelNextRow();
            }

            @Override
            public void remove() {
                throw new UnsupportedOperationException();
            }
        };
    }
}
