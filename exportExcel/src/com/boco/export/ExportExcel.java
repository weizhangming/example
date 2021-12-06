package com.boco.export;

import com.boco.entity.Result;
import com.boco.entity.Student;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExportExcel {
    public static void main(String[] args) throws IOException {
        Result result = new Result();
        result.setColspan(1);
        result.setRowspan(2);
        result.setTitle("测试");

        Result result1 = new Result();
        result1.setColspan(1);
        result1.setRowspan(2);
        result1.setTitle("姓名");

        Result result2 = new Result();
        result2.setColspan(1);
        result2.setRowspan(2);
        result2.setTitle("年龄");

        Result result3 = new Result();
        result3.setColspan(1);
        result3.setRowspan(2);
        result3.setTitle("住址");

        Result result4 = new Result();
        result4.setColspan(1);
        result4.setRowspan(2);
        result4.setTitle("性别");

        Result result5 = new Result();
        result5.setColspan(2);
        result5.setRowspan(1);
        result5.setTitle("aaaa");

        Result result6 = new Result();
        result6.setColspan(1);
        result6.setRowspan(1);
        result6.setTitle("bbb");
        Result result7 = new Result();
        result7.setColspan(1);
        result7.setRowspan(1);
        result7.setTitle("ccc");

        Result result8 = new Result();
        result8.setColspan(2);
        result8.setRowspan(1);
        result8.setTitle("ddd");


        Result result9 = new Result();
        result9.setColspan(1);
        result9.setRowspan(1);
        result9.setTitle("eee");

        Result result10 = new Result();
        result10.setColspan(1);
        result10.setRowspan(1);
        result10.setTitle("fff");



        List<Result> list = new ArrayList<>();
        list.add(result6);
        list.add(result7);
        result5.setList(list);

        List<Result> list1 = new ArrayList<>();
        list1.add(result9);
        list1.add(result10);
        result8.setList(list1);


        List<Result> resultList = new ArrayList<>();

        resultList.add(result);
        resultList.add(result1);
        resultList.add(result2);
        resultList.add(result3);
        resultList.add(result4);
        resultList.add(result5);
        resultList.add(result8);
        List<Student> studentList = new ArrayList<Student>();
        Student student = new Student();
        student.setName("张三");
        student.setAge("28");
        student.setAddress("河北");
        student.setSex("男");
        student.setFirstName("张四");
        student.setLastName("张五");

        Student student1 = new Student();
        student1.setName("李三");
        student1.setAge("29");
        student1.setAddress("河北");
        student1.setSex("女");
        student1.setFirstName("李四");
        student1.setLastName("李五");

        studentList.add(student);
        studentList.add(student1);

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        HSSFSheet sheet = workbook.createSheet("sheet");


        insertHeader(sheet,resultList,0,workbook);
        List<Map<String, Object>> objectToMap =new ArrayList<Map<String, Object>> ();
        try{
           objectToMap = getObjectToMap(studentList);
        }catch (IllegalAccessException ex){

        }
        if(objectToMap!=null){
            fillData(sheet,objectToMap,2,1);
        }
        File file = new File("D:\\demo11.xls");
        FileOutputStream fout = new FileOutputStream(file);
        workbook.write(fout);
        fout.close();
    }


    public static int insertHeader(HSSFSheet sheet, List<Result> resultList,
                                   int rowIndex, HSSFWorkbook wb){

        HSSFRow row = sheet.createRow(rowIndex);
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //横向的单元格的列数,这里设置一个方法外的变量，初始化为零
        int  nextCellIndex  = 0;
        int rowNumber = setSubCellValue(sheet, rowIndex,
                style, row, nextCellIndex, resultList);
        return rowNumber;

    }

    private static int setSubCellValue(HSSFSheet sheet, int rowIndex, HSSFCellStyle style, HSSFRow row, int nextCellIndex, List<Result> resultList) {
        HSSFRow nextRow = null;
        //遍历json格式的字段描述,并横向生成单元格
        for (int i = 0; i < resultList.size(); i++) {
            Result column =  resultList.get(i);
            List<Result> list = column.getList();
            if (list != null) {
                //设置父表格,并返回下一个单元格的列数
                nextCellIndex = setParentCellValue(sheet, style, row,
                        nextCellIndex, list.size(),column);
                nextCellIndex = nextCellIndex + 1;
                //因为存在子表格，所以生成下一行
                if (nextRow == null) {
                    nextRow = sheet.createRow(++rowIndex);
                }
                //使子单元格的位置，回退的到起始位置
                int subCellStartPosition = nextCellIndex - list.size();
                //递归生成子单元格,这里参数可以简化
                setSubCellValue(sheet, rowIndex,
                        style, nextRow, subCellStartPosition, list);
            } else {
                //如果不是复合表头，直接生成并设置单元格
                nextCellIndex = setCellValue(sheet, style, row, nextCellIndex,column);
                CellRangeAddress region = new CellRangeAddress(0, column.getRowspan()-1, nextCellIndex, column.getColspan()-1+nextCellIndex);
                sheet.addMergedRegion(region);
                ++nextCellIndex;
            }
        }
        return rowIndex;
    }

    private static int setParentCellValue(HSSFSheet sheet, HSSFCellStyle style, HSSFRow row, int cellIndex, int size,Result result) {
        cellIndex = setCellValue(sheet,style, row, cellIndex,result);
        int columnFrom = cellIndex;
        int columnTo = columnFrom + size - 1;
        CellRangeAddress region = new CellRangeAddress(0, result.getRowspan()-1, columnFrom, columnTo);
        sheet.addMergedRegion(region);
        return columnTo;
    }

    private static int setCellValue(HSSFSheet sheet, HSSFCellStyle style, HSSFRow row, int cellIndex,Result result) {
        HSSFCell subCell = row.createCell(cellIndex);
        subCell.setCellStyle(style);
        subCell.setCellValue(result.getTitle().toString());
        return cellIndex;
    }


    /**
     * 填充数据
     * @param sheet
     * @param dataList 数据
     * @param rownum 开始的行
     * @param colnum 开始的列
     */
    public static void fillData(HSSFSheet sheet,List<Map<String, Object>> dataList, int rownum,int colnum){
       int total =dataList.size();
       List<String>filedList= new ArrayList<>();
        filedList.add("name");
        filedList.add("age");
        filedList.add("address");
        filedList.add("sex");
        filedList.add("firstName");
        filedList.add("lastName");
       final int size = dataList.get(0).size();
       for(int i=0;i<total;i++){
           HSSFRow row = sheet.createRow(rownum + i);
           for(int j=0;j<size;j++){
               HSSFCell cell = row.createCell(colnum + j);
               Map<String, Object> stringObjectMap = dataList.get(i);
               cell.setCellValue(stringObjectMap.get(filedList.get(j)).toString());
           }

       }
    }

    public static List<Map<String, Object>> getObjectToMap(List<? extends Object> list) throws IllegalAccessException {
        List<Map<String, Object>> dataList=new ArrayList<Map<String, Object>>();
        for(int i=0;i<list.size();i++){
            Map<String, Object> map = new LinkedHashMap<String, Object>();
            final Object obj = list.get(i);
            Class<?> clazz = obj.getClass();
            for (Field field : clazz.getDeclaredFields()) {
                field.setAccessible(true);
                String fieldName = field.getName();
                try{
                    Object value = field.get(obj);
                    if (value == null){
                        value = "";
                    }
                    map.put(fieldName, value);
                }catch (IllegalAccessException ex){
                    ex.printStackTrace();
                }
            }
            dataList.add(map);
        }
        return dataList;
    }
}
