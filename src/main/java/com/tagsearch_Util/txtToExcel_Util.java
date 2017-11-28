package com.tagsearch_Util;

import org.apache.poi.hssf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;

/**
 * @author mushuangcheng@jd.com
 * @date 2017/11/27 19:45
 */
public class txtToExcel_Util {
    @Test
    public void txtToExcel() throws IOException {
        // 第一步，创建一个webbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("Sheet1");
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        HSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

        sheet.setDefaultRowHeightInPoints(12);//设置缺省列高
        sheet.setDefaultColumnWidth(15);//设置缺省列宽
        //设置指定列的列宽，256 * 50这种写法是因为width参数单位是单个字符的256分之一
        // sheet.setColumnWidth(cell.getColumnIndex(), 256 * 50);
        HSSFCell cell = row.createCell(1);
        cell.setCellValue("pin");
        cell.setCellStyle(style);
        cell = row.createCell(2);
        cell.setCellValue("购买商品兴趣标签");
        cell.setCellStyle(style);
        cell = row.createCell(3);
        cell.setCellValue("浏览商品兴趣标签");
        cell.setCellStyle(style);
        cell = row.createCell(4);
        cell.setCellValue("三级品类偏好");
        cell.setCellStyle(style);
        cell = row.createCell(5);
        cell.setCellValue("资讯兴趣标签");
        cell.setCellStyle(style);
        cell = row.createCell(6);
        cell.setCellValue("众筹兴趣标签");
        cell.setCellStyle(style);
        cell = row.createCell(7);
        cell.setCellValue("是否浏览过保险");
        cell.setCellStyle(style);
        cell = row.createCell(8);
        cell.setCellValue("是否浏览过基金");
        cell.setCellStyle(style);
        cell = row.createCell(9);
        cell.setCellValue("是否购买过金融商品");
        cell.setCellStyle(style);
        ArrayList<String[]> list = new ArrayList<String[]>();
        ArrayList<String> imeiList = new ArrayList<String>();
        // 第五步，写入实体数据 实际应用中这些数据从数据库得到，


        BufferedReader imeiKeyBR = null;
        BufferedReader keyValueBR = null;
        FileOutputStream fos = null;
        //设置读取 需求给的用来查询生产数据数据的txt文件路径
        imeiKeyBR = new BufferedReader(new FileReader("E:\\jd标签\\my\\android_imei.txt"));
        //设置读取 key查询出来的结果的txt文本路径
        keyValueBR = new BufferedReader(new FileReader("E:\\jd标签\\my\\新建文件夹\\android_imei_out.txt"));
        //设置输出excel文件路径及名称
        fos = new FileOutputStream("E:\\jd标签\\my\\新建文件夹\\1.xls");




        String imeiLine = "";
        while ((imeiLine = imeiKeyBR.readLine()) != null) {
            imeiList.add(imeiLine);
        }
        for (int i = 0; i < imeiList.size(); i++) {
            row = sheet.createRow(i + 1);
            String par = imeiList.get(i);
            // 第四步，创建单元格，并设置值
            row.createCell(0).setCellValue(par);
        }
        String line = "";
        while ((line = keyValueBR.readLine()) != null) {
            line = line.replace("\"{\\\"result\\\":\\\"0\\\",\\\"code\\\":\\\"2001\\\",\\\"msg\\\":\\\"IDMapping error, input type: mobile, input: 15046513391\\\",\\\"data\\\":\\\"\\\"}\"", "\"null\\tnull\\tnull\\tnull\\tnull\\tnull\\t0\\t0\\t0\"");
            line = line.replace("\\\"", "\"");
            line = line.substring(1, line.length() - 1);
            list.add(line.split("\\\\t"));
        }
        for (int i = 0; i < list.size(); i++) {
            row = sheet.getRow(i + 1);
            String[] par = list.get(i);
            // 第四步，并设置值
            for (int j = 0; j < par.length; j++) {
                row.createCell(j + 1).setCellValue(par[j]);
            }
        }
        //第六步,输出Excel文件
        wb.write(fos);
        fos.close();
    }
}
