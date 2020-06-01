package com.hash.excel;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.hash.excel.dojo.DemoData;
import com.hash.excel.dojo.ReadData;
import com.hash.excel.listiner.DemoDataListener;
import com.hash.excel.listiner.ExcelListener;
import jdk.nashorn.internal.ir.annotations.Ignore;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Ignore
public class ReadTest {
    private static List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setSno(i);
            data.setSname("张三"+i);
            list.add(data);
        }
        return list;
    }

    //写入excle
    @Test
    public  void test1(){
        // 写法1
        String fileName = "D:\\11.xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, DemoData.class).sheet("写入方法一").doWrite(data());
    }

    //写入excel
    @Test
    public void test2(){
        // 写法2，方法二需要手动关闭流
        String fileName = "D:\\112.xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(fileName, DemoData.class).build();
            WriteSheet writeSheet = EasyExcel.writerSheet("写入方法二").build();
            excelWriter.write(data(), writeSheet);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            /// 千万别忘记finish 会帮忙关闭流
            excelWriter.finish();
        }
    }

    //读取excle中的内容
    @Test
    public void test3(){
        // 写法1：
        String fileName = "D:\\11.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, ReadData.class, new ExcelListener()).sheet().doRead();
    }


    //读取excel中的内容
    @Test
    public void test4(){
        // 写法2：
        ExcelReader excelReader = null;
        try {
            InputStream in = new BufferedInputStream(new FileInputStream("D:\\11.xlsx"));
            excelReader = EasyExcel.read(in, ReadData.class, new ExcelListener()).build();
            ReadSheet readSheet = EasyExcel.readSheet(0).build();
            excelReader.read(readSheet);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
            excelReader.finish();
        }

    }

}
