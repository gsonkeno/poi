package com.gsonkeno.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
    }

    @Test
    public void testGenerateExcel() throws IOException {
        byte[] imgbytes1 = ExcelUtils.convertImgToBytes("https://gd4.alicdn.com/imgextra/i2/40284249/TB2Pb0TiyCYBuNkHFCcXXcHtVXa_!!40284249.jpg_400x400.jpg");
        byte[] imgbytes2 = ExcelUtils.convertImgToBytes("https://img.alicdn.com/imgextra/i1/2192788400/TB2j_57ahYaK1RjSZFnXXa80pXa_!!2192788400-0-item_pic.jpg");
        String title = "sheet1";
        String[] headers = {"图片","分数"};
        String[] dataKey = {"img_url", "score"};

        List<Map<String,Object>> excelData = new ArrayList<>();
        Map<String,Object>  excelDataMap1 = new HashMap<>();
        excelDataMap1.put("score", "0.87");
        excelData.add(excelDataMap1);
        Map<String,Object>  excelDataMap2 = new HashMap<>();
        excelDataMap2.put("score", "0.85");
        excelData.add(excelDataMap2);

        List<Map<String,byte[]>> imgList = new ArrayList<>();
        Map<String,byte[]>  imgMap = new HashMap<>();
        imgMap.put("img_url",imgbytes1);
        imgList.add(imgMap);
        Map<String,byte[]>  imgMap1 = new HashMap<>();
        imgMap1.put("img_url",imgbytes2);
        imgList.add(imgMap1);

        HSSFWorkbook workbook = ExcelUtils.exportExcel2Stream(title, headers, dataKey, excelData, imgList);

        ExcelUtils.generateExcelFile(workbook, new File("结果.xls"));
    }

}
