package com.gsonkeno.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.ProtocolException;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
/**
 * excel的poi工具类
 * @author gaosong
 * @since 2018-09-26
 */
public class ExcelUtils {

    public static byte[] convertImgToBytes(String imgUrl) throws IOException {
        URL urlConet = new URL(imgUrl);
        HttpURLConnection con = (HttpURLConnection)urlConet.openConnection();
        con.setRequestMethod("GET");
        con.setConnectTimeout(4 * 1000);
        //通过输入流获取图片数据
        InputStream inStream = con .getInputStream();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[2048];
        int len = 0;
        while( (len=inStream.read(buffer)) != -1 ){
            outStream.write(buffer, 0, len);
        }
        inStream.close();
        byte[] data =  outStream.toByteArray();
        return data;
    }

    public static HSSFWorkbook exportExcel2Stream(String title, String[] headers, String[] dataKey, List<Map<String, Object>> excelData, List<Map<String, byte[]>> imgList)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(title);

        sheet.setDefaultColumnWidth(30);

        HSSFCellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor((short)40);
        style.setFillPattern((short)1);
        style.setBorderBottom((short)1);
        style.setBorderLeft((short)1);
        style.setBorderRight((short)1);
        style.setBorderTop((short)1);
        style.setAlignment((short)2);

        HSSFFont font = workbook.createFont();
        font.setColor((short)20);
        font.setFontHeightInPoints((short)12);
        font.setBoldweight((short)700);

        style.setFont(font);

        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor((short)43);
        style2.setFillPattern((short)1);
        style2.setBorderBottom((short)1);
        style2.setBorderLeft((short)1);
        style2.setBorderRight((short)1);
        style2.setBorderTop((short)1);
        style2.setAlignment((short)2);
        style2.setVerticalAlignment((short)1);

        HSSFFont font2 = workbook.createFont();
        font2.setBoldweight((short)400);

        style2.setFont(font2);

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }

        Iterator excelDataIte = excelData.iterator();
        int excelDataIteIndex = 0;
        try {
            while (excelDataIte.hasNext()) {
                excelDataIteIndex++;
                row = sheet.createRow(excelDataIteIndex);
                Map excelDataMap = (Map)excelDataIte.next();
                for (int dataKeyIndex = 0; dataKeyIndex < dataKey.length; dataKeyIndex++) {
                    HSSFCell cell = row.createCell(dataKeyIndex);
                    cell.setCellStyle(style2);
                    String excelDataMapKey = dataKey[dataKeyIndex];
                    Object value = excelDataMap.get(excelDataMapKey);
                    if ((imgList != null) && (imgList.size() > 0) && ((imgList.get(0)).containsKey(excelDataMapKey)))
                    {
                        row.setHeightInPoints(60.0F);

                        sheet.setColumnWidth(dataKeyIndex, 2856);
                        byte[] bsValue = (byte[])((Map)imgList.get(excelDataIteIndex - 1)).get(excelDataMapKey);
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0,
                                1023, 255, (short)dataKeyIndex,
                                excelDataIteIndex, (short)dataKeyIndex,
                                excelDataIteIndex);
                        anchor.setAnchorType(2);
                        patriarch.createPicture(anchor, workbook.addPicture(bsValue, 5));
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
        return workbook;
    }

    public static void generateExcelFile(Workbook workbook, File file) throws IOException {
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.flush();
        fos.close();
    }
}
