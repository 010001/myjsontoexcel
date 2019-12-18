package com.example.github.fb01001.myjsontoexcel;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.Set;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author fangbin
 * @date 19-7-13
 */

public class ExcelUtil {
  public static void main(String[] args) {
    //File file = new File("/home/fangbin/Desktop/JsonToExcel/response20190712-1");
      BufferedReader reader01 = null;
      BufferedReader reader02 = null;
      BufferedReader reader03 = null;
      try {
          reader01 = new BufferedReader(new InputStreamReader(new FileInputStream("/home/fangbin/Desktop/JsonToExcel/response20190715-1")));
          reader02 = new BufferedReader(new InputStreamReader(new FileInputStream("/home/fangbin/Desktop/JsonToExcel/response20190715-2")));
          reader03 = new BufferedReader(new InputStreamReader(new FileInputStream("/home/fangbin/Desktop/JsonToExcel/response20190715-3")));
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      try{
          String str01 = reader01.readLine();
          String str02 = reader02.readLine();
          String str03 = reader03.readLine();
          JSONObject jsonObject01= JSONObject.parseObject(str01);
          JSONObject jsonObject02= JSONObject.parseObject(str02);
          JSONObject jsonObject03= JSONObject.parseObject(str03);
          JSONArray jSONArray01 = jsonObject01.getJSONArray("data");
          JSONArray jSONArray02 = jsonObject02.getJSONArray("data");
          JSONArray jSONArray03 = jsonObject03.getJSONArray("data");



          int lineNum = 0;

          // 新建文件

          // excel 信息
          try {
              OutputStream outputStream = new FileOutputStream("/home/fangbin/Desktop/JsonToExcel/20190715.xlsx");// 创建工作薄
              WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
              WritableWorkbook writableWorkbook = Workbook.createWorkbook(outputStream);

              // 创建新的一页
              WritableSheet sheet = workbook.createSheet("First Sheet",0);
              Label xuexiao = new Label(0,0,"ID");
              sheet.addCell(xuexiao);
              Label zhuanye = new Label(1,0,"TAG");
              sheet.addCell(zhuanye);


              for(int i = 0; i < jSONArray01.size(); i++){
                  String dataObj = jSONArray01.get(i).toString();
                  JSONObject dataObject= JSONObject.parseObject(dataObj);
                  if("Not an Issue".equals(dataObject.getString("primaryTag"))){
                      lineNum += 1;
                      Label cel01 = new Label(0,lineNum,dataObject.getString("issueInstanceId"));
                      sheet.addCell(cel01);
                      Label cel02 = new Label(1,lineNum,dataObject.getString("primaryTag"));
                      sheet.addCell(cel02);
                      //System.out.println(lineNum + "---" + dataObject.getString("issueInstanceId") + "---" + dataObject.getString("primaryTag"));
                  }

              }
              for(int i = 0; i < jSONArray02.size(); i++){
                  String dataObj = jSONArray02.get(i).toString();
                  JSONObject dataObject= JSONObject.parseObject(dataObj);
                  if("Not an Issue".equals(dataObject.getString("primaryTag"))){
                      lineNum += 1;
                      Label cel01 = new Label(0,lineNum,dataObject.getString("issueInstanceId"));
                      sheet.addCell(cel01);
                      Label cel02 = new Label(1,lineNum,dataObject.getString("primaryTag"));
                      sheet.addCell(cel02);
                      //System.out.println(lineNum + "---" + dataObject.getString("issueInstanceId") + "---" + dataObject.getString("primaryTag"));
                  }
              }
              for(int i = 0; i < jSONArray03.size(); i++){
                  String dataObj = jSONArray03.get(i).toString();
                  JSONObject dataObject= JSONObject.parseObject(dataObj);
                  if("Not an Issue".equals(dataObject.getString("primaryTag"))){
                      lineNum += 1;
                      Label cel01 = new Label(0,lineNum,dataObject.getString("issueInstanceId"));
                      sheet.addCell(cel01);
                      Label cel02 = new Label(1,lineNum,dataObject.getString("primaryTag"));
                      sheet.addCell(cel02);
                      //System.out.println(lineNum + "---" + dataObject.getString("issueInstanceId") + "---" + dataObject.getString("primaryTag"));
                  }
              }

              //把创建的内容写入到输出流中，并关闭输出流
              workbook.write();
              workbook.close();
              outputStream.close();

          } catch (FileNotFoundException e) {
              e.printStackTrace();
          } catch (IOException e) {
              e.printStackTrace();
          } catch (RowsExceededException e) {
              e.printStackTrace();
          } catch (WriteException e) {
              e.printStackTrace();
          }



          /*Set<String> keys = jsonObject.keySet();
          Iterator it = keys.iterator();
          while (it.hasNext()){
              System.out.println(it.next());
          }*/
          /*Object data = jsonObject.get("data");
          JSONObject jsonDataObject= JSONObject.parseObject(String.valueOf(data));
          jsonDataObject.get*/
      } catch (Exception e) {
          e.printStackTrace();
      }




  }
}
