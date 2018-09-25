package com.zy.easyexcelutilexample.controller;

import com.github.zywaiting.easyexcelutil.EasyExcelUtil;
import com.zy.easyexcelutilexample.pojo.ExportInfo;
import com.zy.easyexcelutilexample.pojo.ImportInfo;
import com.zy.easyexcelutilexample.utils.ResponseMessageUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @Author zy
 * @Date 2018-09-24
 */
@RestController
public class TestController {
    /**
     * MultipartFile
     * 读取 Excel（允许多个 sheet）
     * @param excel
     * @return list
     */
    @RequestMapping(value = "readExcelMultipartFile", method = RequestMethod.POST)
    public ResponseMessageUtils readExcelMultipartFile(MultipartFile excel) {
        List<Object> objects = EasyExcelUtil.readExcel(excel, new ImportInfo());
        return ResponseMessageUtils.ok(objects);
    }

    /**
     * MultipartFile
     * 读取 Excel（1个 sheet）
     * @param excel
     * @return list
     * ##读取某个 sheet 的 Excel
     * ###XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
     * ###XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
     * 一下同理
     */
    @RequestMapping(value = "readExcelMultipartFileOne", method = RequestMethod.POST)
    public ResponseMessageUtils readExcelMultipartFileOne(MultipartFile excel) {
        List<Object> objects = EasyExcelUtil.readExcel(excel, new ImportInfo(), 1);
        return ResponseMessageUtils.ok(objects);
    }

    /**
     * File excel
     * @return list
     */
    @RequestMapping(value = "readExcelFile", method = RequestMethod.POST)
    public ResponseMessageUtils readExcelFile() {
        String filePath = "C:\\Users\\zhuyao\\\\Downloads\\一个 Excel 文件.xlsx";
        List<Object> objects = EasyExcelUtil.readExcel(new File(filePath), new ImportInfo());
        return ResponseMessageUtils.ok(objects);
    }

    /**
     * InputStream  and fileName
     * @return list
     */
    @RequestMapping(value = "readExcelInputStream", method = RequestMethod.POST)
    public ResponseMessageUtils readExcelInputStream() {
        String urlString = "http://wq-zy.oss-cn-hangzhou.aliyuncs.com/test/test.xlsx";
        String fileName = "test.xlsx";
        try {
            URL url = new URL(urlString);
            // 打开连接
            URLConnection con = url.openConnection();
            //设置请求超时为5s
            con.setConnectTimeout(5*1000);
            // 输入流
            InputStream is = con.getInputStream();
            List<Object> objects = EasyExcelUtil.readExcel(fileName, is, new ImportInfo());
            for (int i = 0; i < objects.size(); i++) {
                ImportInfo importInfo = (ImportInfo)objects.get(i);
                System.out.println(importInfo.getName() + "..." + importInfo.getAge() + "..." + importInfo.getEmail());
                System.out.println(importInfo.toString());
            }
            return ResponseMessageUtils.ok(objects);
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseMessageUtils.error("异常！");
        }
    }



    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public ResponseMessageUtils writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName, new ExportInfo())
                .finish();
        return ResponseMessageUtils.ok();
    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public ResponseMessageUtils writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
        return ResponseMessageUtils.ok();
    }

    /**
     * 导出 Excel（一个 sheet）到本地
     */

    public static void main(String[] args) throws IOException {
        List<ExportInfo> list = getList();
        String sheetName1 = "第一个 sheet";
        EasyExcelUtil.writeExcelWithSheets("D:\\一个 Excel 文件.xlsx")
                .write(list,sheetName1,new ExportInfo())
                .finish();
    }


    private static List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("howie");
        model1.setAge("19");
        model1.setAddress("123456789");
        model1.setEmail("123456789@gmail.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("harry");
        model2.setAge("20");
        model2.setAddress("198752233");
        model2.setEmail("198752233@gmail.com");
        list.add(model2);
        return list;
    }

}
