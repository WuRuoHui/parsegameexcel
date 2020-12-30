package top.wu.parsegameexcel.controller;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import top.wu.parsegameexcel.parseexcel.MergeExcel;
import top.wu.parsegameexcel.parseexcel.SingleDayGift;
import top.wu.parsegameexcel.parseexcel.TotalMoneyGift;
import top.wu.parsegameexcel.service.ParseJLLService;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Controller
public class UploadController {

    @Autowired
    private ParseJLLService parseJLLService;

    @PostMapping("/merge")
    @ResponseBody
    public String uploadFileForMerge(MultipartFile[] files, HttpServletResponse response) throws IOException {
        if (files == null || files.length < 1){
            return "请选择文件";
        }
        XSSFWorkbook xssfWorkbook = parseJLLService.MergeExcel(files);
        response.setContentType("application/msexcel");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String("合并.xlsx".getBytes("utf-8"), "ISO-8859-1"));
        response.addHeader("Access-Control-Allow-Origin", "*");
        xssfWorkbook.write(response.getOutputStream());
        xssfWorkbook.close();
        return "merge success";
    }

    @PostMapping("/upload")
    @ResponseBody
    public String uploadFile(MultipartFile file, HttpServletResponse response) throws IOException {
        Map<String, XSSFWorkbook> map = parseJLLService.ParseJLLExcel(file);
        String originalFilename = file.getOriginalFilename();
        originalFilename = originalFilename.substring(0,originalFilename.lastIndexOf("."))+".zip";
        //初始化返回
        Set<String> fileNames = map.keySet();
//        ZipInputStream zipInputStream= null;
        response.reset();
        response.setCharacterEncoding("UTF-8");
//        response.setContentType("application/msexcel");
        response.setContentType("application/x-zip-compressed");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String(originalFilename.getBytes("utf-8"), "ISO-8859-1"));
        response.addHeader("Access-Control-Allow-Origin", "*");
//        response.setHeader("Content-Disposition", "attachment; filename=" + testName);
        ZipOutputStream zipOutputStream = new ZipOutputStream(response.getOutputStream());
        for (String fileName : fileNames) {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            XSSFWorkbook xssfWorkbook = map.get(fileName);
            xssfWorkbook.write(byteArrayOutputStream);
            byte[] bytes = byteArrayOutputStream.toByteArray();
            ZipEntry zipEntry = new ZipEntry(fileName+".xlsx");
            zipOutputStream.putNextEntry(zipEntry);
            zipOutputStream.write(bytes);
            xssfWorkbook.close();
            byteArrayOutputStream.close();
        }
        zipOutputStream.flush();
        zipOutputStream.close();
        return "success";
    }

    @GetMapping("/upload")
    public String uploadPage() {
        return "upload";
    }

/*    @PostMapping("/merge")
    @ResponseBody
    public String uploadFileForMerge(MultipartFile[] files, HttpServletResponse response) throws IOException {
        if (files == null || files.length < 1){
            return "请选择文件";
        }
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(files[0].getInputStream());

        response.setContentType("application/msexcel");

//        response.setContentType("application/:octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=" + new String("合并.xlsx".getBytes("utf-8"), "ISO-8859-1"));
        response.addHeader("Access-Control-Allow-Origin", "*");
        xssfWorkbook.write(response.getOutputStream());
        xssfWorkbook.close();
        return "merge success";
    }*/
}
