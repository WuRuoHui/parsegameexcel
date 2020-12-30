package top.wu.parsegameexcel.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.Map;

public interface ParseJLLService {

    public Map<String,XSSFWorkbook> ParseJLLExcel(MultipartFile file);

    XSSFWorkbook MergeExcel(MultipartFile[] files);

}
