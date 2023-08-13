package ruby.api;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import ruby.api.excel.ExcelUtil;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@Slf4j
@RestController
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelUtil excelUtil;

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        log.info("excel download");

        XSSFWorkbook xssfWorkbook = excelUtil.getWorkbookByTemplate("coordinate-template.xlsx");

        // 다운로드
        String fileName = "coordinateList.xlsx";
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8).replaceAll("\\+", "%20");    //한글파일 명 설정

        response.setContentType("application/vsd.ms-excel");
        response.setHeader("Content-disposition", "attachment; filename=\"" + fileName + "\"");

        try (ServletOutputStream outputStream = response.getOutputStream()){
            xssfWorkbook.write(outputStream);
        }
    }

    @PostMapping("/excel/upload")
    public void uploadExcel(@RequestPart(value = "file") MultipartFile file) throws IOException {
        log.info("excel upload {}", file.getName());

        XSSFWorkbook workbook = excelUtil.getWorkbookByMultipartFile(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        log.info("excel lastRowNum {}", lastRowNum);
    }
}
