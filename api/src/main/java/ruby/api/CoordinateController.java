package ruby.api;

import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import ruby.api.excel.ExcelUtil;

@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/coordinates")
public class CoordinateController {

    private final CoordinateService coordinateService;
    private final ExcelUtil excelUtil;

    @GetMapping("/excel/download")
    public void downloadExcel(HttpServletResponse response) {
        log.info("excel download");

        XSSFWorkbook workbook = coordinateService.downloadCoordinateExcel();
        excelUtil.write(workbook, response);
    }

    @PostMapping("/excel/upload")
    public void uploadExcel(@RequestPart(value = "file") MultipartFile file) {
        log.info("excel upload {}", file.getName());

        coordinateService.uploadCoordinateExcel(file);
    }
}
