package ruby.api.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

@Component
public class ExcelUtil {

    public XSSFWorkbook getWorkbookByTemplate(String templateName) throws IOException {
        String templatePath = "/static/excel/" + templateName;
        Resource resource = new ClassPathResource(templatePath);
        try (InputStream inputStream = new FileInputStream(resource.getFile().getPath())){
            return new XSSFWorkbook(inputStream);
        }
    }

    public XSSFWorkbook getWorkbookByMultipartFile(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream()){
            return new XSSFWorkbook(inputStream);
        }
    }
}
