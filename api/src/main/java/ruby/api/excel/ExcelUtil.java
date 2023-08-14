package ruby.api.excel;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Component
public class ExcelUtil {

    public XSSFWorkbook writeDataToExcel(List<List<Object>> data, String templateName, int startRow) {
        XSSFWorkbook workbook = getWorkbookByTemplate(templateName);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow styleRow = getStyleRow(sheet, startRow);
        List<CellStyle> cellStyles = new ArrayList<>();
        for (int i = 0; i < styleRow.getLastCellNum(); i++) {
            cellStyles.add(styleRow.getCell(i).getCellStyle());
        }

        for (int i = startRow; i < startRow + data.size(); i++) {
            List<Object> rowData = data.get(i - startRow);
            CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
            sheet.copyRows(startRow, startRow, i, cellCopyPolicy);
            XSSFRow row = sheet.getRow(i);

            for (int j = 0; j < rowData.size(); j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellStyle(cellStyles.get(j));
                Object value = rowData.get(j);
                if (value.getClass().equals(String.class)) {
                    cell.setCellValue((String) value);
                } else if (value.getClass().equals(Integer.class)) {
                    cell.setCellValue((Integer) value);
                } else if (value.getClass().equals(Long.class)) {
                    cell.setCellValue((Long) value);
                } else if (value.getClass().equals(Short.class)) {
                    cell.setCellValue((Short) value);
                } else if (value.getClass().equals(Double.class)) {
                    cell.setCellValue((Double) value);
                } else if (value.getClass().equals(LocalDate.class)) {
                    cell.setCellValue((LocalDate) value);
                } else if (value.getClass().equals(LocalDateTime.class)) {
                    cell.setCellValue((LocalDateTime) value);
                } else if (value.getClass().equals(Date.class)) {
                    cell.setCellValue((Date) value);
                } else {
                    cell.setCellValue((Boolean) value);
                }
            }
        }

        return workbook;
    }

    public XSSFWorkbook getWorkbookByTemplate(String templateName) {
        String templatePath = "/static/excel/" + templateName;
        Resource resource = new ClassPathResource(templatePath);
        try (InputStream inputStream = new FileInputStream(resource.getFile().getPath())){
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public XSSFRow getStyleRow(XSSFSheet sheet, int rowNum) {
        XSSFRow styleRow = sheet.getRow(rowNum);
        if (styleRow == null) {
            return sheet.createRow(rowNum);
        }

        return styleRow;
    }

    public void write(XSSFWorkbook workbook, HttpServletResponse response) {
        String fileName = URLEncoder.encode("coordinateList.xlsx", StandardCharsets.UTF_8).replaceAll("\\+", "%20");
        response.setContentType("application/vsd.ms-excel");
        response.setHeader("Content-disposition", "attachment; filename=\"" + fileName + "\"");

        try (ServletOutputStream outputStream = response.getOutputStream()){
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public XSSFWorkbook getWorkbookByMultipartFile(MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()){
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
