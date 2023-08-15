package ruby.api.excel;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
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

    /**
     * 템플릿 파일에 데이터를 적용하여 반환
     */
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
            XSSFRow row = sheet.createRow(i);

            for (int j = 0; j < rowData.size(); j++) {
                XSSFCell cell = createCellByRow(row, j, cellStyles.get(j));
                setCellValue(cell, rowData.get(j));
            }
        }

        return workbook;
    }

    /**
     * 템플릿 파일로부터 Workbook 객체 생성
     */
    public XSSFWorkbook getWorkbookByTemplate(String templateName) {
        String templatePath = "/static/excel/" + templateName;
        Resource resource = new ClassPathResource(templatePath);
        try (InputStream inputStream = new FileInputStream(resource.getFile().getPath())){
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 템플릿 파일로부터 Style 이 적용된 Row 반환
     * - 해당 Row 의 Style 읋 재활용하여 새 Row 에 적용
     * - 새 Row 및 Cell 을 생성할 때마다 Style 객체를 생성하면 최대 64000 Style 객체까지만 생성할 수 있으며 비효율적임
     */
    public XSSFRow getStyleRow(XSSFSheet sheet, int rowNum) {
        XSSFRow styleRow = sheet.getRow(rowNum);
        if (styleRow == null) {
            return sheet.createRow(rowNum);
        }

        return styleRow;
    }

    /**
     * Cell 생성 및 스타일 적용
     */
    public XSSFCell createCellByRow(XSSFRow row, int col, CellStyle cellStyle) {
        XSSFCell cell = row.createCell(col);
        cell.setCellStyle(cellStyle);

        return cell;
    }

    /**
     * Cell 에 Value 설정
     */
    public void setCellValue(XSSFCell cell, Object data) {
        if (data.getClass().equals(String.class)) {
            cell.setCellValue((String) data);
        } else if (data.getClass().equals(Integer.class)) {
            cell.setCellValue((Integer) data);
        } else if (data.getClass().equals(Long.class)) {
            cell.setCellValue((Long) data);
        } else if (data.getClass().equals(Short.class)) {
            cell.setCellValue((Short) data);
        } else if (data.getClass().equals(Double.class)) {
            cell.setCellValue((Double) data);
        } else if (data.getClass().equals(LocalDate.class)) {
            cell.setCellValue((LocalDate) data);
        } else if (data.getClass().equals(LocalDateTime.class)) {
            cell.setCellValue((LocalDateTime) data);
        } else if (data.getClass().equals(Date.class)) {
            cell.setCellValue((Date) data);
        } else if (data.getClass().equals(Boolean.class)){
            cell.setCellValue((Boolean) data);
        } else {
            throw new RuntimeException("createCellByRow exception!");
        }
    }

    /**
     * Response 객체에 workbook 쓰기 (다운로드)
     */
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

    /**
     * MultipartFile 객체로부터 데이터를 읽기 위해 Workbook 객체 생성
     */
    public XSSFWorkbook getWorkbookByMultipartFile(MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()){
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
