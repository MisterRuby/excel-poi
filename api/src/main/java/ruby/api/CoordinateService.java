package ruby.api;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import ruby.api.excel.ExcelUtil;

import java.util.ArrayList;
import java.util.List;


@Slf4j
@Service
@RequiredArgsConstructor
public class CoordinateService {

    private final CoordinateRepository coordinateRepository;
    private final ExcelUtil excelUtil;

    public XSSFWorkbook downloadCoordinateExcel()  {
        List<Coordinate> coordinates = coordinateRepository.findAll();
        List<List<Object>> coordinatesRows = coordinates.stream().map(coordinate -> {
                List<Object> row = new ArrayList<>();
                row.add(coordinate.getNodeId());
                row.add(coordinate.getArsId());
                row.add(coordinate.getStationName());
                row.add(coordinate.getLongitude());
                row.add(coordinate.getLatitude());
                return row;
            })
            .toList();

        int startRow = 1;
        return excelUtil.writeDataToExcel(coordinatesRows, "coordinate-template.xlsx", startRow);
    }

    public void uploadCoordinateExcel(MultipartFile file) {
        XSSFWorkbook workbook = excelUtil.getWorkbookByMultipartFile(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int startRowNum = 1;
        int lastRowNum = sheet.getLastRowNum();

        List<Coordinate> coordinates = new ArrayList<>();
        for (int i = startRowNum; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);

            Coordinate coordinate = Coordinate.builder()
                .nodeId(row.getCell(0).getStringCellValue())
                .arsId(row.getCell(1).getStringCellValue())
                .stationName(row.getCell(2).getStringCellValue())
                .longitude(row.getCell(3).getNumericCellValue())
                .latitude(row.getCell(4).getNumericCellValue())
                .build();

            coordinates.add(coordinate);
        }

        coordinateRepository.saveAll(coordinates);
    }
}
