package ar.com.leo.etiquetas.parser;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import ar.com.leo.etiquetas.model.ExcelMapping;

public class ExcelMappingReader {

    public ExcelMapping readMapping(Path excelPath) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(excelPath.toFile(), PackageAccess.READ);
             Workbook workbook = new XSSFWorkbook(pkg)) {
            return readFromWorkbook(workbook);
        }
    }

    private ExcelMapping readFromWorkbook(Workbook workbook) {
        Map<String, String> zoneMapping = new HashMap<>();
        Map<String, String> extCodeMapping = new HashMap<>();
        Sheet sheet = workbook.getSheetAt(0);

        int skuCol = -1;
        int zoneCol = -1;
        int extCodeCol = -1;
        int headerRowIdx = -1;

        Row headerRow = sheet.getRow(2); // Fila 3 (índice 0)
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell == null) continue;
                String header = getCellStringValue(cell).toLowerCase().trim();
                if (header.equals("código producto")) {
                    skuCol = i;
                }
                if (header.contains("unidad")) {
                    zoneCol = i;
                }
                if (header.contains("código externo")) {
                    extCodeCol = i;
                }
            }
            if (skuCol != -1 && zoneCol != -1) {
                headerRowIdx = 2;
            }
        }

        if (headerRowIdx == -1) {
            throw new IllegalArgumentException(
                    "No se encontraron las columnas requeridas en el Excel. "
                    + "Se necesita una columna con 'Código'/'SKU' y otra con 'Unidad'.");
        }

        for (int r = headerRowIdx + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Cell skuCell = row.getCell(skuCol);
            Cell zoneCell = row.getCell(zoneCol);
            if (skuCell == null || zoneCell == null) continue;

            String sku = getCellStringValue(skuCell).trim();
            String zoneName = getCellStringValue(zoneCell).trim();

            if (!sku.isEmpty() && !zoneName.isEmpty()) {
                zoneMapping.put(sku, zoneName);
            }

            if (extCodeCol != -1) {
                Cell extCodeCell = row.getCell(extCodeCol);
                if (extCodeCell != null) {
                    String extCode = getCellStringValue(extCodeCell).trim();
                    if (!sku.isEmpty() && !extCode.isEmpty()) {
                        extCodeMapping.put(sku, extCode);
                    }
                }
            }
        }

        return new ExcelMapping(zoneMapping, extCodeMapping);
    }

    private String getCellStringValue(Cell cell) {
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultType();
        }
        return switch (type) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    yield String.valueOf((long) val);
                }
                yield String.valueOf(val);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
