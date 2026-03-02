package ar.com.leo.pickit.excel;

import ar.com.leo.AppLogger;
import ar.com.leo.pickit.model.CarrosItem;
import ar.com.leo.pickit.model.CarrosOrden;
import ar.com.leo.pickit.model.PickitItem;
import ar.com.leo.pickit.service.PickitGenerator.SlaOrden;
import ar.com.leo.util.Util;

import static ar.com.leo.pickit.service.PickitGenerator.esSkuConError;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class PickitExcelWriter {

    private static final String[] HEADERS = {"SKU", "CANT", "DESCRIPCION", "PROVEEDOR", "SECTOR", "STOCK"};

    public static File generar(List<PickitItem> items, List<CarrosOrden> carrosOrdenes, List<SlaOrden> slaOrdenes, boolean soloHoy) throws Exception {
        Path excelDir = Paths.get(Util.getJarFolder(), "Pickits y Carros");
        Files.createDirectories(excelDir);

        LocalDateTime ahora = LocalDateTime.now();
        String fecha = ahora.format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"));
        File outputFile = excelDir.resolve("PICKIT_" + fecha + ".xlsx").toFile();

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("PICKIT");

            Font fontNormal = workbook.createFont();
            fontNormal.setFontName("Calibri");
            fontNormal.setFontHeightInPoints((short) 14);

            Font fontBoldUnderline = workbook.createFont();
            fontBoldUnderline.setFontName("Calibri");
            fontBoldUnderline.setFontHeightInPoints((short) 14);
            fontBoldUnderline.setBold(true);
            fontBoldUnderline.setUnderline(Font.U_SINGLE);

            Font fontHeader = workbook.createFont();
            fontHeader.setFontName("Calibri");
            fontHeader.setFontHeightInPoints((short) 14);
            fontHeader.setBold(true);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(fontHeader);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(headerStyle);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle normalStyle = workbook.createCellStyle();
            normalStyle.setFont(fontNormal);
            normalStyle.setAlignment(HorizontalAlignment.CENTER);
            normalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(normalStyle);

            CellStyle boldUnderlineStyle = workbook.createCellStyle();
            boldUnderlineStyle.setFont(fontBoldUnderline);
            boldUnderlineStyle.setAlignment(HorizontalAlignment.CENTER);
            boldUnderlineStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(boldUnderlineStyle);

            CellStyle separatorStyle = workbook.createCellStyle();
            separatorStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
            separatorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            setBorders(separatorStyle);

            Font fontError = workbook.createFont();
            fontError.setFontName("Calibri");
            fontError.setFontHeightInPoints((short) 14);
            fontError.setBold(true);

            CellStyle errorStyle = workbook.createCellStyle();
            errorStyle.setFont(fontError);
            errorStyle.setAlignment(HorizontalAlignment.CENTER);
            errorStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(errorStyle);
            errorStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
            errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle warningStyle = workbook.createCellStyle();
            warningStyle.setFont(fontNormal);
            warningStyle.setAlignment(HorizontalAlignment.CENTER);
            warningStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(warningStyle);
            warningStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            warningStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle stockBajoStyle = workbook.createCellStyle();
            stockBajoStyle.setFont(fontNormal);
            stockBajoStyle.setAlignment(HorizontalAlignment.CENTER);
            stockBajoStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorders(stockBajoStyle);
            stockBajoStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            stockBajoStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle normalWrap = workbook.createCellStyle();
            normalWrap.cloneStyleFrom(normalStyle);
            normalWrap.setWrapText(true);

            CellStyle boldUnderlineWrap = workbook.createCellStyle();
            boldUnderlineWrap.cloneStyleFrom(boldUnderlineStyle);
            boldUnderlineWrap.setWrapText(true);

            CellStyle errorWrap = workbook.createCellStyle();
            errorWrap.cloneStyleFrom(errorStyle);
            errorWrap.setWrapText(true);

            CellStyle warningWrap = workbook.createCellStyle();
            warningWrap.cloneStyleFrom(warningStyle);
            warningWrap.setWrapText(true);

            Font fontTitle = workbook.createFont();
            fontTitle.setFontName("Calibri");
            fontTitle.setFontHeightInPoints((short) 18);
            fontTitle.setBold(true);

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setFont(fontTitle);
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            titleStyle.setBorderTop(BorderStyle.THICK);
            titleStyle.setBorderBottom(BorderStyle.THICK);
            titleStyle.setBorderLeft(BorderStyle.THICK);
            titleStyle.setBorderRight(BorderStyle.THICK);

            sheet.setDefaultRowHeightInPoints(36);

            Row titleRow = sheet.createRow(0);
            String fechaHora = ahora.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = titleRow.createCell(i);
                if (i == 0) cell.setCellValue("PICKIT KT - " + fechaHora + " | Despacho ML: " + (soloHoy ? "Hoy" : "Sin límite"));
                cell.setCellStyle(titleStyle);
            }
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, HEADERS.length - 1));

            Row headerRow = sheet.createRow(1);
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowIndex = 2;
            String lastUnidadPrefix = null;

            for (PickitItem item : items) {
                String currentUnidadPrefix = getUnidadPrefix(item.getUnidad());
                if (lastUnidadPrefix != null && !currentUnidadPrefix.equals(lastUnidadPrefix)) {
                    Row sepRow = sheet.createRow(rowIndex++);
                    for (int i = 0; i < HEADERS.length; i++) {
                        sepRow.createCell(i).setCellStyle(separatorStyle);
                    }
                }
                lastUnidadPrefix = currentUnidadPrefix;

                String codigo = item.getCodigo();
                boolean esError = esSkuConError(codigo);
                String descripcion = item.getDescripcion() != null ? item.getDescripcion() : "";
                String unidad = item.getUnidad() != null ? item.getUnidad() : "";
                boolean esAdvertencia = !esError && (descripcion.isBlank() || unidad.isBlank());
                boolean destacar = item.getCantidad() > 1;
                CellStyle style;
                CellStyle descStyle;
                if (esError) { style = errorStyle; descStyle = errorWrap; }
                else if (esAdvertencia) { style = warningStyle; descStyle = warningWrap; }
                else if (destacar) { style = boldUnderlineStyle; descStyle = boldUnderlineWrap; }
                else { style = normalStyle; descStyle = normalWrap; }

                Row row = sheet.createRow(rowIndex++);

                Cell cellCodigo = row.createCell(0);
                cellCodigo.setCellValue(codigo);
                cellCodigo.setCellStyle(style);

                Cell cellCantidad = row.createCell(1);
                double cant = item.getCantidad();
                if (cant == Math.floor(cant)) cellCantidad.setCellValue((int) cant);
                else cellCantidad.setCellValue(cant);
                cellCantidad.setCellStyle(style);

                Cell cellDesc = row.createCell(2);
                cellDesc.setCellValue(descripcion);
                cellDesc.setCellStyle(descStyle);

                Cell cellProv = row.createCell(3);
                cellProv.setCellValue(item.getProveedor() != null ? item.getProveedor() : "");
                cellProv.setCellStyle(style);

                Cell cellUnidad = row.createCell(4);
                cellUnidad.setCellValue(unidad);
                cellUnidad.setCellStyle(style);

                Cell cellStock = row.createCell(5);
                double stock = item.getStockDisponible();
                if (stock == Math.floor(stock)) cellStock.setCellValue((int) stock);
                else cellStock.setCellValue(stock);
                boolean stockInsuficiente = !esError && !esAdvertencia && stock < item.getCantidad();
                cellStock.setCellStyle(stockInsuficiente ? stockBajoStyle : style);
            }

            CellRangeAddress tablaCompleta = new CellRangeAddress(0, rowIndex - 1, 0, HEADERS.length - 1);
            RegionUtil.setBorderTop(BorderStyle.THICK, tablaCompleta, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THICK, tablaCompleta, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THICK, tablaCompleta, sheet);
            RegionUtil.setBorderRight(BorderStyle.THICK, tablaCompleta, sheet);

            for (int i = 0; i < HEADERS.length; i++) sheet.autoSizeColumn(i);

            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
            printSetup.setLandscape(false);
            printSetup.setFitWidth((short) 1);
            printSetup.setFitHeight((short) 0);
            sheet.setFitToPage(true);
            sheet.setMargin(Sheet.LeftMargin, 0.25);
            sheet.setMargin(Sheet.RightMargin, 0.25);
            sheet.setMargin(Sheet.TopMargin, 0.4);
            sheet.setMargin(Sheet.BottomMargin, 0.4);
            sheet.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
            sheet.setHorizontallyCenter(true);
            sheet.getFooter().setCenter("Página &P de &N");

            generarHojaCarros(workbook, carrosOrdenes, ahora, soloHoy);
            generarHojaSla(workbook, slaOrdenes, ahora);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }

        return outputFile;
    }

    private static String getUnidadPrefix(String unidad) {
        if (unidad == null || unidad.isEmpty()) return "";
        return unidad.length() >= 2 ? unidad.substring(0, 2).toUpperCase() : unidad.toUpperCase();
    }

    private static void setBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private static final String[] CARROS_HEADERS = {"# de venta", "Unidades", "SKU", "Producto", "Sector", "Carro"};

    private static void generarHojaCarros(XSSFWorkbook workbook, List<CarrosOrden> carrosOrdenes, LocalDateTime ahora, boolean soloHoy) {
        Sheet sheet = workbook.createSheet("CARROS");

        Font fontNormal = workbook.createFont();
        fontNormal.setFontName("Calibri");
        fontNormal.setFontHeightInPoints((short) 11);

        Font fontBoldUnderline = workbook.createFont();
        fontBoldUnderline.setFontName("Calibri");
        fontBoldUnderline.setFontHeightInPoints((short) 11);
        fontBoldUnderline.setBold(true);
        fontBoldUnderline.setUnderline(Font.U_SINGLE);

        Font fontHeader = workbook.createFont();
        fontHeader.setFontName("Calibri");
        fontHeader.setFontHeightInPoints((short) 13);
        fontHeader.setBold(true);

        byte[] greenRgb = {(byte) 51, (byte) 153, (byte) 51};
        XSSFColor greenColor = new XSSFColor(greenRgb, null);

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(fontHeader);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THICK);
        headerStyle.setBorderBottom(BorderStyle.THICK);
        headerStyle.setBorderLeft(BorderStyle.THICK);
        headerStyle.setBorderRight(BorderStyle.THICK);
        headerStyle.setFillForegroundColor(greenColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font fontBold = workbook.createFont();
        fontBold.setFontName("Calibri");
        fontBold.setFontHeightInPoints((short) 11);
        fontBold.setBold(true);

        XSSFCellStyle orderStyle = workbook.createCellStyle();
        orderStyle.setFont(fontBold);
        orderStyle.setAlignment(HorizontalAlignment.CENTER);
        orderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorders(orderStyle);
        orderStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        orderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        byte[] lightGreyRgb = {(byte) 243, (byte) 243, (byte) 243};
        XSSFColor lightGreyColor = new XSSFColor(lightGreyRgb, null);

        XSSFCellStyle itemStyle = workbook.createCellStyle();
        itemStyle.setFont(fontNormal);
        itemStyle.setAlignment(HorizontalAlignment.CENTER);
        itemStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorders(itemStyle);
        itemStyle.setFillForegroundColor(lightGreyColor);
        itemStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle itemBoldStyle = workbook.createCellStyle();
        itemBoldStyle.setFont(fontBoldUnderline);
        itemBoldStyle.setAlignment(HorizontalAlignment.CENTER);
        itemBoldStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        itemBoldStyle.setFillForegroundColor(lightGreyColor);
        itemBoldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(itemBoldStyle);

        XSSFCellStyle itemErrorStyle = workbook.createCellStyle();
        itemErrorStyle.setFont(fontBold);
        itemErrorStyle.setAlignment(HorizontalAlignment.CENTER);
        itemErrorStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        itemErrorStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
        itemErrorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(itemErrorStyle);

        XSSFCellStyle itemWarningStyle = workbook.createCellStyle();
        itemWarningStyle.setFont(fontNormal);
        itemWarningStyle.setAlignment(HorizontalAlignment.CENTER);
        itemWarningStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        itemWarningStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        itemWarningStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(itemWarningStyle);

        XSSFCellStyle itemWrap = workbook.createCellStyle();
        itemWrap.cloneStyleFrom(itemStyle);
        itemWrap.setWrapText(true);
        XSSFCellStyle itemBoldWrap = workbook.createCellStyle();
        itemBoldWrap.cloneStyleFrom(itemBoldStyle);
        itemBoldWrap.setWrapText(true);
        XSSFCellStyle itemErrorWrap = workbook.createCellStyle();
        itemErrorWrap.cloneStyleFrom(itemErrorStyle);
        itemErrorWrap.setWrapText(true);
        XSSFCellStyle itemWarningWrap = workbook.createCellStyle();
        itemWarningWrap.cloneStyleFrom(itemWarningStyle);
        itemWarningWrap.setWrapText(true);

        Font fontTitle = workbook.createFont();
        fontTitle.setFontName("Calibri");
        fontTitle.setFontHeightInPoints((short) 18);
        fontTitle.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(fontTitle);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setBorderTop(BorderStyle.THICK);
        titleStyle.setBorderBottom(BorderStyle.THICK);
        titleStyle.setBorderLeft(BorderStyle.THICK);
        titleStyle.setBorderRight(BorderStyle.THICK);

        sheet.setDefaultRowHeightInPoints(30);

        Row titleRow = sheet.createRow(0);
        String fechaHora = ahora.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
        for (int i = 0; i < CARROS_HEADERS.length; i++) {
            Cell cell = titleRow.createCell(i);
            if (i == 0) cell.setCellValue("CARROS KT - " + fechaHora + " | Despacho ML: " + (soloHoy ? "Hoy" : "Sin límite"));
            cell.setCellStyle(titleStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, CARROS_HEADERS.length - 1));

        Row headerRow = sheet.createRow(1);
        for (int i = 0; i < CARROS_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(CARROS_HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }

        int rowIndex = 2;
        List<int[]> grupoRanges = new java.util.ArrayList<>();

        for (CarrosOrden orden : carrosOrdenes) {
            int grupoStart = rowIndex;

            Row orderRow = sheet.createRow(rowIndex++);
            Cell cellVenta = orderRow.createCell(0);
            cellVenta.setCellValue(orden.getNumeroVenta());
            cellVenta.setCellStyle(orderStyle);
            for (int i = 1; i < CARROS_HEADERS.length - 1; i++) {
                orderRow.createCell(i).setCellStyle(orderStyle);
            }
            Cell cellCarro = orderRow.createCell(CARROS_HEADERS.length - 1);
            cellCarro.setCellValue(orden.getLetraCarro());
            cellCarro.setCellStyle(orderStyle);

            for (CarrosItem item : orden.getItems()) {
                String sku = item.getSku();
                boolean esError = esSkuConError(sku);
                String descripcion = item.getDescripcion() != null ? item.getDescripcion() : "";
                String sector = item.getSector() != null ? item.getSector() : "";
                boolean esAdvertencia = !esError && (descripcion.isBlank() || sector.isBlank());
                boolean destacar = item.getCantidad() > 1;

                CellStyle style;
                CellStyle prodStyle;
                if (esError) { style = itemErrorStyle; prodStyle = itemErrorWrap; }
                else if (esAdvertencia) { style = itemWarningStyle; prodStyle = itemWarningWrap; }
                else if (destacar) { style = itemBoldStyle; prodStyle = itemBoldWrap; }
                else { style = itemStyle; prodStyle = itemWrap; }

                Row itemRow = sheet.createRow(rowIndex++);
                Cell c0 = itemRow.createCell(0); c0.setCellValue(orden.getNumeroVenta()); c0.setCellStyle(style);
                Cell c1 = itemRow.createCell(1);
                double cant = item.getCantidad();
                if (cant == Math.floor(cant)) c1.setCellValue((int) cant); else c1.setCellValue(cant);
                c1.setCellStyle(style);
                Cell c2 = itemRow.createCell(2); c2.setCellValue(sku); c2.setCellStyle(style);
                Cell c3 = itemRow.createCell(3); c3.setCellValue(descripcion); c3.setCellStyle(prodStyle);
                Cell c4 = itemRow.createCell(4); c4.setCellValue(sector); c4.setCellStyle(style);
                Cell c5 = itemRow.createCell(5); c5.setCellStyle(style);
            }
            grupoRanges.add(new int[]{grupoStart, rowIndex - 1});
        }

        for (int[] range : grupoRanges) {
            CellRangeAddress region = new CellRangeAddress(range[0], range[1], 0, CARROS_HEADERS.length - 1);
            RegionUtil.setBorderTop(BorderStyle.THICK, region, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THICK, region, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THICK, region, sheet);
            RegionUtil.setBorderRight(BorderStyle.THICK, region, sheet);
        }

        for (int i = 0; i < CARROS_HEADERS.length; i++) sheet.autoSizeColumn(i);

        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(false);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setMargin(Sheet.LeftMargin, 0.25);
        sheet.setMargin(Sheet.RightMargin, 0.25);
        sheet.setMargin(Sheet.TopMargin, 0.4);
        sheet.setMargin(Sheet.BottomMargin, 0.4);
        sheet.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
        sheet.setHorizontallyCenter(true);
        sheet.getFooter().setCenter("Página &P de &N");
    }

    private static final String[] SLA_HEADERS = {"# de venta", "Items", "SLA", "Despachar"};
    private static final DateTimeFormatter SLA_DATE_FORMAT = DateTimeFormatter.ofPattern("dd/MM HH:mm");

    private static void generarHojaSla(XSSFWorkbook workbook, List<SlaOrden> slaOrdenes, LocalDateTime ahora) {
        Sheet sheet = workbook.createSheet("SLA");

        Font fontNormal = workbook.createFont();
        fontNormal.setFontName("Calibri");
        fontNormal.setFontHeightInPoints((short) 13);

        Font fontBold = workbook.createFont();
        fontBold.setFontName("Calibri");
        fontBold.setFontHeightInPoints((short) 13);
        fontBold.setBold(true);

        byte[] blueRgb = {(byte) 68, (byte) 114, (byte) 196};
        XSSFColor blueColor = new XSSFColor(blueRgb, null);

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(fontBold);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THICK);
        headerStyle.setBorderBottom(BorderStyle.THICK);
        headerStyle.setBorderLeft(BorderStyle.THICK);
        headerStyle.setBorderRight(BorderStyle.THICK);
        headerStyle.setFillForegroundColor(blueColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle normalStyle = workbook.createCellStyle();
        normalStyle.setFont(fontNormal);
        normalStyle.setAlignment(HorizontalAlignment.CENTER);
        normalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorders(normalStyle);

        Font fontTitle = workbook.createFont();
        fontTitle.setFontName("Calibri");
        fontTitle.setFontHeightInPoints((short) 18);
        fontTitle.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(fontTitle);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setBorderTop(BorderStyle.THICK);
        titleStyle.setBorderBottom(BorderStyle.THICK);
        titleStyle.setBorderLeft(BorderStyle.THICK);
        titleStyle.setBorderRight(BorderStyle.THICK);

        sheet.setDefaultRowHeightInPoints(30);

        Row titleRow = sheet.createRow(0);
        String fechaHora = ahora.format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
        for (int i = 0; i < SLA_HEADERS.length; i++) {
            Cell cell = titleRow.createCell(i);
            if (i == 0) cell.setCellValue("SLA ML - " + fechaHora);
            cell.setCellStyle(titleStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, SLA_HEADERS.length - 1));

        Row headerRow = sheet.createRow(1);
        for (int i = 0; i < SLA_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(SLA_HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }

        int rowIndex = 2;
        for (SlaOrden sla : slaOrdenes) {
            Row row = sheet.createRow(rowIndex++);
            Cell c0 = row.createCell(0); c0.setCellValue(sla.numeroVenta()); c0.setCellStyle(normalStyle);
            Cell c1 = row.createCell(1); c1.setCellValue(sla.cantidadItems()); c1.setCellStyle(normalStyle);
            Cell c2 = row.createCell(2); c2.setCellValue(sla.slaStatus() != null ? sla.slaStatus() : ""); c2.setCellStyle(normalStyle);
            Cell c3 = row.createCell(3);
            if (sla.slaExpectedDate() != null) {
                c3.setCellValue(sla.slaExpectedDate().atZoneSameInstant(ZoneId.of("America/Argentina/Buenos_Aires")).format(SLA_DATE_FORMAT));
            }
            c3.setCellStyle(normalStyle);
        }

        if (rowIndex > 2) {
            CellRangeAddress tabla = new CellRangeAddress(0, rowIndex - 1, 0, SLA_HEADERS.length - 1);
            RegionUtil.setBorderTop(BorderStyle.THICK, tabla, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THICK, tabla, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THICK, tabla, sheet);
            RegionUtil.setBorderRight(BorderStyle.THICK, tabla, sheet);
        }

        for (int i = 0; i < SLA_HEADERS.length; i++) sheet.autoSizeColumn(i);

        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(false);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setMargin(Sheet.LeftMargin, 0.25);
        sheet.setMargin(Sheet.RightMargin, 0.25);
        sheet.setMargin(Sheet.TopMargin, 0.4);
        sheet.setMargin(Sheet.BottomMargin, 0.4);
        sheet.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
        sheet.setHorizontallyCenter(true);
        sheet.getFooter().setCenter("Página &P de &N");
    }
}
