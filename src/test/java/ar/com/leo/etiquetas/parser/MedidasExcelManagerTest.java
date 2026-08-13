package ar.com.leo.etiquetas.parser;

import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.model.MedidaSku;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

class MedidasExcelManagerTest {

    @TempDir
    Path tempDir;

    private final MedidasExcelManager manager = new MedidasExcelManager();

    /** Columnas base, sin ninguna de las de embalaje. */
    private static final String[] HEADERS_BASE = {
            "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%", "Alto +20%", "Profunidad +20%",
            "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
    };

    /** Columnas base más las ocho de embalaje entre SUBIDO y ERROR, como las carga el usuario. */
    private static final String[] HEADERS_CON_EMBALAJE = {
            "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
            "Peso físico\n(empaque + producto)\nkg",
            "Ancho +20%", "Alto +20%", "Profunidad +20%",
            "Peso físico (empaque + producto) +20%", "SUBIDO",
            "ESTANDARIZADO", "ENVASE",
            "TIPO DE ROLLO", "CANT PAÑOS", "OBSERVACIONES",
            "ERROR"
    };

    // -------------------------------------------------------------------------------------------
    // Lectura de las columnas de embalaje
    // -------------------------------------------------------------------------------------------

    @Test
    void leeLasColumnasDeEmbalaje() throws Exception {
        Path excel = crearExcel("completo.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "SI", "BOL-1", "DIAMANTES", "4", "Colchon + Tapa"});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertEquals("BOL-1", datos.envase());
        assertEquals("DIAMANTES", datos.rollo());
        assertEquals("4", datos.cantPanos());
        assertEquals("Colchon + Tapa", datos.observaciones());
        assertTrue(datos.estandarizado());
    }

    @Test
    void laInscripcionSaleDeLaHojaDeEstandarizacion() throws Exception {
        Path excel = crearExcel("con-inscripcion.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "SI", "CAJ-1"});
        agregarHojaEstandarizacion(excel, new String[][]{{"CAJ-1", "9Y"}, {"BOL-1", "AYUDIN"}});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertEquals("CAJ-1", datos.envase());
        assertEquals("9Y", datos.inscripcion());
    }

    @Test
    void sinHojaDeEstandarizacionLaInscripcionQuedaVacia() throws Exception {
        Path excel = crearExcel("sin-hoja-est.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "SI", "CAJ-1"});

        assertEquals("", manager.leerMedidas(excel).porSku().get("1241212").embalaje().inscripcion());
    }

    @Test
    void unCodigoQueNoEstaEnLaHojaNoRompeNada() throws Exception {
        Path excel = crearExcel("codigo-desconocido.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "SI", "CAJ-99"});
        agregarHojaEstandarizacion(excel, new String[][]{{"CAJ-1", "9Y"}});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertEquals("CAJ-99", datos.envase());
        assertEquals("", datos.inscripcion());
    }

    @Test
    void elPesoConMargenDeCincoPorCientoNoSeConfundeConElPesoBase() throws Exception {
        // El encabezado del peso base tambien tiene un "+": "(empaque + producto)".
        String[] headers = {
                "SKU", "PRODUCTO",
                "Peso físico (empaque + producto) kg",
                "Peso físico (empaque + producto) +5%",
                "SUBIDO"
        };
        Path excel = crearExcel("peso-5.xlsx", headers, new String[]{"1241212", "Producto A"});
        escribirNumero(excel, 1, 2, 0.8);
        escribirNumero(excel, 1, 3, 0.84);

        MedidaSku medida = manager.leerMedidas(excel).porSku().get("1241212");

        assertEquals(0.8, medida.pesoKg());
        assertEquals(0.84, medida.pesoMasKg());
    }

    @Test
    void unaCeldaDeTextoNoCuentaComoMedida() throws Exception {
        // Solo se sube lo que esta cargado como numero: un texto parseable puede ser un dato mal
        // pegado, y una medida equivocada llega a la publicacion de ML.
        Path excel = crearExcel("texto.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "38,4", "24", "10", "0,8"});

        MedidaSku medida = manager.leerMedidas(excel).porSku().get("1241212");

        assertNull(medida.anchoMasCm());
        assertFalse(medida.tieneMedidasParaSubir());
    }

    @Test
    void unArchivoSinLaColumnaEstandarizadoNoMarcaNada() throws Exception {
        // Un Excel anterior a esta función no tiene ESTANDARIZADO. Sin esto, todos los SKU
        // saldrían con el cartel NO ESTANDARIZADO y el aviso reclamaría el lote entero.
        Path excel = crearExcel("sin-embalaje.xlsx", HEADERS_BASE,
                new String[]{"1241212", "Producto A"});

        MedidaSku medida = manager.leerMedidas(excel).porSku().get("1241212");

        assertFalse(medida.embalaje().aplica(), "sin la columna, la función no aplica");
        assertEquals("Producto A", medida.producto(), "el resto de la fila se sigue leyendo");
    }

    @Test
    void conLaColumnaEstandarizadoEnNoSiAplica() throws Exception {
        Path excel = crearExcel("estandarizado-no.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO", "NO"});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertTrue(datos.aplica(), "la columna está: el SKU sí se reclama");
        assertFalse(datos.estandarizado());
    }

    @Test
    void lasColumnasSeUbicanPorHeaderNoPorPosicion() throws Exception {
        // Mismo contenido pero con las columnas de embalaje en otro orden y al final.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR",
                "OBSERVACIONES", "ENVASE", "TIPO DE ROLLO"
        };
        Path excel = crearExcel("desordenado.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO", "",
                        "Colchon", "BOL-1", "DIAMANTES"});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertEquals("BOL-1", datos.envase());
        assertEquals("DIAMANTES", datos.rollo());
        assertEquals("Colchon", datos.observaciones());
    }

    @Test
    void aceptaPanosEscritoSinEnie() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO", "ROLLO INFLABLE", "CANT PANOS"};
        Path excel = crearExcel("sin-enie.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO", "DIAMANTE", "4"});

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("1241212").embalaje();

        assertEquals("4", datos.cantPanos());
    }

    @Test
    void reconoceElEncabezadoDeTipoDeRollo() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO", "TIPO DE ROLLO"};
        Path excel = crearExcel("rollo.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO", "DIAMANTES"});

        assertEquals("DIAMANTES", manager.leerMedidas(excel).porSku().get("1241212").embalaje().rollo());
    }

    @Test
    void unaCantidadNumericaSeLeeSinDecimales() throws Exception {
        Path excel = crearExcel("numerico.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        escribirNumero(excel, 1, 14, 3);

        assertEquals("3", manager.leerMedidas(excel).porSku().get("1241212").embalaje().cantPanos());
    }

    // -------------------------------------------------------------------------------------------
    // El resto del Excel sigue funcionando
    // -------------------------------------------------------------------------------------------

    @Test
    void agregarPendientesNoEscribeEnLasColumnasDeEmbalaje() throws Exception {
        Path excel = crearExcel("pendientes.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});

        manager.agregarPendientes(excel, List.of("999999"));

        DatosEmbalaje datos = manager.leerMedidas(excel).porSku().get("999999").embalaje();
        assertEquals("", datos.envase(), "las columnas de embalaje quedan como estaban");
        assertFalse(datos.estandarizado());
    }

    @Test
    void agregarPendientesPreservaLasFormulasExistentes() throws Exception {
        Path excel = crearExcel("con-formula.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        escribirFormula(excel, 1, 6, "C2*1.2");

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(6);
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("C2*1.2", cell.getCellFormula());
        }
    }

    @Test
    void marcarResultadosNoBorraLosDatosDeEmbalaje() throws Exception {
        Path excel = crearExcel("marcar.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "SI", "BOL-1"});

        manager.marcarResultados(excel, List.of("1241212"), Map.of());

        assertEquals("BOL-1", manager.leerMedidas(excel).porSku().get("1241212").embalaje().envase());
    }

    @Test
    void unArchivoNuevoSeCreaConLasColumnasDeEmbalaje() throws Exception {
        Path excel = tempDir.resolve("nuevo.xlsx");

        manager.leerMedidas(excel).porSku();

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertEquals("SUBIDO", header.getCell(10).getStringCellValue());
            assertEquals("ESTANDARIZADO", header.getCell(11).getStringCellValue());
            assertEquals("ENVASE", header.getCell(12).getStringCellValue());
            assertEquals("OBSERVACIONES", header.getCell(15).getStringCellValue());
            assertEquals("ERROR", header.getCell(16).getStringCellValue());
        }
    }

    @Test
    void noSeCreaLaHojaCatalogoDeEmbalajes() throws Exception {
        Path excel = tempDir.resolve("sin-catalogo.xlsx");

        manager.leerMedidas(excel).porSku();

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            assertEquals(1, wb.getNumberOfSheets());
            assertNotNull(wb.getSheet("MEDIDAS"));
        }
    }

    @Test
    void laHojaDeMedidasSeUbicaPorSuColumnaSku() throws Exception {
        // El usuario puede tener hojas propias ("Resumen", "Notas") antes de la de medidas.
        Path excel = crearExcel("hoja-ajena.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        agregarHojaAjenaAlPrincipio(excel, "Resumen");

        MedidaSku medida = manager.leerMedidas(excel).porSku().get("1241212");

        assertNotNull(medida, "debe encontrar la hoja de medidas aunque no sea la primera");
    }

    @Test
    void unaMedidaQueMencionaElEnvaseNoSeConfundeConLaColumnaDeEnvase() throws Exception {
        // "Ancho caja cm" es un nombre natural: lo que se mide es la caja.
        String[] headers = {"SKU", "PRODUCTO", "Ancho envase cm", "Alto del envase", "SUBIDO", "ENVASE"};
        Path excel = crearExcel("medida-envase.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "NO", "BOL-1"});
        escribirNumero(excel, 1, 2, 30);
        escribirNumero(excel, 1, 3, 20);

        MedidaSku medida = manager.leerMedidas(excel).porSku().get("1241212");

        assertEquals(30.0, medida.anchoCm(), "la medida no debe ir a parar a la columna de envase");
        assertEquals(20.0, medida.altoCm());
        assertEquals("BOL-1", medida.embalaje().envase());
    }

    @Test
    void agregarPendientesUbicaLasColumnasPorHeader() throws Exception {
        // Columnas de embalaje intercaladas antes de las +20%, como habilita el README.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "ENVASE", "TIPO DE ROLLO", "CANT PAÑOS",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
        };
        Path excel = crearExcel("intercaladas.xlsx", headers,
                new String[]{"", "", "", "", "", "", "BOL-1", "DIAMANTES", "2"});

        manager.agregarPendientes(excel, List.of("999999"));

        // La fila con SKU vacío se reusa: los datos de embalaje que el usuario cargó ahí no se pisan.
        MedidaSku medida = manager.leerMedidas(excel).porSku().get("999999");
        assertEquals("BOL-1", medida.embalaje().envase());
        assertEquals("DIAMANTES", medida.embalaje().rollo());
        assertEquals("2", medida.embalaje().cantPanos());
    }

    @Test
    void agregarPendientesEscribeElNoEnLaColumnaSubido() throws Exception {
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "N° Bolsa", "Nombre Caja", "N° Caja",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO", "ERROR"
        };
        Path excel = crearExcel("subido.xlsx", headers, new String[]{"1241212", "Producto A"});

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Row fila = wb.getSheetAt(0).getRow(2);
            assertEquals("NO", fila.getCell(13).getStringCellValue(), "el NO va en SUBIDO");
            assertEquals(CellType.BLANK, fila.getCell(9).getCellType(), "no en Ancho +20%");
        }
    }

    @Test
    void noPisaLosEncabezadosCuandoSkuNoEstaEnLaPrimeraColumna() throws Exception {
        // El usuario agregó una columna propia a la izquierda, así que SKU quedó en la B.
        String[] headers = {"ID interno", "SKU", "PRODUCTO", "SUBIDO", "ENVASE", "ERROR"};
        Path excel = crearExcel("sku-corrido.xlsx", headers,
                new String[]{"X-1", "1241212", "Producto A", "NO", "BOL-1", ""});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("ID interno", leerHeader(excel, 0), "no se deben reescribir los encabezados");
        assertEquals("SKU", leerHeader(excel, 1));
        assertEquals("BOL-1", manager.leerMedidas(excel).porSku().get("1241212").embalaje().envase());
    }

    @Test
    void noEligeUnaHojaAuxiliarDelUsuarioComoHojaDeMedidas() throws Exception {
        // Una tabla de búsqueda con columna SKU no es la hoja de medidas: no tiene ni SUBIDO ni
        // las columnas de dimensiones, y escribir ahí destruiría los datos del usuario.
        Path excel = crearExcel("medidas.xlsx", HEADERS_CON_EMBALAJE,
                new String[]{"1241212", "Producto A"});
        agregarHojaAuxiliarAlPrincipio(excel, "Catalogo", new String[]{"SKU", "DESCRIPCION"});

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Sheet auxiliar = wb.getSheet("Catalogo");
            assertEquals(2, auxiliar.getRow(0).getLastCellNum(), "la hoja auxiliar no se toca");
            assertEquals(0, auxiliar.getLastRowNum());
        }
        assertNotNull(manager.leerMedidas(excel).porSku().get("999999"), "el SKU va a la hoja de medidas");
    }

    @Test
    void laColumnaErrorSeCreaAlFinalSinPisarColumnasPropias() throws Exception {
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%", "SUBIDO",
                "ESTANDARIZADO", "ENVASE",
                "TIPO DE ROLLO", "CANT PAÑOS", "OBSERVACIONES",
                "MI COLUMNA"
        };
        Path excel = crearExcel("sin-error.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "", "NO",
                        "", "", "", "", "", "dato mio"});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("MI COLUMNA", leerHeader(excel, 16), "no se pisa la columna del usuario");
        assertEquals("ERROR", leerHeader(excel, 17));
        assertEquals("dato mio", leerCelda(excel, 1, 16));
    }

    @Test
    void unaHojaSinEncabezadosNoRompeLaLectura() throws Exception {
        // Un .xlsx recién creado a mano: la app tiene que poder inicializarlo.
        Path excel = tempDir.resolve("hoja-vacia.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("MEDIDAS");
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }

        assertEquals(Map.of(), manager.leerMedidas(excel).porSku());

        manager.agregarPendientes(excel, List.of("999999"));
        assertEquals("SKU", leerHeader(excel, 0));
        assertNotNull(manager.leerMedidas(excel).porSku().get("999999"));
    }

    @Test
    void siFaltaLaColumnaSubidoSeCreaEnVezDeEscribirPorPosicion() throws Exception {
        // El usuario renombró SUBIDO, así que la app no la reconoce. Escribir el "NO" en el índice
        // 10 caería sobre una de sus columnas de embalaje.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho\ncm", "Alto\ncm", "Profundidad\ncm",
                "Peso físico\n(empaque + producto)\nkg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso físico (empaque + producto) +20%",
                "ENVASE", "TIPO DE ROLLO", "CANT PAÑOS"
        };
        Path excel = crearExcel("sin-subido.xlsx", headers,
                new String[]{"1241212", "Producto A", "", "", "", "", "", "", "", "",
                        "BOL-1", "DIAMANTES", "2"});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("ENVASE", leerHeader(excel, 10), "no se pisa la columna de embalaje");
        assertEquals("BOL-1", manager.leerMedidas(excel).porSku().get("1241212").embalaje().envase());
        assertEquals("SUBIDO", leerHeader(excel, 13));
    }

    @Test
    void unaHojaConFilasPeroSinEncabezadosSigueSiendoUnError() throws Exception {
        // Distinto de un archivo recién creado: acá hay datos, así que el usuario eligió el Excel
        // equivocado y tiene que enterarse en vez de que la app escriba encima.
        Path excel = tempDir.resolve("sin-headers.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("MEDIDAS");
            sheet.createRow(1).createCell(0, CellType.STRING).setCellValue("un dato");
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }

        assertThrows(IllegalArgumentException.class, () -> manager.leerMedidas(excel).porSku());
    }

    @Test
    void prefiereLaHojaConSkuAntesQueLaPrimeraCuandoNingunaTieneColumnasDeLaApp() throws Exception {
        // La hoja de medidas del usuario todavía no tiene SUBIDO ni dimensiones, solo SKU y sus
        // columnas de embalaje. Aun así hay que preferirla sobre una hoja sin SKU.
        Path excel = tempDir.resolve("sin-columnas-app.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet portada = wb.createSheet("Instructivo");
            portada.createRow(0).createCell(0, CellType.STRING).setCellValue("Como usar");

            Sheet medidas = wb.createSheet("MEDIDAS");
            Row header = medidas.createRow(0);
            String[] headers = {"SKU", "PRODUCTO", "ENVASE"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            Row fila = medidas.createRow(1);
            fila.createCell(0, CellType.STRING).setCellValue("1241212");
            fila.createCell(2, CellType.STRING).setCellValue("BOL-1");

            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }

        assertEquals("BOL-1", manager.leerMedidas(excel).porSku().get("1241212").embalaje().envase());
    }

    @Test
    void laColumnaErrorSeCreaPegadaALaUltimaEnUnArchivoViejo() throws Exception {
        // Excel anterior a las columnas de embalaje: 11 columnas y sin ERROR.
        String[] headers = {
                "SKU", "PRODUCTO", "Ancho cm", "Alto cm", "Profundidad cm",
                "Peso fisico kg",
                "Ancho +20%", "Alto +20%", "Profunidad +20%",
                "Peso fisico +20%", "SUBIDO"
        };
        Path excel = crearExcel("legacy.xlsx", headers, new String[]{"1241212", "Producto A"});

        manager.agregarPendientes(excel, List.of("999999"));

        assertEquals("ERROR", leerHeader(excel, 11), "sin dejar columnas en blanco en el medio");
    }

    @Test
    void noEscribeEnUnaHojaQueSoloTieneSku() throws Exception {
        // La tabla de búsqueda del usuario tiene SKU pero no es la hoja de medidas: se puede leer
        // de ahí si no hay nada mejor, pero nunca escribirle.
        Path excel = tempDir.resolve("solo-lookup.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet lookup = wb.createSheet("Catalogo");
            Row header = lookup.createRow(0);
            header.createCell(0, CellType.STRING).setCellValue("SKU");
            header.createCell(1, CellType.STRING).setCellValue("DESCRIPCION");
            Row fila = lookup.createRow(1);
            fila.createCell(0, CellType.STRING).setCellValue("1241212");
            fila.createCell(1, CellType.STRING).setCellValue("Producto A");
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }

        manager.agregarPendientes(excel, List.of("999999"));

        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Sheet lookup = wb.getSheet("Catalogo");
            assertEquals(2, lookup.getRow(0).getLastCellNum(), "no se le agregan columnas");
            assertEquals(1, lookup.getLastRowNum(), "no se le agregan filas");
        }
    }

    @Test
    void laColumnaNuevaNoPisaUnaColumnaConDatosSinEncabezado() throws Exception {
        String[] headers = {"SKU", "PRODUCTO", "SUBIDO"};
        Path excel = crearExcel("dato-sin-header.xlsx", headers,
                new String[]{"1241212", "Producto A", "NO"});
        // Justo donde iría la columna nueva según el ancho del encabezado.
        escribirCelda(excel, 1, 3, "dato mio");

        manager.marcarResultados(excel, List.of(), Map.of("1241212", "error de ML"));

        assertEquals("dato mio", leerCelda(excel, 1, 3), "la columna del usuario queda intacta");
    }

    // -------------------------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------------------------

    // -------------------------------------------------------------------------------------------
    // Si la función de embalaje está en uso
    // -------------------------------------------------------------------------------------------

    @Test
    void conLaColumnaEstandarizadoLaFuncionEstaEnUsoAunqueNoHayaFilas() throws Exception {
        // Es el archivo recién creado por la app: trae la columna y ninguna fila todavía.
        // Deducirlo de las filas leídas diría que está apagada justo en el primer lote, que es
        // donde todos los SKU son nuevos y el aviso es el que más importa.
        Path excel = crearExcel("solo-headers.xlsx", HEADERS_CON_EMBALAJE, new String[]{});

        assertTrue(manager.leerMedidas(excel).embalajeEnUso());
    }

    @Test
    void sinLaColumnaEstandarizadoLaFuncionNoEstaEnUso() throws Exception {
        Path excel = crearExcel("sin-estandarizado.xlsx", HEADERS_BASE,
                new String[]{"1241212", "Producto A"});

        assertFalse(manager.leerMedidas(excel).embalajeEnUso());
    }

    @Test
    void unArchivoQueNoExisteSeCreaConLaFuncionEnUso() throws Exception {
        // La app lo crea con sus HEADERS, que incluyen ESTANDARIZADO.
        Path excel = tempDir.resolve("nuevo-con-flag.xlsx");

        assertTrue(manager.leerMedidas(excel).embalajeEnUso());
    }

    private Path crearExcel(String nombre, String[] headers, String[] fila) throws Exception {
        Path excel = tempDir.resolve(nombre);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("MEDIDAS");
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            Row row = sheet.createRow(1);
            for (int i = 0; i < fila.length; i++) {
                row.createCell(i, CellType.STRING).setCellValue(fila[i]);
            }
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
        return excel;
    }

    private void agregarHojaEstandarizacion(Path excel, String[][] filas) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet hoja = wb.createSheet("ESTANDARIZACION");
            Row header = hoja.createRow(0);
            header.createCell(0, CellType.STRING).setCellValue("N°");
            header.createCell(1, CellType.STRING).setCellValue("INSCRIPCION");
            int r = 1;
            for (String[] fila : filas) {
                Row row = hoja.createRow(r++);
                row.createCell(0, CellType.STRING).setCellValue(fila[0]);
                row.createCell(1, CellType.STRING).setCellValue(fila[1]);
            }
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void agregarHojaAuxiliarAlPrincipio(Path excel, String nombre, String[] headers) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet aux = wb.createSheet(nombre);
            Row header = aux.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i, CellType.STRING).setCellValue(headers[i]);
            }
            wb.setSheetOrder(nombre, 0);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void agregarHojaAjenaAlPrincipio(Path excel, String nombre) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            Sheet ajena = wb.createSheet(nombre);
            ajena.createRow(0).createCell(0, CellType.STRING).setCellValue("dato propio");
            wb.setSheetOrder(nombre, 0);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private String leerHeader(Path excel, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheet("MEDIDAS").getRow(0).getCell(col);
            return cell == null ? null : cell.getStringCellValue();
        }
    }

    private String leerCelda(Path excel, int fila, int col) throws Exception {
        try (Workbook wb = WorkbookFactory.create(excel.toFile(), null, true)) {
            Cell cell = wb.getSheet("MEDIDAS").getRow(fila).getCell(col);
            if (cell == null || cell.getCellType() == CellType.BLANK) return "";
            return cell.getStringCellValue();
        }
    }

    private void escribirCelda(Path excel, int fila, int col, String valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.STRING).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void escribirNumero(Path excel, int fila, int col, double valor) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.NUMERIC).setCellValue(valor);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void escribirFormula(Path excel, int fila, int col, String formula) throws Exception {
        try (Workbook wb = abrirParaEditar(excel)) {
            wb.getSheetAt(0).getRow(fila).createCell(col, CellType.FORMULA).setCellFormula(formula);
            try (FileOutputStream fos = new FileOutputStream(excel.toFile())) {
                wb.write(fos);
            }
        }
    }

    /**
     * Carga el workbook completo en memoria y suelta el archivo, para poder reescribirlo encima.
     * WorkbookFactory.create(file, ..., false) mantiene el archivo tomado y lo corrompe al guardar.
     */
    private Workbook abrirParaEditar(Path excel) throws Exception {
        try (InputStream in = Files.newInputStream(excel)) {
            return new XSSFWorkbook(in);
        }
    }
}
