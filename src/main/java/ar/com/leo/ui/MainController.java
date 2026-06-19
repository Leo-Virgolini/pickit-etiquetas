package ar.com.leo.ui;

import ar.com.leo.AppLogger;
import ar.com.leo.api.ml.MercadoLibreAPI;
import ar.com.leo.api.ml.model.OrdenML;
import ar.com.leo.api.ml.model.Venta;
import ar.com.leo.etiquetas.model.*;
import ar.com.leo.etiquetas.parser.ComboProduct;
import ar.com.leo.etiquetas.parser.ComboExcelReader;
import ar.com.leo.etiquetas.parser.ExcelMappingReader;
import ar.com.leo.etiquetas.parser.MedidasExcelManager;
import ar.com.leo.etiquetas.parser.ZplParser;
import ar.com.leo.etiquetas.ui.ComboPrintDialog;
import ar.com.leo.etiquetas.ui.LabelTableRow;
import ar.com.leo.etiquetas.ui.OrderTableRow;
import ar.com.leo.pickit.excel.ExcelManager;
import ar.com.leo.pedidos.service.PedidosService;
import ar.com.leo.pickit.model.ProductoManual;
import ar.com.leo.pickit.service.PickitService;
import ar.com.leo.etiquetas.printer.PrinterDiscovery;
import ar.com.leo.etiquetas.printer.ZplFileSaver;
import ar.com.leo.etiquetas.printer.ZplPrinterService;
import ar.com.leo.etiquetas.sorter.LabelSorter;
import javafx.application.Platform;
import javafx.geometry.Pos;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.layout.HBox;
import javafx.util.Duration;
import javafx.scene.layout.VBox;
import javafx.scene.media.AudioClip;
import javafx.scene.paint.Color;
import javafx.scene.text.TextFlow;
import javafx.stage.FileChooser;

import javax.print.PrintService;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.prefs.Preferences;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainController {

    @FXML
    private TabPane tabPane;
    @FXML
    private TabPane etiquetasSubTabPane;
    @FXML
    private TextField zplFileField;
    @FXML
    private TextField excelFileField;
    @FXML
    private TextField comboExcelField;
    @FXML
    private TextField medidasExcelField;
    @FXML
    private CheckBox medidasEnabledCheck;
    @FXML
    private Button medidasSelectBtn;
    @FXML
    private Button subirMedidasBtn;
    @FXML
    private Label medidasStatusLabel;

    private volatile String medidasUltimoDetalle = "";
    private volatile boolean medidasUltimoTuvoError = false;
    private volatile boolean subidaMedidasEnCurso = false;
    @FXML
    private VBox excelSelectorsBox;
    @FXML
    private Label meliStatusLabel;
    @FXML
    private ComboBox<String> estadoFilterCombo;
    @FXML
    private ComboBox<String> despachoFilterCombo;
    @FXML
    private Label statsLabel;
    @FXML
    private HBox statsBar;
    @FXML
    private HBox fileLinkBar;
    @FXML
    private HBox searchBar;
    @FXML
    private TextField searchField;
    @FXML
    private TableView<LabelTableRow> labelTable;
    @FXML
    private TableColumn<LabelTableRow, String> labelOrderCol;
    @FXML
    private TableColumn<LabelTableRow, String> zoneCol;
    @FXML
    private TableColumn<LabelTableRow, String> skuCol;
    @FXML
    private TableColumn<LabelTableRow, String> descCol;
    @FXML
    private TableColumn<LabelTableRow, String> detailsCol;
    @FXML
    private TableColumn<LabelTableRow, Integer> countCol;
    @FXML
    private Button fetchOrdersBtn;
    @FXML
    private Button backToOrdersBtn;
    @FXML
    private Button downloadLabelsBtn;
    @FXML
    private Button comboSheetBtn;
    @FXML
    private Button printDirectBtn;
    @FXML
    private ProgressIndicator progressIndicator;
    @FXML
    private TableView<OrderTableRow> orderTable;
    @FXML
    private TableColumn<OrderTableRow, Boolean> orderSelectCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderIdCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderZoneCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderSkuCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderDescCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderQtyCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderStatusCol;
    @FXML
    private TableColumn<OrderTableRow, String> orderSlaCol;

    // ── Pickit Tab ──
    @FXML
    private RadioButton radioPickitSlaHoy;
    @FXML
    private RadioButton radioPickitSlaTodos;
    @FXML
    private CheckBox pickitCheckTurbo;
    @FXML
    private CheckBox pickitCheckML;
    @FXML
    private CheckBox pickitCheckNube;
    @FXML
    private CheckBox pickitCheckManual;
    @FXML
    private HBox pickitSlaSection;
    @FXML
    private VBox pickitManualSection;
    @FXML
    private TextField pickitSkuField;
    @FXML
    private TextField pickitCantidadField;
    @FXML
    private TableView<ProductoManual> pickitManualTable;
    @FXML
    private TableColumn<ProductoManual, String> pickitColSku;
    @FXML
    private TableColumn<ProductoManual, Double> pickitColCantidad;
    @FXML
    private Button pickitBtnAgregarModificar;
    @FXML
    private ScrollPane pickitLogScrollPane;
    @FXML
    private TextFlow pickitLogTextFlow;
    @FXML
    private ProgressIndicator pickitProgressIndicator;
    @FXML
    private Button pickitGenerateBtn;

    // ── Pedidos Tab ──
    @FXML
    private ScrollPane pedidosLogScrollPane;
    @FXML
    private TextFlow pedidosLogTextFlow;
    @FXML
    private ProgressIndicator pedidosProgressIndicator;
    @FXML
    private Button pedidosGenerateBtn;

    private final ZplParser zplParser = new ZplParser();
    private final ExcelMappingReader excelReader = new ExcelMappingReader();
    private final ComboExcelReader comboExcelReader = new ComboExcelReader();
    private final MedidasExcelManager medidasManager = new MedidasExcelManager();
    private final LabelSorter labelSorter = new LabelSorter();
    private final ZplFileSaver fileSaver = new ZplFileSaver();
    private final ZplPrinterService printerService = new ZplPrinterService();
    private final PrinterDiscovery printerDiscovery = new PrinterDiscovery();
    private final Preferences prefs = Preferences.userRoot().node("etiquetas");

    private static final String PREF_EXCEL_PATH = "excelFilePath";
    private static final String PREF_COMBO_EXCEL_PATH = "comboExcelFilePath";
    private static final String PREF_MEDIDAS_EXCEL_PATH = "medidasExcelFilePath";
    private static final String PREF_MEDIDAS_ENABLED = "medidasEnabled";
    private static final String PREF_ZPL_DIR = "zplLastDir";

    private boolean meliInitialized = false;
    private SortResult currentResult;
    private List<OrdenML> fetchedOrders;
    private Set<Long> turboShipmentIds = Set.of();
    private FilteredList<OrderTableRow> filteredOrders;
    private FilteredList<LabelTableRow> filteredLabels;

    // ── Pickit ──
    private final Preferences pickitPrefs = Preferences.userRoot().node("pickit");
    private File pickitImportDir;
    private final ObservableList<ProductoManual> pickitProductosList = FXCollections.observableArrayList();
    private ProductoManual pickitProductoEnEdicion = null;
    private AudioClip errorSound;
    private AudioClip successSound;

    @FXML
    public void initialize() {
        labelTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        labelTable.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        labelOrderCol.setCellValueFactory(new PropertyValueFactory<>("orderIds"));
        labelOrderCol.setCellFactory(col -> new TableCell<>() {
            private final Label prefixLabel = new Label();
            private final Label suffixLabel = new Label();
            private final HBox box = new HBox(0, prefixLabel, suffixLabel);
            {
                suffixLabel.setStyle("-fx-font-weight: bold;");
                box.setAlignment(Pos.CENTER);
                setContentDisplay(javafx.scene.control.ContentDisplay.GRAPHIC_ONLY);
            }
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null || item.isEmpty()) {
                    setGraphic(null);
                } else {
                    if (item.length() > 5) {
                        prefixLabel.setText(item.substring(0, item.length() - 5));
                        suffixLabel.setText(item.substring(item.length() - 5));
                    } else {
                        prefixLabel.setText("");
                        suffixLabel.setText(item);
                    }
                    setGraphic(box);
                }
            }
        });
        zoneCol.setCellValueFactory(new PropertyValueFactory<>("zone"));
        zoneCol.setCellFactory(col -> zoneCellWithUnknownHighlight());
        skuCol.setCellValueFactory(new PropertyValueFactory<>("sku"));
        skuCol.setCellFactory(col -> centeredCell());
        descCol.setCellValueFactory(new PropertyValueFactory<>("productDescription"));
        detailsCol.setCellValueFactory(new PropertyValueFactory<>("details"));
        countCol.setCellValueFactory(new PropertyValueFactory<>("quantity"));
        countCol.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(Integer item, boolean empty) {
                super.updateItem(item, empty);
                setAlignment(Pos.CENTER);
                setText(empty || item == null ? null : String.valueOf(item));
            }
        });

        orderTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        orderTable.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        orderTable.setEditable(true);
        orderSelectCol.setCellValueFactory(cd -> cd.getValue().selectedProperty());
        orderSelectCol.setCellFactory(col -> {
            CheckBoxTableCell<OrderTableRow, Boolean> cell = new CheckBoxTableCell<>(idx -> orderTable.getItems().get(idx).selectedProperty());
            cell.setAlignment(Pos.CENTER);
            return cell;
        });
        CheckBox selectAllCheck = new CheckBox();
        selectAllCheck.setSelected(true);
        selectAllCheck.setOnAction(e -> {
            boolean val = selectAllCheck.isSelected();
            // Afectar TODAS las filas (no solo las visibles del filtro del buscador),
            // para ser consistente con el contador y la generación.
            List<? extends OrderTableRow> todasLasFilas = filteredOrders != null
                    ? filteredOrders.getSource()
                    : orderTable.getItems();
            for (OrderTableRow row : todasLasFilas) {
                row.setSelected(val);
            }
        });
        orderSelectCol.setGraphic(selectAllCheck);
        orderIdCol.setCellValueFactory(new PropertyValueFactory<>("orderId"));
        orderZoneCol.setCellValueFactory(new PropertyValueFactory<>("zone"));
        orderSkuCol.setCellValueFactory(new PropertyValueFactory<>("sku"));
        orderDescCol.setCellValueFactory(new PropertyValueFactory<>("productDescription"));
        orderQtyCol.setCellValueFactory(new PropertyValueFactory<>("quantity"));
        orderStatusCol.setCellValueFactory(new PropertyValueFactory<>("status"));
        orderStatusCol.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(String status, boolean empty) {
                super.updateItem(status, empty);
                setAlignment(Pos.CENTER);
                if (empty || status == null || status.isEmpty()) {
                    setText(null);
                    setStyle("");
                } else {
                    String label = switch (status) {
                        case "ready_to_print" -> "\ud83d\udfe1 Pendiente";
                        case "printed" -> "\u2705 Impresa";
                        case "ready_for_dropoff", "ready_for_pickup" -> "\ud83d\udce6 Lista p/ despacho";
                        case "dropped_off" -> "\ud83d\udce5 Despachada";
                        case "picked_up", "in_hub", "in_transit" -> "\ud83d\ude9a En camino";
                        case "shipped" -> "\ud83d\ude9a Enviada";
                        case "delivered" -> "\u2714 Entregada";
                        default -> "\u2753 " + status;
                    };
                    setText(label);
                    String bg = switch (status) {
                        case "ready_to_print" -> "-fx-background-color: #C8E6C9;";
                        case "printed", "ready_for_dropoff", "ready_for_pickup" -> "-fx-background-color: #FFCDD2;";
                        default -> "";
                    };
                    setStyle(bg);
                }
            }
        });
        orderSlaCol.setCellValueFactory(new PropertyValueFactory<>("slaDate"));

        // Celdas multilínea para columnas que pueden tener varios productos
        orderIdCol.setCellFactory(col -> new TableCell<>() {
            private final Label prefixLabel = new Label();
            private final Label suffixLabel = new Label();
            private final HBox box = new HBox(0, prefixLabel, suffixLabel);
            {
                suffixLabel.setStyle("-fx-font-weight: bold;");
                box.setAlignment(Pos.CENTER);
                setContentDisplay(javafx.scene.control.ContentDisplay.GRAPHIC_ONLY);
            }
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setGraphic(null);
                } else {
                    if (item.length() > 5) {
                        prefixLabel.setText(item.substring(0, item.length() - 5));
                        suffixLabel.setText(item.substring(item.length() - 5));
                    } else {
                        prefixLabel.setText("");
                        suffixLabel.setText(item);
                    }
                    setGraphic(box);
                }
            }
        });
        orderZoneCol.setCellFactory(col -> zoneCellWithUnknownHighlight());
        orderSkuCol.setCellFactory(col -> centeredCell());
        orderQtyCol.setCellFactory(col -> centeredCell());
        orderSlaCol.setCellFactory(col -> centeredCell());

        // Centrar headers de ambas tablas
        centerColumnHeaders(orderTable);
        centerColumnHeaders(labelTable);

        // Bloquear reordenamiento de columnas
        lockColumns(orderTable);
        lockColumns(labelTable);

        // Placeholder dinámico según sub-tab seleccionado
        Label placeholderLocal = new Label("\uD83D\uDCE6 Cargue etiquetas ZPL para ver el resultado ordenado");
        placeholderLocal.setStyle("-fx-font-size: 14px; -fx-text-fill: #888;");
        Label placeholderApi = new Label("\uD83D\uDCCB Haga clic en 'Obtener Órdenes' para cargar las órdenes de ML");
        placeholderApi.setStyle("-fx-font-size: 14px; -fx-text-fill: #888;");
        etiquetasSubTabPane.getSelectionModel().selectedIndexProperty().addListener((obs, oldIdx, newIdx) -> {
            if (labelTable.getItems() == null || labelTable.getItems().isEmpty()) {
                labelTable.setPlaceholder(newIdx.intValue() == 0 ? placeholderApi : placeholderLocal);
            }
        });
        labelTable.setPlaceholder(placeholderApi);

        // Copiar al portapapeles con Ctrl+C (fila) y click derecho (celda)
        setupTableCopyHandler(orderTable);
        setupTableCopyHandler(labelTable);
        setupCellCopyMenu(orderTable);
        setupCellCopyMenu(labelTable);

        searchField.textProperty().addListener((obs, oldVal, newVal) -> {
            String filter = newVal == null ? "" : newVal.trim().toLowerCase();
            if (filteredOrders != null) {
                filteredOrders.setPredicate(row ->
                        filter.isEmpty() || row.getOrderId().toLowerCase().contains(filter)
                                || (row.getSku() != null && row.getSku().toLowerCase().contains(filter))
                                || (row.getZone() != null && row.getZone().toLowerCase().contains(filter))
                                || (row.getProductDescription() != null && row.getProductDescription().toLowerCase().contains(filter)));
            }
            if (filteredLabels != null) {
                filteredLabels.setPredicate(row ->
                        filter.isEmpty() || (row.getOrderIds() != null && row.getOrderIds().toLowerCase().contains(filter))
                                || (row.getSku() != null && row.getSku().toLowerCase().contains(filter))
                                || (row.getZone() != null && row.getZone().toLowerCase().contains(filter))
                                || (row.getProductDescription() != null && row.getProductDescription().toLowerCase().contains(filter)));
            }
        });

        estadoFilterCombo.setItems(FXCollections.observableArrayList("Todas", "Solo impresas", "Solo pendientes"));
        estadoFilterCombo.setValue("Solo pendientes");
        setupComboIcons(estadoFilterCombo, Map.of(
                "Todas", "\uD83D\uDCCB",
                "Solo impresas", "✅",
                "Solo pendientes", "\uD83D\uDD51"
        ));
        despachoFilterCombo.setItems(FXCollections.observableArrayList("Solo para hoy", "Hoy + futuro"));
        despachoFilterCombo.setValue("Solo para hoy");
        setupComboIcons(despachoFilterCombo, Map.of(
                "Solo para hoy", "\uD83D\uDCC5",
                "Hoy + futuro", "\uD83D\uDCC6"
        ));

        String savedExcelPath = prefs.get(PREF_EXCEL_PATH, "");
        if (!savedExcelPath.isBlank() && new File(savedExcelPath).exists()) {
            excelFileField.setText(savedExcelPath);
        }

        String savedComboPath = prefs.get(PREF_COMBO_EXCEL_PATH, "");
        if (!savedComboPath.isBlank() && new File(savedComboPath).exists()) {
            comboExcelField.setText(savedComboPath);
        }

        String savedMedidasPath = prefs.get(PREF_MEDIDAS_EXCEL_PATH, "");
        if (!savedMedidasPath.isBlank()) {
            medidasExcelField.setText(savedMedidasPath);
        }

        boolean medidasEnabled = prefs.getBoolean(PREF_MEDIDAS_ENABLED, false);
        medidasEnabledCheck.setSelected(medidasEnabled);
        medidasExcelField.setDisable(!medidasEnabled);
        medidasSelectBtn.setDisable(!medidasEnabled);
        actualizarBotonSubirMedidas();
        medidasEnabledCheck.selectedProperty().addListener((obs, oldVal, newVal) -> {
            boolean on = newVal != null && newVal;
            prefs.putBoolean(PREF_MEDIDAS_ENABLED, on);
            medidasExcelField.setDisable(!on);
            medidasSelectBtn.setDisable(!on);
            actualizarBotonSubirMedidas();
        });
        medidasExcelField.textProperty().addListener((obs, oldVal, newVal) -> actualizarBotonSubirMedidas());

        try {
            meliInitialized = MercadoLibreAPI.inicializar();
            if (meliInitialized) {
                meliStatusLabel.setText("\ud83d\udfe2 Estado: Conectado");
            }
        } catch (Exception e) {
            meliStatusLabel.setText("\u26aa Estado: No conectado");
        }

        // ── Pickit Tab init ──
        initPickitTab();

        // ── Pedidos Tab init ──
        initPedidosTab();
    }

    @FXML
    private void onSelectZplFile() {
        FileChooser fc = new FileChooser();
        fc.setTitle("Seleccionar archivo ZPL");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Archivos ZPL", "*.txt", "*.zpl"));
        String lastDir = prefs.get(PREF_ZPL_DIR, "");
        if (!lastDir.isBlank()) {
            File dir = new File(lastDir);
            if (dir.isDirectory()) {
                fc.setInitialDirectory(dir);
            }
        }
        File file = fc.showOpenDialog(getWindow());
        if (file != null) {
            zplFileField.setText(file.getAbsolutePath());
            prefs.put(PREF_ZPL_DIR, file.getParent());
        }
    }

    @FXML
    private void onSelectExcelFile() {
        FileChooser fc = new FileChooser();
        fc.setTitle("Seleccionar archivo Excel");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xls"));
        File file = fc.showOpenDialog(getWindow());
        if (file != null) {
            excelFileField.setText(file.getAbsolutePath());
            prefs.put(PREF_EXCEL_PATH, file.getAbsolutePath());
        }
    }

    @FXML
    private void onSelectComboExcelFile() {
        FileChooser fc = new FileChooser();
        fc.setTitle("Seleccionar Excel de composición de combos");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xls"));
        File file = fc.showOpenDialog(getWindow());
        if (file != null) {
            comboExcelField.setText(file.getAbsolutePath());
            prefs.put(PREF_COMBO_EXCEL_PATH, file.getAbsolutePath());
        }
    }

    @FXML
    private void onSelectMedidasExcelFile() {
        FileChooser fc = new FileChooser();
        fc.setTitle("Seleccionar Excel madre de medidas (ML)");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
        String current = medidasExcelField.getText();
        if (current != null && !current.isBlank()) {
            File f = new File(current);
            if (f.getParentFile() != null && f.getParentFile().isDirectory()) {
                fc.setInitialDirectory(f.getParentFile());
                fc.setInitialFileName(f.getName());
            }
        }
        File file = fc.showOpenDialog(getWindow());
        if (file != null) {
            String path = file.getAbsolutePath();
            medidasExcelField.setText(path);
            prefs.put(PREF_MEDIDAS_EXCEL_PATH, path);
        }
    }

    /**
     * Handler del botón "Subir Medidas". Valida precondiciones y dispara la subida.
     */
    @FXML
    private void onSubirMedidas() {
        if (medidasEnabledCheck == null || !medidasEnabledCheck.isSelected()) {
            AlertHelper.showError("Medidas ML", "Habilite primero el checkbox de 'Archivo Excel (medidas ML)'.");
            return;
        }
        String path = medidasExcelField == null ? null : medidasExcelField.getText();
        if (path == null || path.isBlank() || !new File(path).isFile()) {
            AlertHelper.showError("Medidas ML", "Seleccione un archivo Excel de medidas válido.");
            return;
        }
        if (!meliInitialized) {
            AlertHelper.showError("Medidas ML", "ML no está inicializado. Inicie sesión en MercadoLibre para subir medidas.");
            return;
        }
        if (subidaMedidasEnCurso) return;
        subirMedidasPendientesAsync(path);
    }

    /**
     * Sube en background las medidas pendientes (SUBIDO=NO con +20% completas) a ML.
     * Se invoca manualmente desde el botón "Subir Medidas". Actualiza el label de estado,
     * marca los resultados en el Excel y abre un diálogo al finalizar.
     */
    private void subirMedidasPendientesAsync(String path) {
        if (path == null || path.isBlank()) return;

        subidaMedidasEnCurso = true;
        Platform.runLater(this::actualizarBotonSubirMedidas);

        new Thread(() -> {
            try {
                Map<String, ar.com.leo.etiquetas.model.MedidaSku> medidas;
                try {
                    medidas = medidasManager.leerMedidas(Path.of(path));
                } catch (Exception e) {
                    setMedidasStatus("⚠ Error leyendo Excel", "-fx-text-fill: #b91c1c;", "No se pudo leer el Excel de medidas: " + e.getMessage());
                    Platform.runLater(() -> AlertHelper.showError("Medidas ML", "No se pudo leer el Excel de medidas: " + e.getMessage()));
                    return;
                }

                List<ar.com.leo.etiquetas.model.MedidaSku> aSubir = medidas.values().stream()
                        .filter(m -> !m.subido())
                        .filter(ar.com.leo.etiquetas.model.MedidaSku::tieneMedidasParaSubir)
                        .toList();
                if (aSubir.isEmpty()) {
                    setMedidasStatus("✓ Sin pendientes", "-fx-text-fill: #15803d;", "No hay SKUs con medidas completas pendientes de subir.");
                    Platform.runLater(() -> AlertHelper.showInfo("Medidas ML", "No hay SKUs con medidas completas pendientes de subir."));
                    return;
                }

                String userId;
                try {
                    userId = MercadoLibreAPI.getUserId();
                } catch (Exception e) {
                    setMedidasStatus("⚠ Error ML", "-fx-text-fill: #b91c1c;", "No se pudo obtener userId: " + e.getMessage());
                    Platform.runLater(() -> AlertHelper.showError("Medidas ML", "No se pudo obtener userId: " + e.getMessage()));
                    return;
                }

                final int total = aSubir.size();
                StringBuilder detalles = new StringBuilder();
                detalles.append("Subida de ").append(total).append(" SKU(s) pendientes:\n\n");
                List<String> ok = new ArrayList<>();
                Map<String, String> errores = new LinkedHashMap<>();
                int procesados = 0;
                setMedidasStatus("⏳ Subiendo 0/" + total, "-fx-text-fill: #1d4ed8;", detalles.toString());

                for (ar.com.leo.etiquetas.model.MedidaSku m : aSubir) {
                    try {
                        MercadoLibreAPI.UploadResult r = MercadoLibreAPI.actualizarDimensionesPaquete(
                                userId, m.sku(),
                                m.anchoMasCm(), m.altoMasCm(), m.profundidadMasCm(), m.pesoMasKg());
                        if (r.ok()) {
                            ok.add(m.sku());
                            detalles.append("✓ ").append(m.sku()).append(" (item ").append(r.itemId()).append(")\n");
                        } else {
                            errores.put(m.sku(), r.mensaje());
                            detalles.append("✗ ").append(m.sku()).append(" → ").append(r.mensaje()).append("\n");
                        }
                    } catch (Exception e) {
                        String msg = "Excepción: " + e.getMessage();
                        errores.put(m.sku(), msg);
                        detalles.append("✗ ").append(m.sku()).append(" → ").append(msg).append("\n");
                    }
                    procesados++;
                    final int p = procesados;
                    final int f = errores.size();
                    final int o = ok.size();
                    final String snap = detalles.toString();
                    Platform.runLater(() -> {
                        medidasUltimoDetalle = snap;
                        medidasStatusLabel.setText("⏳ Subiendo " + p + "/" + total
                                + "  (OK " + o + (f > 0 ? " · FAIL " + f : "") + ")");
                    });
                }

                try {
                    medidasManager.marcarResultados(Path.of(path), ok, errores);
                } catch (Exception e) {
                    detalles.append("\n⚠ No se pudo persistir resultados en el Excel: ").append(e.getMessage());
                }

                final int finalOk = ok.size();
                final int finalFail = errores.size();
                detalles.append("\nFinalizado. OK=").append(finalOk).append(" · FAIL=").append(finalFail);
                final String detalleFinal = detalles.toString();

                Platform.runLater(() -> {
                    medidasUltimoDetalle = detalleFinal;
                    medidasUltimoTuvoError = finalFail > 0;
                    if (finalFail == 0) {
                        medidasStatusLabel.setGraphic(null);
                        medidasStatusLabel.setText("✓ " + finalOk + " subido" + (finalOk == 1 ? "" : "s") + " a ML (click para detalles)");
                        medidasStatusLabel.setStyle("-fx-font-size: 12px; -fx-cursor: hand; -fx-text-fill: #15803d; -fx-font-weight: bold;");
                    } else {
                        medidasStatusLabel.setGraphic(crearIconoAdvertencia());
                        medidasStatusLabel.setText(finalOk + " OK · " + finalFail + " FAIL (click para detalles)");
                        medidasStatusLabel.setStyle("-fx-font-size: 12px; -fx-cursor: hand; -fx-text-fill: #b91c1c; -fx-font-weight: bold;");
                    }
                    // Auto-abrir diálogo al finalizar la subida (errores en rojo, éxito normal).
                    if (finalFail > 0) {
                        AlertHelper.showErrorScrollable("Subida de medidas a ML", detalleFinal);
                    } else {
                        AlertHelper.showInfoScrollable("Subida de medidas a ML", detalleFinal);
                    }
                });
            } finally {
                subidaMedidasEnCurso = false;
                Platform.runLater(this::actualizarBotonSubirMedidas);
            }
        }, "subir-medidas-ml").start();
    }

    /**
     * Habilita el botón "Subir Medidas" solo si: el checkbox está activo, hay un path configurado,
     * el archivo existe, y no se está subiendo actualmente.
     */
    private void actualizarBotonSubirMedidas() {
        if (subirMedidasBtn == null) return;
        if (subidaMedidasEnCurso) {
            subirMedidasBtn.setDisable(true);
            return;
        }
        boolean habilitable = medidasEnabledCheck != null && medidasEnabledCheck.isSelected()
                && medidasExcelField != null
                && medidasExcelField.getText() != null && !medidasExcelField.getText().isBlank()
                && new File(medidasExcelField.getText()).isFile();
        subirMedidasBtn.setDisable(!habilitable);
    }

    private javafx.scene.image.ImageView crearIconoAdvertencia() {
        javafx.scene.image.Image img = new javafx.scene.image.Image(
                getClass().getResourceAsStream("/ar/com/leo/ui/icons8-señal-de-advertencia-general-100.png"),
                20, 20, true, true);
        return new javafx.scene.image.ImageView(img);
    }

    private void setMedidasStatus(String text, String colorStyle, String detalle) {
        // true si el color indica error (rojo/amarillo) — simple heurística para decidir el ícono.
        boolean esWarn = colorStyle != null && (colorStyle.contains("#b91c1c") || colorStyle.contains("#b45309"));
        Platform.runLater(() -> {
            medidasUltimoDetalle = detalle;
            medidasUltimoTuvoError = esWarn;
            medidasStatusLabel.setGraphic(esWarn ? crearIconoAdvertencia() : null);
            medidasStatusLabel.setText(text);
            medidasStatusLabel.setStyle("-fx-font-size: 12px; -fx-cursor: hand; " + colorStyle + " -fx-font-weight: bold;");
        });
    }

    @FXML
    private void onMedidasStatusClicked() {
        if (medidasUltimoDetalle == null || medidasUltimoDetalle.isBlank()) return;
        if (medidasUltimoTuvoError) {
            AlertHelper.showErrorScrollable("Subida de medidas a ML", medidasUltimoDetalle);
        } else {
            AlertHelper.showInfoScrollable("Subida de medidas a ML", medidasUltimoDetalle);
        }
    }

    private void validateExcelFiles() {
        String excelPath = excelFileField.getText();
        if (excelPath == null || excelPath.isBlank()) {
            throw new IllegalArgumentException("Seleccione el archivo Excel de stock (SKU \u2192 Zona).");
        }
        String comboPath = comboExcelField.getText();
        if (comboPath == null || comboPath.isBlank()) {
            throw new IllegalArgumentException("Seleccione el archivo Excel de composición de combos.");
        }
        if (medidasEnabledCheck.isSelected()) {
            String medidasPath = medidasExcelField.getText();
            if (medidasPath == null || medidasPath.isBlank()) {
                throw new IllegalArgumentException(
                        "Habilitó 'medidas ML' pero no seleccionó archivo. Elija uno (si no existe se crea automáticamente) o desmarque el checkbox.");
            }
        }
    }

    private ExcelMapping loadExcelMapping() throws Exception {
        validateExcelFiles();
        return excelReader.readMapping(Path.of(excelFileField.getText()));
    }

    @FXML
    private void onProcessLocal() {
        String zplPath = zplFileField.getText();
        if (zplPath == null || zplPath.isBlank()) {
            AlertHelper.showError("Error", "Seleccione un archivo ZPL.");
            return;
        }

        try {
            ExcelMapping excelMapping = loadExcelMapping();
            List<ZplLabel> labels = zplParser.parseFile(Path.of(zplPath));
            Map<String, ar.com.leo.etiquetas.model.MedidaSku> medidas = loadMedidasMap();
            Map<String, String> skusPendientes = new LinkedHashMap<>();
            currentResult = injectZplHeaders(
                    labelSorter.sort(labels, excelMapping.skuToZone()), excelMapping, medidas, skusPendientes);
            int agregadosExcel = guardarSkusPendientesMedicion(skusPendientes);
            showLabelTable();
            displayResult(currentResult);
            mostrarMensajeSkusFaltantes(skusPendientes.size(), agregadosExcel, new ArrayList<>(skusPendientes.keySet()));
        } catch (Exception e) {
            AlertHelper.showError("Error al procesar", e.getMessage(), e);
        }
    }

    @FXML
    private void onFetchMeliOrders() {
        if (!meliInitialized) {
            AlertHelper.showError("Error", "Primero inicie sesi\u00f3n en MercadoLibre.");
            return;
        }

        ExcelMapping excelMapping;
        try {
            excelMapping = loadExcelMapping();
        } catch (Exception e) {
            AlertHelper.showError("Error", e.getMessage(), e);
            return;
        }

        String estadoFiltro = estadoFilterCombo.getValue();
        String despachoFiltro = despachoFilterCombo.getValue();
        boolean incluirImpresas = !"Solo pendientes".equals(estadoFiltro);
        boolean soloPendientes = "Solo pendientes".equals(estadoFiltro);
        boolean soloImpresas = "Solo impresas".equals(estadoFiltro);
        boolean soloSlaHoy = "Solo para hoy".equals(despachoFiltro);

        setLoading(true);

        new Thread(() -> {
            try {
                String userId = MercadoLibreAPI.getUserId();
                MercadoLibreAPI.MLOrderResult result = MercadoLibreAPI.obtenerVentasReadyToPrint(userId, incluirImpresas);
                List<OrdenML> ordenes = result.ordenes();

                // Obtener SLAs en paralelo
                List<Long> shipmentIds = new ArrayList<>();
                for (OrdenML orden : ordenes) {
                    Long shipId = orden.getShipmentId();
                    if (shipId != null && shipId > 0) {
                        shipmentIds.add(shipId);
                    }
                }

                // Substatus ya viene del search (asignado en searchAndCollect)
                Map<Long, MercadoLibreAPI.SlaInfo> slaMap = new HashMap<>();
                Map<Long, String> substatusMap = new HashMap<>();
                for (OrdenML orden : ordenes) {
                    Long shipId = orden.getShipmentId();
                    if (shipId != null && shipId > 0) {
                        substatusMap.put(shipId, orden.getShippingSubstatus());
                    }
                }
                // Solo consultar SLAs (fecha de despacho) en paralelo
                if (!shipmentIds.isEmpty()) {
                    slaMap = MercadoLibreAPI.obtenerSlasParalelo(shipmentIds);
                }

                if (soloSlaHoy && !shipmentIds.isEmpty()) {
                    OffsetDateTime hoyFin = java.time.LocalDate.now()
                            .atTime(23, 59, 59).atZone(java.time.ZoneId.systemDefault()).toOffsetDateTime();

                    List<OrdenML> filtradas = new ArrayList<>();
                    for (OrdenML orden : ordenes) {
                        Long shipId = orden.getShipmentId();
                        if (shipId == null || shipId <= 0) {
                            filtradas.add(orden);
                            continue;
                        }
                        MercadoLibreAPI.SlaInfo sla = slaMap.get(shipId);
                        if (sla == null || sla.expectedDate() == null) {
                            filtradas.add(orden);
                            continue;
                        }
                        OffsetDateTime expected = sla.expectedDate();
                        if (expected.isBefore(hoyFin) || expected.isEqual(hoyFin)) {
                            filtradas.add(orden); // SLA hoy o antes
                        }
                    }
                    ordenes = filtradas;
                }

                // Filtro por estado (solo impresas / solo pendientes)
                // "Solo pendientes" = substatus ready_to_print
                // "Solo impresas" = cualquier substatus que NO sea ready_to_print (printed, ready_for_dropoff, etc.)
                if (soloImpresas || soloPendientes) {
                    List<OrdenML> filtradasEstado = new ArrayList<>();
                    for (OrdenML orden : ordenes) {
                        Long shipId = orden.getShipmentId();
                        String substatus = shipId != null ? substatusMap.getOrDefault(shipId, "") : "";
                        boolean esPendiente = "ready_to_print".equals(substatus);
                        if (soloImpresas && !esPendiente) {
                            filtradasEstado.add(orden);
                        } else if (soloPendientes && esPendiente) {
                            filtradasEstado.add(orden);
                        }
                    }
                    ordenes = filtradasEstado;
                }

                // Extraer shipment IDs turbo
                Set<Long> turboIds = new HashSet<>();
                for (var slaEntry : slaMap.entrySet()) {
                    if (slaEntry.getValue().turbo()) {
                        turboIds.add(slaEntry.getKey());
                    }
                }
                final List<OrdenML> finalOrdenes = ordenes;
                final Map<Long, MercadoLibreAPI.SlaInfo> finalSlaMap = slaMap;
                final Map<Long, String> finalSubstatusMap = substatusMap;
                final Set<Long> finalTurboIds = turboIds;

                Platform.runLater(() -> {
                    setLoading(false);
                    fetchedOrders = finalOrdenes;
                    turboShipmentIds = finalTurboIds;
                    displayOrders(finalOrdenes, excelMapping.skuToZone(), finalSlaMap, finalSubstatusMap);
                    showOrderTable();
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    setLoading(false);
                    AlertHelper.showError("Error API ML", e.getMessage(), e);
                });
            }
        }).start();
    }

    @FXML
    private void onDownloadSelectedLabels() {
        if (fetchedOrders == null || fetchedOrders.isEmpty()) {
            AlertHelper.showError("Error", "No hay \u00f3rdenes cargadas.");
            return;
        }

        ExcelMapping excelMapping;
        try {
            excelMapping = loadExcelMapping();
        } catch (Exception e) {
            AlertHelper.showError("Error", e.getMessage(), e);
            return;
        }

        // Filtrar solo las órdenes seleccionadas (deduplicar por orderId).
        // Importante: iterar la lista completa (no orderTable.getItems(), que solo trae
        // las filas visibles según el buscador) para no perder selecciones ocultas por el filtro.
        List<? extends OrderTableRow> todasLasFilas = filteredOrders != null
                ? filteredOrders.getSource()
                : orderTable.getItems();
        LinkedHashSet<Long> seenOrderIds = new LinkedHashSet<>();
        List<OrdenML> seleccionadas = new ArrayList<>();
        for (OrderTableRow row : todasLasFilas) {
            if (row.isSelected()) {
                for (OrdenML o : row.getOrdenes()) {
                    if (seenOrderIds.add(o.getOrderId())) {
                        seleccionadas.add(o);
                    }
                }
            }
        }

        if (seleccionadas.isEmpty()) {
            AlertHelper.showError("Error", "No hay \u00f3rdenes seleccionadas.");
            return;
        }

        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("Confirmar descarga");
        long totalEtiquetas = seleccionadas.stream()
                .map(OrdenML::getShipmentId)
                .filter(id -> id != null && id > 0)
                .distinct()
                .count();
        confirm.setHeaderText("Se descargarán " + totalEtiquetas + " etiqueta(s)");
        boolean hayPendientes = seleccionadas.stream()
                .anyMatch(o -> "ready_to_print".equals(o.getShippingSubstatus()));
        StringBuilder advertencia = new StringBuilder();
        if (hayPendientes) {
            advertencia.append("Al descargar, el estado de las órdenes pendientes pasará a \"Impresa\" en MercadoLibre.\n\n");
        }
        advertencia.append("¿Desea continuar?");
        confirm.setContentText(advertencia.toString());
        confirm.setGraphic(new javafx.scene.image.ImageView(
                new javafx.scene.image.Image(getClass().getResourceAsStream("/ar/com/leo/ui/icons8-señal-de-advertencia-general-100.png"), 48, 48, true, true)));
        ((javafx.stage.Stage) confirm.getDialogPane().getScene().getWindow()).getIcons().add(
                new javafx.scene.image.Image(getClass().getResourceAsStream("/ar/com/leo/ui/icons8-etiqueta-100.png")));
        Optional<ButtonType> confirmResult = confirm.showAndWait();
        if (confirmResult.isEmpty() || confirmResult.get() != ButtonType.OK) return;

        // Capturar referencias/paths en el hilo UI antes de lanzar el thread background.
        final Map<String, ar.com.leo.etiquetas.model.MedidaSku> medidas = loadMedidasMap();
        final String medidasPath = medidasExcelField != null ? medidasExcelField.getText() : null;

        setLoading(true);

        new Thread(() -> {
            try {
                List<ZplLabel> labels = MercadoLibreAPI.descargarEtiquetasZplParaOrdenes(seleccionadas, turboShipmentIds);
                Map<String, String> skusPendientes = new LinkedHashMap<>();
                SortResult result = injectZplHeaders(
                        labelSorter.sort(labels, excelMapping.skuToZone()), excelMapping, medidas, skusPendientes);
                int agregadosExcel = guardarSkusPendientesMedicion(skusPendientes, medidasPath);

                // Guardar automáticamente en carpeta "Etiquetas"
                String saveError = null;
                File savedFile = null;
                try {
                    Path etiquetasDir = Path.of("Etiquetas");
                    Files.createDirectories(etiquetasDir);
                    String fechaHora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));
                    Path outputFile = etiquetasDir.resolve("etiquetas_ordenadas_" + fechaHora + ".txt");
                    fileSaver.save(interleaveForPrint(result.sortedFlatList()), outputFile);
                    savedFile = outputFile.toFile();
                } catch (Exception ex) {
                    AppLogger.error("Error al guardar automáticamente", ex);
                    saveError = ex.getMessage();
                }

                final String finalSaveError = saveError;
                final File finalSavedFile = savedFile;
                final int skusFaltantesCount = skusPendientes.size();
                final int agregadosCount = agregadosExcel;
                final List<String> skusFaltantesList = new ArrayList<>(skusPendientes.keySet());
                Platform.runLater(() -> {
                    setLoading(false);
                    currentResult = result;
                    showLabelTable();
                    displayResult(result);
                    if (finalSavedFile != null) {
                        fileLinkBar.getChildren().clear();
                        LogHelper.addFileLink(fileLinkBar, finalSavedFile);
                        fileLinkBar.setVisible(true);
                        fileLinkBar.setManaged(true);
                    }
                    if (finalSaveError != null) {
                        AlertHelper.showError("Error al guardar", "No se pudo guardar el archivo automáticamente:\n" + finalSaveError);
                    }
                    mostrarMensajeSkusFaltantes(skusFaltantesCount, agregadosCount, skusFaltantesList);
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    setLoading(false);
                    AlertHelper.showError("Error API ML", e.getMessage(), e);
                });
            }
        }).start();
    }

    /**
     * Muestra un diálogo informativo con la cantidad de SKU detectados sin medidas completas
     * en el lote de etiquetas recién procesado. Solo se muestra si hay al menos uno.
     */
    private void mostrarMensajeSkusFaltantes(int faltantesCount, int agregadosExcel, List<String> skus) {
        if (faltantesCount <= 0) return;
        int yaExistentes = Math.max(0, faltantesCount - agregadosExcel);
        StringBuilder msg = new StringBuilder();
        msg.append(faltantesCount).append(" SKU(s) sin medidas detectados en este lote.\n\n");
        if (agregadosExcel > 0) {
            msg.append("• ").append(agregadosExcel).append(" nuevo(s) agregado(s) al Excel de medidas.\n");
        }
        if (yaExistentes > 0) {
            msg.append("• ").append(yaExistentes).append(" ya figuraba(n) en el Excel.\n");
        }
        if (skus != null && !skus.isEmpty()) {
            msg.append("\nSKUs:\n");
            for (String sku : skus) {
                msg.append("  ").append(sku).append("\n");
            }
        }
        AlertHelper.showInfoScrollable("Medidas pendientes", msg.toString());
    }


    @FXML
    private void onPrintDirect() {
        if (currentResult == null || currentResult.groups().isEmpty()) {
            AlertHelper.showError("Error", "No hay etiquetas para imprimir.");
            return;
        }

        // 1. Seleccionar zonas a imprimir
        Map<String, Long> zoneCounts = new LinkedHashMap<>();
        for (SortedLabelGroup group : currentResult.groups()) {
            zoneCounts.merge(group.zone(), (long) group.labels().size(), Long::sum);
        }

        Dialog<List<String>> zoneDialog = new Dialog<>();
        zoneDialog.setTitle("Seleccionar zonas");
        zoneDialog.setHeaderText("Seleccione las zonas a imprimir:");
        zoneDialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);

        VBox zoneBox = new VBox(8);
        zoneBox.setStyle("-fx-padding: 10;");
        List<CheckBox> checkBoxes = new ArrayList<>();
        for (var entry : zoneCounts.entrySet()) {
            CheckBox cb = new CheckBox(entry.getKey() + "  (" + entry.getValue() + " etiquetas)");
            cb.setSelected(true);
            cb.setUserData(entry.getKey());
            checkBoxes.add(cb);
            zoneBox.getChildren().add(cb);
        }

        Button toggleBtn = new Button("Deseleccionar todas");
        toggleBtn.setOnAction(e -> {
            boolean allSelected = checkBoxes.stream().allMatch(CheckBox::isSelected);
            checkBoxes.forEach(cb -> cb.setSelected(!allSelected));
            toggleBtn.setText(allSelected ? "Seleccionar todas" : "Deseleccionar todas");
        });
        zoneBox.getChildren().add(toggleBtn);

        zoneDialog.getDialogPane().setContent(zoneBox);
        zoneDialog.setResultConverter(btn -> {
            if (btn == ButtonType.OK) {
                return checkBoxes.stream()
                        .filter(CheckBox::isSelected)
                        .map(cb -> (String) cb.getUserData())
                        .toList();
            }
            return null;
        });

        Optional<List<String>> zonesResult = zoneDialog.showAndWait();
        if (zonesResult.isEmpty() || zonesResult.get().isEmpty()) return;
        Set<String> selectedZones = new LinkedHashSet<>(zonesResult.get());

        // Filtrar etiquetas por zonas seleccionadas
        List<ZplLabel> labelsToPrint = currentResult.groups().stream()
                .filter(g -> selectedZones.contains(g.zone()))
                .flatMap(g -> g.labels().stream())
                .toList();

        if (labelsToPrint.isEmpty()) {
            AlertHelper.showError("Error", "No hay etiquetas en las zonas seleccionadas.");
            return;
        }

        // 2. Seleccionar impresora
        List<PrintService> printers = printerDiscovery.findAll();
        if (printers.isEmpty()) {
            AlertHelper.showError("Error", "No se encontraron impresoras.");
            return;
        }

        ChoiceDialog<String> dialog = new ChoiceDialog<>(
                printers.getFirst().getName(),
                printers.stream().map(PrintService::getName).toList());
        dialog.setTitle("Seleccionar impresora");
        dialog.setHeaderText("Seleccione la impresora para enviar " + labelsToPrint.size() + " etiqueta(s):");

        Optional<String> selected = dialog.showAndWait();
        if (selected.isEmpty()) return;

        PrintService selectedPrinter = printers.stream()
                .filter(p -> p.getName().equals(selected.get()))
                .findFirst()
                .orElse(null);

        if (selectedPrinter == null) return;

        try {
            List<ZplLabel> reordered = interleaveForPrint(labelsToPrint);
            printerService.printViaPrintService(reordered, selectedPrinter);
            AlertHelper.showInfo("\ud83d\udda8 Impresi\u00f3n", labelsToPrint.size() + " etiquetas enviadas a " + selectedPrinter.getName());
            showComboSheetIfNeeded();
        } catch (Exception e) {
            AlertHelper.showError("Error al imprimir", e.getMessage(), e);
        }
    }

    @FXML
    private void onShowComboSheet() {
        showComboSheetIfNeeded();
    }

    private List<ComboProduct> findMatchingCombos() {
        String comboPath = comboExcelField.getText();
        if (comboPath == null || comboPath.isBlank()) return List.of();
        if (currentResult == null || currentResult.groups().isEmpty()) return List.of();

        try {
            Map<String, ComboProduct> allCombos = comboExcelReader.read(Path.of(comboPath));
            if (allCombos.isEmpty()) return List.of();

            // Recolectar SKUs del lote actual (separar multi-SKU de CARROS)
            Set<String> batchSkus = new HashSet<>();
            for (SortedLabelGroup group : currentResult.groups()) {
                for (String sku : group.sku().split("\n")) {
                    String trimmed = sku.trim();
                    if (!trimmed.isEmpty()) batchSkus.add(trimmed);
                }
            }

            // Crear mapa normalizado de combos para matchear por SKU numérico
            Map<String, ComboProduct> normalizedCombos = new LinkedHashMap<>();
            for (var entry : allCombos.entrySet()) {
                normalizedCombos.put(entry.getKey(), entry.getValue());
                // También indexar por SKU normalizado (solo dígitos)
                String normalized = ZplParser.normalizeSku(entry.getKey());
                if (normalized != null && !normalized.startsWith("SKU INVALIDO")) {
                    normalizedCombos.putIfAbsent(normalized, entry.getValue());
                }
            }

            List<ComboProduct> matchingCombos = new ArrayList<>();
            Set<String> matched = new HashSet<>();
            for (var entry : normalizedCombos.entrySet()) {
                if (batchSkus.contains(entry.getKey()) && matched.add(entry.getValue().codigoCompuesto())) {
                    matchingCombos.add(entry.getValue());
                }
            }

            matchingCombos.sort(Comparator.comparing(ComboProduct::codigoCompuesto));
            return matchingCombos;
        } catch (Exception e) {
            AppLogger.warn("Error al leer Excel de combos: " + e.getMessage());
            return List.of();
        }
    }

    private void showComboSheetIfNeeded() {
        List<ComboProduct> combos = findMatchingCombos();
        if (combos.isEmpty()) {
            AlertHelper.showInfo("Combos", "No se encontraron combos para las etiquetas actuales.");
            return;
        }
        new ComboPrintDialog(getWindow(), combos).show();
    }

    private static <T> TableCell<T, String> centeredCell() {
        return new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setAlignment(Pos.CENTER);
                setText(empty || item == null ? null : item);
            }
        };
    }

    private static <T> TableCell<T, String> zoneCellWithUnknownHighlight() {
        return new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setAlignment(Pos.CENTER);
                if (empty || item == null) {
                    setText(null);
                    setStyle("");
                } else {
                    setText(item);
                    if ("???".equals(item)) {
                        setStyle("-fx-background-color: #FFCDD2; -fx-text-fill: #B71C1C; -fx-font-weight: bold;");
                    } else {
                        setStyle("");
                    }
                }
            }
        };
    }

    private <T> void lockColumns(TableView<T> table) {
        for (TableColumn<T, ?> col : table.getColumns()) {
            col.setReorderable(false);
        }
    }

    private <T> void centerColumnHeaders(TableView<T> table) {
        for (TableColumn<T, ?> col : table.getColumns()) {
            if (col.getGraphic() != null) continue; // ya tiene graphic (ej. checkbox)
            Label headerLabel = new Label(col.getText());
            headerLabel.setStyle("-fx-font-weight: bold;");
            headerLabel.setMaxWidth(Double.MAX_VALUE);
            headerLabel.setAlignment(Pos.CENTER);
            col.setGraphic(headerLabel);
            col.setText("");
        }
    }

    @SuppressWarnings("unchecked")
    private <T> void setupCellCopyMenu(TableView<T> table) {
        MenuItem copiarCelda = new MenuItem("Copiar celda");
        copiarCelda.setOnAction(e -> {
            var pos = table.getFocusModel().getFocusedCell();
            if (pos != null && pos.getRow() >= 0 && pos.getTableColumn() != null) {
                TableColumn<T, ?> col = (TableColumn<T, ?>) pos.getTableColumn();
                Object val = col.getCellObservableValue(pos.getRow()).getValue();
                if (val != null) {
                    ClipboardContent content = new ClipboardContent();
                    content.putString(val.toString());
                    Clipboard.getSystemClipboard().setContent(content);
                }
            }
        });
        table.setContextMenu(new ContextMenu(copiarCelda));
    }

    private <T> void setupTableCopyHandler(TableView<T> table) {
        KeyCodeCombination ctrlC = new KeyCodeCombination(KeyCode.C, KeyCombination.CONTROL_DOWN);
        table.setOnKeyPressed(event -> {
            if (ctrlC.match(event)) {
                StringBuilder sb = new StringBuilder();
                for (T item : table.getSelectionModel().getSelectedItems()) {
                    if (item == null) continue;
                    StringJoiner line = new StringJoiner("\t");
                    for (TableColumn<T, ?> col : table.getColumns()) {
                        Object val = col.getCellObservableValue(item).getValue();
                        line.add(val != null ? val.toString().replace("\n", " / ") : "");
                    }
                    sb.append(line).append("\n");
                }
                if (!sb.isEmpty()) {
                    ClipboardContent content = new ClipboardContent();
                    content.putString(sb.toString());
                    Clipboard.getSystemClipboard().setContent(content);
                }
            }
        });
    }

    private void setLoading(boolean loading) {
        progressIndicator.setVisible(loading);
        progressIndicator.setManaged(loading);
        excelSelectorsBox.setDisable(loading);
        tabPane.setDisable(loading);
        orderTable.setDisable(loading);
        labelTable.setDisable(loading);
        statsBar.setDisable(loading);
        fileLinkBar.setDisable(loading);
        if (loading) {
            fileLinkBar.setVisible(false);
            fileLinkBar.setManaged(false);
        }
        searchBar.setDisable(loading);
        downloadLabelsBtn.setDisable(loading);
        comboSheetBtn.setDisable(loading);
        printDirectBtn.setDisable(loading);
        backToOrdersBtn.setDisable(loading);
        searchField.setDisable(loading);
        estadoFilterCombo.setDisable(loading);
        despachoFilterCombo.setDisable(loading);
        fetchOrdersBtn.setDisable(loading);
    }

    private void showOrderTable() {
        orderTable.setVisible(true);
        orderTable.setManaged(true);
        labelTable.setVisible(false);
        labelTable.setManaged(false);
        searchField.clear();
        boolean hayOrdenes = !orderTable.getItems().isEmpty();
        downloadLabelsBtn.setDisable(!hayOrdenes);
        comboSheetBtn.setDisable(true);
        printDirectBtn.setDisable(true);
        backToOrdersBtn.setVisible(false);
        backToOrdersBtn.setManaged(false);
    }

    private void showLabelTable() {
        labelTable.setVisible(true);
        labelTable.setManaged(true);
        orderTable.setVisible(false);
        orderTable.setManaged(false);
        searchField.clear();
        downloadLabelsBtn.setDisable(true);
        boolean hayEtiquetas = currentResult != null && !currentResult.groups().isEmpty();
        comboSheetBtn.setDisable(!hayEtiquetas);
        printDirectBtn.setDisable(!hayEtiquetas);
        // Mostrar botón volver solo si hay órdenes cargadas
        boolean hayOrdenes = !orderTable.getItems().isEmpty();
        backToOrdersBtn.setVisible(hayOrdenes);
        backToOrdersBtn.setManaged(hayOrdenes);
    }

    @FXML
    private void onBackToOrders() {
        showOrderTable();
    }

    private void displayOrders(List<OrdenML> ordenes, Map<String, String> skuToZone,
                               Map<Long, MercadoLibreAPI.SlaInfo> slaMap,
                               Map<Long, String> substatusMap) {
        DateTimeFormatter slaFormatter = DateTimeFormatter.ofPattern("dd/MM HH:mm");
        ObservableList<OrderTableRow> rows = FXCollections.observableArrayList();

        // Agrupar órdenes por pack_id (o por order_id si no tiene pack)
        Map<String, List<OrdenML>> grouped = new LinkedHashMap<>();
        for (OrdenML orden : ordenes) {
            String groupKey = orden.getPackId() != null
                    ? "P" + orden.getPackId()
                    : "O" + orden.getOrderId();
            grouped.computeIfAbsent(groupKey, k -> new ArrayList<>()).add(orden);
        }

        int totalOrdenes = grouped.size();

        for (var entry : grouped.entrySet()) {
            List<OrdenML> group = entry.getValue();
            OrdenML firstOrden = group.getFirst();

            String orderIdStr = firstOrden.getPackId() != null
                    ? String.valueOf(firstOrden.getPackId())
                    : String.valueOf(firstOrden.getOrderId());

            // SLA y status del primer envío del grupo
            String slaDate = "";
            String status = "";
            for (OrdenML o : group) {
                Long shipId = o.getShipmentId();
                if (shipId != null) {
                    if (slaDate.isEmpty() && slaMap.containsKey(shipId)) {
                        MercadoLibreAPI.SlaInfo sla = slaMap.get(shipId);
                        if (sla.expectedDate() != null) {
                            slaDate = sla.expectedDate().format(slaFormatter);
                        }
                    }
                    if (status.isEmpty() && substatusMap.containsKey(shipId)) {
                        status = substatusMap.get(shipId);
                    }
                }
            }

            // Recolectar todos los productos de todas las órdenes del grupo
            StringJoiner skuJoiner = new StringJoiner("\n");
            StringJoiner descJoiner = new StringJoiner("\n");
            StringJoiner qtyJoiner = new StringJoiner("\n");

            for (OrdenML o : group) {
                for (Venta v : o.getItems()) {
                    String itemSku = v.getSku() != null ? v.getSku() : "?";
                    String desc = v.getTitulo() != null && !v.getTitulo().isEmpty() ? v.getTitulo() : itemSku;
                    String qty = String.valueOf((int) v.getCantidad());
                    skuJoiner.add(itemSku);
                    descJoiner.add(desc);
                    qtyJoiner.add(qty);
                }
            }

            // Detectar si el envío es turbo
            boolean esTurbo = false;
            for (OrdenML o : group) {
                Long shipId = o.getShipmentId();
                if (shipId != null && slaMap.containsKey(shipId) && slaMap.get(shipId).turbo()) {
                    esTurbo = true;
                    break;
                }
            }

            // Determinar zona: TURBOS si es turbo, CARROS si hay 2+ SKUs distintos
            Set<String> distinctSkus = new HashSet<>();
            for (OrdenML o : group) {
                for (Venta v : o.getItems()) {
                    String s = v.getSku() != null ? v.getSku() : "";
                    if (!s.isEmpty()) distinctSkus.add(s);
                }
            }
            String zone;
            if (esTurbo) {
                zone = "TURBOS";
            } else if (distinctSkus.size() > 1) {
                zone = "CARROS";
            } else {
                Venta firstItem = firstOrden.getItems().getFirst();
                String firstSku = firstItem.getSku() != null ? firstItem.getSku() : "";
                zone = !firstSku.isEmpty() ? skuToZone.getOrDefault(firstSku, "???") : "???";
            }

            rows.add(new OrderTableRow(true, orderIdStr, zone, skuJoiner.toString(),
                    descJoiner.toString(), qtyJoiner.toString(), status, slaDate, group));
        }

        // Prioridad: J*, T*, COMBOS, CARROS, TURBOS, RETIROS, resto
        rows.sort(Comparator
                .<OrderTableRow, Integer>comparing(r -> {
                    String z = r.getZone().toUpperCase();
                    if (z.startsWith("J")) return 0;
                    if (z.startsWith("TURBOS")) return 4;
                    if (z.startsWith("T")) return 1;
                    if (z.startsWith("COMBOS")) return 2;
                    if (z.startsWith("CARROS")) return 3;
                    if (z.startsWith("RETIROS")) return 5;
                    return Integer.MAX_VALUE;
                })
                .thenComparing(r -> r.getZone().toUpperCase())
                .thenComparing(OrderTableRow::getSku));

        filteredOrders = new FilteredList<>(rows, p -> true);
        SortedList<OrderTableRow> sortedOrders = new SortedList<>(filteredOrders);
        sortedOrders.comparatorProperty().bind(orderTable.comparatorProperty());
        orderTable.setItems(sortedOrders);
        searchField.clear();

        if (rows.isEmpty()) {
            orderTable.setPlaceholder(new Label("No se encontraron ordenes con los filtros seleccionados"));
        }

        // Estadísticas
        int printedCount = 0;
        int readyCount = 0;
        int totalProductos = 0;
        Map<String, Integer> countByZone = new LinkedHashMap<>();
        Set<String> uniqueSkus = new HashSet<>();
        for (OrderTableRow r : rows) {
            if ("printed".equals(r.getStatus())) printedCount++;
            else readyCount++;
            totalProductos += r.getProductCount();
            countByZone.merge(r.getZone(), 1, Integer::sum);
            for (String s : r.getSku().split("\n")) {
                String trimmed = s.trim();
                if (!trimmed.isEmpty()) uniqueSkus.add(trimmed);
            }
        }
        final int printed = printedCount;
        final int readyToPrint = readyCount;
        final int ordCount = totalOrdenes;
        final int prodCount = totalProductos;
        final int skuCount = uniqueSkus.size();

        Runnable updateStats = () -> {
            long selected = rows.stream().filter(OrderTableRow::isSelected).count();
            StringJoiner sj = new StringJoiner("  \u2502  ");
            sj.add("Ordenes: " + ordCount);
            sj.add("Productos: " + prodCount);
            sj.add("SKUs: " + skuCount);
            sj.add("Seleccionados: " + selected);
            if (readyToPrint > 0) sj.add("Pendientes: " + readyToPrint);
            if (printed > 0) sj.add("Impresas: " + printed);
            for (Map.Entry<String, Integer> entry : countByZone.entrySet()) {
                if (entry.getValue() > 0) sj.add(entry.getKey() + ": " + entry.getValue());
            }
            String text = sj.toString();
            statsLabel.setText(text);
            Tooltip tip = new Tooltip(text.replace(" | ", "\n"));
            tip.setShowDelay(Duration.millis(200));
            statsLabel.setTooltip(tip);
        };

        if (rows.isEmpty()) {
            statsLabel.setText("No hay ordenes para mostrar");
            statsLabel.setTooltip(null);
        } else {
            updateStats.run();
        }
        statsBar.setVisible(true);
        statsBar.setManaged(true);
        searchBar.setVisible(true);
        searchBar.setManaged(true);

        // Actualizar stats cuando cambia la selección
        for (OrderTableRow r : rows) {
            r.selectedProperty().addListener((obs, oldVal, newVal) -> updateStats.run());
        }
    }

    private static final Pattern UNIT_PATTERN = Pattern.compile(
            "(\\^FO(\\d+),(\\d+)\\^A0N,70,70\\^FB160,1,0,C\\^FD)(\\d+)(\\^FS)");
    private static final Pattern FO_PATTERN = Pattern.compile("\\^FO(\\d+),(\\d+)");
    private static final Pattern LH_PATTERN = Pattern.compile("\\^LH(\\d+),(\\d+)");
    private static final Pattern FONT_PATTERN = Pattern.compile("\\^A0N,(\\d+),(\\d+)");
    private static final Pattern FB_PATTERN = Pattern.compile("\\^FB(\\d+),(\\d+)");
    private static final Pattern CARROS_SKU_FIELD = Pattern.compile(
            "(\\^FD[^^]*?SKU:\\s*)(\\d+)([^^]*?)(\\^FS)");

    // Anclas de texto que dependen del formato de etiqueta de ML: la inyección de
    // ZONA se posiciona debajo del campo "Unidad" y la de COD.EXT. debajo del de "SKU:".
    // Si ML cambia estos textos, la inyección se omite y se registra una advertencia
    // (en vez de fallar en silencio produciendo etiquetas incompletas).
    private static final String ANCHOR_UNIDAD = "Unidad";
    private static final String ANCHOR_SKU = "SKU:";

    private SortResult injectZplHeaders(SortResult result, ExcelMapping excelMapping,
                                        Map<String, ar.com.leo.etiquetas.model.MedidaSku> medidas,
                                        Map<String, String> skusPendientesOut) {
        Map<String, String> skuToExtCode = excelMapping.skuToExternalCode();
        Map<String, ComboProduct> normalizedCombos = loadNormalizedCombos();
        List<SortedLabelGroup> newGroups = new ArrayList<>();
        int labelPosition = 1;
        Set<String> skusYaMarcados = new HashSet<>();
        for (SortedLabelGroup group : result.groups()) {
            String zone = group.zone();
            String sku = group.sku();
            String zoneText = "ZONA: " + zone;
            String extCodeText = null;
            if (!"CARROS".equals(zone)) {
                String extCode = skuToExtCode.getOrDefault(sku, "");
                // Fallback: si el SKU es un combo con un solo componente, usar el COD.EXT.
                // del componente. Aplica a cualquier zona (el combo puede estar mapeado
                // a J*, T*, COMBOS, etc. en el Excel de stock).
                if (normalizedCombos != null) {
                    String componentExt = resolveSingleComponentExtCode(sku, normalizedCombos, skuToExtCode);
                    if (componentExt != null) extCode = componentExt;
                }
                extCodeText = "COD.EXT.: " + (extCode.isEmpty() ? "-" : extCode);
            }

            // Detectar si el SKU individual está pendiente de medición (MEDIR tag).
            // Solo aplica a productos individuales (no CARROS) con SKU numérico válido,
            // y solo se marca/guarda cuando el pedido es de 1 unidad (se evalúa por etiqueta).
            boolean skuPendienteMedicion = false;
            if (medidas != null && !"CARROS".equals(zone) && sku != null && !sku.isBlank()
                    && !sku.contains("\n") && sku.matches("\\d+")) {
                ar.com.leo.etiquetas.model.MedidaSku m = medidas.get(sku);
                skuPendienteMedicion = (m == null || !m.estaMedido());
            }

            List<ZplLabel> newLabels = new ArrayList<>();
            for (ZplLabel label : group.labels()) {
                String raw = label.rawZpl();
                boolean necesitaMedir = skuPendienteMedicion && label.quantity() == 1;
                if (necesitaMedir) {
                    skusPendientesOut.putIfAbsent(sku, group.productDescription() != null ? group.productDescription() : "");
                }
                // Inyectar número de posición (#1, #2, ...) arriba a la izquierda en negrita
                // Se inserta antes de ^LH (si existe) para que use coordenadas absolutas (top-left del label)
                String posText = "#" + labelPosition;
                int lhIdx = raw.indexOf("^LH");
                int insertIdx = lhIdx >= 0 ? lhIdx : raw.indexOf("^XA") + 3;
                // ^LH0,0 resetea el label home a (0,0) para usar coordenadas absolutas. Y=30 para no ser cortado por el borde superior
                String posField1 = "^FO45,30^A0N,35,35^FD" + posText + "^FS";
                String posField2 = "^FO46,30^A0N,35,35^FD" + posText + "^FS";
                String posField3 = "^FO45,31^A0N,35,35^FD" + posText + "^FS";
                String medirPrefix = "";
                // Solo marcamos una etiqueta por SKU aunque haya varias elegibles (todas de 1 unidad).
                // Alcanza con una sola medición para cargar las dimensiones del SKU.
                if (necesitaMedir && skusYaMarcados.add(sku)) {
                    // Banner MEDIR: [SKU] en video inverso (blanco sobre negro), bien visible.
                    // Se ubica a la derecha del #X (x>=180) para no taparlo.
                    String medirText = "MEDIR: " + sku;
                    medirPrefix =
                            "^FO200,15^GB580,65,65^FS\n"
                            + "^FO200,22^A0N,50,50^FB580,1,0,C^FR^FD" + medirText + "^FS\n";
                }
                raw = raw.substring(0, insertIdx) + "^LH0,0\n" + posField1 + "\n" + posField2 + "\n" + posField3 + "\n" + medirPrefix + raw.substring(insertIdx);
                labelPosition++;

                // Parsear ^LH original para convertir coordenadas relativas a absolutas
                int origLhX = 0, origLhY = 0;
                Matcher lhMatcher = LH_PATTERN.matcher(raw);
                // Buscar el segundo ^LH (el primero es el ^LH0,0 inyectado para #X)
                if (lhMatcher.find() && lhMatcher.find()) {
                    origLhX = Integer.parseInt(lhMatcher.group(1));
                    origLhY = Integer.parseInt(lhMatcher.group(2));
                }

                // 1. Inyectar ZONA siempre debajo de "Unidades"
                int unidadIdx = raw.indexOf(ANCHOR_UNIDAD);
                if (unidadIdx < 0) {
                    AppLogger.warn("ZPL - No se encontró el ancla '" + ANCHOR_UNIDAD
                            + "' para inyectar ZONA (sku=" + sku + ", zona=" + zone
                            + "). ¿Cambió el formato de etiqueta de ML?");
                } else {
                    int zoneAnchorFsIdx = raw.indexOf("^FS", unidadIdx);
                    int zoneAnchorFoIdx = raw.lastIndexOf("^FO", unidadIdx);
                    if (zoneAnchorFoIdx >= 0 && zoneAnchorFsIdx >= 0) {
                        String segment = raw.substring(zoneAnchorFoIdx, zoneAnchorFsIdx);
                        Matcher foMatcher = FO_PATTERN.matcher(segment);
                        Matcher fontMatcher = FONT_PATTERN.matcher(segment);
                        Matcher fbMatcher = FB_PATTERN.matcher(segment);
                        if (foMatcher.find()) {
                            int y = Integer.parseInt(foMatcher.group(2));
                            int fontH = fontMatcher.find() ? Integer.parseInt(fontMatcher.group(1)) : 28;
                            int fbLines = fbMatcher.find() ? Integer.parseInt(fbMatcher.group(2)) : 1;
                            int newY = y + (fontH * fbLines) + 4;
                            int fontSize = 25;
                            // Usar coordenadas absolutas (^LH0,0) para alinear con la tijera/logo
                            int absZoneX = 20;
                            int absZoneY = origLhY + newY;
                            String field1 = "^LH0,0\n^FO" + absZoneX + "," + absZoneY + "^A0N," + fontSize + "," + fontSize + "^FD" + zoneText + "^FS";
                            String field2 = "^FO" + (absZoneX + 1) + "," + absZoneY + "^A0N," + fontSize + "," + fontSize + "^FD" + zoneText + "^FS";
                            String restoreLh = "\n^LH" + origLhX + "," + origLhY;
                            raw = raw.substring(0, zoneAnchorFsIdx + 3) + "\n" + field1 + "\n" + field2 + restoreLh + raw.substring(zoneAnchorFsIdx + 3);
                        }
                    }
                }

                // 2. Inyectar COD.EXT. debajo del último SKU (solo para zonas que no son CARROS)
                if (extCodeText != null) {
                    int lastSkuIdx = raw.lastIndexOf(ANCHOR_SKU);
                    if (lastSkuIdx < 0) {
                        AppLogger.warn("ZPL - No se encontró el ancla '" + ANCHOR_SKU
                                + "' para inyectar COD.EXT. (sku=" + sku + ", zona=" + zone
                                + "). ¿Cambió el formato de etiqueta de ML?");
                    } else {
                        int extAnchorFsIdx = raw.indexOf("^FS", lastSkuIdx);
                        int extAnchorFoIdx = raw.lastIndexOf("^FO", lastSkuIdx);
                        if (extAnchorFoIdx >= 0 && extAnchorFsIdx >= 0) {
                            String segment = raw.substring(extAnchorFoIdx, extAnchorFsIdx);
                            Matcher foMatcher = FO_PATTERN.matcher(segment);
                            Matcher fontMatcher = FONT_PATTERN.matcher(segment);
                            Matcher fbMatcher = FB_PATTERN.matcher(segment);
                            if (foMatcher.find()) {
                                int x = Integer.parseInt(foMatcher.group(1));
                                int y = Integer.parseInt(foMatcher.group(2));
                                int fontH = fontMatcher.find() ? Integer.parseInt(fontMatcher.group(1)) : 28;
                                int fbLines = fbMatcher.find() ? Integer.parseInt(fbMatcher.group(2)) : 1;
                                int newY = y + (fontH * fbLines) + 4;
                                int fontSize = 25;
                                String field1 = "^FO" + x + "," + newY + "^A0N," + fontSize + "," + fontSize + "^FD" + extCodeText + "^FS";
                                String field2 = "^FO" + (x + 1) + "," + newY + "^A0N," + fontSize + "," + fontSize + "^FD" + extCodeText + "^FS";
                                raw = raw.substring(0, extAnchorFsIdx + 3) + "\n" + field1 + "\n" + field2 + raw.substring(extAnchorFsIdx + 3);
                            }
                        }
                    }
                }
                // Resaltar número de unidad (video inverso) si > 1 y zona no es CARROS ni RETIROS
                raw = highlightUnitIfNeeded(raw, zone);
                // Para CARROS con productos listados, resaltar cantidades individuales > 1 (ej: "| 3 u.")
                // y agregar COD.EXT. inline junto a cada SKU listado
                if ("CARROS".equals(zone)) {
                    raw = highlightCarrosProductQuantities(raw);
                    raw = injectCarrosExtCodes(raw, skuToExtCode);
                }
                newLabels.add(new ZplLabel(raw, label.sku(), label.productDescription(), label.details(), label.quantity(), label.turbo(), label.orderIds()));
            }
            newGroups.add(new SortedLabelGroup(zone, group.sku(), group.productDescription(), group.details(), group.orderIds(), newLabels));
        }
        return new SortResult(newGroups, result.statistics());
    }

    private String highlightUnitIfNeeded(String rawZpl, String zone) {
        String zoneUpper = zone.toUpperCase();
        if (zoneUpper.startsWith("CARROS") || zoneUpper.startsWith("RETIROS")) {
            return rawZpl;
        }
        Matcher m = UNIT_PATTERN.matcher(rawZpl);
        if (m.find()) {
            int unitNum = Integer.parseInt(m.group(4));
            if (unitNum > 1) {
                int x = Integer.parseInt(m.group(2));
                int y = Integer.parseInt(m.group(3));
                // Caja negra rellena detrás del número, tamaño ajustado a la cantidad de dígitos
                int digits = String.valueOf(unitNum).length();
                int boxW = digits * 50 + 28;
                int boxH = 76;
                int boxX = x + (160 - boxW) / 2;
                String box = "^FO" + boxX + "," + (y - 3) + "^GB" + boxW + "," + boxH + "," + boxH + "^FS\n";
                // ^FR debe ir ANTES de ^FD para invertir el campo (blanco sobre negro)
                String prefix = m.group(1); // termina en ^FD
                String correctedPrefix = prefix.substring(0, prefix.length() - 3) + "^FR^FD";
                String replacement = box + correctedPrefix + m.group(4) + m.group(5);
                rawZpl = m.replaceFirst(Matcher.quoteReplacement(replacement));
            }
        }
        return rawZpl;
    }

    private static final Pattern CHECKBOX_PATTERN = Pattern.compile("\\^FO(\\d+),(\\d+)\\^GB30,30,3\\^FS");
    private static final Pattern PRODUCT_QTY_PATTERN = Pattern.compile("\\|\\s*(\\d+)\\s*u\\.");
    /** Convierte texto ZPL con ^FH hex a forma renderizada: cada secuencia UTF-8 (_C3_A9 etc.) se reemplaza por 'X'. */
    private static String toRenderedForm(String fdText) {
        StringBuilder sb = new StringBuilder();
        int i = 0;
        while (i < fdText.length()) {
            if (i + 2 < fdText.length() && fdText.charAt(i) == '_'
                    && isHexDigit(fdText.charAt(i + 1)) && isHexDigit(fdText.charAt(i + 2))) {
                int firstByte = Integer.parseInt(fdText.substring(i + 1, i + 3), 16);
                i += 3;
                int extraBytes = 0;
                if (firstByte >= 0xC0 && firstByte < 0xE0) extraBytes = 1;
                else if (firstByte >= 0xE0 && firstByte < 0xF0) extraBytes = 2;
                else if (firstByte >= 0xF0) extraBytes = 3;
                for (int b = 0; b < extraBytes; b++) {
                    if (i + 2 < fdText.length() && fdText.charAt(i) == '_'
                            && isHexDigit(fdText.charAt(i + 1)) && isHexDigit(fdText.charAt(i + 2))) {
                        i += 3;
                    } else break;
                }
                sb.append('X');
            } else {
                sb.append(fdText.charAt(i));
                i++;
            }
        }
        return sb.toString();
    }

    private static boolean isHexDigit(char c) {
        return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
    }

    private String highlightCarrosProductQuantities(String rawZpl) {
        // Buscar checkboxes de productos y detectar cuáles tienen "| N u." con N > 1
        Matcher m = CHECKBOX_PATTERN.matcher(rawZpl);
        List<int[]> mods = new ArrayList<>();

        while (m.find()) {
            int afterCheckbox = m.end();
            int fdStart = rawZpl.indexOf("^FD", afterCheckbox);
            int fsEnd = fdStart >= 0 ? rawZpl.indexOf("^FS", fdStart) : -1;
            if (fdStart >= 0 && fsEnd >= 0 && (fdStart - afterCheckbox) < 200) {
                String fdContent = rawZpl.substring(fdStart + 3, fsEnd);
                Matcher qtyM = PRODUCT_QTY_PATTERN.matcher(fdContent);
                if (qtyM.find()) {
                    int qty = Integer.parseInt(qtyM.group(1));
                    if (qty > 1) {
                        int removeStart = fdStart + 3 + qtyM.start();
                        int removeEnd = fdStart + 3 + qtyM.end();
                        // Incluir espacio previo al |
                        if (removeStart > 0 && rawZpl.charAt(removeStart - 1) == ' ') {
                            removeStart--;
                        }

                        // Parsear posición y fuente del campo de texto del producto
                        String fieldSetup = rawZpl.substring(afterCheckbox, fdStart);
                        Matcher foM = FO_PATTERN.matcher(fieldSetup);
                        int textX = 200, textY = Integer.parseInt(m.group(2));
                        if (foM.find()) {
                            textX = Integer.parseInt(foM.group(1));
                            textY = Integer.parseInt(foM.group(2));
                        }
                        Matcher fontM = FONT_PATTERN.matcher(fieldSetup);
                        int fontH = 22, fontW = 22;
                        if (fontM.find()) {
                            fontH = Integer.parseInt(fontM.group(1));
                            fontW = Integer.parseInt(fontM.group(2));
                            if (fontW == 0) fontW = fontH;
                        }
                        Matcher fbM = FB_PATTERN.matcher(fieldSetup);
                        int fbWidth = 570;
                        if (fbM.find()) {
                            fbWidth = Integer.parseInt(fbM.group(1));
                        }

                        // Calcular texto restante (sin " | N u.")
                        int fdContentRemoveStart = removeStart - (fdStart + 3);
                        int fdContentRemoveEnd = removeEnd - (fdStart + 3);
                        String remainingText = fdContent.substring(0, fdContentRemoveStart) + fdContent.substring(fdContentRemoveEnd);

                        // Convertir a forma renderizada (hex → char placeholder) para simular word-wrap
                        // Usar remainingText completo (con "...") para que el wrap refleje lo que realmente se renderiza
                        String rendered = toRenderedForm(remainingText);
                        // Factor para word-wrap (ajustado para coincidir con wrapping real de ZPL A0)
                        double wrapCharW = fontW * 0.46;
                        int charsPerLine = Math.max(1, (int) (fbWidth / wrapCharW));

                        // Simular word-wrap de ^FB para encontrar posición en la última línea
                        int lineNum = 0;
                        int pos = 0;
                        int lastLineChars = 0;
                        while (pos < rendered.length()) {
                            int remaining = rendered.length() - pos;
                            if (remaining <= charsPerLine) {
                                lastLineChars = remaining;
                                break;
                            }
                            int maxEnd = pos + charsPerLine;
                            int lastSpace = rendered.lastIndexOf(' ', maxEnd);
                            if (lastSpace > pos) {
                                pos = lastSpace + 1;
                            } else {
                                pos = maxEnd;
                            }
                            lineNum++;
                        }

                        // Badge inline después del texto visible
                        String qtyTextTemp = qty + " u.";
                        int boxWTemp = qtyTextTemp.length() * 13 + 8;
                        double posCharW = fontW * 0.50;
                        int qtyX = textX + (int) (lastLineChars * posCharW) + 16;

                        // Si no cabe en la línea, mover a la siguiente alineado a la izquierda
                        if (qtyX + boxWTemp > textX + fbWidth) {
                            qtyX = textX;
                            lineNum++;
                        }
                        int qtyY = textY + lineNum * fontH;

                        mods.add(new int[]{qty, qtyX, qtyY, fontH, removeStart, removeEnd, fsEnd + 3});
                    }
                }
            }
        }

        if (mods.isEmpty()) return rawZpl;

        // Aplicar de atrás hacia adelante para no desplazar índices
        StringBuilder sb = new StringBuilder(rawZpl);
        for (int i = mods.size() - 1; i >= 0; i--) {
            int[] mod = mods.get(i);
            int qty = mod[0], qtyX = mod[1], qtyY = mod[2];
            int fontH = mod[3];
            int removeStart = mod[4], removeEnd = mod[5];
            int insertPos = mod[6];

            // Superponer "N u." resaltado (video inverso) sobre el texto original
            String qtyText = qty + " u.";
            int fontSize = Math.min(fontH, 22);
            int boxW = qtyText.length() * 13 + 8;
            int boxH = fontSize + 4;
            String boldField = "\n^FO" + qtyX + "," + (qtyY - 1)
                    + "^GB" + boxW + "," + boxH + "," + boxH + "^FS"
                    + "\n^FO" + qtyX + "," + (qtyY + 1)
                    + "^A0N," + fontSize + "," + fontSize + "^FB" + boxW + ",1,0,C^FR^FD" + qtyText + "^FS";
            // Primero insertar el campo resaltado (posición posterior), luego borrar " | N u." del texto original
            sb.insert(insertPos, boldField);
            sb.delete(removeStart, removeEnd);
        }

        return sb.toString();
    }

    private String injectCarrosExtCodes(String rawZpl, Map<String, String> skuToExtCode) {
        Matcher m = CARROS_SKU_FIELD.matcher(rawZpl);
        StringBuilder sb = new StringBuilder();
        while (m.find()) {
            String prefix = m.group(1);
            String sku = m.group(2);
            String suffix = m.group(3);
            String fsTag = m.group(4);
            String extCode = skuToExtCode.getOrDefault(sku, "");
            String ceText = " | COD.EXT.: " + (extCode.isEmpty() ? "-" : extCode);
            m.appendReplacement(sb, Matcher.quoteReplacement(prefix + sku + suffix + ceText + fsTag));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    private Map<String, ar.com.leo.etiquetas.model.MedidaSku> loadMedidasMap() {
        if (medidasEnabledCheck == null || !medidasEnabledCheck.isSelected()) return null;
        String path = medidasExcelField == null ? null : medidasExcelField.getText();
        if (path == null || path.isBlank()) return null;
        try {
            return medidasManager.leerMedidas(Path.of(path));
        } catch (Exception e) {
            AppLogger.warn("No se pudo leer el Excel de medidas: " + e.getMessage());
            return null;
        }
    }

    private int guardarSkusPendientesMedicion(Map<String, String> skusPendientes) {
        if (medidasEnabledCheck == null || !medidasEnabledCheck.isSelected()) return 0;
        String path = medidasExcelField == null ? null : medidasExcelField.getText();
        return guardarSkusPendientesMedicion(skusPendientes, path);
    }

    /**
     * Appendea al Excel de medidas los SKU detectados como pendientes. Devuelve cuántos se agregaron
     * efectivamente (los que ya figuraban se omiten).
     */
    private int guardarSkusPendientesMedicion(Map<String, String> skusPendientes, String path) {
        if (path == null || path.isBlank()) return 0;
        int agregados = 0;
        try {
            if (skusPendientes != null && !skusPendientes.isEmpty()) {
                agregados = medidasManager.agregarPendientes(Path.of(path), skusPendientes.keySet());
                if (agregados > 0) {
                    AppLogger.info("MEDIDAS - " + agregados + " SKU(s) pendientes agregados al Excel madre.");
                }
            }
        } catch (Exception e) {
            AppLogger.warn("No se pudo actualizar el Excel de medidas: " + e.getMessage());
        }
        Platform.runLater(this::actualizarBotonSubirMedidas);
        return agregados;
    }

    private Map<String, ComboProduct> loadNormalizedCombos() {
        String comboPath = comboExcelField.getText();
        if (comboPath == null || comboPath.isBlank()) return null;
        try {
            Map<String, ComboProduct> all = comboExcelReader.read(Path.of(comboPath));
            if (all.isEmpty()) return null;
            Map<String, ComboProduct> normalized = new LinkedHashMap<>(all);
            for (var entry : all.entrySet()) {
                String norm = ZplParser.normalizeSku(entry.getKey());
                if (norm != null && !norm.startsWith("SKU INVALIDO")) {
                    normalized.putIfAbsent(norm, entry.getValue());
                }
            }
            return normalized;
        } catch (Exception e) {
            AppLogger.warn("Error al leer Excel de combos para COD.EXT.: " + e.getMessage());
            return null;
        }
    }

    private String resolveSingleComponentExtCode(String sku, Map<String, ComboProduct> combos, Map<String, String> skuToExtCode) {
        ComboProduct combo = combos.get(sku);
        if (combo == null) {
            String norm = ZplParser.normalizeSku(sku);
            if (norm != null && !norm.startsWith("SKU INVALIDO")) {
                combo = combos.get(norm);
            }
        }
        if (combo == null || combo.componentes().size() != 1) return null;
        String componentSku = combo.componentes().getFirst().codigoComponente();
        if (componentSku == null || componentSku.isBlank()) return null;
        String ext = skuToExtCode.getOrDefault(componentSku, "");
        if (ext.isEmpty()) {
            String normComp = ZplParser.normalizeSku(componentSku);
            if (normComp != null && !normComp.startsWith("SKU INVALIDO")) {
                ext = skuToExtCode.getOrDefault(normComp, "");
            }
        }
        return ext.isEmpty() ? null : ext;
    }

    private int extractQuantityFromLabels(List<ZplLabel> labels) {
        int total = 0;
        for (ZplLabel label : labels) {
            total += label.quantity();
        }
        return total;
    }

    private void setupComboIcons(ComboBox<String> combo, Map<String, String> icons) {
        javafx.util.Callback<ListView<String>, ListCell<String>> factory = lv -> new ListCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    String icon = icons.getOrDefault(item, "");
                    setText(icon.isEmpty() ? item : icon + " " + item);
                }
            }
        };
        combo.setCellFactory(factory);
        combo.setButtonCell(factory.call(null));
    }

    private void displayResult(SortResult result) {
        ObservableList<LabelTableRow> rows = FXCollections.observableArrayList();
        for (SortedLabelGroup group : result.groups()) {
            rows.add(new LabelTableRow(
                    group.orderIds(),
                    group.zone(),
                    group.sku(),
                    group.productDescription(),
                    group.details(),
                    extractQuantityFromLabels(group.labels())));
        }
        filteredLabels = new FilteredList<>(rows, p -> true);
        SortedList<LabelTableRow> sortedLabels = new SortedList<>(filteredLabels);
        sortedLabels.comparatorProperty().bind(labelTable.comparatorProperty());
        labelTable.setItems(sortedLabels);
        searchField.clear();

        LabelStatistics stats = result.statistics();
        int totalProductos = result.groups().stream()
                .mapToInt(g -> extractQuantityFromLabels(g.labels()))
                .sum();
        StringJoiner sj = new StringJoiner("  \u2502  ");
        sj.add("Etiquetas: " + stats.totalLabels());
        sj.add("Productos: " + totalProductos);
        sj.add("SKUs: " + stats.uniqueSkus());
        if (stats.unmappedLabels() > 0) {
            sj.add("\u26a0 Sin zona: " + stats.unmappedLabels());
        }
        for (Map.Entry<String, Integer> entry : stats.countByZone().entrySet()) {
            if (entry.getValue() > 0) {
                sj.add(entry.getKey() + ": " + entry.getValue());
            }
        }

        String text = sj.toString();
        statsLabel.setText(text);
        Tooltip tip = new Tooltip(text.replace(" | ", "\n"));
        tip.setShowDelay(Duration.millis(200));
        statsLabel.setTooltip(tip);
        statsBar.setVisible(true);
        statsBar.setManaged(true);
        searchBar.setVisible(true);
        searchBar.setManaged(true);
    }

    /**
     * Reordena las etiquetas intercalando primera y segunda mitad para compensar
     * el doblado en acordeón y corte al medio.
     * Ej: [1,2,3,4,5,6,7,8,9,10,11] con mitad=6 → [1,7,2,8,3,9,4,10,5,11,6]
     */
    static <T> List<T> interleaveForPrint(List<T> labels) {
        int n = labels.size();
        if (n <= 1) return labels;
        int mitad = (n + 1) / 2; // ceil(N / 2)
        List<T> result = new ArrayList<>(n);
        for (int i = 0; i < mitad; i++) {
            result.add(labels.get(i));
            int j = i + mitad;
            if (j < n) {
                result.add(labels.get(j));
            }
        }
        return result;
    }

    private javafx.stage.Window getWindow() {
        return labelTable.getScene().getWindow();
    }

    // ═══════════════════════════════════════════════════════════════
    //  Pickit Tab
    // ═══════════════════════════════════════════════════════════════

    private void initPickitTab() {
        // Audio clips
        try {
            errorSound = new AudioClip(getClass().getResource("/ar/com/leo/audios/error.mp3").toExternalForm());
            successSound = new AudioClip(getClass().getResource("/ar/com/leo/audios/success.mp3").toExternalForm());
            errorSound.setVolume(0.1);
            successSound.setVolume(0.1);
        } catch (Exception e) {
            AppLogger.warn("No se pudieron cargar los audios de Pickit: " + e.getMessage());
        }

        // ToggleGroup para los radio de SLA
        ToggleGroup slaGroup = new ToggleGroup();
        radioPickitSlaHoy.setToggleGroup(slaGroup);
        radioPickitSlaTodos.setToggleGroup(slaGroup);

        // Checkbox ML habilita/deshabilita sección de despacho ML
        pickitCheckML.selectedProperty().addListener((obs, old, val) -> pickitSlaSection.setDisable(!val));
        // Checkbox Manual habilita/deshabilita sección de productos manuales
        pickitCheckManual.selectedProperty().addListener((obs, old, val) -> pickitManualSection.setDisable(!val));

        // Solo Turbo: desactiva Nube y Manual cuando está marcado
        pickitCheckTurbo.selectedProperty().addListener((obs, old, val) -> {
            pickitCheckNube.setDisable(val);
            pickitCheckManual.setDisable(val);
            if (val) {
                pickitCheckNube.setSelected(false);
                pickitCheckManual.setSelected(false);
            }
        });

        // Tabla de productos manuales
        pickitColSku.setCellValueFactory(new PropertyValueFactory<>("sku"));
        pickitColCantidad.setCellValueFactory(new PropertyValueFactory<>("cantidad"));
        pickitColSku.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item);
                setStyle("-fx-alignment: CENTER;");
            }
        });
        pickitColCantidad.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(Double item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    setText(item == Math.floor(item) ? String.valueOf(item.intValue()) : String.valueOf(item));
                    setStyle("-fx-alignment: CENTER;");
                }
            }
        });
        pickitManualTable.setItems(pickitProductosList);
        pickitManualTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        pickitManualTable.setFixedCellSize(-1);
        centerColumnHeaders(pickitManualTable);
        lockColumns(pickitManualTable);

        // Listener para editar producto al seleccionar fila
        pickitManualTable.getSelectionModel().selectedItemProperty().addListener((obs, oldVal, newVal) -> {
            if (newVal != null) {
                pickitProductoEnEdicion = newVal;
                pickitSkuField.setText(newVal.getSku());
                double cant = newVal.getCantidad();
                pickitCantidadField.setText(cant == Math.floor(cant) ? String.valueOf((int) cant) : String.valueOf(cant));
                pickitBtnAgregarModificar.setText("\u270F Modificar");
            } else {
                pickitProductoEnEdicion = null;
                pickitBtnAgregarModificar.setText("\u2795 Agregar");
            }
        });

        // Menú contextual para copiar el log
        ContextMenu logContextMenu = new ContextMenu();
        MenuItem copiarTodo = new MenuItem("Copiar todo");
        copiarTodo.setOnAction(e -> {
            ClipboardContent content = new ClipboardContent();
            content.putString(LogHelper.extractText(pickitLogTextFlow));
            Clipboard.getSystemClipboard().setContent(content);
        });
        logContextMenu.getItems().add(copiarTodo);
        pickitLogScrollPane.setContextMenu(logContextMenu);

        // Cargar preferencias
        loadPickitPreferences();
    }

    private void loadPickitPreferences() {
        String pathImportDir = pickitPrefs.get("pathImportDir", null);
        if (pathImportDir != null && !pathImportDir.isBlank()) {
            File dir = new File(pathImportDir);
            if (dir.isDirectory()) {
                pickitImportDir = dir;
            }
        }
        boolean slaHoy = pickitPrefs.getBoolean("slaHoy", true);
        radioPickitSlaHoy.setSelected(slaHoy);
        radioPickitSlaTodos.setSelected(!slaHoy);
    }

    private void savePickitPreferences() {
        if (pickitImportDir != null) {
            pickitPrefs.put("pathImportDir", pickitImportDir.getAbsolutePath());
        }
        pickitPrefs.putBoolean("slaHoy", radioPickitSlaHoy.isSelected());
    }

    private void pickitAppendLog(String message, Color color) {
        LogHelper.appendLog(pickitLogTextFlow, pickitLogScrollPane, message, color);
    }

    @FXML
    private void onPickitAgregarProducto() {
        String sku = pickitSkuField.getText();
        if (sku == null || sku.isBlank()) return;
        sku = sku.trim();

        if (!sku.matches("\\d+")) {
            pickitLogTextFlow.getChildren().clear();
            pickitAppendLog("Error: el SKU debe ser numérico.", Color.FIREBRICK);
            return;
        }

        double cantidad = 1;
        String cantidadText = pickitCantidadField.getText();
        if (cantidadText != null && !cantidadText.isBlank()) {
            try {
                cantidad = Double.parseDouble(cantidadText.trim());
            } catch (NumberFormatException e) {
                pickitLogTextFlow.getChildren().clear();
                pickitAppendLog("Error: cantidad inválida.", Color.FIREBRICK);
                return;
            }
        }
        if (cantidad <= 0) {
            pickitLogTextFlow.getChildren().clear();
            pickitAppendLog("Error: la cantidad debe ser mayor a 0.", Color.FIREBRICK);
            return;
        }

        final String skuFinal = sku;
        boolean duplicado = pickitProductosList.stream()
                .anyMatch(p -> p.getSku().equalsIgnoreCase(skuFinal) && p != pickitProductoEnEdicion);
        if (duplicado) {
            pickitLogTextFlow.getChildren().clear();
            pickitAppendLog("Error: ya existe un producto con SKU " + sku, Color.FIREBRICK);
            return;
        }

        if (pickitProductoEnEdicion != null) {
            pickitProductoEnEdicion.setSku(sku);
            pickitProductoEnEdicion.setCantidad(cantidad);
            pickitManualTable.refresh();
            pickitProductoEnEdicion = null;
        } else {
            pickitProductosList.add(new ProductoManual(sku, cantidad));
        }

        pickitManualTable.getSelectionModel().clearSelection();
        pickitBtnAgregarModificar.setText("\u2795 Agregar");
        pickitSkuField.clear();
        pickitCantidadField.clear();
        pickitSkuField.requestFocus();
    }

    @FXML
    private void onPickitEliminarProducto() {
        ProductoManual selected = pickitManualTable.getSelectionModel().getSelectedItem();
        if (selected != null) {
            pickitProductosList.remove(selected);
            if (selected == pickitProductoEnEdicion) {
                pickitProductoEnEdicion = null;
            }
            pickitManualTable.getSelectionModel().clearSelection();
            pickitBtnAgregarModificar.setText("\u2795 Agregar");
            pickitSkuField.clear();
            pickitCantidadField.clear();
        }
    }

    @FXML
    private void onPickitLimpiarProductos() {
        pickitProductosList.clear();
        pickitProductoEnEdicion = null;
        pickitManualTable.getSelectionModel().clearSelection();
        pickitBtnAgregarModificar.setText("\u2795 Agregar");
        pickitSkuField.clear();
        pickitCantidadField.clear();
    }

    @FXML
    private void onPickitImportarExcel() {
        FileChooser fc = new FileChooser();
        fc.setTitle("Importar productos manuales desde Excel");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Archivo XLSX", "*.xlsx"));
        File initialDir = (pickitImportDir != null && pickitImportDir.isDirectory())
                ? pickitImportDir : new File(System.getProperty("user.dir"));
        fc.setInitialDirectory(initialDir);

        File file = fc.showOpenDialog(getWindow());
        if (file == null) return;

        pickitImportDir = file.getParentFile();
        savePickitPreferences();
        pickitLogTextFlow.getChildren().clear();

        try {
            List<ProductoManual> importados = ExcelManager.obtenerProductosManualesDesdeExcel(file);
            int nuevos = 0;
            int sumados = 0;

            for (ProductoManual importado : importados) {
                ProductoManual existente = pickitProductosList.stream()
                        .filter(p -> p.getSku().equalsIgnoreCase(importado.getSku()))
                        .findFirst().orElse(null);
                if (existente != null) {
                    existente.setCantidad(existente.getCantidad() + importado.getCantidad());
                    sumados++;
                } else {
                    pickitProductosList.add(importado);
                    nuevos++;
                }
            }

            pickitManualTable.refresh();
            pickitAppendLog("Importación completada: " + nuevos + " nuevos, " + sumados + " sumados a existentes.", Color.web("#2E7D32"));
        } catch (Exception e) {
            pickitAppendLog("Error al importar: " + e.getMessage(), Color.FIREBRICK);
            AppLogger.error("Error al importar Excel: " + e.getMessage(), e);
        }
    }

    @FXML
    private void onPickitGenerar() {
        pickitLogTextFlow.getChildren().clear();

        String stockPath = excelFileField.getText();
        if (stockPath == null || stockPath.isBlank()) {
            pickitAppendLog("Error: seleccionar el archivo Excel de stock primero (selector general).", Color.FIREBRICK);
            return;
        }
        File stockFile = new File(stockPath);
        if (!stockFile.isFile()) {
            pickitAppendLog("Error: el archivo Excel de stock no existe: " + stockPath, Color.FIREBRICK);
            return;
        }

        String combosPath = comboExcelField.getText();
        if (combosPath == null || combosPath.isBlank()) {
            pickitAppendLog("Error: seleccionar el archivo Excel de combos primero (selector general).", Color.FIREBRICK);
            return;
        }
        File combosFile = new File(combosPath);
        if (!combosFile.isFile()) {
            pickitAppendLog("Error: el archivo Excel de combos no existe: " + combosPath, Color.FIREBRICK);
            return;
        }

        savePickitPreferences();

        boolean soloHoy = radioPickitSlaHoy.isSelected();
        boolean soloTurbo = pickitCheckTurbo.isSelected();
        boolean useML = pickitCheckML.isSelected();
        boolean useNube = pickitCheckNube.isSelected();
        boolean useManual = pickitCheckManual.isSelected();

        if (!useML && !useNube && !useManual) {
            pickitAppendLog("Error: seleccionar al menos un canal.", Color.FIREBRICK);
            return;
        }

        PickitService service = new PickitService(stockFile, combosFile, pickitProductosList, soloHoy, soloTurbo, useML, useNube, useManual, pickitLogTextFlow, pickitLogScrollPane);

        service.setOnRunning(e -> {
            pickitGenerateBtn.setDisable(true);
            pickitCheckML.setDisable(true);
            pickitCheckNube.setDisable(true);
            pickitCheckManual.setDisable(true);
            pickitSlaSection.setDisable(true);
            pickitManualSection.setDisable(true);
            tabPane.getTabs().forEach(t -> {
                if (t != tabPane.getSelectionModel().getSelectedItem()) t.setDisable(true);
            });
            pickitProgressIndicator.setVisible(true);
            pickitProgressIndicator.setManaged(true);
        });
        service.setOnSucceeded(e -> {
            if (successSound != null) successSound.play();
            pickitSetInputsEnabled();
            pickitProgressIndicator.setVisible(false);
            pickitProgressIndicator.setManaged(false);
        });
        service.setOnFailed(e -> {
            if (errorSound != null) errorSound.play();
            Throwable ex = service.getException();
            String mensaje = ex != null ? ex.getLocalizedMessage() : "Error desconocido";
            pickitAppendLog("\nERROR: " + mensaje, Color.FIREBRICK);
            AppLogger.error("Error Pickit: " + mensaje, ex);
            pickitSetInputsEnabled();
            pickitProgressIndicator.setVisible(false);
            pickitProgressIndicator.setManaged(false);
        });
        service.start();
    }

    private void pickitSetInputsEnabled() {
        pickitGenerateBtn.setDisable(false);
        pickitCheckML.setDisable(false);
        pickitCheckNube.setDisable(false);
        pickitCheckManual.setDisable(false);
        // Restaurar estado según checkboxes (los bindings se re-evalúan)
        pickitSlaSection.setDisable(!pickitCheckML.isSelected());
        pickitManualSection.setDisable(!pickitCheckManual.isSelected());
        tabPane.getTabs().forEach(t -> t.setDisable(false));
    }

    // ══════════════════════════════════════════════════════════════
    // ══ Pedidos Tab ══
    // ══════════════════════════════════════════════════════════════

    private void initPedidosTab() {
        ContextMenu logContextMenu = new ContextMenu();
        MenuItem copiarTodo = new MenuItem("Copiar todo");
        copiarTodo.setOnAction(e -> {
            ClipboardContent content = new ClipboardContent();
            content.putString(LogHelper.extractText(pedidosLogTextFlow));
            Clipboard.getSystemClipboard().setContent(content);
        });
        logContextMenu.getItems().add(copiarTodo);
        pedidosLogScrollPane.setContextMenu(logContextMenu);
    }

    private void pedidosAppendLog(String message, Color color) {
        LogHelper.appendLog(pedidosLogTextFlow, pedidosLogScrollPane, message, color);
    }

    @FXML
    private void onPedidosGenerar() {
        pedidosLogTextFlow.getChildren().clear();

        String stockPath = excelFileField.getText();
        File stockFile = (stockPath != null && !stockPath.isBlank()) ? new File(stockPath) : null;
        if (stockFile != null && !stockFile.isFile()) stockFile = null;

        PedidosService service = new PedidosService(stockFile, pedidosLogTextFlow, pedidosLogScrollPane);

        service.setOnRunning(e -> {
            pedidosGenerateBtn.setDisable(true);
            tabPane.getTabs().forEach(t -> {
                if (t != tabPane.getSelectionModel().getSelectedItem()) t.setDisable(true);
            });
            pedidosProgressIndicator.setVisible(true);
            pedidosProgressIndicator.setManaged(true);
        });
        service.setOnSucceeded(e -> {
            if (successSound != null) successSound.play();
            pedidosGenerateBtn.setDisable(false);
            tabPane.getTabs().forEach(t -> t.setDisable(false));
            pedidosProgressIndicator.setVisible(false);
            pedidosProgressIndicator.setManaged(false);
        });
        service.setOnFailed(e -> {
            if (errorSound != null) errorSound.play();
            Throwable ex = service.getException();
            String mensaje = ex != null ? ex.getLocalizedMessage() : "Error desconocido";
            pedidosAppendLog("\nERROR: " + mensaje, Color.FIREBRICK);
            AppLogger.error("Error Pedidos: " + mensaje, ex);
            pedidosGenerateBtn.setDisable(false);
            tabPane.getTabs().forEach(t -> t.setDisable(false));
            pedidosProgressIndicator.setVisible(false);
            pedidosProgressIndicator.setManaged(false);
        });
        service.start();
    }
}
