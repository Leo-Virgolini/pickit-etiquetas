package ar.com.leo.ui;

import ar.com.leo.AppLogger;
import ar.com.leo.api.ml.MercadoLibreAPI;
import ar.com.leo.api.ml.model.OrdenML;
import ar.com.leo.api.ml.model.ShippingType;
import ar.com.leo.api.ml.model.Venta;
import ar.com.leo.etiquetas.model.*;
import ar.com.leo.etiquetas.parser.ComboProduct;
import ar.com.leo.etiquetas.model.DatosEmbalaje;
import ar.com.leo.etiquetas.parser.EmbalajeRenderer;
import ar.com.leo.etiquetas.parser.ComboExcelReader;
import ar.com.leo.etiquetas.parser.ExcelMappingReader;
import ar.com.leo.etiquetas.parser.MedidasExcelManager;
import ar.com.leo.etiquetas.parser.ZplParser;
import ar.com.leo.etiquetas.ui.ComboPrintDialog;
import ar.com.leo.etiquetas.ui.EstadoDato;
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
import ar.com.leo.etiquetas.sorter.CarrosOrdering;
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
    private CheckBox filterFlexCheck;
    @FXML
    private CheckBox filterColectaCheck;
    @FXML
    private CheckBox filterTurboCheck;
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
    private TableColumn<LabelTableRow, String> printNumCol;
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
    private TableColumn<LabelTableRow, EstadoDato> estandarizadoCol;
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
    @FXML
    private TableColumn<OrderTableRow, ShippingType> orderShippingCol;

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

    /** Los dos sub-tabs de Etiquetas, en el orden en que están en el FXML. */
    private static final int SUBTAB_API = 0;
    private static final int SUBTAB_LOCAL = 1;

    /**
     * Lo que un sub-tab tiene cargado. Sin esto hay un solo juego de estado y cambiar de pestaña
     * deja a la vista el resultado del otro flujo, con sus botones y sus estadísticas.
     */
    private record VistaEtiquetas(SortResult resultado, ObservableList<LabelTableRow> filas,
                                  String estadisticas, File archivoGuardado) {
    }

    private final Map<Integer, VistaEtiquetas> vistasEtiquetas = new HashMap<>();
    /** El sub-tab de API alterna entre las órdenes buscadas y las etiquetas ya descargadas. */
    private boolean apiMostrandoOrdenes = true;

    private Label placeholderApi;
    private Label placeholderLocal;
    private FilteredList<OrderTableRow> filteredOrders;
    private Runnable orderStatsUpdater;
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
        printNumCol.setCellValueFactory(new PropertyValueFactory<>("printNumber"));
        printNumCol.setCellFactory(col -> centeredCell());
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

        estandarizadoCol.setCellValueFactory(cd -> cd.getValue().estandarizadoProperty());
        estandarizadoCol.setCellFactory(col -> estadoDatoCell());

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
            // Afectar las filas que pasan el filtro de tipo de envío (ignorando el buscador),
            // para ser consistente con el contador y la descarga.
            List<? extends OrderTableRow> todasLasFilas = filteredOrders != null
                    ? filteredOrders.getSource()
                    : orderTable.getItems();
            for (OrderTableRow row : todasLasFilas) {
                if (passesTypeFilter(row)) {
                    row.setSelected(val);
                }
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
        orderShippingCol.setCellValueFactory(cd -> new javafx.beans.property.ReadOnlyObjectWrapper<>(cd.getValue().getShippingType()));
        orderShippingCol.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(ShippingType item, boolean empty) {
                super.updateItem(item, empty);
                setAlignment(Pos.CENTER);
                setText(empty || item == null ? null : item.label());
            }
        });

        // Centrar headers de ambas tablas
        centerColumnHeaders(orderTable);
        centerColumnHeaders(labelTable);

        // Bloquear reordenamiento de columnas
        lockColumns(orderTable);
        lockColumns(labelTable);

        // Cada pestaña muestra lo suyo. Placeholder dinámico según sub-tab seleccionado
        placeholderLocal = new Label("\uD83D\uDCE6 Cargue etiquetas ZPL para ver el resultado ordenado");
        placeholderLocal.setStyle("-fx-font-size: 14px; -fx-text-fill: #888;");
        placeholderApi = new Label("\uD83D\uDCCB Haga clic en 'Obtener Órdenes' para cargar las órdenes de ML");
        placeholderApi.setStyle("-fx-font-size: 14px; -fx-text-fill: #888;");
        labelTable.setPlaceholder(placeholderApi);
        // Al cambiar de pestaña se reemplazan la tabla, las estadísticas, el link al archivo y los
        // botones por los de la pestaña a la que se entra.
        etiquetasSubTabPane.getSelectionModel().selectedIndexProperty().addListener(
                (obs, anterior, actual) -> aplicarVistaDeSubTab(actual.intValue()));

        // Copiar al portapapeles con Ctrl+C (fila) y click derecho (celda)
        setupTableCopyHandler(orderTable);
        setupTableCopyHandler(labelTable);
        setupCellCopyMenu(orderTable);
        setupCellCopyMenu(labelTable);

        searchField.textProperty().addListener((obs, oldVal, newVal) -> {
            String filter = newVal == null ? "" : newVal.trim().toLowerCase();
            applyOrderFilters();
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

        filterFlexCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());
        filterColectaCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());
        filterTurboCheck.selectedProperty().addListener((o, a, b) -> applyOrderFilters());

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
                    medidas = medidasManager.leerMedidas(Path.of(path)).porSku();
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

    /**
     * Las rutas y el interruptor tal como están en la UI.
     *
     * Se toman en el hilo de JavaFX antes de arrancar el trabajo pesado: leer los Excel tarda lo
     * suficiente como para congelar la ventana, y desde un hilo no se pueden tocar los controles.
     */
    private record ConfigExcel(String stock, String combos, String medidas, boolean medidasActivo) {
    }

    private ConfigExcel configExcel() {
        return new ConfigExcel(
                excelFileField.getText(),
                comboExcelField.getText(),
                medidasExcelField == null ? null : medidasExcelField.getText(),
                medidasEnabledCheck != null && medidasEnabledCheck.isSelected());
    }

    /** Muestra un error venga del hilo de JavaFX o de uno de trabajo. */
    private void mostrarError(String titulo, String mensaje) {
        if (Platform.isFxApplicationThread()) AlertHelper.showError(titulo, mensaje);
        else Platform.runLater(() -> AlertHelper.showError(titulo, mensaje));
    }

    private ExcelMapping loadExcelMapping(ConfigExcel config) throws Exception {
        return excelReader.readMapping(Path.of(config.stock()));
    }

    @FXML
    private void onProcessLocal() {
        String zplPath = zplFileField.getText();
        if (zplPath == null || zplPath.isBlank()) {
            AlertHelper.showError("Error", "Seleccione un archivo ZPL.");
            return;
        }

        final ConfigExcel config;
        try {
            validateExcelFiles();
            config = configExcel();
        } catch (Exception e) {
            AlertHelper.showError("Error al procesar", e.getMessage(), e);
            return;
        }

        // Lee dos Excel, parsea el archivo y escribe en el Excel de medidas: en el hilo de JavaFX
        // la ventana queda congelada y el spinner ni siquiera llega a pintarse.
        setLoading(true);

        new Thread(() -> {
            try {
                ExcelMapping excelMapping = loadExcelMapping(config);
                List<ZplLabel> labels = zplParser.parseFile(Path.of(zplPath));
                MedidasExcelManager.Medidas medidas = loadMedidas(config);
                Map<String, String> skusPendientes = new LinkedHashMap<>();
                Set<String> embalajesFaltantes = new LinkedHashSet<>();
                SortResult result = injectZplHeaders(
                        labelSorter.sort(labels, excelMapping.skuToZone()), excelMapping, medidas,
                        skusPendientes, embalajesFaltantes, config.combos());
                int agregadosExcel = guardarSkusPendientesMedicion(skusPendientes, config);

                Platform.runLater(() -> {
                    setLoading(false);
                    currentResult = result;
                    showLabelTable();
                    displayResult(result, medidas);
                    mostrarMensajeSkusFaltantes(agregadosExcel, embalajesFaltantes);
                });
            } catch (Exception e) {
                AppLogger.error("Error al procesar el archivo local", e);
                Platform.runLater(() -> {
                    setLoading(false);
                    AlertHelper.showError("Error al procesar", e.getMessage(), e);
                });
            }
        }).start();
    }

    @FXML
    private void onFetchMeliOrders() {
        if (!meliInitialized) {
            AlertHelper.showError("Error", "Primero inicie sesi\u00f3n en MercadoLibre.");
            return;
        }

        final ConfigExcel config;
        try {
            validateExcelFiles();
            config = configExcel();
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
                ExcelMapping excelMapping = loadExcelMapping(config);
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

        final ConfigExcel config;
        try {
            validateExcelFiles();
            config = configExcel();
        } catch (Exception e) {
            AlertHelper.showError("Error", e.getMessage(), e);
            return;
        }

        // Filtrar las órdenes seleccionadas que además pasan el filtro de tipo de envío
        // (deduplicar por orderId). Se itera la lista completa (no orderTable.getItems(),
        // que solo trae las filas visibles según el buscador) para no perder selecciones
        // ocultas por el buscador; el filtro de tipo SÍ se respeta vía passesTypeFilter.
        List<? extends OrderTableRow> todasLasFilas = filteredOrders != null
                ? filteredOrders.getSource()
                : orderTable.getItems();
        LinkedHashSet<Long> seenOrderIds = new LinkedHashSet<>();
        List<OrdenML> seleccionadas = new ArrayList<>();
        for (OrderTableRow row : todasLasFilas) {
            if (row.isSelected() && passesTypeFilter(row)) {
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

        setLoading(true);

        new Thread(() -> {
            try {
                // Leer los Excel también tarda, así que va adentro del hilo: si no, la ventana
                // queda congelada un rato antes de que el spinner alcance a pintarse.
                ExcelMapping excelMapping = loadExcelMapping(config);
                MedidasExcelManager.Medidas medidas = loadMedidas(config);
                List<ZplLabel> labels = MercadoLibreAPI.descargarEtiquetasZplParaOrdenes(seleccionadas, turboShipmentIds);
                Map<String, String> skusPendientes = new LinkedHashMap<>();
                Set<String> embalajesFaltantes = new LinkedHashSet<>();
                SortResult result = injectZplHeaders(
                        labelSorter.sort(labels, excelMapping.skuToZone()), excelMapping, medidas,
                        skusPendientes, embalajesFaltantes, config.combos());
                int agregadosExcel = guardarSkusPendientesMedicion(skusPendientes, config);

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
                final int agregadosCount = agregadosExcel;
                final Set<String> embalajesFaltantesFinal = new LinkedHashSet<>(embalajesFaltantes);
                Platform.runLater(() -> {
                    setLoading(false);
                    currentResult = result;
                    showLabelTable();
                    displayResult(result, medidas);
                    if (finalSavedFile != null) {
                        registrarArchivoGuardado(finalSavedFile);
                    }
                    if (finalSaveError != null) {
                        AlertHelper.showError("Error al guardar", "No se pudo guardar el archivo automáticamente:\n" + finalSaveError);
                    }
                    mostrarMensajeSkusFaltantes(agregadosCount, embalajesFaltantesFinal);
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
     * Diálogo del final del lote: los SKU que salieron avisados como sin estandarizar y las filas
     * nuevas que se agregaron al Excel, que son justamente las que hay que completar con el envase.
     * No se muestra si no hay ninguna de las dos cosas.
     */
    private void mostrarMensajeSkusFaltantes(int agregadosExcel, Set<String> embalajesFaltantes) {
        boolean hayEmbalajes = embalajesFaltantes != null && !embalajesFaltantes.isEmpty();
        if (agregadosExcel <= 0 && !hayEmbalajes) return;

        StringBuilder msg = new StringBuilder();
        if (agregadosExcel > 0) {
            msg.append(agregadosExcel).append(" SKU(s) nuevo(s) agregado(s) al Excel de medidas.\n");
        }

        if (hayEmbalajes) {
            if (!msg.isEmpty()) msg.append("\n");
            msg.append(embalajesFaltantes.size())
                    .append(" SKU(s) sin estandarizar en este lote:\n");
            for (String sku : embalajesFaltantes) msg.append("  ").append(sku).append("\n");
        }

        String titulo = hayEmbalajes ? "Estandarización pendiente" : "Excel de medidas actualizado";
        AlertHelper.showInfoScrollable(titulo, msg.toString());
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

    /** Los SKU del lote a la vista, separando los multi-SKU de los CARROS. */
    private Set<String> skusDelLote() {
        Set<String> batchSkus = new HashSet<>();
        if (currentResult == null) return batchSkus;
        for (SortedLabelGroup group : currentResult.groups()) {
            for (String sku : group.sku().split("\n")) {
                String trimmed = sku.trim();
                if (!trimmed.isEmpty()) batchSkus.add(trimmed);
            }
        }
        return batchSkus;
    }

    private List<ComboProduct> findMatchingCombos(String comboPath, Set<String> batchSkus) {
        if (comboPath == null || comboPath.isBlank()) return List.of();
        if (batchSkus.isEmpty()) return List.of();

        try {
            Map<String, ComboProduct> allCombos = comboExcelReader.read(Path.of(comboPath));
            if (allCombos.isEmpty()) return List.of();

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

    /**
     * Busca los combos del lote y abre la hoja. La lectura del Excel de combos se hace en un hilo:
     * en el de JavaFX la ventana queda congelada y el spinner no llega a pintarse.
     */
    private void showComboSheetIfNeeded() {
        String comboPath = comboExcelField.getText();
        Set<String> batchSkus = skusDelLote();

        setLoading(true);
        new Thread(() -> {
            List<ComboProduct> combos = findMatchingCombos(comboPath, batchSkus);
            Platform.runLater(() -> {
                setLoading(false);
                if (combos.isEmpty()) {
                    AlertHelper.showInfo("Combos", "No se encontraron combos para las etiquetas actuales.");
                    return;
                }
                new ComboPrintDialog(getWindow(), combos).show();
            });
        }).start();
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
        filterFlexCheck.setDisable(loading);
        filterColectaCheck.setDisable(loading);
        filterTurboCheck.setDisable(loading);
        fetchOrdersBtn.setDisable(loading);
    }

    private void showOrderTable() {
        // Solo se llega acá desde el sub-tab de API: la tabla de órdenes es de ese flujo.
        apiMostrandoOrdenes = true;
        // El resumen de las órdenes es otro que el del lote de etiquetas. Se recalcula por si la
        // selección cambió mientras se estaba en la otra pestaña.
        boolean hayOrdenesBuscadas = orderStatsUpdater != null;
        if (hayOrdenesBuscadas) {
            if (orderTable.getItems().isEmpty()) {
                statsLabel.setText("No hay ordenes para mostrar");
                statsLabel.setTooltip(null);
            } else {
                orderStatsUpdater.run();
            }
        }
        statsBar.setVisible(hayOrdenesBuscadas);
        statsBar.setManaged(hayOrdenesBuscadas);
        searchBar.setVisible(hayOrdenesBuscadas);
        searchBar.setManaged(hayOrdenesBuscadas);
        // El link apunta al lote descargado, que no es lo que se está mostrando.
        mostrarLinkArchivo(null);
        downloadLabelsBtn.setVisible(true);
        downloadLabelsBtn.setManaged(true);
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
        setShippingFilterDisabled(false);
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
        // Descargar etiquetas y volver a las órdenes son pasos del flujo de la API: en Archivo
        // Local no tienen sentido.
        boolean enApi = subTabActual() == SUBTAB_API;
        downloadLabelsBtn.setVisible(enApi);
        downloadLabelsBtn.setManaged(enApi);
        boolean hayOrdenes = enApi && !orderTable.getItems().isEmpty();
        backToOrdersBtn.setVisible(hayOrdenes);
        backToOrdersBtn.setManaged(hayOrdenes);
        // El filtro de tipo de envío es un control de la vista de órdenes (pre-descarga);
        // en la vista de etiquetas se deshabilita para que no modifique el resumen.
        setShippingFilterDisabled(true);
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

            // Detectar turbo y logistic_type del envío
            boolean esTurbo = false;
            String logisticType = "";
            for (OrdenML o : group) {
                Long shipId = o.getShipmentId();
                if (shipId != null && slaMap.containsKey(shipId)) {
                    MercadoLibreAPI.SlaInfo sla = slaMap.get(shipId);
                    if (sla.turbo()) esTurbo = true;
                    if (logisticType.isEmpty() && sla.logisticType() != null && !sla.logisticType().isEmpty()) {
                        logisticType = sla.logisticType();
                    }
                }
            }
            ar.com.leo.api.ml.model.ShippingType shippingType =
                    ar.com.leo.api.ml.model.ShippingType.from(esTurbo, logisticType);

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
                    descJoiner.toString(), qtyJoiner.toString(), status, slaDate, group, shippingType));
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
                .thenComparingInt(r -> CarrosOrdering.bucket(r.getZone(), r.getSku()))
                .thenComparing(OrderTableRow::getSku));

        filteredOrders = new FilteredList<>(rows, p -> true);
        SortedList<OrderTableRow> sortedOrders = new SortedList<>(filteredOrders);
        sortedOrders.comparatorProperty().bind(orderTable.comparatorProperty());
        orderTable.setItems(sortedOrders);
        searchField.clear();
        // Se resetea antes de reasignar el updateStats de este lote (evita que el
        // applyOrderFilters inicial ejecute un actualizador de un lote anterior).
        orderStatsUpdater = null;
        applyOrderFilters();

        if (rows.isEmpty()) {
            orderTable.setPlaceholder(new Label("No se encontraron ordenes con los filtros seleccionados"));
        }

        // Estadísticas
        Runnable updateStats = () -> {
            List<OrderTableRow> scoped = rows.stream().filter(this::passesTypeFilter).toList();
            int ordCount = scoped.size();
            int prodCount = scoped.stream().mapToInt(OrderTableRow::getProductCount).sum();
            long selected = scoped.stream().filter(OrderTableRow::isSelected).count();
            int printed = 0;
            int readyToPrint = 0;
            Set<String> uniqueSkus = new HashSet<>();
            Map<String, Integer> countByZone = new LinkedHashMap<>();
            for (OrderTableRow r : scoped) {
                if ("printed".equals(r.getStatus())) printed++;
                else readyToPrint++;
                for (String s : r.getSku().split("\n")) {
                    String trimmed = s.trim();
                    if (!trimmed.isEmpty()) uniqueSkus.add(trimmed);
                }
                countByZone.merge(r.getZone(), 1, Integer::sum);
            }
            int skuCount = uniqueSkus.size();
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
        orderStatsUpdater = updateStats;

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
    // Fragmento del texto "Recortá esta parte..." de ML, sin acentos ni mayúsculas para no depender
    // de cómo venga codificada la tilde.
    private static final String ANCHOR_RECORTE = "ecort";

    /**
     * El banner MEDIR quedó fuera de uso. El código que lo arma se conserva entero: alcanza con
     * poner esto en true para que vuelva. Lo que sigue activo es lo demás que dispara la detección
     * de pendientes: el alta en el Excel y el diálogo del final del lote.
     */
    private static final boolean BANNER_MEDIR = false;

    /**
     * @param skusPendientesOut     se completa con los SKU sin medidas de una etiqueta de una
     *                              unidad, para darlos de alta en el Excel.
     * @param embalajesFaltantesOut se completa con los SKU que salieron con el aviso
     *                              "NO ESTANDARIZADO" impreso: los que tienen NO en esa columna y
     *                              los que todavía no figuran en el Excel. Un SKU cuyas etiquetas
     *                              son todas de dos o más unidades no entra, porque ahí el envase
     *                              sale como referencia y no se reclama nada.
     */
    private SortResult injectZplHeaders(SortResult result, ExcelMapping excelMapping,
                                        MedidasExcelManager.Medidas medidas,
                                        Map<String, String> skusPendientesOut,
                                        Set<String> embalajesFaltantesOut,
                                        String comboPath) {
        Map<String, String> skuToExtCode = excelMapping.skuToExternalCode();
        Map<String, ComboProduct> normalizedCombos = loadNormalizedCombos(comboPath);
        // Hace falta para los SKU que todavía no figuran en el Excel: sin ninguna fila propia, es
        // lo único que dice si la función está en uso o si no hay nada que reclamar.
        boolean moduloEmbalajeActivo = medidas != null && medidas.embalajeEnUso();
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

            // Un SKU es elegible para las líneas de embalaje y para la detección de pendientes de
            // medición si tiene número propio y su etiqueta corresponde a un solo producto. Los
            // carros listan varios; los no numéricos son sentinelas del parser ("SKU INVALIDO:
            // ...") que nunca llegan al Excel de medidas, así que no se les puede cargar ni medida
            // ni embalaje. La condición es una sola para los dos usos, para que no diverjan.
            boolean skuElegible = medidas != null && EstadoDato.esSkuElegible(zone, sku);
            ar.com.leo.etiquetas.model.MedidaSku medidaSku = skuElegible ? medidas.porSku().get(sku) : null;

            boolean skuPendienteMedicion = skuElegible && (medidaSku == null || !medidaSku.estaMedido());

            // Datos de embalaje del SKU. Son los mismos para todas las etiquetas del grupo, pero
            // las líneas dependen de la cantidad de cada una, así que se arman más abajo.
            DatosEmbalaje datosEmbalaje = skuElegible
                    ? EstadoDato.embalajeDe(medidaSku, moduloEmbalajeActivo)
                    : DatosEmbalaje.VACIO;

            List<ZplLabel> newLabels = new ArrayList<>();
            for (ZplLabel label : group.labels()) {
                String raw = label.rawZpl();
                // Un pedido de dos o más unidades no es un producto suelto, así que el envase sale
                // como referencia, y no sale nada si todavía no está cargado.
                List<String> lineasEmbalaje = EmbalajeRenderer.lineas(datosEmbalaje, label.quantity());
                String embalajeZpl = EmbalajeRenderer.campoZpl(lineasEmbalaje);
                // El aviso final lista exactamente los SKU que salieron avisados en papel.
                if (embalajesFaltantesOut != null
                        && EmbalajeRenderer.avisaSinEstandarizar(lineasEmbalaje)) {
                    embalajesFaltantesOut.add(sku);
                }
                boolean necesitaMedir = skuPendienteMedicion && label.quantity() == 1;
                if (necesitaMedir) {
                    skusPendientesOut.putIfAbsent(sku, group.productDescription() != null ? group.productDescription() : "");
                }
                // Las líneas de embalaje ocupan la franja donde ML imprime "Recortá esta parte...",
                // que no le sirve al operario. Se quita antes de calcular el punto de inserción para
                // que el índice corresponda al texto que efectivamente se va a partir.
                if (!embalajeZpl.isEmpty()) {
                    raw = quitarTextoRecorte(raw, sku);
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
                if (BANNER_MEDIR && necesitaMedir && skusYaMarcados.add(sku)) {
                    // Banner MEDIR: [SKU] en video inverso (blanco sobre negro), bien visible.
                    // Va debajo del #X (que termina en y=65) y encima del "Pack ID:" de ML (y=129),
                    // ocupando el ancho hasta x=400. El margen superior derecho quedó para las
                    // líneas de embalaje, que necesitan todo el alto para las observaciones largas.
                    // Los SKU son de 7 dígitos, así que el texto mide 14 caracteres: con fuente 42
                    // ocupa ~322 de los 380 de ancho y entra sin necesidad de una segunda línea.
                    String medirText = "MEDIR: " + sku;
                    medirPrefix =
                            "^FO20,70^GB380,52,52^FS\n"
                            + "^FO20,75^A0N,42,42^FB380,1,0,C^FR^FD" + medirText + "^FS\n";
                }
                String inyectado = "^LH0,0\n" + posField1 + "\n" + posField2 + "\n" + posField3 + "\n" + embalajeZpl + medirPrefix;
                raw = raw.substring(0, insertIdx) + inyectado + raw.substring(insertIdx);
                // A partir de acá arranca el ZPL de ML. Las anclas se buscan solo ahí: OBS es
                // texto libre del Excel y una observación que diga "2 Unidades por caja" haría
                // que ZONA se posicione tomando como referencia la línea de embalaje.
                final int inicioMl = insertIdx + inyectado.length();
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
                int unidadIdx = raw.indexOf(ANCHOR_UNIDAD, inicioMl);
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
                    if (lastSkuIdx < inicioMl) lastSkuIdx = -1;
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
            newGroups.add(new SortedLabelGroup(zone, group.sku(), group.productDescription(),
                    group.details(), newLabels));
        }
        return new SortResult(newGroups, result.statistics());
    }

    /**
     * Elimina el campo de ML "Recortá esta parte de la etiqueta para que tu paquete viaje seguro",
     * que ocupa la franja donde van las líneas de embalaje. Se busca por el fragmento "ecort" —sin
     * acentos ni mayúsculas— para no depender de cómo ML codifique la tilde, y se corta desde su
     * ^FO hasta su ^FS.
     *
     * Si el ancla no aparece (ML cambió el texto) se registra una advertencia y la etiqueta sale
     * con los dos textos encimados: visible, en vez de perder el embalaje en silencio.
     */
    private String quitarTextoRecorte(String rawZpl, String sku) {
        int idx = rawZpl.indexOf(ANCHOR_RECORTE);
        if (idx < 0) {
            AppLogger.warn("ZPL - No se encontró el ancla '" + ANCHOR_RECORTE
                    + "' para quitar el texto de recorte (sku=" + sku
                    + "). Las líneas de embalaje pueden encimarse. ¿Cambió el formato de ML?");
            return rawZpl;
        }
        int inicio = rawZpl.lastIndexOf("^FO", idx);
        int fin = rawZpl.indexOf("^FS", idx);
        if (inicio < 0 || fin < 0) return rawZpl;
        return rawZpl.substring(0, inicio) + rawZpl.substring(fin + 3);
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

    /**
     * Lo que dice el Excel de medidas. Devuelve null si el módulo está apagado o el archivo no es
     * usable; en ese caso no se inyectan las líneas de embalaje ni se detectan pendientes.
     */
    private MedidasExcelManager.Medidas loadMedidas(ConfigExcel config) {
        if (!config.medidasActivo()) return null;
        String path = config.medidas();
        if (path == null || path.isBlank()) return null;
        try {
            return medidasManager.leerMedidas(Path.of(path));
        } catch (Exception e) {
            // Se avisa además de loguear: sin medidas el lote sale sin banner MEDIR, sin líneas de
            // embalaje y con el diálogo final informando cero pendientes. El operario embalaría
            // creyendo que la etiqueta está completa.
            AppLogger.warn("No se pudo leer el Excel de medidas: " + e.getMessage());
            mostrarError("Excel de medidas",
                    "No se pudo leer el Excel de medidas, así que las etiquetas de este lote van a "
                    + "salir sin los datos de embalaje ni el aviso de medición.\n\n" + e.getMessage());
            return null;
        }
    }

    private int guardarSkusPendientesMedicion(Map<String, String> skusPendientes, ConfigExcel config) {
        if (!config.medidasActivo()) return 0;
        return guardarSkusPendientesMedicion(skusPendientes, config.medidas());
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
            // Se avisa además de loguear. El caso habitual es tener el Excel abierto en otra
            // ventana: sin aviso, el lote termina sin diálogo y el operario supone que los SKU
            // nuevos quedaron cargados cuando en realidad no se escribió nada.
            AppLogger.warn("No se pudo actualizar el Excel de medidas: " + e.getMessage());
            // POI tira varias excepciones sin mensaje, y ahí el diálogo terminaría en "null".
            String detalle = e.getMessage() != null ? e.getMessage() : e.toString();
            mostrarError("Excel de medidas",
                    "No se pudieron agregar los SKU nuevos al Excel de medidas. Si lo tenés abierto, "
                    + "cerralo y volvé a procesar el lote.\n\n" + detalle);
        }
        Platform.runLater(this::actualizarBotonSubirMedidas);
        return agregados;
    }

    private Map<String, ComboProduct> loadNormalizedCombos(String comboPath) {
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

    private Set<ShippingType> checkedShippingTypes() {
        Set<ShippingType> checked = new HashSet<>();
        if (filterFlexCheck.isSelected()) checked.add(ShippingType.FLEX);
        if (filterColectaCheck.isSelected()) checked.add(ShippingType.COLECTA);
        if (filterTurboCheck.isSelected()) checked.add(ShippingType.TURBO);
        return checked;
    }

    private boolean passesTypeFilter(OrderTableRow row) {
        return ShippingType.passes(row.getShippingType(), checkedShippingTypes());
    }

    private void setShippingFilterDisabled(boolean disabled) {
        filterFlexCheck.setDisable(disabled);
        filterColectaCheck.setDisable(disabled);
        filterTurboCheck.setDisable(disabled);
    }

    private void applyOrderFilters() {
        String filter = searchField.getText() == null ? "" : searchField.getText().trim().toLowerCase();

        Set<ShippingType> checked = checkedShippingTypes();

        if (filteredOrders != null) {
            filteredOrders.setPredicate(row -> {
                boolean matchesSearch = filter.isEmpty()
                        || row.getOrderId().toLowerCase().contains(filter)
                        || (row.getSku() != null && row.getSku().toLowerCase().contains(filter))
                        || (row.getZone() != null && row.getZone().toLowerCase().contains(filter))
                        || (row.getProductDescription() != null && row.getProductDescription().toLowerCase().contains(filter));
                boolean matchesType = ShippingType.passes(row.getShippingType(), checked);
                return matchesSearch && matchesType;
            });
        }
        // Solo refrescar el resumen de órdenes cuando su tabla está visible: en la vista de
        // etiquetas el buscador también llama a este método y no debe pisar el resumen de etiquetas.
        if (orderStatsUpdater != null && orderTable.isVisible()) orderStatsUpdater.run();
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

    /**
     * Celda de MEDIDAS/EMBALAJE. El NO va con fondo rosa pálido y texto rojo oscuro, los mismos
     * colores que la columna ERROR del Excel de medidas, para que el código de color sea el mismo
     * en la app y en la planilla.
     */
    private TableCell<LabelTableRow, EstadoDato> estadoDatoCell() {
        return new TableCell<>() {
            @Override
            protected void updateItem(EstadoDato item, boolean empty) {
                super.updateItem(item, empty);
                setAlignment(Pos.CENTER);
                if (empty || item == null) {
                    setText(null);
                    setStyle("");
                    return;
                }
                setText(item.texto());
                setStyle(switch (item) {
                    case NO -> "-fx-background-color: #FEE2E2; -fx-text-fill: #991B1B; -fx-font-weight: bold;";
                    case SI -> "-fx-text-fill: #15803d; -fx-font-weight: bold;";
                    case NO_APLICA -> "-fx-text-fill: #9ca3af;";
                });
            }
        };
    }

    private void displayResult(SortResult result, MedidasExcelManager.Medidas medidas) {
        ObservableList<LabelTableRow> rows = FXCollections.observableArrayList();
        boolean moduloEmbalajeActivo = medidas != null && medidas.embalajeEnUso();
        // Una fila por etiqueta, no por grupo: cada fila es un paquete, con su número impreso, su
        // venta y las unidades que van adentro. Agrupar obligaba a mostrar el # como rango, la
        // cantidad como suma y la orden como lista, tres cosas distintas en la misma fila.
        //
        // La agrupación se mantiene para el ZPL: es la que define el orden de impresión.
        int posicion = 1;
        for (SortedLabelGroup group : result.groups()) {
            // Sin módulo de medidas activo no hay nada que informar: las dos columnas quedan en "—".
            boolean elegible = medidas != null && EstadoDato.esSkuElegible(group.zone(), group.sku());
            ar.com.leo.etiquetas.model.MedidaSku medida = elegible ? medidas.porSku().get(group.sku()) : null;
            // La columna informa sobre el SKU, no sobre una etiqueta puntual: un grupo puede tener
            // etiquetas de una unidad y de varias, y el dato cargado es el mismo para todas.
            DatosEmbalaje embalaje = elegible
                    ? EstadoDato.embalajeDe(medida, moduloEmbalajeActivo)
                    : DatosEmbalaje.VACIO;

            // El estado es del SKU, así que se resuelve una vez para todas sus etiquetas.
            EstadoDato estandarizado = EstadoDato.estandarizadoDe(embalaje);

            for (ZplLabel label : group.labels()) {
                // El número sale del mismo contador que inyecta el #N en la etiqueta, así que la
                // fila y el papel dicen lo mismo.
                rows.add(new LabelTableRow(
                        String.valueOf(posicion++),
                        label.orderIds(),
                        group.zone(),
                        group.sku(),
                        group.productDescription(),
                        group.details(),
                        label.quantity(),
                        estandarizado));
            }
        }
        mostrarFilas(rows);

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
        // El link apunta al archivo del resultado anterior. Lo repone quien haya guardado uno.
        mostrarLinkArchivo(null);

        int tab = subTabActual();
        if (tab == SUBTAB_API) apiMostrandoOrdenes = false;
        vistasEtiquetas.put(tab, new VistaEtiquetas(result, rows, text, null));
    }

    private int subTabActual() {
        return etiquetasSubTabPane == null
                ? SUBTAB_API
                : etiquetasSubTabPane.getSelectionModel().getSelectedIndex();
    }

    private void mostrarFilas(ObservableList<LabelTableRow> filas) {
        filteredLabels = new FilteredList<>(filas, p -> true);
        SortedList<LabelTableRow> ordenadas = new SortedList<>(filteredLabels);
        ordenadas.comparatorProperty().bind(labelTable.comparatorProperty());
        labelTable.setItems(ordenadas);
        searchField.clear();
    }

    private void mostrarLinkArchivo(File archivo) {
        fileLinkBar.getChildren().clear();
        if (archivo != null) LogHelper.addFileLink(fileLinkBar, archivo);
        fileLinkBar.setVisible(archivo != null);
        fileLinkBar.setManaged(archivo != null);
    }

    /** Deja el archivo guardado registrado en la vista, para reponer el link al volver a la pestaña. */
    private void registrarArchivoGuardado(File archivo) {
        mostrarLinkArchivo(archivo);
        VistaEtiquetas vista = vistasEtiquetas.get(subTabActual());
        if (vista != null) {
            vistasEtiquetas.put(subTabActual(), new VistaEtiquetas(
                    vista.resultado(), vista.filas(), vista.estadisticas(), archivo));
        }
    }

    /**
     * Muestra lo que la pestaña tenga cargado. Si todavía no corrió nada queda vacía con su
     * placeholder, en vez de dejar a la vista el resultado del otro flujo.
     */
    private void aplicarVistaDeSubTab(int tab) {
        if (tab == SUBTAB_API && apiMostrandoOrdenes) {
            showOrderTable();
            return;
        }

        VistaEtiquetas vista = vistasEtiquetas.get(tab);
        if (vista == null) {
            currentResult = null;
            labelTable.setPlaceholder(tab == SUBTAB_LOCAL ? placeholderLocal : placeholderApi);
            // Se rearma el envoltorio igual que con datos: si no, el buscador seguiría filtrando
            // la lista de la pestaña anterior.
            mostrarFilas(FXCollections.observableArrayList());
            statsBar.setVisible(false);
            statsBar.setManaged(false);
            searchBar.setVisible(false);
            searchBar.setManaged(false);
            mostrarLinkArchivo(null);
        } else {
            currentResult = vista.resultado();
            mostrarFilas(vista.filas());
            statsLabel.setText(vista.estadisticas());
            Tooltip tip = new Tooltip(vista.estadisticas().replace(" | ", "\n"));
            tip.setShowDelay(Duration.millis(200));
            statsLabel.setTooltip(tip);
            statsBar.setVisible(true);
            statsBar.setManaged(true);
            searchBar.setVisible(true);
            searchBar.setManaged(true);
            mostrarLinkArchivo(vista.archivoGuardado());
        }
        showLabelTable();
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
