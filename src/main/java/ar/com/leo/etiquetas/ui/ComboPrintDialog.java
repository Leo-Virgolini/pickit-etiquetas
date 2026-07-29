package ar.com.leo.etiquetas.ui;

import ar.com.leo.etiquetas.parser.ComboComponent;
import ar.com.leo.etiquetas.parser.ComboProduct;
import ar.com.leo.ui.AlertHelper;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.print.PageLayout;
import javafx.print.PageOrientation;
import javafx.print.Paper;
import javafx.print.PrinterJob;
import javafx.scene.Group;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.*;
import javafx.scene.transform.Scale;
import javafx.scene.image.Image;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class ComboPrintDialog {

    private static final String CELL_BORDER = "-fx-border-color: #999; -fx-border-width: 0 1 1 0;";
    private static final String HEADER_STYLE = "-fx-font-weight: bold; -fx-font-size: 12px; -fx-padding: 6 8; "
            + "-fx-background-color: #d0d0d0; -fx-border-color: #999; -fx-border-width: 0 1 1 0;";
    private static final String COMBO_HEADER_STYLE = "-fx-font-weight: bold; -fx-font-size: 12px; -fx-padding: 6 8; "
            + "-fx-background-color: #e8e8e8; -fx-border-color: #999; -fx-border-width: 0 1 1 0;";
    private static final String CELL_STYLE = "-fx-font-size: 12px; -fx-padding: 4 8; " + CELL_BORDER;
    private static final String FOOTER_STYLE = "-fx-font-size: 10px; -fx-text-fill: #666;";
    private static final double COMBO_SPACING = 8;
    private static final double FOOTER_HEIGHT = 20;
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");

    private final Stage stage;
    private final List<ComboProduct> combos;
    private final String titleText;

    public ComboPrintDialog(Window owner, List<ComboProduct> combos) {
        this.combos = combos;
        this.titleText = "Composición de Combos  -  " + LocalDateTime.now().format(DATE_FMT);
        stage = new Stage();
        stage.initModality(Modality.WINDOW_MODAL);
        stage.initOwner(owner);
        stage.setTitle("Composición de Combos");
        stage.getIcons().add(new Image(getClass().getResourceAsStream("/ar/com/leo/ui/icons8-etiqueta-100.png")));

        VBox displayContent = buildContent(combos);

        ScrollPane scrollPane = new ScrollPane(displayContent);
        scrollPane.setFitToWidth(true);
        scrollPane.setPadding(new Insets(10));

        Button printBtn = new Button("\uD83D\uDDA8 Imprimir");
        printBtn.setStyle("-fx-font-size: 14px; -fx-padding: 8 20;");
        printBtn.setOnAction(e -> print());

        Button closeBtn = new Button("Cerrar");
        closeBtn.setStyle("-fx-font-size: 14px; -fx-padding: 8 20;");
        closeBtn.setOnAction(e -> stage.close());

        HBox buttons = new HBox(15, printBtn, closeBtn);
        buttons.setAlignment(Pos.CENTER_RIGHT);
        buttons.setPadding(new Insets(10, 15, 10, 15));

        BorderPane root = new BorderPane();
        root.setCenter(scrollPane);
        root.setBottom(buttons);

        Scene scene = new Scene(root, 950, 800);
        stage.setScene(scene);
    }

    public void show() {
        stage.show();
    }

    private Label makeLabel(String text, String style, boolean forPrint) {
        Label label = new Label(text);
        label.setStyle(style);
        label.setMaxHeight(Double.MAX_VALUE);
        if (forPrint) {
            label.setWrapText(true);
        } else {
            label.setTextOverrun(OverrunStyle.CLIP);
            label.setMinWidth(Region.USE_PREF_SIZE);
        }
        return label;
    }

    private GridPane buildComboGrid(ComboProduct combo) {
        return buildComboGrid(combo, false);
    }

    /**
     * Construye una grilla individual para un combo (para vista en pantalla).
     */
    private GridPane buildComboGrid(ComboProduct combo, boolean forPrint) {
        GridPane grid = new GridPane();
        grid.setStyle("-fx-border-color: #999; -fx-border-width: 1 0 0 1;");

        ColumnConstraints col0 = new ColumnConstraints();
        col0.setMinWidth(60);
        col0.setHgrow(Priority.SOMETIMES);
        ColumnConstraints col1 = new ColumnConstraints();
        col1.setHgrow(Priority.ALWAYS);
        ColumnConstraints col2 = new ColumnConstraints();
        col2.setPrefWidth(45);
        col2.setMinWidth(35);
        grid.getColumnConstraints().addAll(col0, col1, col2);

        Label comboLabel = makeLabel(combo.codigoCompuesto() + "  -  " + combo.productoCompuesto(), COMBO_HEADER_STYLE, false);
        comboLabel.setMaxWidth(Double.MAX_VALUE);
        grid.add(comboLabel, 0, 0, 3, 1);

        grid.add(headerCell("Código"), 0, 1);
        grid.add(headerCell("Producto"), 1, 1);
        grid.add(headerCell("Cant."), 2, 1);

        int row = 2;
        for (ComboComponent comp : combo.componentes()) {
            grid.add(makeLabel(comp.codigoComponente(), CELL_STYLE, false), 0, row);
            grid.add(makeLabel(comp.productoComponente(), CELL_STYLE, false), 1, row);
            grid.add(makeLabel(String.valueOf(comp.cantidad()), CELL_STYLE + "-fx-font-weight: bold; -fx-alignment: center;", false), 2, row);
            row++;
        }

        return grid;
    }

    /**
     * Construye un bloque para impresión: header en negrita + líneas simples por componente.
     */
    private VBox buildComboPrintBlock(ComboProduct combo) {
        VBox block = new VBox(2);
        block.setStyle("-fx-border-color: #999; -fx-border-width: 0 0 1 0; -fx-padding: 4 0 6 0;");

        Label header = new Label(combo.codigoCompuesto() + "  -  " + combo.productoCompuesto());
        header.setStyle("-fx-font-weight: bold; -fx-font-size: 12px; -fx-background-color: #e8e8e8; -fx-padding: 4 8;");
        header.setMaxWidth(Double.MAX_VALUE);
        header.setWrapText(true);
        block.getChildren().add(header);

        for (ComboComponent comp : combo.componentes()) {
            Label line = new Label("    " + comp.codigoComponente() + "  -  " + comp.productoComponente() + "    x" + comp.cantidad());
            line.setStyle("-fx-font-size: 11px; -fx-border-color: #ccc; -fx-border-width: 0 0 1 0; -fx-padding: 2 0;");
            line.setWrapText(true);
            block.getChildren().add(line);
        }

        return block;
    }

    /**
     * Construye el contenido con todos los combos en una sola columna (para vista en pantalla).
     */
    private VBox buildContent(List<ComboProduct> comboList) {
        VBox content = new VBox(10);
        content.setPadding(new Insets(15));
        content.setStyle("-fx-background-color: white;");

        Label title = new Label(titleText);
        title.setStyle("-fx-font-size: 18px; -fx-font-weight: bold; -fx-padding: 0 0 8 0;");
        title.setMinWidth(Region.USE_PREF_SIZE);
        content.getChildren().add(title);

        VBox column = new VBox(COMBO_SPACING);
        for (ComboProduct combo : comboList) {
            column.getChildren().add(buildComboGrid(combo));
        }
        content.getChildren().add(column);

        return content;
    }

    private Label headerCell(String text) {
        Label label = makeLabel(text, HEADER_STYLE, false);
        label.setMaxWidth(Double.MAX_VALUE);
        return label;
    }

    /**
     * Mide la altura de un nodo renderizándolo en una Scene temporal.
     */
    private double measureNodeHeight(Node node, double width) {
        Group group = new Group(node);
        new Scene(group);
        if (node instanceof Region region) {
            region.setPrefWidth(width);
        }
        node.applyCss();
        if (node instanceof Region region) {
            region.layout();
            double h = region.prefHeight(width);
            group.getChildren().remove(node);
            return h;
        }
        group.getChildren().remove(node);
        return 0;
    }

    /**
     * Construye una página con combos y footer de numeración.
     */
    private VBox buildPage(List<ComboProduct> pageCombos, boolean firstPage, int pageNum, int totalPages, double width) {
        VBox page = new VBox(COMBO_SPACING);
        page.setPadding(new Insets(10, 15, 5, 15));
        page.setStyle("-fx-background-color: white;");
        page.setPrefWidth(width);

        if (firstPage) {
            Label title = new Label(titleText);
            title.setStyle("-fx-font-size: 16px; -fx-font-weight: bold; -fx-padding: 0 0 6 0;");
            page.getChildren().add(title);
        }

        for (ComboProduct combo : pageCombos) {
            page.getChildren().add(buildComboPrintBlock(combo));
        }

        // Spacer para empujar el footer al fondo
        Region spacer = new Region();
        VBox.setVgrow(spacer, Priority.ALWAYS);
        page.getChildren().add(spacer);

        // Footer con número de página
        Label footer = new Label("Página " + pageNum + " de " + totalPages);
        footer.setStyle(FOOTER_STYLE);
        footer.setMaxWidth(Double.MAX_VALUE);
        footer.setAlignment(Pos.CENTER);
        page.getChildren().add(footer);

        return page;
    }

    private void print() {
        PrinterJob job = PrinterJob.createPrinterJob();
        if (job == null) {
            AlertHelper.showError("Error", "No se pudo crear el trabajo de impresión.");
            return;
        }

        boolean proceed = job.showPrintDialog(stage);
        if (!proceed) {
            job.endJob();
            return;
        }

        PageLayout pageLayout = job.getPrinter().createPageLayout(
                Paper.A4, PageOrientation.PORTRAIT,
                javafx.print.Printer.MarginType.HARDWARE_MINIMUM);
        job.getJobSettings().setPageLayout(pageLayout);

        double printableWidth = pageLayout.getPrintableWidth();
        double printableHeight = pageLayout.getPrintableHeight();

        // Medir altura del título
        Label titleMeasure = new Label(titleText);
        titleMeasure.setStyle("-fx-font-size: 16px; -fx-font-weight: bold; -fx-padding: 0 0 6 0;");
        double titleHeight = measureNodeHeight(titleMeasure, printableWidth) + COMBO_SPACING;

        // Medir altura de cada combo individualmente (usando formato de impresión)
        double contentWidth = printableWidth - 30; // 15+15 padding
        List<Double> comboHeights = new ArrayList<>();
        for (ComboProduct combo : combos) {
            VBox block = buildComboPrintBlock(combo);
            double h = measureNodeHeight(block, contentWidth);
            comboHeights.add(h);
        }

        // Espacio disponible por página (descontando padding y footer)
        double pageTopBottomPadding = 15; // 10 top + 5 bottom
        double availableFirst = printableHeight - pageTopBottomPadding - titleHeight - FOOTER_HEIGHT;
        double availableOther = printableHeight - pageTopBottomPadding - FOOTER_HEIGHT;

        // Distribuir combos en páginas
        List<List<ComboProduct>> pages = new ArrayList<>();
        List<ComboProduct> currentPage = new ArrayList<>();
        double currentHeight = 0;
        double available = availableFirst;

        for (int i = 0; i < combos.size(); i++) {
            double comboH = comboHeights.get(i) + COMBO_SPACING;

            if (!currentPage.isEmpty() && currentHeight + comboH > available) {
                // Página llena, iniciar nueva
                pages.add(currentPage);
                currentPage = new ArrayList<>();
                currentHeight = 0;
                available = availableOther;
            }

            currentPage.add(combos.get(i));
            currentHeight += comboH;
        }
        if (!currentPage.isEmpty()) {
            pages.add(currentPage);
        }

        int totalPages = pages.size();

        // Imprimir cada página
        boolean allOk = true;
        for (int p = 0; p < totalPages; p++) {
            VBox pageContent = buildPage(pages.get(p), p == 0, p + 1, totalPages, printableWidth);

            Group root = new Group(pageContent);
            new Scene(root);
            pageContent.applyCss();
            pageContent.layout();

            // Verificar si el ancho real excede el imprimible y escalar si es necesario
            double actualWidth = pageContent.prefWidth(-1);
            double scaleX = actualWidth > printableWidth ? printableWidth / actualWidth : 1.0;

            // Usar la altura de la página imprimible para que el footer quede abajo
            pageContent.setPrefHeight(printableHeight / scaleX);
            pageContent.layout();

            Scale scaleTransform = null;
            if (scaleX < 1.0) {
                scaleTransform = new Scale(scaleX, scaleX);
                pageContent.getTransforms().add(scaleTransform);
            }

            boolean printed = job.printPage(pageLayout, pageContent);

            if (scaleTransform != null) {
                pageContent.getTransforms().remove(scaleTransform);
            }

            if (!printed) {
                allOk = false;
                break;
            }
        }

        if (allOk) {
            job.endJob();
        }
    }
}
