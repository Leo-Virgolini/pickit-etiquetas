package ar.com.leo;

import ar.com.leo.api.ml.MercadoLibreAPI;
import ar.com.leo.cli.PickitCli;
import ar.com.leo.pedidos.service.PedidosGenerator;
import ar.com.leo.pickit.service.PickitGenerator;
import atlantafx.base.theme.PrimerLight;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.util.Arrays;

public class EtiquetasApp extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {
        Application.setUserAgentStylesheet(new PrimerLight().getUserAgentStylesheet());

        FXMLLoader loader = new FXMLLoader(getClass().getResource("/ar/com/leo/ui/MainView.fxml"));
        Scene scene = new Scene(loader.load());

        primaryStage.setTitle("Pickit y Etiquetas");
        primaryStage.setScene(scene);
        primaryStage.getIcons().add(new Image(getClass().getResourceAsStream("/ar/com/leo/ui/icons8-productos-64.png")));
        primaryStage.setMinWidth(800);
        primaryStage.setMinHeight(600);
        primaryStage.show();
    }

    @Override
    public void stop() {
        MercadoLibreAPI.shutdownExecutors();
        PickitGenerator.shutdownExecutors();
        PedidosGenerator.shutdownExecutors();
    }

    public static void main(String[] args) {
        // Modo CLI: si recibimos --pickit-manual, no levantamos GUI ni JavaFX
        // toolkit. El control vuelve al SO con un exit code adecuado vía
        // System.exit(...). Pensado para integraciones automatizadas
        // (showroom-backend invoca el jar con ProcessBuilder).
        if (Arrays.asList(args).contains("--pickit-manual")) {
            PickitCli.run(args);
            return; // PickitCli ya llamó a System.exit, esto es defensa.
        }
        launch(args);
    }
}
