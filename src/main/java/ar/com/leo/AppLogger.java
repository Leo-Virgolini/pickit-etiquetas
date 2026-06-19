package ar.com.leo;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.function.Consumer;

public final class AppLogger {

    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

    // Logging a archivo/consola vía log4j2 (config en log4j2.xml: rotación por tamaño/fecha).
    private static final Logger LOG = LogManager.getLogger("app");

    // Logger de la UI (panel de Pickit/Pedidos). El flujo de Etiquetas no tiene panel,
    // pero todo queda persistido en logs/app.log igualmente.
    private static volatile Consumer<String> uiLogger;

    private AppLogger() {
    }

    public static void setUiLogger(Consumer<String> logger) {
        uiLogger = logger;
    }

    public static void info(String message) {
        sendToUi("[" + LocalTime.now().format(TIME_FMT) + "] " + message);
        LOG.info(message);
    }

    public static void success(String message) {
        sendToUi("[" + LocalTime.now().format(TIME_FMT) + "] [OK] " + message);
        LOG.info("[OK] {}", message);
    }

    public static void warn(String message) {
        sendToUi("[" + LocalTime.now().format(TIME_FMT) + "] [WARN] " + message);
        LOG.warn(message);
    }

    public static void error(String message, Throwable t) {
        sendToUi("[" + LocalTime.now().format(TIME_FMT) + "] [ERROR] " + message);
        if (t != null) {
            sendToUi("[" + LocalTime.now().format(TIME_FMT) + "] [ERROR] " + t.getMessage());
            LOG.error(message, t);
        } else {
            LOG.error(message);
        }
    }

    private static void sendToUi(String timestamped) {
        Consumer<String> logger = uiLogger;
        if (logger != null) {
            logger.accept(timestamped);
        }
    }
}
