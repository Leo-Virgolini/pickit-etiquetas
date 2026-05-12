package ar.com.leo.cli;

import ar.com.leo.AppLogger;
import ar.com.leo.pickit.excel.ExcelManager;
import ar.com.leo.pickit.model.ProductoManual;
import ar.com.leo.pickit.service.PickitGenerator;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Modo CLI del programa para integraciones automatizadas.
 *
 * <p>Permite generar el Excel Pickit a partir de SKUs+cantidades sin
 * levantar la GUI JavaFX. Pensado para que otros sistemas (ej. showroom-backend)
 * disparen la generación tras crear un pedido.
 *
 * <p>Args esperados:
 * <pre>
 *   --pickit-manual           (modo, requerido)
 *   --input &lt;skus.xlsx&gt;       Excel con headers SKU/CANTIDAD (requerido)
 *   --stock &lt;Stock.xlsx&gt;      mapping SKU→sector/codigoExterno (requerido)
 *   --combos &lt;Combos.xlsx&gt;    expansión de productos compuestos (requerido)
 *   --output-dir &lt;dir&gt;        carpeta de salida; si se omite, usa la default
 *                              ({@code Pickits y Carros} adyacente al jar)
 * </pre>
 *
 * <p>Al terminar, imprime al stdout una sola línea con el path absoluto del
 * Excel generado. Exit 0 si OK, exit 1 si hubo error (el detalle va a stderr).
 *
 * <p>Las APIs externas (MercadoLibre, TiendaNube, DUX) no se tocan en este
 * modo — solo se procesan los items del Excel manual ({@code useManual=true},
 * los demás flags en {@code false}).
 */
public final class PickitCli {

    private PickitCli() {}

    public static void run(String[] args) {
        try {
            Map<String, String> opts = parseArgs(args);
            File input = requerirArchivo(opts, "--input", "Excel de SKUs");
            File stock = requerirArchivo(opts, "--stock", "Excel de stock");
            File combos = requerirArchivo(opts, "--combos", "Excel de combos");
            File outputDir = opts.containsKey("--output-dir") ? new File(opts.get("--output-dir")) : null;

            AppLogger.info("CLI - Leyendo SKUs desde " + input.getAbsolutePath());
            List<ProductoManual> productos = ExcelManager.obtenerProductosManualesDesdeExcel(input);
            if (productos.isEmpty()) {
                System.err.println("ERROR: el Excel de SKUs no contiene productos válidos.");
                System.exit(1);
            }

            AppLogger.info("CLI - Generando pickit para " + productos.size() + " SKUs...");
            File resultado = PickitGenerator.generarPickit(
                    stock, combos, productos,
                    false,  // soloHoy
                    false,  // soloTurbo
                    false,  // useML
                    false,  // useNube
                    true,   // useManual
                    outputDir);

            // El caller (showroom-backend) lee el path desde stdout.
            System.out.println(resultado.getAbsolutePath());
            System.exit(0);
        } catch (IllegalArgumentException e) {
            System.err.println("ERROR: " + e.getMessage());
            System.exit(2);
        } catch (Exception e) {
            System.err.println("ERROR generando pickit: " + e.getMessage());
            e.printStackTrace(System.err);
            System.exit(1);
        } finally {
            PickitGenerator.shutdownExecutors();
        }
    }

    private static Map<String, String> parseArgs(String[] args) {
        Map<String, String> opts = new HashMap<>();
        for (int i = 0; i < args.length; i++) {
            String a = args[i];
            if (a.startsWith("--")) {
                if (i + 1 < args.length && !args[i + 1].startsWith("--")) {
                    opts.put(a, args[i + 1]);
                    i++;
                } else {
                    // flag sin valor (ej. --pickit-manual)
                    opts.put(a, "");
                }
            }
        }
        return opts;
    }

    private static File requerirArchivo(Map<String, String> opts, String flag, String descripcion) {
        String path = opts.get(flag);
        if (path == null || path.isBlank()) {
            throw new IllegalArgumentException(descripcion + " requerido — pasar con " + flag + " <path>");
        }
        File f = new File(path);
        if (!f.exists() || !f.isFile()) {
            throw new IllegalArgumentException(descripcion + " no existe o no es un archivo: " + path);
        }
        return f;
    }
}
