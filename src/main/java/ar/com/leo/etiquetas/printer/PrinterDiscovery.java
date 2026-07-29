package ar.com.leo.etiquetas.printer;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import java.util.Arrays;
import java.util.List;

public class PrinterDiscovery {

    public List<PrintService> findAll() {
        return Arrays.asList(PrintServiceLookup.lookupPrintServices(null, null));
    }

    public List<PrintService> findZebraPrinters() {
        return findAll().stream()
                .filter(ps -> {
                    String name = ps.getName().toLowerCase();
                    return name.contains("zebra") || name.contains("zpl") || name.contains("thermal");
                })
                .toList();
    }

    public PrintService getDefault() {
        return PrintServiceLookup.lookupDefaultPrintService();
    }
}
