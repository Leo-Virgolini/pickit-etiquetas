package ar.com.leo.etiquetas.printer;

import ar.com.leo.etiquetas.model.ZplLabel;

import javax.print.*;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.Socket;
import java.nio.charset.StandardCharsets;
import java.util.List;

public class ZplPrinterService {

    private static final int DEFAULT_ZPL_PORT = 9100;
    public void printViaSocket(List<ZplLabel> labels, String host, int port) throws IOException {
        String zplData = buildZplString(labels);
        try (Socket socket = new Socket(host, port);
             OutputStream out = socket.getOutputStream()) {
            out.write(zplData.getBytes(StandardCharsets.UTF_8));
            out.flush();
        }
    }

    public void printViaSocket(List<ZplLabel> labels, String host) throws IOException {
        printViaSocket(labels, host, DEFAULT_ZPL_PORT);
    }

    public void printViaPrintService(List<ZplLabel> labels, PrintService printService) throws PrintException {
        String zplData = buildZplString(labels);
        DocPrintJob job = printService.createPrintJob();
        Doc doc = new SimpleDoc(
                new ByteArrayInputStream(zplData.getBytes(StandardCharsets.UTF_8)),
                DocFlavor.INPUT_STREAM.AUTOSENSE,
                null);
        job.print(doc, null);
    }

    private String buildZplString(List<ZplLabel> labels) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < labels.size(); i++) {
            sb.append(labels.get(i).rawZpl());
            if (!labels.get(i).rawZpl().endsWith("\n")) {
                sb.append("\n");
            }
        }
        return sb.toString();
    }
}
