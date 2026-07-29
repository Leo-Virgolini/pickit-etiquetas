package ar.com.leo.etiquetas.parser;

import java.util.List;

public record ComboProduct(String codigoCompuesto, String productoCompuesto, List<ComboComponent> componentes) {
}
