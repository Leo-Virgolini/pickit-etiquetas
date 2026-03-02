package ar.com.leo.pedidos.model;

import java.time.OffsetDateTime;

public record PedidoML(long orderId, OffsetDateTime fecha, String usuario,
                       String nombreApellido, String sku, double cantidad, String detalle, long buyerId) {}
