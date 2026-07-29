package ar.com.leo.etiquetas.model;

public record ZplLabel(String rawZpl, String sku, String productDescription, String details, int quantity, boolean turbo, String orderIds) {

    public ZplLabel(String rawZpl, String sku, String productDescription, String details) {
        this(rawZpl, sku, productDescription, details, 1, false, "");
    }

    public ZplLabel(String rawZpl, String sku, String productDescription, String details, int quantity) {
        this(rawZpl, sku, productDescription, details, quantity, false, "");
    }

    public ZplLabel(String rawZpl, String sku, String productDescription, String details, int quantity, boolean turbo) {
        this(rawZpl, sku, productDescription, details, quantity, turbo, "");
    }
}
