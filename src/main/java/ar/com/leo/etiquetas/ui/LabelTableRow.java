package ar.com.leo.etiquetas.ui;

import javafx.beans.property.*;

public class LabelTableRow {

    private final StringProperty printNumber;
    private final StringProperty orderIds;
    private final StringProperty zone;
    private final StringProperty sku;
    private final StringProperty productDescription;
    private final StringProperty details;
    private final IntegerProperty quantity;

    public LabelTableRow(String printNumber, String orderIds, String zone, String sku, String productDescription, String details, int quantity) {
        this.printNumber = new SimpleStringProperty(printNumber);
        this.orderIds = new SimpleStringProperty(orderIds);
        this.zone = new SimpleStringProperty(zone);
        this.sku = new SimpleStringProperty(sku);
        this.productDescription = new SimpleStringProperty(productDescription);
        this.details = new SimpleStringProperty(details);
        this.quantity = new SimpleIntegerProperty(quantity);
    }

    public StringProperty printNumberProperty() { return printNumber; }
    public StringProperty orderIdsProperty() { return orderIds; }
    public StringProperty zoneProperty() { return zone; }
    public StringProperty skuProperty() { return sku; }
    public StringProperty productDescriptionProperty() { return productDescription; }
    public StringProperty detailsProperty() { return details; }
    public IntegerProperty quantityProperty() { return quantity; }

    public String getPrintNumber() { return printNumber.get(); }
    public String getOrderIds() { return orderIds.get(); }
    public String getZone() { return zone.get(); }
    public String getSku() { return sku.get(); }
    public String getProductDescription() { return productDescription.get(); }
    public String getDetails() { return details.get(); }
    public int getQuantity() { return quantity.get(); }
}
