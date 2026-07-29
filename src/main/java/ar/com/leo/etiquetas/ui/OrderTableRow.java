package ar.com.leo.etiquetas.ui;

import ar.com.leo.api.ml.model.OrdenML;
import ar.com.leo.api.ml.model.ShippingType;
import javafx.beans.property.*;

import java.util.List;

public class OrderTableRow {

    private final BooleanProperty selected;
    private final StringProperty orderId;
    private final StringProperty zone;
    private final StringProperty sku;
    private final StringProperty productDescription;
    private final StringProperty quantity;
    private final StringProperty status;
    private final StringProperty slaDate;
    private final List<OrdenML> ordenes;
    private final ShippingType shippingType;

    public OrderTableRow(boolean selected, String orderId, String zone, String sku, String productDescription,
                         String quantity, String status, String slaDate, List<OrdenML> ordenes,
                         ShippingType shippingType) {
        this.selected = new SimpleBooleanProperty(selected);
        this.orderId = new SimpleStringProperty(orderId);
        this.zone = new SimpleStringProperty(zone);
        this.sku = new SimpleStringProperty(sku);
        this.productDescription = new SimpleStringProperty(productDescription);
        this.quantity = new SimpleStringProperty(quantity);
        this.status = new SimpleStringProperty(status);
        this.slaDate = new SimpleStringProperty(slaDate);
        this.ordenes = ordenes;
        this.shippingType = shippingType;
    }

    public BooleanProperty selectedProperty() { return selected; }
    public StringProperty orderIdProperty() { return orderId; }
    public StringProperty zoneProperty() { return zone; }
    public StringProperty skuProperty() { return sku; }
    public StringProperty productDescriptionProperty() { return productDescription; }
    public StringProperty quantityProperty() { return quantity; }
    public StringProperty statusProperty() { return status; }
    public StringProperty slaDateProperty() { return slaDate; }

    public boolean isSelected() { return selected.get(); }
    public void setSelected(boolean value) { selected.set(value); }
    public String getOrderId() { return orderId.get(); }
    public String getZone() { return zone.get(); }
    public String getSku() { return sku.get(); }
    public String getProductDescription() { return productDescription.get(); }
    public String getQuantity() { return quantity.get(); }
    public String getStatus() { return status.get(); }
    public String getSlaDate() { return slaDate.get(); }
    public List<OrdenML> getOrdenes() { return ordenes; }
    public ShippingType getShippingType() { return shippingType; }

    public int getProductCount() {
        int count = 0;
        for (OrdenML o : ordenes) {
            for (var v : o.getItems()) {
                count += (int) v.getCantidad();
            }
        }
        return count;
    }
}
