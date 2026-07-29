package ar.com.leo.api.ml.model;

import org.junit.jupiter.api.Test;

import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ShippingTypeTest {

    @Test
    void fromClasificaSegunTurboYLogisticType() {
        assertEquals(ShippingType.TURBO, ShippingType.from(true, "self_service"));
        assertEquals(ShippingType.TURBO, ShippingType.from(true, "cross_docking"));
        assertEquals(ShippingType.FLEX, ShippingType.from(false, "self_service"));
        assertEquals(ShippingType.COLECTA, ShippingType.from(false, "cross_docking"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "xd_drop_off"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "drop_off"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, "fulfillment"));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, ""));
        assertEquals(ShippingType.OTRO, ShippingType.from(false, null));
    }

    @Test
    void passesSinFiltroDejaPasarTodo() {
        assertTrue(ShippingType.passes(ShippingType.OTRO, Set.of()));
        assertTrue(ShippingType.passes(ShippingType.FLEX, Set.of()));
    }

    @Test
    void passesConFiltroSoloDejaPasarLosTildados() {
        Set<ShippingType> checked = Set.of(ShippingType.FLEX, ShippingType.TURBO);
        assertTrue(ShippingType.passes(ShippingType.FLEX, checked));
        assertTrue(ShippingType.passes(ShippingType.TURBO, checked));
        assertFalse(ShippingType.passes(ShippingType.COLECTA, checked));
        assertFalse(ShippingType.passes(ShippingType.OTRO, checked));
    }

    @Test
    void labelDevuelveIconoYTexto() {
        assertEquals("📦 Flex", ShippingType.FLEX.label());
        assertEquals("🚚 Colecta", ShippingType.COLECTA.label());
        assertEquals("⚡ Turbo", ShippingType.TURBO.label());
        assertEquals("❓ Otro", ShippingType.OTRO.label());
    }
}
