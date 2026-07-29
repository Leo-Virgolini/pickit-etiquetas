package ar.com.leo.etiquetas.model;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class PrintNumberingTest {

    @Test
    void gruposDeUnaEtiquetaNumeranSecuencial() {
        assertEquals(List.of("1", "2", "3"), PrintNumbering.compute(List.of(1, 1, 1)));
    }

    @Test
    void gruposMultiEtiquetaMuestranRango() {
        // grupo de 1 → "1"; grupo de 4 → "2–5"; grupo de 2 → "6–7"
        assertEquals(List.of("1", "2–5", "6–7"), PrintNumbering.compute(List.of(1, 4, 2)));
    }

    @Test
    void listaVaciaDevuelveVacio() {
        assertEquals(List.of(), PrintNumbering.compute(List.of()));
    }
}
