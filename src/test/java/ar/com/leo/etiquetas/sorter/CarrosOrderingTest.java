package ar.com.leo.etiquetas.sorter;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class CarrosOrderingTest {

    @Test
    void distinctSkuCountCuentaCodigosDistintos() {
        assertEquals(2, CarrosOrdering.distinctSkuCount("10\n20\n10"));
        assertEquals(3, CarrosOrdering.distinctSkuCount("10\n20\n30"));
        assertEquals(1, CarrosOrdering.distinctSkuCount("10"));
        assertEquals(0, CarrosOrdering.distinctSkuCount(""));
        assertEquals(0, CarrosOrdering.distinctSkuCount(null));
    }

    @Test
    void bucketDevuelve0ParaZonasNoCarros() {
        assertEquals(0, CarrosOrdering.bucket("J1", "10\n20\n30"));
        assertEquals(0, CarrosOrdering.bucket("TURBOS", "10"));
    }

    @Test
    void bucketSeparaCarrosPorCantidadDeSkus() {
        assertEquals(0, CarrosOrdering.bucket("CARROS", "10\n20"));      // 2 distintos
        assertEquals(0, CarrosOrdering.bucket("CARROS", "10\n20\n10"));  // 2 distintos (con dup)
        assertEquals(1, CarrosOrdering.bucket("CARROS", "10\n20\n30"));  // 3 distintos
        assertEquals(1, CarrosOrdering.bucket("CARROS", "10\n20\n30\n40"));
    }
}
