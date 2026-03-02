# Pickit y Etiquetas

Aplicacion de escritorio JavaFX para gestionar despachos de e-commerce: generacion de listas de picking (pickit), armado de carros, ordenamiento e impresion de etiquetas ZPL, y generacion de pedidos/etiquetas de envio. Integra MercadoLibre, Tienda Nube y DUX ERP.

## Configuracion global

Dos selectores de archivo persistentes (guardados en `Preferences`) disponibles en todas las pestanas:

- **Excel de stock** (`Stock.xlsx`): mapea SKU a zona de almacen (J1, J2, T1, T2, etc.) y opcionalmente codigo externo. Lee desde fila 3 las columnas "Codigo Producto", "Unidad" y "Codigo Externo".
- **Excel de combos**: define productos compuestos y sus componentes. Columnas: "Codigo Compuesto", "Codigo Componente", "Cantidad".

## Funcionalidades

### Pickit

Genera un Excel de picking para el deposito con todos los pedidos pendientes de todos los canales.

- **Fuentes de datos**:
  - **ML ready_to_print**: `/orders/search` con `shipping.status=ready_to_ship`, `shipping.substatus=ready_to_print`. Excluye ordenes con tag `delivered`.
  - **ML acuerdo (seller_agreement)**: `/orders/search` con `tags=no_shipping`, `order.status=paid`, ultimos 7 dias. Excluye ordenes entregadas, cumplidas (`fulfilled`) y con notas (`/orders/{id}/notes`).
  - **TN HOGAR / TN GASTRO**: `/v1/{storeId}/orders` con `payment_status=paid`, `shipping_status=unpacked`, `status=open`. Excluye ordenes pickup con nota del vendedor.
  - **Productos manuales**: ingreso directo de SKU + cantidad.
- **Filtro Despacho ML**: "Hasta hoy" (solo ordenes ML con SLA para hoy o antes) o "Sin limite" (todas las pendientes). Aplica solo a MercadoLibre; las ventas de Tienda Nube se incluyen siempre.
- **Expansion de combos**: los SKU compuestos se expanden automaticamente en sus componentes con cantidades multiplicadas.
- **Consulta de stock**: busca descripcion, proveedor, sector y stock actual en DUX ERP.
- **Excel generado** (`Pickits y Carros/PICKIT_*.xlsx`) con 3 hojas:
  - **PICKIT**: lista ordenada por sector con columnas SKU, CANT, DESCRIPCION, PROVEEDOR, SECTOR, STOCK. Resaltado: coral=SKU invalido, amarillo=datos faltantes, bold=cantidad>1, naranja=stock insuficiente.
  - **CARROS**: ordenes con 3+ SKUs distintos agrupadas con letra (A, B, C...) para identificar carros fisicos.
  - **SLA**: listado de ordenes ML con fecha/hora de despacho esperado.
- **Resumen en log**: al finalizar muestra desglose por seccion (ML ready_to_print, ML acuerdo, KT HOGAR, KT GASTRO, Manuales) con conteo de ordenes y productos, SKUs OK vs problemas.

### Etiquetas ML

Dos sub-pestanas para obtener etiquetas ZPL, procesarlas y enviarlas a la impresora Zebra.

#### API MercadoLibre

1. **Obtener ordenes**: `/orders/search` con `shipping.status=ready_to_ship` y filtros por substatus (pendientes: `ready_to_print` / impresas: `printed,ready_for_dropoff,ready_for_pickup` / todas) y despacho ML (solo hoy/todas). Muestra tabla con columnas: Orden, Zona, SKU, Producto, Cantidad, Estado, Despacho.
2. **Descargar etiquetas**: `/shipment_labels?shipment_ids={ids}&response_type=zpl2` en lotes de 50. Nota: ML cambia automaticamente el substatus de `ready_to_print` a `printed` al descargar.
3. **Procesamiento automatico**: parseo ZPL, asignacion de zona, ordenamiento e inyeccion de headers.

#### Archivo Local

- Carga un archivo `.txt`/`.zpl` con etiquetas ZPL crudas y las procesa con la misma pipeline.

#### Procesamiento de etiquetas ZPL

- **Parseo**: extrae bloques `^XA...^XZ`, decodifica hex (`_XX`), extrae SKU, producto, cantidad y detalles.
- **Asignacion de zona**: cruza cada SKU contra el Excel de stock para determinar la zona de almacen.
- **Ordenamiento** por prioridad: J* > T* > COMBOS > CARROS > TURBOS > RETIROS > ??? (sin mapear).
- **Inyeccion de headers en ZPL**: agrega al codigo ZPL de cada etiqueta:
  - Numero de posicion (#1, #2...) en bold (triple render)
  - Zona ("ZONA: J5")
  - Codigo externo ("COD.EXT.: 12345")
  - Resaltado de cantidad >1 (rectangulo negro con texto inverso)
- **Interleave para impresion**: reordena las etiquetas para compensar el plegado en acordeon de la impresora termica, de modo que al cortar el stack queden en orden.
- **Impresion directa**: dialog de seleccion de zonas a imprimir + seleccion de impresora. Envia ZPL crudo via `javax.print`.
- **Combos**: muestra desglose de productos compuestos presentes en el lote para facilitar el armado.

### Pedidos

Genera un Excel con tarjetas recortables de todos los pedidos pendientes, listas para pegar en los paquetes.

- **Fuentes**:
  - **ML retiro**: `/orders/search` con `tags=no_shipping`, `order.status=paid`, ultimos 7 dias. Excluye ordenes entregadas, cumplidas (`fulfilled`) y con notas. Obtiene nickname del comprador via `/users/{buyerId}` en paralelo (la API de ordenes ya no devuelve nombre/apellido).
  - **TN HOGAR / TN GASTRO**: `/v1/{storeId}/orders` con `payment_status=paid`, `shipping_status=unpacked`, `status=open`. Excluye ordenes pickup con nota del vendedor. Genera etiquetas LLEGA HOY para envios que contengan "LLEGA HOY" en el nombre (excepto Zippin).
- **Excel generado** (`Pedidos/PEDIDOS_*.xlsx`) con hasta 3 hojas:

#### ML PEDIDOS RETIRO (violeta)
- Tarjetas recortables con borde grueso, 2 columnas por pagina.
- Cada tarjeta: N de venta (grande), fecha, nickname del comprador, tabla de productos (SKU, CANT, DETALLE).
- Productos con cantidad >1 resaltados en amarillo.
- Altura dinamica: la tarjeta crece segun la cantidad de productos.

#### TN PEDIDOS (verde)
- Mismo layout de tarjetas recortables.
- Badge de tienda (KT HOGAR / KT GASTRO) con tipo de envio simplificado (RETIRO, CABA - LLEGA HOY, etc.; se omite el detalle entre parentesis).

#### TN ETIQUETAS (naranja)
- Etiquetas de envio para pedidos "LLEGA HOY" (no Zippin), 10 por pagina.
- Cada etiqueta: nombre grande, domicilio, localidad, CP, telefono, observaciones.

Todas las hojas: A4 portrait, margenes estrechos, page breaks inteligentes basados en altura, numeracion de pagina.

- **Resumen en log**: al finalizar muestra desglose por seccion (ML retiro, KT HOGAR, KT GASTRO, etiquetas LLEGA HOY) con conteo de ordenes.

## Integraciones

| Servicio | Uso | Credenciales |
|---|---|---|
| **MercadoLibre** | Ordenes ME2, etiquetas ZPL, SLA | `ml_credentials.json`, `ml_tokens.json` |
| **Tienda Nube** | Pedidos KT HOGAR y KT GASTRO | `nube_tokens.json` |
| **DUX ERP** | Stock, descripciones, proveedores | `dux_tokens.json` |

Credenciales almacenadas en `%PROGRAMDATA%\SuperMaster\secrets\`. Tokens ML se renuevan automaticamente al expirar.

`HttpRetryHandler` implementa: rate limiting (Guava `RateLimiter`), refresh automatico de token en 401, backoff exponencial con jitter en 429/503/5xx.

## Requisitos

- Java 25+
- Maven 3.9+
- Impresora Zebra (opcional, para impresion directa)

## Compilacion y ejecucion

```bash
mvn clean compile
mvn javafx:run
```

## Generar JAR

```bash
mvn clean package
# Genera: target/Pickit y Etiquetas.jar
```

## Estructura del proyecto

```
src/main/java/ar/com/leo/
├── api/
│   ├── ml/              # Integracion con API de MercadoLibre
│   ├── nube/            # Integracion con Tienda Nube
│   └── dux/             # Integracion con DUX ERP (stock)
├── pickit/
│   ├── api/             # API ML especifica del pickit
│   ├── excel/           # Lectura de stock/combos y escritura del Excel pickit
│   ├── model/           # Modelos: PickitItem, CarrosOrden, Venta, etc.
│   └── service/         # PickitGenerator y PickitService
├── pedidos/
│   ├── api/             # API ML especifica de pedidos
│   ├── excel/           # PedidosExcelWriter (tarjetas recortables)
│   ├── model/           # PedidoML, PedidoTN, EtiquetaTN, PedidosResult
│   └── service/         # PedidosGenerator y PedidosService
├── model/               # Records: ZplLabel, ExcelMapping, ComboProduct, etc.
├── parser/              # Parseo de archivos ZPL y Excel
├── printer/             # Descubrimiento de impresoras y envio ZPL
├── sorter/              # Ordenamiento de etiquetas por zona
├── ui/                  # Controladores JavaFX y dialogos
├── util/                # Utilidades
├── AppLogger.java
└── EtiquetasApp.java
```

## Tecnologias

- **JavaFX 25** + AtlantaFX (tema PrimerLight)
- **Apache POI 5.5** - Lectura y escritura de archivos Excel
- **Jackson 3** - Procesamiento JSON (APIs)
- **Guava** - RateLimiter para llamadas a las APIs
