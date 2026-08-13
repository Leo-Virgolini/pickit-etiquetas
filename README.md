# Pickit y Etiquetas

Aplicacion de escritorio JavaFX para gestionar despachos de e-commerce: generacion de listas de picking (pickit), armado de carros, ordenamiento e impresion de etiquetas ZPL, y generacion de pedidos/etiquetas de envio. Integra MercadoLibre, Tienda Nube y DUX ERP.

## Configuracion global

Tres selectores de archivo persistentes (guardados en `Preferences`) disponibles en todas las pestañas:

- **Excel de stock** (`Stock.xlsx`): mapea SKU a zona de almacen (J1, J2, T1, T2, etc.) y opcionalmente codigo externo. Lee desde fila 3 las columnas "Codigo Producto", "Unidad" y "Codigo Externo".
- **Excel de combos**: define productos compuestos y sus componentes. Columnas: "Codigo Compuesto", "Codigo Componente", "Cantidad".
- **Excel de medidas ML** (opcional, se activa con checkbox): base "madre" de medidas de embalaje por SKU, usada para marcar etiquetas pendientes de medir y para cargar las dimensiones de paquete en ML (atributos `SELLER_PACKAGE_*`) y para indicar en la etiqueta con que embalar. 17 columnas en fila 1:

  | # | Columna | Uso |
  |---|---|---|
  | 0 | `SKU` | Clave. |
  | 1 | `PRODUCTO` | Descripcion del producto. La app no escribe esta columna: usualmente se delega a una formula del usuario (ej: `=BUSCARX(A2;Hoja2!A:A;Hoja2!B:B)`) que la resuelve automaticamente al escribir el SKU. |
  | 2 | `Ancho cm` | Medida real (base). |
  | 3 | `Alto cm` | Medida real (base). |
  | 4 | `Profundidad cm` | Medida real (base). |
  | 5 | `Peso físico (empaque + producto) kg` | Medida real (base). |
  | 6 | `Ancho +20%` | Valor que se sube a ML. |
  | 7 | `Alto +20%` | Valor que se sube a ML. |
  | 8 | `Profunidad +20%` | Valor que se sube a ML. |
  | 9 | `Peso físico (empaque + producto) +20%` | Valor que se sube a ML. |
  | 10 | `SUBIDO` | `NO` al agregar (rojo tenue), `SI` al subir OK (verde tenue). |
  | 11 | `ESTANDARIZADO` | `SI`/`NO`, calculado por una formula del usuario: resume si completo envase, tipo de rollo y cantidad de paños. Es la unica fuente de verdad sobre si el embalaje esta cargado. |
  | 12 | `ENVASE` | Codigo del envase (`BOL-1`, `CAJ-1`), o `NO` si el producto no lleva. |
  | 13 | `TIPO DE ROLLO` | Tipo de rollo (`DIAMANTES`, `CUADRADOS`), o `NO`. |
  | 14 | `CANT PAÑOS` | Cantidad de paños; solo se imprime si es mayor a cero. |
  | 15 | `OBSERVACIONES` | Texto libre. |
  | 16 | `ERROR` | Mensaje de ML en rojo cuando falla la subida. Se limpia al pasar a `SUBIDO=SI` en un reintento exitoso. |

  - Las 4 columnas base cm/kg son los valores reales medidos por el deposito. Las `+20%` son los valores efectivos declarados a ML (margen por variaciones de armado).
  - Si el archivo no existe se crea automaticamente con headers en la primera ejecucion. Los SKUs nuevos se insertan primero en filas con SKU vacio (reutilizando slots pre-cargados con formulas) y si se agotan se appendean al final. En ambos casos las celdas de medidas faltantes quedan en amarillo y `SUBIDO=NO`. Las celdas que contengan una formula se preservan intactas.
  - El lector tolera variantes: "Largo" o "Profundidad", espacios y saltos de linea dentro del header, y el typo "Profunidad" en la columna +20%.
  - Si el archivo existente no tiene columna `ERROR`, se agrega automaticamente en la primera escritura (migracion silenciosa).
  - Las columnas de embalaje (11 a 15) **las crea y carga el usuario a mano**: la app solo las lee. Se ubican por su encabezado, no por posicion, asi que se pueden reordenar o intercalar columnas propias. Si alguna falta, ese dato no se muestra y el resto sigue funcionando. Solo se crean automaticamente cuando la app genera el archivo desde cero.
  - Con el checkbox desactivado se saltea el marcado MEDIR, las lineas de embalaje y la subida a ML.

- **Hoja `ESTANDARIZACION`** (dentro del mismo archivo de medidas): catalogo de envases. La app usa dos columnas: `N°` con el codigo (`CAJ-1`, `BOL-1`) e `INSCRIPCION` con el texto escrito en el envase fisico (`9Y`, `AYUDIN`). Las bolsas no llevan inscripcion y las filas que no la tienen traen un guion, que se ignora. Si falta la hoja o el codigo no figura, la etiqueta muestra solo el codigo.
  - Escritura serializada con lock interno y reintentos con backoff (500/1000/1500/2000 ms) si el archivo esta abierto en Excel (sharing violation).
  - Los decimales con coma ("3,006" = 3.006 kg) se leen correctamente tanto si la celda es numerica (POI devuelve el valor crudo) como si es texto (se normaliza `,` → `.`).

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

Dos sub-pestañas para obtener etiquetas ZPL, procesarlas y enviarlas a la impresora Zebra.

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
- **Marcado MEDIR y autocarga al Excel** (durante la descarga/procesamiento de etiquetas): si esta configurado el Excel de medidas:
  1. **Banner MEDIR en la etiqueta**: cada etiqueta individual (no CARROS) con SKU numerico, **de pedido de 1 unidad**, cuyo SKU no tenga las 4 columnas base cm/kg cargadas, recibe un banner "MEDIR: [SKU]" en negro invertido sobre el encabezado. Las ordenes de 2+ unidades no se marcan (esos embalajes se miden aparte).
  2. **Autocarga al Excel**: los SKU detectados como pendientes se insertan en el Excel con SUBIDO=NO. El inserter primero **reusa filas pre-existentes con SKU vacio** (tipicamente filas con formulas pre-cargadas, ej: `=BUSCARX(...)` en PRODUCTO o `=base*1.2` en las +20%) y recien appendea al final cuando se agotan. Preserva todas las formulas existentes (celdas tipo FORMULA se dejan intactas; Excel las recalcula al abrir gracias a `setForceFormulaRecalculation(true)`). **No escribe la columna PRODUCTO**: queda delegada a la formula que el usuario tenga configurada. No se duplican si el SKU ya existe.
  3. **Datos de embalaje en la etiqueta**: cada etiqueta individual (no CARROS) con SKU numerico lleva, en el margen superior derecho, entre 1 y 3 lineas segun lo cargado en el Excel (si `ESTANDARIZADO` dice `NO`, la unica linea es `NO ESTANDARIZADO`, mas grande y en negrita):

     | Condicion | Linea |
     |---|---|
     | `ESTANDARIZADO` en `NO` | `NO ESTANDARIZADO`, y es la unica linea |
     | `ENVASE` | `ENVASE: CAJ-1 - 9Y` (la inscripcion sale de la hoja `ESTANDARIZACION`) |
     | `ENVASE` en `NO` | `ENVASE: NO` |
     | `TIPO DE ROLLO` | `ROLLO: DIAMANTES - 2 paños` (con 0 paños: `ROLLO: DIAMANTES`) |
     | `OBSERVACIONES` | `OBS: Colchon + Tapa` |

     Caja y bolsa no se combinan. La linea de observaciones se queda con el alto libre que sobra, asi que un texto largo sigue en las lineas de abajo en vez de imprimirse encima de si mismo. El banner MEDIR se ubica debajo del numero de etiqueta, a la izquierda, para dejarle ese espacio. Los rotulos (`CAJA:`, `BOLSA:`, `ROLLO:`, `OBS:`) van en negrita. Las lineas que no son la ultima se cortan con `...` si no entran, y las palabras muy largas (codigos, URLs) se parten para que no quede el rotulo solo arriba. Para hacerles lugar se elimina el texto de ML "Recorta esta parte de la etiqueta...", que ocupa esa franja. Si ML cambiara ese texto y no se encontrara, queda una advertencia en el log y los textos se encimarian.
  4. **Columnas Medidas y Embalaje en la tabla**: la tabla de etiquetas descargadas muestra, ademas de #, Orden, Zona, SKU, Producto, Detalles y Cantidad, el estado de cada SKU en el Excel de medidas: `✓ SI`, `✘ NO` (fondo rosa palido) o `—` cuando no aplica (carros y SKU no numericos, o modulo de medidas apagado). Medidas mira las 4 columnas base cm/kg; Estandarizado mira esa columna del Excel. Permite ver el estado del lote sin depender del dialogo final.
  5. **Mensaje de pendientes al finalizar**: al terminar la descarga se abre un dialogo scrollable con la cantidad de SKUs sin medidas detectados en el lote, cuantos se agregaron efectivamente al Excel y cuantos ya figuraban, ademas del listado de SKUs. Se suman ahi los SKU **sin estandarizar**, segun la columna `ESTANDARIZADO`.
  - Durante la descarga **no** se sube nada a ML: el flujo de descarga solo marca y escribe en el Excel.

- **Subida manual a ML** (boton "⬆ Subir Medidas" al lado del selector del Excel de medidas): la subida a ML es una accion independiente, disparada a demanda. Requisitos para que el boton este habilitado: checkbox activo + archivo existente. El handler valida ademas que la sesion ML este inicializada.
  - Al ejecutarse, recorre el Excel, filtra filas con `SUBIDO=NO` **y** las 4 columnas de margen (`+20%` en las dimensiones, `+5%` en el peso) cargadas como numero y mayores a cero, y para cada una:
     - Resuelve `SKU → item_id` via `GET /users/{uid}/items/search?seller_sku=...` con fallback a `?sku=...`.
     - Hace `PUT /items/{item_id}` con body `{"attributes":[...]}` y los 4 atributos `SELLER_PACKAGE_WIDTH`, `SELLER_PACKAGE_HEIGHT`, `SELLER_PACKAGE_LENGTH`, `SELLER_PACKAGE_WEIGHT`. Formato requerido por ML: enteros, `cm` para dimensiones, `g` para peso. El codigo convierte `kg × 1000 → g` y redondea con `Math.round` (evita sesgo de truncado y el ruido de floats de Excel).
     - Si HTTP 200/201, marca `SUBIDO=SI` en verde tenue y limpia la celda `ERROR`. Si falla, deja `SUBIDO=NO` (rojo) y escribe el mensaje parseado en la columna `ERROR` (rojo oscuro sobre rosa palido, con wrap).
     - Parseo del error: del JSON de ML se extrae `cause[0].cause_id` + `cause[0].message` para dejar un mensaje legible (ej: `HTTP 400 · 5401 · The packaging attributes [seller_package_height] are too small for...`). Si no es JSON, se usa el body crudo.
  - **UI**:
     - Corre en thread de background: el boton se deshabilita mientras dura la subida y se reactiva al finalizar.
     - Un label al lado del selector del Excel muestra el progreso en vivo (`Subiendo 2/5 (OK 1 · FAIL 1)`) y al finalizar resume con icono verde o rojo.
     - El dialogo con el detalle por SKU se abre automaticamente al finalizar: estilo `ERROR` (icono rojo, texto rojo oscuro monospace) si hubo fallas, `INFORMATION` si fue todo OK. Se puede reabrir haciendo click en el label de estado.
     - Si no hay pendientes para subir, se informa con un dialogo en lugar de subir nada.
  - Los atributos `SELLER_PACKAGE_*` son los que documenta ML para cuentas ME2 (obligatorios para `cross_docking`/`xd_drop_off`, aceptados en el resto).
  - **Alcance: siempre a nivel ítem (MLA), no por variacion**. Según la doc de ML, los `SELLER_PACKAGE_*` se declaran en cada publicación y no están tageados como `variation_attribute` en ninguna categoría visible. Testeos empíricos confirmaron que intentar subirlos a nivel variación genera errores recurrentes (`cause 146` de duplicación ítem↔variación, `cause 161` "invalid in variation attributes for category"). Por eso el PUT siempre es `PUT /items/{id}` con `{"attributes":[4 SELLER_PACKAGE_*]}` sin wrapper `variations`.
  - **Consecuencia para MLAs con variaciones**: todas las variaciones del mismo MLA comparten estas 4 medidas. Si tu Excel tiene varias filas cuyo SKU mapea al mismo MLA (ej: talles S/M/L distintos), la última que se procese define el valor final a nivel ítem. Esto alinea con el modelo que ML ofrece hoy para `SELLER_PACKAGE_*`.

### Pedidos

Genera un Excel con tarjetas recortables de todos los pedidos pendientes, listas para pegar en los paquetes.

- **Fuentes**:
  - **ML retiro**: `/orders/search` con `tags=no_shipping`, `order.status=paid`, ultimos 7 dias. Excluye ordenes entregadas, cumplidas (`fulfilled`) y con notas. Obtiene nombre, apellido y nickname del comprador via GET `/orders/{orderId}` en paralelo (el search solo devuelve `buyer.id` y `nickname`, el GET directo agrega `first_name` y `last_name`).
  - **TN HOGAR / TN GASTRO**: `/v1/{storeId}/orders` con `payment_status=paid`, `shipping_status=unpacked`, `status=open`. Excluye ordenes pickup con nota del vendedor. Genera etiquetas LLEGA HOY para envios que contengan "LLEGA HOY" en el nombre (excepto Zippin).
- **Excel generado** (`Pedidos/PEDIDOS_*.xlsx`) con hasta 3 hojas:

#### ML PEDIDOS RETIRO (violeta)
- Tarjetas recortables con borde grueso, 2 columnas por pagina.
- Cada tarjeta: N de venta (grande), fecha, nombre y apellido del comprador con nickname entre parentesis, tabla de productos (SKU, CANT, DETALLE).
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
| **MercadoLibre** | Ordenes ME2, etiquetas ZPL, SLA, atributos `SELLER_PACKAGE_*` (dimensiones de paquete) | `ml_credentials.json`, `ml_tokens.json` |
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
