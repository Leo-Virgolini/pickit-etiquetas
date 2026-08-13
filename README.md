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
  | 2 | `Largo cm` | Medida real (base). |
  | 3 | `Ancho cm` | Medida real (base). |
  | 4 | `Alto cm` | Medida real (base). |
  | 5 | `Peso físico (empaque + producto) kg` | Medida real (base). |
  | 6 | `Largo +20%` | Valor que se sube a ML. |
  | 7 | `Ancho +20%` | Valor que se sube a ML. |
  | 8 | `Alto +20%` | Valor que se sube a ML. |
  | 9 | `Peso físico (empaque + producto) +5%` | Valor que se sube a ML. |
  | 10 | `SUBIDO` | `NO` al agregar (rojo tenue), `SI` al subir OK (verde tenue). |
  | 11 | `ESTANDARIZADO` | `SI`/`NO`, calculado por una formula del usuario: resume si completo envase, tipo de rollo y cantidad de paños. Es la unica fuente de verdad sobre si el embalaje esta cargado. |
  | 12 | `ENVASE` | Codigo del envase (`BOL-1`, `CAJ-1`), o `NO` si el producto no lleva. |
  | 13 | `TIPO DE ROLLO` | Tipo de rollo (`DIAMANTES`, `CUADRADOS`), o `NO`. |
  | 14 | `CANT PAÑOS` | Cantidad de paños; solo se imprime si es mayor a cero. Si no es un numero, se imprime tal cual. |
  | 15 | `OBSERVACIONES` | Texto libre. |
  | 16 | `ERROR` | Mensaje de ML en rojo cuando falla la subida. Se limpia al pasar a `SUBIDO=SI` en un reintento exitoso. |

  - Las 4 columnas base cm/kg son los valores reales medidos por el deposito. Las columnas con porcentaje (`+20%` en las dimensiones, `+5%` en el peso) son los valores efectivos declarados a ML. Se ubican por ese porcentaje en el encabezado, no por el texto exacto.
  - Si el archivo no existe se crea automaticamente con headers en la primera ejecucion. Los SKUs nuevos se insertan primero en filas con SKU vacio (reutilizando slots pre-cargados con formulas) y si se agotan se appendean al final. En ambos casos las celdas de medidas faltantes quedan en amarillo y `SUBIDO=NO`. Las celdas que contengan una formula se preservan intactas.
  - El lector tolera variantes: "Largo" o "Profundidad", espacios y saltos de linea dentro del header, y el typo "Profunidad" en la columna +20%.
  - Si el archivo existente no tiene columna `ERROR`, se agrega automaticamente en la primera escritura (migracion silenciosa).
  - Las columnas de embalaje (11 a 15) **las crea y carga el usuario a mano**: la app solo las lee. Se ubican por su encabezado, no por posicion, asi que se pueden reordenar o intercalar columnas propias. Si alguna falta, ese dato no se muestra y el resto sigue funcionando. Solo se crean automaticamente cuando la app genera el archivo desde cero.
  - Escritura serializada con lock interno y reintentos con backoff (500/1000/1500/2000 ms) si el archivo esta abierto en Excel (sharing violation).
  - Las medidas se leen **solo de celdas numericas**, incluidas las formulas con resultado numerico (que es como estan cargadas las columnas con porcentaje). Un valor escrito como texto —aunque parezca un numero, como `"3,006"`— se ignora y ese SKU no se sube a ML: suele ser un dato mal pegado, y de ahi sale la medida que se publica.
  - Si falta la columna `ESTANDARIZADO`, la funcion de embalaje se considera apagada: no se imprimen lineas ni se reclama nada, ni siquiera para los SKU que no figuran en el archivo.
  - Con el checkbox desactivado se saltea el marcado MEDIR, las lineas de embalaje y la subida a ML.

- **Hoja `ESTANDARIZACION`** (dentro del mismo archivo de medidas): catalogo de envases. La app usa dos columnas: `N°` con el codigo (`CAJ-1`, `BOL-1`) e `INSCRIPCION` con el texto escrito en el envase fisico (`9Y`, `AYUDIN`). Las bolsas no llevan inscripcion y las filas que no la tienen traen un guion, que se ignora. Si falta la hoja o el codigo no figura, la etiqueta muestra solo el codigo. La app nunca le escribe.

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
  1. **Banner MEDIR en la etiqueta** (*desactivado*): imprimia un banner "MEDIR: [SKU]" en negro invertido sobre el encabezado de cada etiqueta individual de 1 unidad cuyo SKU no tuviera las 4 columnas base cm/kg cargadas. Quedo fuera de uso: el codigo se conserva entero detras de la constante `BANNER_MEDIR` de `MainController`, que alcanza con poner en `true` para que vuelva. La **deteccion** de pendientes sigue activa y es la que alimenta los dos puntos siguientes.
  2. **Autocarga al Excel**: los SKU detectados como pendientes se insertan en el Excel con SUBIDO=NO. El inserter primero **reusa filas pre-existentes con SKU vacio** (tipicamente filas con formulas pre-cargadas, ej: `=BUSCARX(...)` en PRODUCTO o `=base*1.2` en las +20%) y recien appendea al final cuando se agotan. Preserva todas las formulas existentes (celdas tipo FORMULA se dejan intactas; Excel las recalcula al abrir gracias a `setForceFormulaRecalculation(true)`). **No escribe la columna PRODUCTO**: queda delegada a la formula que el usuario tenga configurada. No se duplican si el SKU ya existe.
  3. **Datos de embalaje en la etiqueta**: cada etiqueta individual (no CARROS) con SKU numerico lleva, en el margen superior derecho, entre 1 y 4 lineas segun lo cargado en el Excel. Aplica igual a las etiquetas turbo, que se tratan como cualquier otra zona.

     | Condicion | Linea |
     |---|---|
     | `ESTANDARIZADO` en `NO` | `NO ESTANDARIZADO`, y es la unica linea |
     | SKU que todavia no figura en el Excel | `NO ESTANDARIZADO` (se lo agrega al Excel en el mismo lote) |
     | pedido de 2+ unidades, con embalaje cargado | `REFERENCIA` en negrita como primera linea, y debajo las demas |
     | pedido de 2+ unidades, sin embalaje cargado | no se imprime nada |
     | `ENVASE` | `ENVASE: CAJ-1 "9Y"` (la inscripcion sale de la hoja `ESTANDARIZACION` y va entre comillas: es lo que esta escrito en el envase fisico) |
     | `ENVASE` en `NO` | `ENVASE: NO` |
     | `TIPO DE ROLLO` | `ROLLO: DIAMANTES - 2 paños` (con 0 paños: `ROLLO: DIAMANTES`) |
     | `OBSERVACIONES` | `OBS: Colchon + Tapa` |

     En un pedido de 2+ unidades el operario no esta embalando un producto suelto, asi que el envase es orientativo: sale encabezado por `REFERENCIA` y no se reclama lo que falte. `NO ESTANDARIZADO` sale en un recuadro negro con el texto en blanco y centrado, para que frene al operario. La linea de observaciones se queda con el alto libre que sobra, asi que un texto largo sigue en las lineas de abajo en vez de imprimirse encima de si mismo. Los rotulos (`ENVASE:`, `ROLLO:`, `OBS:`) van en negrita. `REFERENCIA` va en negrita, subrayado y en la fuente residual B de Zebra (`^ABN,22,14`), de trazo mas cuadrado que el resto, para que se lea como un rotulo y no como un dato mas. ZPL no tiene subrayado como atributo: es una linea dibujada, y se puede calcular exacta porque esa fuente es monoespaciada. Las lineas que no son la ultima se cortan con `...` si no entran, y las palabras muy largas (codigos, URLs) se parten para que no quede el rotulo solo arriba. Para hacerles lugar se elimina el texto de ML "Recorta esta parte de la etiqueta...", que ocupa esa franja. Si ML cambiara ese texto y no se encontrara, queda una advertencia en el log y los textos se encimarian.
  4. **Columna Estandarizado en la tabla**: la tabla de etiquetas descargadas muestra, ademas de #, Orden (el numero que ML imprime como `Pack ID:` o `Venta ID:`, que sale de la API o del propio ZPL al procesar un archivo local), Zona, SKU, Producto, Detalles y Cantidad, el valor de la columna `ESTANDARIZADO` del Excel: `✓ SI`, `✘ NO` (fondo rosa palido) o `—` cuando no aplica (carros y SKU no numericos, o modulo de medidas apagado). Informa sobre el **SKU** y no sobre una etiqueta puntual, asi que no mira la cantidad: una fila puede decir `✘ NO` aunque ninguna de sus etiquetas lleve el aviso impreso por ser de 2+ unidades. Permite ver el estado del lote sin depender del dialogo final.
  5. **Mensaje al finalizar**: al terminar la descarga se abre un dialogo scrollable con cuantos SKU nuevos se agregaron al Excel —son las filas que hay que completar con el envase— y el listado de SKU **sin estandarizar**: los que tienen `NO` en la columna `ESTANDARIZADO` y los que todavia no figuran en el Excel. Se listan solo los que efectivamente salieron con el aviso impreso, asi que un SKU cuyas etiquetas son todas de 2+ unidades no aparece.
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

## Logs

`logs/app.log`, junto al jar. Rota al cambiar el dia o al llegar a 5 MB, y comprime el anterior en
`app-AAAA-MM-DD-N.log.gz`. Los comprimidos con mas de **7 dias** se borran solos.

La limpieza corre **al rotar**, no al arrancar: si la app estuvo sin usarse un tiempo, los archivos
viejos siguen ahi hasta la primera rotacion. Ojo con el `max="10"` de `DefaultRolloverStrategy`, que
no son 10 dias sino el tope del contador `N` dentro de una misma fecha.

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
