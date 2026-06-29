# Validador de creación de artículos — Thermia / Accesorios Dimac S.L.

## Objetivo del proyecto
Automatizar el repaso de artículos creados manualmente en el ERP **Geinfor**
(SQL Server, accedido vía Access con tablas vinculadas, prefijo `DB2ADMIN_`).
Hasta ahora el repaso se hacía campo a campo a mano. Este validador cruza las
tablas exportadas de Geinfor + el Excel de seguimiento de José María y devuelve
una lista de incidencias accionables.

Regla de oro de todo el proyecto: **solo LECTURA de Geinfor**. Nunca escribir en
las tablas del ERP. Las correcciones las hace una persona dentro de Geinfor.

## Principios de trabajo (cómo quiere trabajar José María)
- Consultor senior con criterio propio, no complaciente. Si algo está mal
  enfocado, decirlo claramente.
- Distinguir falsos positivos de errores reales. Si una regla salta en el 100%
  de los casos, sospechar de la REGLA antes que de los datos.
- Validar solo lo que aporta valor. No añadir campos ni complejidad sin justificar.
- Pareto: atacar primero los pocos errores que concentran la mayoría del trabajo.
- Responder siempre en español.

## Propósito y uso
- El validador se usa **SOLO para el ALTA de artículos nuevos**, NO para auditar el
  catálogo existente. Costes y PVP oscilan con el tiempo y ya se recomprueban al hacer
  tarifas nuevas, así que las desviaciones de margen/coste en artículos antiguos quedan
  FUERA de alcance. En uso real se lanza sobre las altas recientes (filtro por FECHACREACION).
- No hace falta mecanismo de excepción de margen: en artículos antiguos puede haber
  decisiones comerciales (precio mantenido a la espera de consumos) que NO son errores;
  como no se revalidan, no generan ruido.

## Universo a validar
- **Accesorios**: familia de producto DISTINTA de 1 y 2 (las familias 1 y 2 son perfiles).
  NOTA: este universo incluye la serie X (familia 105) y catálogos maestros de color (familia 0),
  que tienen casos de TIPO atípicos (60 con estructura, 92 con catálogos != CATALOGOINTERIORES).
  Por eso la regla de TIPO se mantiene CONSERVADORA (ver sección TIPO).
- **Perfiles**: familia 1 o 2 (pendiente de abordar; reglas distintas).
- Validar solo **altas (artículos CREADOS)**, filtrando por **FECHACREACION** dentro de un
  rango. NO valida artículos modificados (decisión de JM: el validador es para creación de
  artículos nuevos). NO filtrar por usuario.
- Los códigos de 13 dígitos se generan automáticamente por el programa → se ignoran.
  Se revisan los de 8 (base) y 10 (base + acabado).

## Estructura del código de artículo
- **Código base = 8 caracteres** (95% de los casos; longitudes distintas → AVISO).
- Perfiles, desglose del base: díg. 1-4 = serie; 5-6 = familia de pieza
  (41=marco, 42=hoja…); 7-8 = tipo concreto de pieza.
- `A000…` = artículo universal/complementario (no de serie concreta).
- **Sufijo de acabado** = 2 dígitos a partir del carácter 9 (solo si longitud > 8):
  - (sin sufijo) bruto/sin acabado — perfil y accesorio
  - 01 plata (perfil y accesorio)
  - 02 bronce (solo perfil)
  - 03 blanco (perfil y accesorio)
  - 04 negro (perfil y accesorio)
  - 05 anod. inox/negro/oro (solo perfil)
  - 06 RAL estándar (solo perfil)
  - 07 RAL especial (solo perfil)
  - 08 bicolor — SOLO perfiles familia 2 (rotura), y solo cuando se decide
  - 09 madera (solo perfil)
  - Tras 06 puede haber 3 dígitos más (código de 13) = acabado concreto del grupo
    (ej. 061 = RAL 7016 TX). Esos son automáticos → no se revisan.
- IMPORTANTE: el acabado SOLO existe si longitud > 8. Los dígitos 7-8 del base
  NO son acabado aunque coincidan con "02", etc.

## Reglas de validación — ACCESORIOS (familia ≠ 1 y 2)

### Campos del maestro (MAESTRO DE ARTICULOS)
- TIPO: depende del artículo (ver sección TIPO).
- CLASE: código numérico de serie. Se valida cruzando la columna CLASE(SERIE) del Excel
  de JM (texto, p.ej. "CF40/CR40") con la DENOMINACION de la hoja CLASE ARTICULO del repaso.
  Basta con que alguna de las series (separadas por "/") aparezca en la descripción de la clase.
  Normalización: se ignoran guiones y espacios extra. Alias conocidos: DOPLO = PR77.
  Caso real a corregir por JM: poner "ACCESORIOS DIVERSOS" en vez de "UNIVERSAL" (clase 0).
  CLASE=53 era solo para artículos FW de prueba, no es una regla general.
- COEFICIENTE = 1 (salvo casos de unidad con coeficiente, ver Unidades)
- ESTADO = 0 o 70 al crear (0 = no sale en tarifa; 70 = sale en tarifa). El valor correcto lo indica JM en su Excel (columna pendiente de añadir). Validar contra esa columna cuando esté disponible.
- NORMA: NO se valida. Se solicita y se añade manualmente por JM (puede ser EU, CHINA,
  TURQUIA, etc. según origen). Regla "siempre EU" eliminada (daba falsos positivos).
- FAMILIAPRODUCTO: cualquier valor excepto 1 y 2 (perfiles). La regla FAMILIAPRODUCTO=104 era solo para FW de prueba, eliminada.
- UNIDADES_ARTICULO = 1
- DENOMINACION (base): debe tener texto.
- MEDIDA = 1; excepción: si UNIDAD_INTERNA=4 y UNIDAD_EXTERNA=4 → MEDIDA=2.

### DENOMINACION_2 según GRUPO
- Grupo por defecto (sin acabado): vacío
- Grupo 10 → "NEGRO"
- Grupo 11 → "BLANCO"
- Grupo 12 → "PLATA"
- Grupo 14 → "RAL STD"
- Grupo 16 → "PVD"
(Es un fallo frecuente. A futuro se podría automatizar el rellenado.)

### GRUPO según sufijo del código (accesorios)
- sufijo 04 → grupo 10
- sufijo 03 → grupo 11
- sufijo 01 → grupo 12
- sufijo 06 → grupo 14
- sufijo 05 → grupo 16

### Característica 0 (TABLA CARACTERISTICAS, formato largo ARTICULO/CARACTERISTICA/VALOR)
- Debe existir característica 0 con valor (es la imagen JPG del artículo).
- OJO con el JOIN en Access: usar siempre LEFT/equivalente para no ocultar
  artículos que NO tienen característica (esos son justo los que fallan).

### Proveedor (ARTICULOS PROVEEDOR) — SIEMPRE proveedor PRINCIPAL (PRINCIPAL=1)
- Debe existir fila con PRINCIPAL=1 (proveedor principal asignado).
- PLAZOENTREGA > 0 (informado; el valor varía según proveedor).
- LOTEHABITUAL > 0 (informado; varía).
- Campo de coste a usar: **COSTEINTERNO** del principal (= coste ACTUAL neto: el
  COSTEBRUTO con DTO_ACTUAL y RECARGO_ACTUAL ya aplicados). NO usar COSTEBRUTO
  (es el precio de tarifa del proveedor antes de descuento). Verificado: COSTEINTERNO
  = COSTEBRUTO*(1-DTO/100)*(1+RECARGO/100) salvo casos con recargos por cantidad.
- (El secundario, PRINCIPAL=0, se IGNORA en todo.)

### Margen / PVP (ARTICULOS TARIFABLES, tarifa 40)
- Fórmula: **COSTEINTERNO (principal) ÷ DIVISOR ACCESORIO (Excel José María) = PVP**.
- PVP a 2 decimales. Comparar con PRECIO_VENTA_PUBLICO de tarifa 40.
- Tolerancia HÍBRIDA: solo se considera desviación si supera A LA VEZ 0,05 € Y 5%
  (absorbe el band sistemático ~3-4% validado con JM en el artículo 072040).
- Regla ASIMÉTRICA (confirmada por JM): superada la tolerancia,
  - PVP real POR DEBAJO del calculado → **ERROR** (vende barato, poco margen; más grave).
  - PVP real POR ENCIMA del calculado → **AVISO** (más margen, solo revisar).
- El divisor se lee de la FILA EXACTA del artículo en el Excel de José María
  (no de la base de 8). Cada código tiene su divisor.
- Otras tarifas (no 40) cogen la info de la 40 → no se validan.

### Almacén (ALMACEN ARTICULOS) — stock mínimo cruzado con Excel JM
- Se compara MINIMO ACCESORIO (col. M) del Excel JM contra STOCKMINIMO de almacén:
  - JM = 0 y almacén > 0 → ERROR (no debe tener mínimo).
  - JM > 0 y almacén > 0 → OK (el valor exacto se ajusta por estadística de ventas;
    NO se exige coincidencia exacta).
  - JM > 0 y almacén = 0 y ESTADO(JM)=70 → ERROR (en tarifa: debe tener mínimo fijado).
  - JM > 0 y almacén = 0 y ESTADO(JM)≠70 → AVISO (aún no en tarifa). Usa el estado de JM.

### Unidades — coherencia MAESTRO ↔ PROVEEDOR ↔ TARIFABLE
La unidad EXTERNA del maestro = unidad del proveedor (UNIDADEXTERNA en proveedor).
La unidad de venta del tarifable (UNIDADVENTA) coherente con la interna.
Cuatro casos:
- **A — Unidades (unitario):** interna 1 / externa 1; proveedor unidad 1;
  tarifable UNIDADVENTA 1; sin coeficiente especial.
- **B — Barras / PVC:** interna 9 / externa 4; COEFICIENTE (maestro) = 1 ÷ UNIDADES_ARTICULO
  (la longitud; ej. 6,3 m → 1/6,3 = 0,15873). UNIDADES_ARTICULO = longitud de la barra
  (no se valida contra 1). tarifable UNIDADVENTA 4 (metro). [VALIDADO en código]
- **C — Gomas/juntas/felpas (metro lineal):** interna 4 / externa 4; SIN coeficiente
  (coef=1, porque interna=externa); MEDIDA=2 (son rollos); UNIDADES_ARTICULO debe ser
  = LOTEHABITUAL del proveedor principal (longitud del rollo). tarifable UNIDADVENTA 4. [VALIDADO]
- **D — Color / lacado (metro cuadrado):** interna 5 / externa 5; va ligado a
  TIPO 2. Son artículos de color (códigos con terminación de 3 díg. tras el 06).
- Anomalías detectadas a revisar: combinaciones 9/1 y 4/5 no encajan en ningún caso.

### Campo TIPO (accesorios) — CERRADO (reglas confirmadas con JM, jun 2026)
Reglas implementadas en el validador:
- **0, 16, 91 NO deben existir** en accesorios → ERROR.
- Con estructura que contiene **CATALOGOINTERIORES** → debe ser **92** (modular).
- Con estructura que contiene **otro catálogo** (artículo TIPO 91: MADERA, RALESTANDAR,
  CATALOGOEXTERIORES, ANODIZADO…) → es de PERFILES, no debería estar en accesorios → ERROR.
- Con estructura de **códigos de compraventa** (sin catálogo) → **15** (kit, solo accesorios).
- **Sin estructura** → **1** (compraventa), **2** (lacado), **5** (muestra) o **60**.
Detección de catálogos: se construye el set de artículos con TIPO 91 en el maestro; el único
catálogo válido como componente en accesorios es CATALOGOINTERIORES.

CAVEAT de FAMILIA: el universo (familia ≠ 1 y 2) incluye dos líneas especiales que NO siguen
estas reglas y por eso generan incidencias TIPO que son ruido (no aparecen en altas FW reales):
- **Familia 105 (serie X)**: usa TIPO 60 con estructura base+color → salta como "debe ser 15".
- **Familia 0 (catálogos y MQ especiales)**: TIPO 91 (catálogos) y artículos muy específicos.
En accesorios FW reales (familia 104) la regla deja muy pocas incidencias y son genuinas
(p.ej. TIPO 92 SIN estructura = modular incoherente).

## Tablas y campos clave (nombres reales de columna)
- MAESTRO DE ARTICULOS: CODIGO_ARTICULO, DENOMINACION, DENOMINACION_2, TIPO,
  CLASE, COEFICIENTE, ESTADO, NORMA, FAMILIAPRODUCTO, GRUPO, UNIDAD_INTERNA,
  UNIDAD_EXTERNA, UNIDADES_ARTICULO, MEDIDA, FECHACREACION, FECHAMODIFICACION,
  CREADOPOR, MODIFICADOPOR.
- ARTICULOS TARIFABLES (llave ARTICULO): TARIFA, PRECIO_VENTA_PUBLICO,
  UNIDADVENTA, COEFICIENTE.
- ARTICULOS PROVEEDOR (llave ARTICULO): PROVEEDOR, PRINCIPAL, COSTEBRUTO,
  PLAZOENTREGA, LOTEHABITUAL, UNIDADEXTERNA, COEFICIENTE, CODIGO_PROVEEDOR.
- ALMACEN ARTICULOS (llave ARTICULO): STOCKMINIMO.
- TABLA CARACTERISTICAS (llave ARTICULO): CARACTERISTICA, VALOR (formato largo).
- Excel José María, hoja 'RELLENAR JOSE MARIA' (llave ARTÍCULO): DIVISOR ACCESORIO,
  MINIMO ACCESORIO, PARTIDA ARANCELARIA (recién añadida), divisores por acabado.

## Estado del proyecto (act. jun 2026)
- HECHO y validado: campos fijos, denominación base, medida, grupo, imagen, proveedor
  principal + plazo + lote, CLASE (cruce con tabla), ESTADO (vs Excel JM), margen
  (COSTEINTERNO, tolerancia híbrida + asimétrica), stock (vs Excel JM), y COHERENCIA
  DE UNIDADES casos B (barras: coef=1/longitud) y C (metro lineal: unidades=lote habitual).
- TIPO: conservador (solo TIPO 0 y CATALOGOINTERIORES⇒92). Taxonomía completa abierta.
- PENDIENTE: cerrar taxonomía TIPO con JM; casos de unidad A (unitario) y D (m²/lacado)
  si hace falta más detalle; abordar PERFILES (familia 1 y 2) con su Excel propio.
- Casos de excepción que el validador marca para que JM valore (no son bugs): p.ej.
  ESTADO=99 hecho a conciencia (FWA00057). Cuando aparezcan, preguntar a JM.
- FUTURO: empaquetar para las dos personas que crean artículos; conexión directa
  (solo lectura) a Geinfor; posible autorrelleno de denominación 2.

## Cómo se ejecuta
`python validador_articulos_dimac.py REPASO.xlsx JOSEMARIA.xlsx [FECHA_DESDE]`
Genera INCIDENCIAS_FINAL.xlsx con columnas CODIGO, CAMPO, NIVEL (ERROR/AVISO), DETALLE.
Requiere: Python con pandas y openpyxl.

## Conexión directa a Geinfor (DB2) — VERIFICADA (jun 2026)
- Backend = IBM DB2 (Access es solo front-end). DSN correcta = **`GeinprodINGENRED`**
  (confirmado: PM40310704, FWM0001401, 072040 existen con sus datos; 91.825 artículos).
- **Conecta SIN credenciales**: `pyodbc.connect('DSN=GeinprodINGENRED')`. Solo requiere estar
  en la red de la oficina (servidor 10.0.0.4; desde fuera da timeout). Esquema: DB2ADMIN.
- Ventaja demostrada: datos LIVE. PM40310704 ya tenía DENOMINACION_2='NEGRO' en DB2 (corregido
  en el programa) mientras el Excel exportado seguía con 'nan'.
- **Mapeo de tablas (hoja Excel → tabla DB2, columnas idénticas salvo donde se indica):**
  - MAESTRO DE ARTICULOS → `DB2ADMIN.MAESTRO_DE_ARTICULOS` (mismas columnas)
  - ARTICULOS TARIFABLES → `DB2ADMIN.ARTICULOS_TARIFABLES`
  - MAESTRO DE ESTRUCTURAS → `DB2ADMIN.MAESTRO_ESTRUCTURAS` (ARTSUPERIOR/ARTCOMPONENTE; tiene ESBAJA/FECHA_BAJA)
  - ARTICULOS PROVEEDOR → `DB2ADMIN.ARTICULOS_PROVEEDOR` (incluye COSTEINTERNO, LOTEHABITUAL, UNIDADEXTERNA)
  - ALMACEN ARTICULOS → `DB2ADMIN.ALMACEN_ARTICULOS`
  - CLASE ARTICULO → `DB2ADMIN.TABLA_CLASE_ARTICULO` (CODIGO, DENOMINACION, …)
  - TABLA CARACTERISTICAS → `DB2ADMIN.TCARAC_ART` (ARTICULO, CARACTERISTICA, VALOR; caract 0 = ruta de la imagen JPG)
- **Modo --db OPERATIVO y rápido (~3,5 s).** El Excel de JM se localiza SOLO en la red
  (no hay que copiarlo): `JM_CARPETA` = `\\dimac2023\DATOS\DATOS DIMAC\USERS\INTERCAMBIO\RAUL\
  xavi\PROCES DE COMPRES ADRIA\EXTRUSORES\MATRICES EXTRUSIÓN`, patrón `*Matrices nuevas*.xlsx`
  (coge el más reciente). Usos:
  - Rango: `python validador_articulos_dimac.py --db [FECHA_DESDE] [FECHA_HASTA]`
  - Un artículo: `python validador_articulos_dimac.py --db --articulo CODIGO` (~2,7 s)
  - Se puede pasar un .xlsx explícito como primer arg tras --db (alternativa/fallback).
  Accesos directos (doble clic): `VALIDAR ALTAS NUEVAS.bat` (pide fecha) y `VALIDAR UN
  ARTICULO.bat` (pide código). Ya NO copian el Excel: lo leen de la red.
  IMPORTANTE: la fuente de verdad es el Excel de la RED; JM lo mantiene ahí.
- SALIDA VISUAL (HTML) en ambos modos, se abre solo en el navegador:
  - Un artículo: `INFORME_<cod>.html` (banner verde si correcto; tarjetas rojas/ámbar por
    incidencia con nombre de campo legible). NO escribe Excel en este modo.
  - Lote: `INFORME_ALTAS.html` (un bloque por artículo con incidencias, errores primero) +
    `INCIDENCIAS_FINAL.xlsx` (tabla para análisis). Si el Excel está abierto/bloqueado, guarda
    en `INCIDENCIAS_<fecha_hora>.xlsx` en vez de crashear.
  - Estilos compartidos en la constante _ESTILO_HTML; nombres legibles en NOMBRES_CAMPO.
  - Modo un-artículo SIEMPRE genera HTML: si el código es un PERFIL (familia 1/2) o no existe,
    el informe lo dice claramente (no se queda en blanco). El validador de accesorios NO valida
    perfiles (tendrán su propia validación).
  Lee Geinfor en directo (solo lectura, readonly=True + solo SELECT). El Excel de JM sigue
  siendo fichero. Mantiene el modo Excel como alternativa (primer arg = ruta .xlsx).
- Optimización: el maestro trae solo accesorios (familia!=1,2) o catálogos (tipo 91) y solo las
  columnas usadas; las tablas relacionadas solo de los códigos EN ALCANCE y con los filtros que
  el validador ya aplica empujados a SQL (proveedor PRINCIPAL=1, tarifa 40, característica 0).
  Consultas IN troceadas de 500. Para altas de un día = casi instantáneo.
