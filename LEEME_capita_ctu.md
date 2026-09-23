# Cápita asignada CTU · circuito mensual

## Qué cambió
La cápita ya no se plancha dentro de los HTML. Los cuatro tableros (general `adendas-375`, consulta `consulta-375`, tablero PHC y padrón) leen al abrir el archivo madre:

`https://raw.githubusercontent.com/ivogomez85/padron-datos/main/capita_ctu.json`

Si el JSON no se puede leer (sin red), general, consulta y PHC usan el bloque planchado de respaldo (jul+ago+sep 2026) y el padrón muestra la cápita solo en la solapa Adendas (mapa de UGLs), que también tiene respaldo.

## Qué hay que subir ahora
1. `capita_ctu.json` → repo **padron-datos**, en la raíz, con ese nombre exacto.
2. `adendas-375_index.html` → repo adendas-375 como `index.html`.
3. `consulta-375_index.html` → repo consulta-375 como `index.html`.
4. `phc-tablero_index.html` → repo del tablero PHC como `index.html`.
5. `padron_index.html` → repo del padrón como `index.html`.

## Cada mes (o cuando PAMI reextrae un mes)
Solo se regenera y se reemplaza `capita_ctu.json`. Los HTML no se tocan.

    python3 gen_capita_ctu.py --padron padron_2026-09.csv.gz \
        --previo capita_ctu.json \
        --excel Capitas_asignadas_CTU_202610.xlsx \
        --out capita_ctu.json

- `--previo` conserva los meses ya cargados; el mes que viene en el Excel se reemplaza entero (por eso un Excel reextraído, como "202608 al 22 de sept", pisa al anterior sin duplicar).
- `--padron` es la base más reciente del repo padron-datos: sirve para traducir C_PRESTADOR a SAP, legajo y CUIT (los Excel CTU no traen SAP).
- `--agencias maestro.xlsx` (opcional): maestro oficial de agencias con columnas UGL, código y nombre. Si se pasa, sus nombres mandan sobre los inferidos y quedan guardados en el JSON para los meses siguientes.
- Alternativa sin script: me pasás el Excel nuevo y te devuelvo el JSON.

## Cómo se lee el Excel CTU (regla validada)
- Grano real: prestador × módulo × agencia asignada, identificada por (UGL asignada, código de agencia). La columna D_AGENCIA viene cartesiana (para cada código trae todas las agencias del país con ese código). El nombre real se infiere cruzando la localidad del nombre con las localidades de esa UGL en el padrón, exigiendo que la localidad sea dominante en la UGL y que cada código quede con un solo nombre por UGL (`agencias` en el JSON, con `fuente` unico/inferido/maestro). Lo que no se puede identificar con certeza se muestra solo con el código. "Padrón agencia" = N_CANT_AGENCIA (afiliados de la agencia).
- N_CAPITA_REF = % asignado × afiliados de la agencia. La cápita del módulo es la suma de sus agencias distintas (Arellano SAP 116041: 25 % de 18.835 = 4.709; Clínica Integral SAP 35928 mód. 543: 3.586 + 2.822 = 6.408, iguales a la cápita de referencia de agosto).
- Entre módulos de un mismo bloque no se suma: se promedia (eso lo hace cada HTML).

## Estructura del archivo madre
```
{ esquema, generado, descripcion, clave,
  meses: { "202609": { archivo, filas, prestadores, modulos_prestador, cargado } },
  ugl: { "01": "TUCUMAN", ... },
  modulos: { "36": "OFTALMOLOGIA - CONSULTAS Y PRACTICAS", ... },
  prestadores: [ { clave:"116041||TUCUMAN", sap, legajo, c_prestador, cuit, ugl, ugl_nombre, nombre, c_red,
      meses: { "202609": { id_proceso, fuente, bloques:{ "Oftalmología": 4707 },
               modulos:[ { mod:"36", cap:4707, rubro, n_agencias, cruza_ugl,
                           agencias:[ { ugl:"01", cod:"0002", pct:25, afiliados:18827, cap:4707 } ] } ] } } } ] }
```

## Dónde se ve en cada HTML
- General y consulta: igual que antes (chip 👥 en el resumen, columna y filtros en firmadas, exports) más un desplegable "📍 Agencias asignadas" en el detalle por módulo.
- PHC: chips por bloque y tabla en la ficha, más el desplegable de agencias.
- Padrón: chips en el encabezado de la ficha del prestador y sección desplegable con módulo × mes y agencias; misma sección en la ficha resumen; columna "👥 Cápita" y filtro Con/Sin cápita en Exp RE375 (también en XLSX y PDF); columna "👥 Cápita" en el Lector 375 (también en el export); y la solapa Adendas del mapa de UGLs pasa a leer el JSON.
