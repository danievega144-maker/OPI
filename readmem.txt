Quiero que reescribas mi parser como un módulo robusto para Python 3.7 en CDSW que haga lo siguiente, sin perder ninguna fila:

Entradas

sharepoint_dir: str → carpeta local donde se descargan los PDFs desde SharePoint (ruta absoluta).

pdf_name: str → nombre del PDF a procesar (ej. "2024.10.07 Statesville.pdf").

Objetivo

Abrir sharepoint_dir/pdf_name (PDF).

Convertirlo a DOCX usando pdf2docx.Converter y guardar el DOCX en la misma carpeta con el mismo nombre y extensión .docx (no sobrescribir si ya existe y tiene tamaño > 0).

Reabrir el DOCX con python-docx y leer TODO: párrafos, tablas y subtablas anidadas (no se vale perder bloques).

Detectar una sola fila de header por bloque usando tokens conocidos y mapear a los 16 nombres destino:
COLUMN_NAMES = [
  "product_number","formula_code","product_name","product_form","unit_weight",
  "fob_or_dlv","price_change","list_price","full_pallet_price","full_load_best_price",
  "pallet_quantity","min_order_quantity","days_lead_time",
  "stocking_status","price_per_unit","price_in_us_dollars"
]

Usa coincidencia por tokens y alias (normaliza a mayúscula, quita símbolos, divide por espacios).

Score = 0.6 * SequenceMatcher.ratio + 0.4 * Jaccard(tokens).

Umbral mínimo 0.35; si ninguna columna supera el umbral, crea la columna vacía para no descuadrar.

Nunca elimines filas internas; primero identifica header_row_index, luego toma todas las filas debajo y reestructura por mapeo.

species por bloque:

Mantén una lista SPECIES_LIST y su versión CANONICAL_SPECIES (nivel superior).

Para cada párrafo/título detectado antes de una tabla, aplica fuzzy (umbral 0.88 a la lista completa; si falla, 0.75 contra canónicas).

Aplica forward-fill por bloque y solo resetea cuando aparece un nuevo título; si no hay match, usa "DESCONOCIDO".

Metadatos (agregar 4 columnas al final y en este orden):

plant_location: extraer por regex contra lista de plantas conocidas; si no hay match, "PLANTA DESCONOCIDA".

date_inserted: fecha efectiva detectada en el documento (dd/mm/aaaa, mm/dd/aa, etc.). Si no hay, usar mtime del PDF.

source: Path(pdf_path).name (nombre del PDF).

species: valor asignado por bloque (punto 5).

Normalizaciones:

product_form: mapear contra VALID_FORMS = {"EXTRUDED","LIQUID","MEAL","PELLETS","CRUMBLES","BLOCKS","BISCUIT","TEXTURED","CUBED","POWDER","SEED","GRANULES"} con limpieza básica; si multi-palabra, toma el último token válido.

Columnas numéricas: pallet_quantity, min_order_quantity, days_lead_time + las columnas 10–15; función segura que quite símbolos, maneje paréntesis como negativos y devuelva NaN si no se puede convertir.

Salida final

Un único DataFrame concatenando todos los bloques, con exactamente COLUMN_NAMES + ["plant_location","date_inserted","source","species"] en ese orden.

Sin pérdida de filas: cuenta filas antes/después y repórtalo por print.

Reporta también # de bloques, # de tablas y % de species == "DESCONOCIDO".

Arquitectura que debes generar

convert_pdf_to_docx(pdf_path) -> Path

iter_body_blocks(doc) -> Iterator[Union[Paragraph, Table]] que incluya tablas anidadas (CT_Tbl) con un generador iter_nested_tables.

table_to_dataframe(tbl) -> DataFrame (sin filtrar filas).

find_header_index(df_raw) -> Optional[int] usando HEADER_TOKENS.

build_columns_map(header_row) -> Dict[target_col, Optional[int]] con el scoring indicado.

restructure_from_header(df_raw) -> DataFrame que:

use header_row_index,

seleccione todas las filas debajo,

cree un DataFrame nuevo con COLUMN_NAMES, rellenando vacíos si faltan columnas.

best_fuzzy(txt, pool, threshold) -> Optional[str] para species.

_normalize_product_form, _fix_numeric(df) y _parse_date_any(s) como utilidades.

read_file(pdf_path) -> DataFrame que orqueste todo.

main(sharepoint_dir, pdf_name) -> DataFrame que:

arme rutas, convierta, parsee y retorne el DF;

opcionalmente guarde un .xlsx junto al PDF con el mismo nombre.

Reglas de oro

No borres encabezados antes de identificar header_row_index.

No uses ventanas de 16 columnas para cortar; todo el re-mapeo se hace por tokens/alias.

No pierdas subtables: recorre CT_Tbl recursivamente.

Si algo no mapea, crea la columna vacía; nunca descartes filas.

Agrega assert/print de control: filas totales detectadas vs filas del DataFrame final.

Devuelve el DF con el orden de columnas exacto, listo para to_excel.

Aceptación (qué debe cumplirse)

Mismo número de filas de datos en el resultado que la suma de todas las filas de todas las subtablas (sin contar la fila header de cada bloque).

len(df.columns) == 20 en el orden indicado.

El DOCX se crea en sharepoint_dir y el parser siempre trabaja sobre ese DOCX.

% DESCONOCIDO razonable (< 10% típico, dejarlo parametrizable).

El código debe estar en un solo archivo y con funciones nombradas exactamente como arriba.

Genera el código completo con esas funciones y un bloque if __name__ == "__main__": que reciba sharepoint_dir y pdf_name por argparse, ejecute main, imprima los conteos y, si paso --excel, guarde sharepoint_dir/<pdf_name_sin_ext>.xlsx.