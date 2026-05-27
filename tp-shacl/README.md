# TP SHACL — Validación de restricciones sobre OVS-inmontology

Este directorio contiene el trabajo práctico de la materia Gestión del Conocimiento: un grafo de restricciones SHACL sobre la ontología inmobiliaria OVS-inmontology, junto con grafos de datos de ejemplo (válido e inválido) para demostrar el funcionamiento.

## Contenido

| Archivo | Descripción |
|---|---|
| `shapes.ttl` | Grafo SHACL con 7 NodeShape y 16 sh:property. Incluye jerarquía `HouseShape → RealEstateShape` vía `sh:node`. |
| `data-valid.ttl` | Grafo de datos que cumple todas las restricciones. |
| `data-invalid.ttl` | Grafo de datos que viola 8 restricciones distintas a propósito. |
| `reports/valid-report.txt` | Reporte de pyshacl para el caso válido (Conforms: True). |
| `reports/invalid-report.txt` | Reporte de pyshacl para el caso inválido (Conforms: False, 9 violaciones). |
| `sparql/` | Consultas SPARQL usadas para explorar la ontología. |
| `seguimiento.docx` | Documento interno con el historial de lo que fuimos haciendo (no es entregable). |

## Cómo validar en GraphDB (paso a paso)

### Requisitos previos
1. **GraphDB Free** instalado (https://graphdb.ontotext.com/ → versión Desktop, gratis).
2. Repositorio clonado.

### Paso 1 — Crear el repositorio en GraphDB con SHACL activado
1. Abrir GraphDB en http://localhost:7200
2. Menú **Setup → Repositories → Create new repository**
3. Tipo: **GraphDB Repository**
4. Repository ID: `inmontology-tp`
5. **⚠️ IMPORTANTE**: marcar el checkbox **"Supports SHACL validation"** (no se puede activar después).
6. Create.

### Paso 2 — Cargar la ontología base
1. Menú **Import → User data → Upload RDF files**
2. Subir `inmontology.owl` (raíz del repo).
3. Esperar "Imported successfully".

### Paso 3 — Cargar las shapes SHACL
1. Menú **Import → User data → Upload RDF files**
2. Subir `tp-shacl/shapes.ttl`
3. En **Import settings → Target graphs**: seleccionar **Named graph** y usar el IRI `http://rdf4j.org/schema/rdf4j#SHACLShapeGraph`.
4. Import.

### Paso 4 — Probar el caso VÁLIDO
1. **Import → User data → Upload RDF files** → subir `tp-shacl/data-valid.ttl`.
2. ✅ Resultado esperado: "Imported successfully".
3. Sacar screenshot del éxito.

### Paso 5 — Probar el caso INVÁLIDO
1. **Import → User data → Upload RDF files** → subir `tp-shacl/data-invalid.ttl`.
2. ❌ Resultado esperado: GraphDB rechaza la carga con una *SHACL Validation Exception* y muestra las violaciones.
3. Sacar screenshot del error y reporte.

## Cómo validar localmente con pyshacl (opcional)

```bash
pip3 install --user pyshacl

# Caso válido (esperable: Conforms: True)
python3 -m pyshacl -s tp-shacl/shapes.ttl \
                   -d tp-shacl/data-valid.ttl \
                   -e inmontology.owl \
                   --inference rdfs -f human

# Caso inválido (esperable: Conforms: False, 9 violaciones)
python3 -m pyshacl -s tp-shacl/shapes.ttl \
                   -d tp-shacl/data-invalid.ttl \
                   -e inmontology.owl \
                   --inference rdfs -f human
```

El flag `-e inmontology.owl` carga la ontología, e `--inference rdfs` aplica las inferencias de subclases (necesario para que pyshacl sepa que `io:Price` es subclase de `io:Feature`, por ejemplo).

## Resumen de las restricciones (shapes.ttl)

| # | Shape | targetClass | Restricciones principales |
|---|---|---|---|
| 1 | `:RealEstateShape` | `rec:RealEstate` | ≥1 `io:hasFeature`, ≥1 `rec:includes` |
| 2 | `:HouseShape` | `io:House` | Hereda de RealEstateShape (`sh:node`) + ≥1 feature de tipo `io:RoomAmnt` |
| 3 | `:RealEstateListingShape` | `pronto:RealEstateListing` | Exactamente 1 `sioc:about`, ≥1 `sioc:has_creator`, `gr:hasBusinessFunction` ∈ {`gr:Sell`, `gr:LeaseOut`}, `sioc:id` |
| 4 | `:PriceShape` | `io:Price` | ≥1 `io:hasDetail` |
| 5 | `:UnitPriceSpecificationShape` | `gr:UnitPriceSpecification` | Currency ISO 4217 (regex), valor float > 0, priceType ∈ {BASE, TOTAL} |
| 6 | `:PostalAddressShape` | `io:PostalAddress` | `io:address` no vacío, ≥1 `rec:city`, ≥1 `io:province`, `io:neighborhood` opcional pero tipado |
| 7 | `:OriginShape` | `io:Origin` | Debe ser IRI nombrado y de tipo `io:Scraper`, `io:AVE` o `io:Curation` |

## Violaciones diseñadas en data-invalid.ttl

| ID | Nodo | Qué se rompió | Shape afectada |
|---|---|---|---|
| V1 | `ex:price_value_bad` | `gr:hasCurrency` = "dolares" (no es ISO 4217) | UnitPriceSpecificationShape |
| V2 | `ex:price_value_bad` | `gr:hasCurrencyValue` = -50000.0 | UnitPriceSpecificationShape |
| V3 | `ex:listing_bad` | sin `sioc:has_creator` | RealEstateListingShape |
| V4 | `ex:listing_bad` | `gr:hasBusinessFunction` = `gr:Repair` | RealEstateListingShape |
| V5 | `ex:house_bad` | Casa sin feature `io:RoomAmnt` | HouseShape |
| V6 | `ex:postal_address_bad` | sin `io:province` | PostalAddressShape |
| V7 | `ex:postal_address_bad` | `io:neighborhood` apunta a un string | PostalAddressShape |
| V8 | blank node de `io:Origin` | blank node + no es Scraper/AVE/Curation | OriginShape |

## Troubleshooting

| Problema | Solución |
|---|---|
| GraphDB no acepta shapes.ttl | El repo no fue creado con SHACL habilitado. Borrar y rehacer marcando "Supports SHACL". |
| `data-valid.ttl` es rechazado | Falta cargar `inmontology.owl` primero. SHACL necesita las jerarquías de subclase. |
| `data-invalid.ttl` es aceptado | `shapes.ttl` no se cargó en el named graph correcto (paso 3). |
