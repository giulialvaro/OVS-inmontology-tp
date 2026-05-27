const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, BorderStyle, WidthType, ShadingType,
  AlignmentType, LevelFormat
} = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function codeBlock(text) {
  return text.split('\n').map(line =>
    new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: line || ' ', font: 'Courier New', size: 20 })]
    })
  );
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}

function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}

function p(text) {
  return new Paragraph({ children: [new TextRun(text)] });
}

// Resultado Query 1 como tabla de 2 columnas (nro + clase)
const clases = [
  "gr:BusinessFunction","gr:Offering","gr:PriceSpecification","gr:QuantitativeValue","gr:UnitPriceSpecification",
  "sioc:Container","sioc:Forum","sioc:Item","sioc:Post","sioc:Site","sioc:Space","sioc:UserAccount",
  "io:AVE","io:Address","io:Apartment","io:City","io:CondominiumUnit","io:CornerLot","io:Curation",
  "io:Dimensions","io:FOT","io:FacadesAmnt","io:Farm","io:Feature","io:Hotel","io:House","io:IrregularLot",
  "io:Land","io:Neighborhood","io:Origin","io:ParkingLot","io:PostalAddress","io:Price","io:PropertyAge",
  "io:Province","io:RoomAmnt","io:Scraper","io:Store","io:Surface","io:SwimmingPool","io:Warehouse",
  "rdfs:Literal","rdfs:Resource",
  "time:Instant","time:Interval","time:ProperInterval","time:TemporalEntity",
  "foaf:Agent","foaf:Document",
  "pronto:RealEstateListing","pronto:SizeSpecification",
  "rec:Agent","rec:Architecture","rec:Bathroom","rec:Bedroom","rec:Building","rec:Collection",
  "rec:Geometry","rec:OutdoorSpace","rec:Point","rec:RealEstate","rec:Region","rec:Room","rec:Site","rec:Space"
];

function resultTable(items) {
  const headerCell = (txt) => new TableCell({
    borders,
    width: { size: txt === '#' ? 720 : 8640, type: WidthType.DXA },
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true })] })]
  });
  const cell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, font: 'Courier New', size: 20 })] })]
  });
  const rows = [
    new TableRow({ children: [headerCell('#'), headerCell('Clase')] }),
    ...items.map((it, i) => new TableRow({ children: [cell(String(i+1), 720), cell(it, 8640)] }))
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [720, 8640],
    rows
  });
}

const query1 = `PREFIX owl: <http://www.w3.org/2002/07/owl#>
SELECT DISTINCT ?clase WHERE {
  ?clase a owl:Class .
  FILTER(isIRI(?clase))
} ORDER BY ?clase`;

const query2 = `PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>
SELECT ?sub ?super WHERE {
  ?sub rdfs:subClassOf ?super .
  FILTER(isIRI(?sub) && isIRI(?super))
} ORDER BY ?super ?sub`;

// Jerarquías relevantes (filtradas del resultado de 100 filas)
const jerarquiaRealEstate = [
  "io:Apartment","io:CondominiumUnit","io:Farm","io:Hotel","io:House",
  "io:Land","io:ParkingLot","io:Store","io:Warehouse"
];
const jerarquiaFeature = [
  "io:Address","io:CornerLot","io:Dimensions","io:FOT","io:FacadesAmnt",
  "io:IrregularLot","io:Neighborhood","io:Price","io:PropertyAge",
  "io:RoomAmnt","io:Surface","io:SwimmingPool"
];
const jerarquiaOrigin = ["io:AVE","io:Curation","io:Scraper"];
const jerarquiaRegion = ["io:City","io:Neighborhood","io:Province"];
const jerarquiaArchitecture = ["rec:Bathroom","rec:Bedroom","rec:Building","rec:OutdoorSpace","rec:Room","rec:Site"];

const query3 = `PREFIX owl: <http://www.w3.org/2002/07/owl#>
PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>
SELECT ?prop ?dom ?ran WHERE {
  ?prop a ?t .
  FILTER(?t IN (owl:ObjectProperty, owl:DatatypeProperty))
  OPTIONAL { ?prop rdfs:domain ?dom }
  OPTIONAL { ?prop rdfs:range ?ran }
} ORDER BY ?prop`;

// Propiedades clave con domain/range, foco en lo que se usará en SHACL
const propsClave = [
  ["gr:hasBusinessFunction", "gr:Offering", "gr:BusinessFunction"],
  ["gr:hasCurrency", "gr:PriceSpecification", "xsd:string"],
  ["gr:hasUnitOfMeasurement", "gr:QuantitativeValue", "xsd:string"],
  ["gr:hasValue", "gr:QuantitativeValue", "xsd:float"],
  ["gr:priceType", "gr:UnitPriceSpecification", "xsd:string"],
  ["sioc:about", "sioc:Item", "(sin rango)"],
  ["sioc:has_creator", "sioc:Post", "sioc:UserAccount"],
  ["sioc:has_space", "(sin dominio)", "sioc:Space"],
  ["sioc:id", "(sin dominio)", "rdfs:Literal"],
  ["sioc:read_at", "(sin dominio)", "rdfs:Literal"],
  ["io:hasDetail", "io:Feature", "(sin rango)"],
  ["io:hasFeature", "(union)", "io:Feature"],
  ["io:neighborhood", "io:PostalAddress", "io:Neighborhood"],
  ["io:province", "io:PostalAddress", "io:Province"],
  ["rec:city", "io:PostalAddress", "io:City"],
  ["rec:geometry", "rec:Space", "rec:Geometry"],
  ["rec:includes", "rec:RealEstate", "rec:Architecture"],
  ["rec:locatedIn", "(sin dominio)", "rec:Space"],
  ["time:hasTime", "(sin dominio)", "time:TemporalEntity"],
];

const propsSinDeclarar = [
  "io:curatedBy", "io:hasAVETime", "io:hasCurationTime", "io:hasScraperTime",
  "io:reviewedBy", "sioc:container_of", "sioc:has_owner", "sioc:space_of",
  "brick:isPartOf", "rec:includedIn", "rec:isLocationOf", "rec:coordinates",
  "foaf:account", "pronto:orientation", "pronto:size_type"
];

function propsTable(items) {
  const headerCell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true })] })]
  });
  const cell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, font: 'Courier New', size: 20 })] })]
  });
  const rows = [
    new TableRow({ children: [headerCell('Propiedad', 3640), headerCell('Dominio', 2860), headerCell('Rango', 2860)] }),
    ...items.map(it => new TableRow({ children: [cell(it[0], 3640), cell(it[1], 2860), cell(it[2], 2860)] }))
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3640, 2860, 2860],
    rows
  });
}

function violationsTable(items) {
  const headerCell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: "F5D6D6", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true })] })]
  });
  const cell = (txt, w, mono) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun(mono ? { text: txt, font: 'Courier New', size: 18 } : { text: txt, size: 18 })] })]
  });
  const widths = [500, 1900, 2200, 2360, 2400];
  const rows = [
    new TableRow({ children: [
      headerCell('ID', widths[0]),
      headerCell('Nodo', widths[1]),
      headerCell('Qué se rompió', widths[2]),
      headerCell('Componente SHACL', widths[3]),
      headerCell('Shape', widths[4])
    ] }),
    ...items.map(it => new TableRow({ children: [
      cell(it[0], widths[0], false),
      cell(it[1], widths[1], true),
      cell(it[2], widths[2], false),
      cell(it[3], widths[3], false),
      cell(it[4], widths[4], true)
    ] }))
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows
  });
}

function shapesTable(items) {
  const headerCell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true })] })]
  });
  const cell = (txt, w, mono) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun(mono ? { text: txt, font: 'Courier New', size: 20 } : { text: txt, size: 20 })] })]
  });
  const rows = [
    new TableRow({ children: [
      headerCell('#', 500),
      headerCell('Shape', 2500),
      headerCell('targetClass', 2360),
      headerCell('Restricciones principales', 4000)
    ] }),
    ...items.map(it => new TableRow({ children: [
      cell(it[0], 500, false),
      cell(it[1], 2500, true),
      cell(it[2], 2360, true),
      cell(it[3], 4000, false)
    ] }))
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [500, 2500, 2360, 4000],
    rows
  });
}

function subclasesTable(items, padre) {
  const headerCell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true })] })]
  });
  const cell = (txt, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, font: 'Courier New', size: 20 })] })]
  });
  const rows = [
    new TableRow({ children: [headerCell('Subclase', 4680), headerCell('Superclase', 4680)] }),
    ...items.map(it => new TableRow({ children: [cell(it, 4680), cell(padre, 4680)] }))
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    rows
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F3A5F" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E5C8A" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("TP SHACL — Seguimiento de queries")] }),
      p("Documento interno (no para entregar). Vamos registrando cada query SPARQL ejecutada en GraphDB, su resultado y la conclusión que sacamos."),
      p(""),
      p("Repositorio GraphDB: inmontology-tp"),
      p("Ontología cargada: inmontology.owl (del repo cientopolis/OVS-inmontology)"),

      h2("Query 1 — Listar todas las clases"),
      h3("SPARQL"),
      ...codeBlock(query1),
      p(""),
      h3("Resultado (65 clases)"),
      resultTable(clases),
      p(""),
      h3("Conclusión"),
      p("Las clases de la ontología están nombradas en INGLÉS, no en español como aparecía en el diagrama del README. Equivalencias clave:"),
      p("• io:House = Casa   • io:Apartment = Departamento   • io:Land = Terreno"),
      p("• io:CondominiumUnit = PH   • io:Store = Local   • io:Price = Precio"),
      p("• io:Address = Dirección   • io:Dimensions = Dimensiones   • io:Origin = Origen"),
      p(""),
      p("Subclases de RealEstate disponibles (más de las que mostraba el diagrama): House, Apartment, Land, CondominiumUnit, Store, Hotel, Farm, Warehouse, ParkingLot."),
      p("Subclases de Feature: Price, Address, Dimensions, Surface, RoomAmnt, FacadesAmnt, PropertyAge, SwimmingPool, CornerLot, IrregularLot, FOT."),
      p("Subclases de Origin: Scraper, AVE, Curation."),
      p("Subclases de Room/Space: Bathroom, Bedroom, OutdoorSpace, Building, Site."),
      p(""),
      p("→ Implicancia para el TP: tenemos material de sobra para definir 5+ NodeShape y 15+ sh:property con jerarquías reales (RealEstate → House/Apartment/..., Feature → Price/Address/..., Origin → Scraper/AVE/Curation)."),
      p(""),

      h2("Query 2 — Jerarquía de clases (rdfs:subClassOf)"),
      h3("SPARQL"),
      ...codeBlock(query2),
      p(""),
      h3("Resultado filtrado (100 filas totales; aquí las jerarquías relevantes para SHACL)"),
      h3("Subclases de rec:RealEstate"),
      subclasesTable(jerarquiaRealEstate, "rec:RealEstate"),
      p(""),
      h3("Subclases de io:Feature"),
      subclasesTable(jerarquiaFeature, "io:Feature"),
      p(""),
      h3("Subclases de io:Origin"),
      subclasesTable(jerarquiaOrigin, "io:Origin"),
      p(""),
      h3("Subclases de rec:Region"),
      subclasesTable(jerarquiaRegion, "rec:Region"),
      p(""),
      h3("Subclases de rec:Architecture"),
      subclasesTable(jerarquiaArchitecture, "rec:Architecture"),
      p(""),
      h3("Conclusión"),
      p("Tenemos jerarquías ricas y reales en la ontología, perfectas para cumplir el requisito \"analizar una jerarquía de clases\" del TP:"),
      p("• RealEstate → 9 subclases (House, Apartment, Land, CondominiumUnit, Store, Hotel, Farm, Warehouse, ParkingLot)."),
      p("• Feature → 12 subclases (Price, Address, Dimensions, Surface, RoomAmnt, FacadesAmnt, PropertyAge, SwimmingPool, CornerLot, IrregularLot, FOT, Neighborhood)."),
      p("• Origin → 3 subclases (Scraper, AVE, Curation)."),
      p("• Region → 3 subclases (City, Neighborhood, Province)."),
      p("• Architecture → 6 subclases (Bathroom, Bedroom, Building, OutdoorSpace, Room, Site)."),
      p(""),
      p("Hallazgos interesantes:"),
      p("• io:Neighborhood tiene HERENCIA MÚLTIPLE: es subClassOf io:Feature Y rec:Region. Buena curiosidad para mostrar en el video."),
      p("• Todas las subclases de RealEstate también son subClassOf rec:Collection (no solo RealEstate)."),
      p("• pronto:RealEstateListing hereda de gr:Offering, sioc:Post, sioc:Item y foaf:Document a la vez."),
      p(""),
      p("→ Decisión para shapes: usar RealEstateShape como shape \"padre\" y CasaShape (sobre io:House) o ApartmentShape como shape hija con sh:node para heredar restricciones. Esto evidencia la jerarquía en SHACL."),
      p(""),

      h2("Query 3 — Propiedades con dominio y rango"),
      h3("SPARQL"),
      ...codeBlock(query3),
      p(""),
      h3("Resultado: propiedades clave para SHACL (con domain/range declarados)"),
      propsTable(propsClave),
      p(""),
      h3("Propiedades SIN dominio o rango declarados"),
      p("Estas propiedades existen en la ontología pero no tienen rdfs:domain ni rdfs:range — quedan a interpretación. Justamente acá es donde más sentido tiene meter restricciones SHACL."),
      p(propsSinDeclarar.join(", ")),
      p(""),
      h3("Conclusión"),
      p("Mapeo directo a constraints SHACL (anotaciones para escribir shapes.ttl):"),
      p("• gr:hasBusinessFunction → sh:in (gr:Sell gr:LeaseOut) en RealEstateListingShape."),
      p("• gr:hasCurrency → sh:pattern \"^[A-Z]{3}$\" (ISO 4217) en PriceShape."),
      p("• gr:hasValue → sh:datatype xsd:float + sh:minExclusive 0 (precios positivos)."),
      p("• sioc:has_creator → sh:class sioc:UserAccount, sh:minCount 1 en RealEstateListingShape."),
      p("• sioc:about → sh:class rec:RealEstate, sh:minCount 1, sh:maxCount 1 (un listing habla de UN inmueble)."),
      p("• io:hasFeature → sh:class io:Feature, sh:minCount 1 (un inmueble debe tener al menos una feature)."),
      p("• rec:city/io:neighborhood/io:province → sh:class correspondiente, minCount apropiado en PostalAddressShape."),
      p("• rec:includes → sh:class rec:Architecture, sh:minCount 1 en RealEstateShape."),
      p(""),
      p("Observación importante para el video y el documento:"),
      p("La ontología tiene MUCHAS propiedades sin rdfs:domain/rdfs:range declarados (io:curatedBy, io:hasAVETime, io:reviewedBy, rec:coordinates, etc.). Esto significa que la ontología es PERMISIVA: en RDF cualquier nodo puede usarlas con cualquier tipo de valor. Esa es exactamente la justificación de fondo para escribir un SHACL — apretar lo que la ontología deja suelto. Buen punto para abrir el video."),
      p(""),

      h2("Query 4 — Verificación: subclases por superclase (consulta dirigida)"),
      h3("SPARQL"),
      ...codeBlock(`PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>
PREFIX rec: <https://w3id.org/rec#>
PREFIX io: <http://www.semanticweb.org/luciana/ontologies/2024/8/inmontology#>

# Repetir para rec:RealEstate, io:Feature, rec:Architecture
SELECT ?sub WHERE { ?sub rdfs:subClassOf rec:RealEstate }`),
      p(""),
      h3("Resultado — Subclases de rec:RealEstate (9)"),
      subclasesTable(jerarquiaRealEstate, "rec:RealEstate"),
      p(""),
      h3("Resultado — Subclases de io:Feature (12)"),
      subclasesTable(jerarquiaFeature, "io:Feature"),
      p(""),
      h3("Resultado — Subclases de rec:Architecture (6)"),
      subclasesTable(jerarquiaArchitecture, "rec:Architecture"),
      p(""),
      h3("Conclusión"),
      p("Verificación cruzada exitosa: las jerarquías coinciden exactamente con lo deducido en la Query 2. Listos para escribir shapes.ttl."),
      p(""),

      h2("Cierre del bloque de exploración"),
      p("Con las 4 queries cubrimos: clases existentes (Q1), jerarquías (Q2 y Q4 cruzada), y propiedades con sus dominios/rangos (Q3)."),
      p("Próximo paso: redactar shapes.ttl usando los nombres reales en inglés (House, Apartment, Price, Address, PostalAddress, etc.) y reflejar la jerarquía RealEstate → House con sh:node."),
      p(""),

      h2("Etapa 2 — Redacción de shapes.ttl"),
      h3("Archivo generado"),
      p("Ruta: tp-shacl/shapes.ttl"),
      p("Estado: parsea correctamente (128 tripletas). Verificado con rdflib."),
      p(""),
      h3("NodeShapes incluidas (7 en total — requisito mínimo: 5)"),
      shapesTable([
        ["1", ":RealEstateShape", "rec:RealEstate", "Shape padre. Todo inmueble debe tener ≥1 io:Feature y ≥1 rec:Architecture."],
        ["2", ":HouseShape", "io:House", "Shape HIJA: hereda de RealEstateShape vía sh:node + exige io:RoomAmnt (cumple jerarquía)."],
        ["3", ":RealEstateListingShape", "pronto:RealEstateListing", "Aviso: sioc:about (1..1 rec:RealEstate), sioc:has_creator, gr:hasBusinessFunction ∈ {Sell, LeaseOut}, sioc:id, ≥1 feature."],
        ["4", ":PriceShape", "io:Price", "Todo io:Price debe tener al menos un io:hasDetail (TemporalFeature)."],
        ["5", ":UnitPriceSpecificationShape", "gr:UnitPriceSpecification", "Moneda ISO 4217 (regex), monto float > 0, priceType ∈ {BASE, TOTAL}."],
        ["6", ":PostalAddressShape", "io:PostalAddress", "io:address no vacío, rec:city y io:province obligatorias, io:neighborhood opcional pero tipada."],
        ["7", ":OriginShape", "io:Origin", "Debe ser IRI nombrado y de tipo io:Scraper, io:AVE o io:Curation (sh:or)."],
      ]),
      p(""),
      h3("Conteo de restricciones"),
      p("• Total sh:property declaradas: 16 (requisito mínimo: 15) ✅"),
      p("• Constraints usadas (variedad): sh:minCount, sh:maxCount, sh:datatype, sh:class, sh:nodeKind, sh:pattern, sh:minExclusive, sh:minLength, sh:in, sh:or, sh:node, sh:qualifiedValueShape, sh:qualifiedMinCount, sh:message."),
      p("• Jerarquía: HouseShape sh:node :RealEstateShape — cumple el requisito 'analizar una jerarquía de clases' ✅"),
      p(""),
      h3("Decisiones de diseño tomadas"),
      p("• Prefijo de las shapes: : → http://example.org/shapes/inmontology# (namespace propio, no choca con OVS)."),
      p("• Se NO restringió io:Feature como NodeShape directa porque cada subclase (Price, Address, Dimensions, etc.) tiene su propia lógica — se atacan individualmente (PriceShape) o vía la subclase de RealEstate (HouseShape exige io:RoomAmnt)."),
      p("• El detalle temporal (io:TemporalFeature, io:hasScraperValue, io:hasScraperTime) queda implícito: PriceShape exige io:hasDetail pero no se restringe la estructura interna del detalle para no acoplarse a un origen específico (Scraper vs AVE vs Curation). Buen punto para trabajos futuros."),
      p("• OriginShape usa sh:or con sh:class para forzar que el origen sea una de las 3 subclases definidas (Scraper, AVE, Curation)."),
      p(""),
      h3("Cosas que quedaron como TRABAJOS FUTUROS"),
      p("• Coherencia superficie cubierta ≤ total (requiere sh:sparql, no constraints básicos)."),
      p("• Validar que sioc:read_at sea posterior o igual a dc:date del listing."),
      p("• Coherencia origen-timestamp: si origen es AVE debe haber io:hasAVETime; si es Scraper debe haber io:hasScraperTime; si es Curation debe haber io:hasCurationTime."),
      p("• Validar que las coordenadas geográficas (rec:coordinates) caigan dentro del bounding box de la io:Province declarada."),
      p(""),
      h3("Verificación realizada"),
      ...codeBlock(`$ python3 -c "import rdflib; g=rdflib.Graph(); g.parse('shapes.ttl'); print(len(g))"
128

NodeShapes encontradas: 7
  - :RealEstateShape
  - :HouseShape
  - :RealEstateListingShape
  - :PriceShape
  - :UnitPriceSpecificationShape
  - :PostalAddressShape
  - :OriginShape

sh:property declaradas: 16`),
      p(""),
      h3("Ajuste aplicado durante la validación"),
      p("En PriceShape la restricción sh:in para gr:priceType combinada con sh:datatype xsd:string generaba un falso positivo en pyshacl. Se reemplazó por:"),
      ...codeBlock(`sh:in ( "BASE"^^xsd:string "TOTAL"^^xsd:string )`),
      p("Funcionalmente equivalente, sin el quirk del motor."),
      p(""),

      h2("Etapa 3 — Datos de prueba y validación"),
      h3("Archivos generados"),
      p("• tp-shacl/data-valid.ttl — grafo que cumple todas las restricciones."),
      p("• tp-shacl/data-invalid.ttl — grafo que viola 8 restricciones distintas a propósito."),
      p("• tp-shacl/reports/valid-report.txt — reporte pyshacl del caso válido."),
      p("• tp-shacl/reports/invalid-report.txt — reporte pyshacl del caso inválido."),
      p(""),
      h3("Comando de validación usado"),
      ...codeBlock(`python3 -m pyshacl -s tp-shacl/shapes.ttl \\
                  -d tp-shacl/data-valid.ttl \\
                  -e inmontology.owl \\
                  --inference rdfs \\
                  -f human`),
      p(""),
      p("Notar el flag -e inmontology.owl: pyshacl necesita la ontología base para saber, por ejemplo, que io:Price es subclase de io:Feature. Sin eso falla por error de tipos. Con --inference rdfs aplica las inferencias rdfs:subClassOf automáticamente."),
      p(""),
      h3("Resultado del caso VÁLIDO"),
      ...codeBlock(`Validation Report
Conforms: True`),
      p("→ El grafo data-valid.ttl pasa las 16 restricciones definidas en shapes.ttl."),
      p(""),
      h3("Resultado del caso INVÁLIDO (8 violaciones diseñadas)"),
      ...codeBlock(`Validation Report
Conforms: False
Results (9)`),
      p("(9 = 8 violaciones diseñadas, pero la del Origin genera 2 reportes porque sh:or registra cada sh:class fallido por separado)"),
      p(""),
      h3("Tabla: violación diseñada → componente SHACL detectado"),
      violationsTable([
        ["V1", "ex:price_value_bad", "gr:hasCurrency = 'dolares'", "PatternConstraintComponent (sh:pattern)", "UnitPriceSpecificationShape"],
        ["V2", "ex:price_value_bad", "gr:hasCurrencyValue = -50000.0", "MinExclusiveConstraintComponent", "UnitPriceSpecificationShape"],
        ["V3", "ex:listing_bad", "sin sioc:has_creator", "MinCountConstraintComponent", "RealEstateListingShape"],
        ["V4", "ex:listing_bad", "gr:hasBusinessFunction = gr:Repair", "InConstraintComponent (sh:in)", "RealEstateListingShape"],
        ["V5", "ex:house_bad", "Casa sin io:RoomAmnt entre sus features", "QualifiedMinCountConstraintComponent", "HouseShape"],
        ["V6", "ex:postal_address_bad", "sin io:province", "MinCountConstraintComponent", "PostalAddressShape"],
        ["V7", "ex:postal_address_bad", "io:neighborhood = 'Centro' (string)", "ClassConstraintComponent (sh:class)", "PostalAddressShape"],
        ["V8", "blank node de io:Origin", "blank node + no es Scraper/AVE/Curation", "NodeKindConstraintComponent + OrConstraintComponent", "OriginShape"],
      ]),
      p(""),
      h3("Conclusión"),
      p("Las dos validaciones funcionan como se esperaba:"),
      p("• El grafo válido es aceptado (Conforms: True)."),
      p("• El grafo inválido es rechazado y SHACL identifica exactamente qué restricción se violó, en qué nodo y por qué."),
      p("• El uso de variedad de constraints (pattern, minExclusive, minCount, in, class, qualifiedMinCount, nodeKind, or) demuestra el alcance de SHACL para validar grafos de conocimiento."),
      p(""),
      h3("Próximo paso"),
      p("• Repetir esta validación dentro de GraphDB (cargar shapes.ttl como SHACL shapes graph + importar los dos data.ttl) para tener los screenshots para el video."),
      p("• Armar el documento final con: integrantes, URL del video, tabla de restricciones con justificación (sale de los comentarios de shapes.ttl), y los screenshots."),
      p("• Grabar el video con los 4 integrantes."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const out = '/Users/giulialvaro/Library/CloudStorage/OneDrive-Personal/repositorios/OVS-inmontology-tp/tp-shacl/seguimiento.docx';
  fs.writeFileSync(out, buffer);
  console.log('Wrote ' + out);
});
