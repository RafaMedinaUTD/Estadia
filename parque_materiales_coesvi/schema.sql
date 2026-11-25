-- Catálogo de Materiales / Existencias
CREATE TABLE IF NOT EXISTS materiales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT,
    nombre TEXT NOT NULL,
    unidad TEXT NOT NULL,
    stock_actual REAL DEFAULT 0.0
);

-- Encabezado de Notas de Entrada (FOR-CSV-DAD-PQM-01)
CREATE TABLE IF NOT EXISTS entradas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folio TEXT NOT NULL,
    fecha TEXT NOT NULL,
    proveedor TEXT,
    factura_no TEXT,
    remision_no TEXT
);

-- Detalle de los materiales que ingresan
CREATE TABLE IF NOT EXISTS detalle_entradas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entrada_id INTEGER NOT NULL,
    cantidad REAL NOT NULL,
    unidad TEXT NOT NULL,
    descripcion TEXT NOT NULL,
    importe REAL DEFAULT 0.0,
    FOREIGN KEY (entrada_id) REFERENCES entradas(id) ON DELETE CASCADE
);

-- Encabezado de Notas de Salida (FOR-CSV-DAD-PQM-02)
CREATE TABLE IF NOT EXISTS salidas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folio TEXT NOT NULL,
    fecha TEXT NOT NULL,
    beneficiario TEXT NOT NULL,
    recibe_en TEXT
);

-- Detalle de los materiales que salen
CREATE TABLE IF NOT EXISTS detalle_salidas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    salida_id INTEGER NOT NULL,
    cantidad REAL NOT NULL,
    unidad TEXT NOT NULL,
    descripcion TEXT NOT NULL,
    importe REAL DEFAULT 0.0,
    FOREIGN KEY (salida_id) REFERENCES salidas(id) ON DELETE CASCADE
);