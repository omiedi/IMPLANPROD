-- Script to create RECALMAQ and RECALMOD tables

CREATE TABLE RECALMAQ (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Equipo VARCHAR(6) NOT NULL,
    DescripcionEquipo VARCHAR(48) NULL,
    CodigoPieza INT NULL,
    DescripcionPieza VARCHAR(64) NULL,
    NumeroOperacion INT NULL,
    DescripcionOperacion VARCHAR(48) NULL,
    CantidadProgramada DECIMAL(18,2) NULL,
    CodigoProductoDestino INT NULL,
    HorasNecesarias DECIMAL(18,4) NULL,
    TotalHorasPorEquipo DECIMAL(18,4) NULL,
    NumeroPrograma INT NOT NULL,
    FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
    -- Campos de auditoría
    AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL
);

CREATE TABLE RECALMOD (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Categoria INT NOT NULL,
    DescripcionCategoria VARCHAR(48) NULL,
    HorasNecesarias DECIMAL(18,4) NULL,
    TotalHorasPorCategoria DECIMAL(18,4) NULL,
    NumeroPrograma INT NOT NULL,
    FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
    -- Campos de auditoría
    AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL
);
