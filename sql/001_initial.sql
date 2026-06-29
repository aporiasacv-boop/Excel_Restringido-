-- Esquema base: tabla de control de versiones del sistema
CREATE TABLE IF NOT EXISTS system_info (
    id SERIAL PRIMARY KEY,
    database_version VARCHAR(50) NOT NULL,
    application_version VARCHAR(50) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMP NOT NULL DEFAULT NOW()
);
