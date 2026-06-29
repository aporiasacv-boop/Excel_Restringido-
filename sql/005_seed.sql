-- Datos iniciales: solo aplica en instalación nueva (base vacía)
CREATE EXTENSION IF NOT EXISTS pgcrypto;

INSERT INTO roles (name, description)
VALUES ('admin', 'Administrador del sistema');

INSERT INTO users (username, password_hash, full_name, role_id, must_change_password)
SELECT
    'admin',
    crypt('Admin123!', gen_salt('bf')),
    'Administrador',
    r.id,
    TRUE
FROM roles r
WHERE r.name = 'admin'
  AND NOT EXISTS (SELECT 1 FROM users WHERE username = 'admin');

INSERT INTO system_info (database_version, application_version)
SELECT '1.1.0', '0.3.0'
WHERE NOT EXISTS (SELECT 1 FROM system_info);
