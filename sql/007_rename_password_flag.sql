-- Migración: renombrar force_password_change a must_change_password
-- Solo ejecutar si la columna anterior aún existe
ALTER TABLE users RENAME COLUMN force_password_change TO must_change_password;
