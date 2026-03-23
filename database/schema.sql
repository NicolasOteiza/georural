CREATE DATABASE IF NOT EXISTS geo_rural CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE geo_rural;

CREATE TABLE IF NOT EXISTS registros (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    num_ingreso VARCHAR(20) NOT NULL,
    correlativo INT UNSIGNED NOT NULL,
    anio SMALLINT UNSIGNED NOT NULL,
    nombre VARCHAR(255) NOT NULL,
    rut VARCHAR(20) NOT NULL,
    telefono VARCHAR(40) NULL,
    correo VARCHAR(255) NULL,
    region VARCHAR(120) NOT NULL,
    comuna VARCHAR(120) NOT NULL,
    nombre_predio VARCHAR(255) NULL,
    rol VARCHAR(120) NOT NULL,
    num_lotes INT NULL,
    estado VARCHAR(40) NULL,
    comentario TEXT NULL,
    created_by VARCHAR(120) NULL,
    created_by_sucursal VARCHAR(120) NULL,
    updated_by VARCHAR(120) NULL,
    updated_by_sucursal VARCHAR(120) NULL,
    documentos JSON NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_num_ingreso (num_ingreso),
    UNIQUE KEY uq_anio_correlativo (anio, correlativo),
    KEY idx_nombre (nombre),
    KEY idx_rut (rut),
    KEY idx_rol (rol)
);

CREATE TABLE IF NOT EXISTS usuarios (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    username VARCHAR(80) NOT NULL,
    nombre VARCHAR(120) NOT NULL,
    sucursal VARCHAR(120) NOT NULL,
    role VARCHAR(20) NOT NULL DEFAULT 'OPERADOR',
    password_salt CHAR(32) NOT NULL,
    password_hash CHAR(128) NOT NULL,
    must_change_password TINYINT(1) NOT NULL DEFAULT 0,
    activo TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_usuarios_username (username)
);

CREATE TABLE IF NOT EXISTS sesiones (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    user_id INT UNSIGNED NOT NULL,
    token_hash CHAR(64) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_used_at DATETIME NULL DEFAULT NULL,
    expires_at DATETIME NOT NULL,
    revoked_at DATETIME NULL DEFAULT NULL,
    PRIMARY KEY (id),
    UNIQUE KEY uq_sesiones_token_hash (token_hash),
    KEY idx_sesiones_user (user_id),
    KEY idx_sesiones_expires (expires_at),
    CONSTRAINT fk_sesiones_user_id FOREIGN KEY (user_id) REFERENCES usuarios(id)
        ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS registro_historial (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    registro_id INT UNSIGNED NOT NULL,
    num_ingreso VARCHAR(20) NOT NULL,
    accion VARCHAR(40) NOT NULL,
    comentario TEXT NOT NULL,
    usuario_nombre VARCHAR(120) NOT NULL,
    sucursal VARCHAR(120) NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    KEY idx_historial_registro_fecha (registro_id, created_at),
    KEY idx_historial_num_ingreso (num_ingreso),
    CONSTRAINT fk_historial_registro_id FOREIGN KEY (registro_id) REFERENCES registros(id)
        ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS sucursales (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    nombre VARCHAR(120) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_sucursales_nombre (nombre)
);

CREATE TABLE IF NOT EXISTS utm_mensual (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    anio SMALLINT UNSIGNED NOT NULL,
    mes TINYINT UNSIGNED NOT NULL,
    valor DECIMAL(14,2) NOT NULL,
    registrado_por_id INT UNSIGNED NULL,
    registrado_por_nombre VARCHAR(120) NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_utm_mensual_periodo (anio, mes),
    KEY idx_utm_mensual_registrado_por_id (registrado_por_id)
);

CREATE TABLE IF NOT EXISTS cotizacion_servicios_estandar (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    nombre_servicio VARCHAR(255) NOT NULL,
    valor_utm DECIMAL(14,2) NOT NULL,
    orden_visual SMALLINT UNSIGNED NOT NULL DEFAULT 1,
    activo TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    KEY idx_cotizacion_servicios_activo_orden (activo, orden_visual)
);

CREATE TABLE IF NOT EXISTS correo_configuracion (
    id TINYINT UNSIGNED NOT NULL,
    smtp_host VARCHAR(255) NOT NULL DEFAULT '',
    smtp_port SMALLINT UNSIGNED NOT NULL DEFAULT 587,
    smtp_secure TINYINT(1) NOT NULL DEFAULT 0,
    smtp_user VARCHAR(255) NULL,
    smtp_password VARCHAR(255) NULL,
    smtp_from_email VARCHAR(255) NOT NULL DEFAULT '',
    smtp_from_name VARCHAR(120) NULL,
    updated_by_id INT UNSIGNED NULL,
    updated_by_nombre VARCHAR(120) NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    KEY idx_correo_config_updated_by (updated_by_id)
);
