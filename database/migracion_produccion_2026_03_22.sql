/*
  Migracion de estructura para Geo Rural
  Fecha: 2026-03-22
  Objetivo: aplicar cambios de BD sin borrar registros existentes.
  Nota: este script NO hace DROP/TRUNCATE/DELETE de tablas de negocio.
*/

SET NAMES utf8mb4;
SET SESSION sql_mode = REPLACE(@@sql_mode, 'ONLY_FULL_GROUP_BY', '');

CREATE DATABASE IF NOT EXISTS `geo_rural` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `geo_rural`;

/* -------------------------------------------------------------------------- */
/*  Compatibilidad: renombrar tablas antiguas antes de CREATE TABLE            */
/* -------------------------------------------------------------------------- */

DELIMITER $$

DROP PROCEDURE IF EXISTS `rename_table_if_needed`$$
CREATE PROCEDURE `rename_table_if_needed`(
    IN p_old VARCHAR(64),
    IN p_new VARCHAR(64)
)
BEGIN
    DECLARE v_old_exists INT DEFAULT 0;
    DECLARE v_new_exists INT DEFAULT 0;

    SELECT COUNT(*)
      INTO v_old_exists
      FROM information_schema.TABLES
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = p_old;

    SELECT COUNT(*)
      INTO v_new_exists
      FROM information_schema.TABLES
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = p_new;

    IF v_old_exists = 1 AND v_new_exists = 0 THEN
        SET @sql = CONCAT('RENAME TABLE `', p_old, '` TO `', p_new, '`');
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    ELSEIF v_old_exists = 1 AND v_new_exists = 1 THEN
        SELECT CONCAT(
            'Aviso: existen ',
            p_old,
            ' y ',
            p_new,
            '. No se renombro para evitar sobrescribir datos.'
        ) AS aviso;
    END IF;
END$$

DELIMITER ;

CALL rename_table_if_needed('historial_registros', 'registro_historial');
CALL rename_table_if_needed('registros_historial', 'registro_historial');
CALL rename_table_if_needed('facturas_solicitudes', 'factura_solicitudes');
CALL rename_table_if_needed('factura_solicitud', 'factura_solicitudes');
CALL rename_table_if_needed('cotizacion_servicios', 'cotizacion_servicios_estandar');
CALL rename_table_if_needed('servicios_cotizacion', 'cotizacion_servicios_estandar');
CALL rename_table_if_needed('configuracion_correo', 'correo_configuracion');
CALL rename_table_if_needed('correo_config', 'correo_configuracion');
CALL rename_table_if_needed('sesiones_usuarios', 'sesiones');
CALL rename_table_if_needed('usuario_sesiones', 'sesiones');

DROP PROCEDURE IF EXISTS `rename_table_if_needed`;

/* -------------------------------------------------------------------------- */
/*  Tablas principales (create if not exists)                                  */
/* -------------------------------------------------------------------------- */

CREATE TABLE IF NOT EXISTS `registros` (
    `id` INT UNSIGNED NOT NULL AUTO_INCREMENT,
    `num_ingreso` VARCHAR(20) NOT NULL,
    `correlativo` INT UNSIGNED NOT NULL,
    `anio` SMALLINT UNSIGNED NOT NULL,
    `nombre` VARCHAR(255) NOT NULL,
    `rut` VARCHAR(20) NOT NULL,
    `telefono` VARCHAR(40) NULL,
    `correo` VARCHAR(255) NULL,
    `region` VARCHAR(120) NOT NULL,
    `comuna` VARCHAR(120) NOT NULL,
    `nombre_predio` VARCHAR(255) NULL,
    `rol` VARCHAR(120) NOT NULL,
    `num_lotes` INT NULL,
    `estado` VARCHAR(40) NULL,
    `comentario` TEXT NULL,
    `factura_nombre_razon` VARCHAR(255) NULL,
    `factura_rut` VARCHAR(20) NULL,
    `factura_giro` VARCHAR(255) NULL,
    `factura_direccion` VARCHAR(255) NULL,
    `factura_comuna` VARCHAR(120) NULL,
    `factura_ciudad` VARCHAR(120) NULL,
    `factura_contacto` VARCHAR(255) NULL,
    `factura_observacion` TEXT NULL,
    `factura_monto` DECIMAL(14,2) NULL,
    `created_by` VARCHAR(120) NULL,
    `created_by_sucursal` VARCHAR(120) NULL,
    `updated_by` VARCHAR(120) NULL,
    `updated_by_sucursal` VARCHAR(120) NULL,
    `documentos` JSON NOT NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    UNIQUE KEY `uq_num_ingreso` (`num_ingreso`),
    UNIQUE KEY `uq_anio_correlativo` (`anio`, `correlativo`),
    KEY `idx_nombre` (`nombre`),
    KEY `idx_rut` (`rut`),
    KEY `idx_rol` (`rol`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `usuarios` (
    `id` INT UNSIGNED NOT NULL AUTO_INCREMENT,
    `username` VARCHAR(80) NOT NULL,
    `nombre` VARCHAR(120) NOT NULL,
    `sucursal` VARCHAR(120) NOT NULL,
    `role` VARCHAR(20) NOT NULL DEFAULT 'OPERADOR',
    `password_salt` CHAR(32) NOT NULL,
    `password_hash` CHAR(128) NOT NULL,
    `must_change_password` TINYINT(1) NOT NULL DEFAULT 0,
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    UNIQUE KEY `uq_usuarios_username` (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `sesiones` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `user_id` INT UNSIGNED NOT NULL,
    `token_hash` CHAR(64) NOT NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `last_used_at` DATETIME NULL DEFAULT NULL,
    `expires_at` DATETIME NOT NULL,
    `revoked_at` DATETIME NULL DEFAULT NULL,
    PRIMARY KEY (`id`),
    UNIQUE KEY `uq_sesiones_token_hash` (`token_hash`),
    KEY `idx_sesiones_user` (`user_id`),
    KEY `idx_sesiones_expires` (`expires_at`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `registro_historial` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `registro_id` INT UNSIGNED NOT NULL,
    `num_ingreso` VARCHAR(20) NOT NULL,
    `accion` VARCHAR(40) NOT NULL,
    `comentario` TEXT NOT NULL,
    `usuario_nombre` VARCHAR(120) NOT NULL,
    `sucursal` VARCHAR(120) NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    KEY `idx_historial_registro_fecha` (`registro_id`, `created_at`),
    KEY `idx_historial_num_ingreso` (`num_ingreso`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `sucursales` (
    `id` INT UNSIGNED NOT NULL AUTO_INCREMENT,
    `nombre` VARCHAR(120) NOT NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    UNIQUE KEY `uq_sucursales_nombre` (`nombre`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `factura_solicitudes` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `num_ingreso` VARCHAR(20) NULL,
    `nombre_razon_social` VARCHAR(255) NOT NULL,
    `rut` VARCHAR(20) NOT NULL,
    `giro` VARCHAR(255) NULL,
    `direccion` VARCHAR(255) NULL,
    `comuna` VARCHAR(120) NULL,
    `ciudad` VARCHAR(120) NULL,
    `contacto` VARCHAR(255) NULL,
    `observacion` TEXT NULL,
    `monto_facturar` DECIMAL(14,2) NULL,
    `destino_email` VARCHAR(255) NULL,
    `estado` VARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
    `created_by` VARCHAR(120) NULL,
    `created_by_sucursal` VARCHAR(120) NULL,
    `updated_by` VARCHAR(120) NULL,
    `updated_by_sucursal` VARCHAR(120) NULL,
    `sent_at` DATETIME NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    KEY `idx_factura_solicitudes_estado` (`estado`),
    KEY `idx_factura_solicitudes_created_at` (`created_at`),
    KEY `idx_factura_solicitudes_num_ingreso` (`num_ingreso`),
    KEY `idx_factura_solicitudes_created_sucursal` (`created_by_sucursal`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `utm_mensual` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `anio` SMALLINT UNSIGNED NOT NULL,
    `mes` TINYINT UNSIGNED NOT NULL,
    `valor` DECIMAL(14,2) NOT NULL,
    `registrado_por_id` INT UNSIGNED NULL,
    `registrado_por_nombre` VARCHAR(120) NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    UNIQUE KEY `uq_utm_mensual_periodo` (`anio`, `mes`),
    KEY `idx_utm_mensual_registrado_por_id` (`registrado_por_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `cotizacion_servicios_estandar` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `nombre_servicio` VARCHAR(255) NOT NULL,
    `valor_utm` DECIMAL(14,2) NOT NULL,
    `orden_visual` SMALLINT UNSIGNED NOT NULL DEFAULT 1,
    `activo` TINYINT(1) NOT NULL DEFAULT 1,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    KEY `idx_cotizacion_servicios_activo_orden` (`activo`, `orden_visual`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `correo_configuracion` (
    `id` TINYINT UNSIGNED NOT NULL,
    `smtp_host` VARCHAR(255) NOT NULL DEFAULT '',
    `smtp_port` SMALLINT UNSIGNED NOT NULL DEFAULT 587,
    `smtp_secure` TINYINT(1) NOT NULL DEFAULT 0,
    `smtp_user` VARCHAR(255) NULL,
    `smtp_password` VARCHAR(255) NULL,
    `smtp_from_email` VARCHAR(255) NOT NULL DEFAULT '',
    `smtp_from_name` VARCHAR(120) NULL,
    `updated_by_id` INT UNSIGNED NULL,
    `updated_by_nombre` VARCHAR(120) NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    KEY `idx_correo_config_updated_by` (`updated_by_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

/* -------------------------------------------------------------------------- */
/*  Helpers de migracion                                                       */
/* -------------------------------------------------------------------------- */

DELIMITER $$

DROP PROCEDURE IF EXISTS `ensure_column`$$
CREATE PROCEDURE `ensure_column`(
    IN p_table VARCHAR(64),
    IN p_column VARCHAR(64),
    IN p_definition TEXT
)
BEGIN
    DECLARE v_exists INT DEFAULT 0;

    SELECT COUNT(*)
      INTO v_exists
      FROM information_schema.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = p_table
       AND COLUMN_NAME = p_column;

    IF v_exists = 0 THEN
        SET @sql = CONCAT('ALTER TABLE `', p_table, '` ADD COLUMN ', p_definition);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$

DROP PROCEDURE IF EXISTS `ensure_index`$$
CREATE PROCEDURE `ensure_index`(
    IN p_table VARCHAR(64),
    IN p_index VARCHAR(64),
    IN p_add_clause TEXT
)
BEGIN
    DECLARE v_exists INT DEFAULT 0;

    SELECT COUNT(*)
      INTO v_exists
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = p_table
       AND INDEX_NAME = p_index;

    IF v_exists = 0 THEN
        SET @sql = CONCAT('ALTER TABLE `', p_table, '` ', p_add_clause);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$

DROP PROCEDURE IF EXISTS `ensure_fk_safe`$$
CREATE PROCEDURE `ensure_fk_safe`(
    IN p_table VARCHAR(64),
    IN p_constraint VARCHAR(64),
    IN p_add_clause TEXT
)
BEGIN
    DECLARE v_exists INT DEFAULT 0;
    DECLARE v_failed INT DEFAULT 0;
    DECLARE CONTINUE HANDLER FOR SQLEXCEPTION SET v_failed = 1;

    SELECT COUNT(*)
      INTO v_exists
      FROM information_schema.TABLE_CONSTRAINTS
     WHERE CONSTRAINT_SCHEMA = DATABASE()
       AND TABLE_NAME = p_table
       AND CONSTRAINT_NAME = p_constraint
       AND CONSTRAINT_TYPE = 'FOREIGN KEY';

    IF v_exists = 0 THEN
        SET @sql = CONCAT('ALTER TABLE `', p_table, '` ', p_add_clause);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;

        IF v_failed = 1 THEN
            SELECT CONCAT('Aviso: no se pudo crear FK ', p_constraint, ' en ', p_table, '. Revisar datos huerfanos.') AS aviso;
        END IF;
    END IF;
END$$

DELIMITER ;

/* -------------------------------------------------------------------------- */
/*  Compatibilidad: agregar columnas faltantes en instalaciones antiguas       */
/* -------------------------------------------------------------------------- */

CALL ensure_column('registros', 'correlativo', '`correlativo` INT UNSIGNED NULL AFTER `num_ingreso`');
CALL ensure_column('registros', 'anio', '`anio` SMALLINT UNSIGNED NULL AFTER `correlativo`');
CALL ensure_column('registros', 'created_by', '`created_by` VARCHAR(120) NULL AFTER `comentario`');
CALL ensure_column('registros', 'created_by_sucursal', '`created_by_sucursal` VARCHAR(120) NULL AFTER `created_by`');
CALL ensure_column('registros', 'updated_by', '`updated_by` VARCHAR(120) NULL AFTER `created_by_sucursal`');
CALL ensure_column('registros', 'updated_by_sucursal', '`updated_by_sucursal` VARCHAR(120) NULL AFTER `updated_by`');
CALL ensure_column('registros', 'factura_nombre_razon', '`factura_nombre_razon` VARCHAR(255) NULL AFTER `comentario`');
CALL ensure_column('registros', 'factura_rut', '`factura_rut` VARCHAR(20) NULL AFTER `factura_nombre_razon`');
CALL ensure_column('registros', 'factura_giro', '`factura_giro` VARCHAR(255) NULL AFTER `factura_rut`');
CALL ensure_column('registros', 'factura_direccion', '`factura_direccion` VARCHAR(255) NULL AFTER `factura_giro`');
CALL ensure_column('registros', 'factura_comuna', '`factura_comuna` VARCHAR(120) NULL AFTER `factura_direccion`');
CALL ensure_column('registros', 'factura_ciudad', '`factura_ciudad` VARCHAR(120) NULL AFTER `factura_comuna`');
CALL ensure_column('registros', 'factura_contacto', '`factura_contacto` VARCHAR(255) NULL AFTER `factura_ciudad`');
CALL ensure_column('registros', 'factura_observacion', '`factura_observacion` TEXT NULL AFTER `factura_contacto`');
CALL ensure_column('registros', 'factura_monto', '`factura_monto` DECIMAL(14,2) NULL AFTER `factura_observacion`');
CALL ensure_column('registros', 'documentos', '`documentos` JSON NULL AFTER `updated_by_sucursal`');

CALL ensure_column('usuarios', 'role', '`role` VARCHAR(20) NOT NULL DEFAULT ''OPERADOR'' AFTER `sucursal`');
CALL ensure_column('usuarios', 'must_change_password', '`must_change_password` TINYINT(1) NOT NULL DEFAULT 0 AFTER `password_hash`');
CALL ensure_column('usuarios', 'activo', '`activo` TINYINT(1) NOT NULL DEFAULT 1 AFTER `must_change_password`');

CALL ensure_column('sesiones', 'last_used_at', '`last_used_at` DATETIME NULL DEFAULT NULL AFTER `created_at`');
CALL ensure_column('sesiones', 'expires_at', '`expires_at` DATETIME NULL DEFAULT NULL AFTER `last_used_at`');
CALL ensure_column('sesiones', 'revoked_at', '`revoked_at` DATETIME NULL DEFAULT NULL AFTER `expires_at`');

CALL ensure_column('registro_historial', 'sucursal', '`sucursal` VARCHAR(120) NULL AFTER `usuario_nombre`');

CALL ensure_column('factura_solicitudes', 'num_ingreso', '`num_ingreso` VARCHAR(20) NULL AFTER `id`');
CALL ensure_column('factura_solicitudes', 'nombre_razon_social', '`nombre_razon_social` VARCHAR(255) NOT NULL AFTER `num_ingreso`');
CALL ensure_column('factura_solicitudes', 'rut', '`rut` VARCHAR(20) NOT NULL AFTER `nombre_razon_social`');
CALL ensure_column('factura_solicitudes', 'giro', '`giro` VARCHAR(255) NULL AFTER `rut`');
CALL ensure_column('factura_solicitudes', 'direccion', '`direccion` VARCHAR(255) NULL AFTER `giro`');
CALL ensure_column('factura_solicitudes', 'comuna', '`comuna` VARCHAR(120) NULL AFTER `direccion`');
CALL ensure_column('factura_solicitudes', 'ciudad', '`ciudad` VARCHAR(120) NULL AFTER `comuna`');
CALL ensure_column('factura_solicitudes', 'contacto', '`contacto` VARCHAR(255) NULL AFTER `ciudad`');
CALL ensure_column('factura_solicitudes', 'observacion', '`observacion` TEXT NULL AFTER `contacto`');
CALL ensure_column('factura_solicitudes', 'monto_facturar', '`monto_facturar` DECIMAL(14,2) NULL AFTER `observacion`');
CALL ensure_column('factura_solicitudes', 'destino_email', '`destino_email` VARCHAR(255) NULL AFTER `monto_facturar`');
CALL ensure_column('factura_solicitudes', 'estado', '`estado` VARCHAR(20) NOT NULL DEFAULT ''PENDIENTE'' AFTER `destino_email`');
CALL ensure_column('factura_solicitudes', 'created_by', '`created_by` VARCHAR(120) NULL AFTER `estado`');
CALL ensure_column('factura_solicitudes', 'created_by_sucursal', '`created_by_sucursal` VARCHAR(120) NULL AFTER `created_by`');
CALL ensure_column('factura_solicitudes', 'updated_by', '`updated_by` VARCHAR(120) NULL AFTER `created_by_sucursal`');
CALL ensure_column('factura_solicitudes', 'updated_by_sucursal', '`updated_by_sucursal` VARCHAR(120) NULL AFTER `updated_by`');
CALL ensure_column('factura_solicitudes', 'sent_at', '`sent_at` DATETIME NULL AFTER `updated_by_sucursal`');

CALL ensure_column('utm_mensual', 'registrado_por_id', '`registrado_por_id` INT UNSIGNED NULL AFTER `valor`');
CALL ensure_column('utm_mensual', 'registrado_por_nombre', '`registrado_por_nombre` VARCHAR(120) NULL AFTER `registrado_por_id`');

CALL ensure_column(
    'cotizacion_servicios_estandar',
    'nombre_servicio',
    '`nombre_servicio` VARCHAR(255) NOT NULL DEFAULT '''' AFTER `id`'
);
CALL ensure_column(
    'cotizacion_servicios_estandar',
    'valor_utm',
    '`valor_utm` DECIMAL(14,2) NOT NULL DEFAULT 1.00 AFTER `nombre_servicio`'
);
CALL ensure_column('cotizacion_servicios_estandar', 'orden_visual', '`orden_visual` SMALLINT UNSIGNED NOT NULL DEFAULT 1 AFTER `valor_utm`');
CALL ensure_column('cotizacion_servicios_estandar', 'activo', '`activo` TINYINT(1) NOT NULL DEFAULT 1 AFTER `orden_visual`');

CALL ensure_column('correo_configuracion', 'smtp_host', '`smtp_host` VARCHAR(255) NOT NULL DEFAULT '''' AFTER `id`');
CALL ensure_column('correo_configuracion', 'smtp_port', '`smtp_port` SMALLINT UNSIGNED NOT NULL DEFAULT 587 AFTER `smtp_host`');
CALL ensure_column('correo_configuracion', 'smtp_secure', '`smtp_secure` TINYINT(1) NOT NULL DEFAULT 0 AFTER `smtp_port`');
CALL ensure_column('correo_configuracion', 'smtp_user', '`smtp_user` VARCHAR(255) NULL AFTER `smtp_secure`');
CALL ensure_column('correo_configuracion', 'smtp_password', '`smtp_password` VARCHAR(255) NULL AFTER `smtp_user`');
CALL ensure_column('correo_configuracion', 'smtp_from_email', '`smtp_from_email` VARCHAR(255) NOT NULL DEFAULT '''' AFTER `smtp_password`');
CALL ensure_column('correo_configuracion', 'smtp_from_name', '`smtp_from_name` VARCHAR(120) NULL AFTER `smtp_from_email`');
CALL ensure_column('correo_configuracion', 'updated_by_id', '`updated_by_id` INT UNSIGNED NULL AFTER `smtp_from_name`');
CALL ensure_column('correo_configuracion', 'updated_by_nombre', '`updated_by_nombre` VARCHAR(120) NULL AFTER `updated_by_id`');

/* -------------------------------------------------------------------------- */
/*  Normalizacion de datos (sin perdida)                                       */
/* -------------------------------------------------------------------------- */

UPDATE `usuarios`
   SET `role` = UPPER(TRIM(COALESCE(`role`, '')))
 WHERE `role` IS NOT NULL;

UPDATE `usuarios`
   SET `role` = 'OPERADOR'
 WHERE `role` IS NULL
    OR TRIM(`role`) = ''
    OR `role` NOT IN ('SUPER', 'ADMIN', 'SECRETARIA', 'OPERADOR', 'SUPERVISOR');

UPDATE `usuarios`
   SET `must_change_password` = 0
 WHERE `must_change_password` IS NULL;

SET @first_super_id := (
    SELECT `id`
      FROM `usuarios`
     WHERE UPPER(COALESCE(`role`, '')) = 'SUPER'
     ORDER BY `id` ASC
     LIMIT 1
);

UPDATE `usuarios`
   SET `role` = 'ADMIN'
 WHERE UPPER(COALESCE(`role`, '')) = 'SUPER'
   AND `id` <> COALESCE(@first_super_id, -1);

UPDATE `usuarios`
   SET `activo` = 1
 WHERE `activo` IS NULL;

UPDATE `registros`
   SET `correlativo` = CAST(SUBSTRING_INDEX(`num_ingreso`, '-', 1) AS UNSIGNED),
       `anio` = CAST(SUBSTRING_INDEX(`num_ingreso`, '-', -1) AS UNSIGNED)
 WHERE (`correlativo` IS NULL OR `correlativo` = 0 OR `anio` IS NULL OR `anio` = 0)
   AND `num_ingreso` REGEXP '^[0-9]+-[0-9]{4}$';

UPDATE `registros`
   SET `documentos` = JSON_ARRAY()
 WHERE `documentos` IS NULL;

UPDATE `factura_solicitudes`
   SET `num_ingreso` = NULL
 WHERE `num_ingreso` IS NOT NULL
   AND TRIM(`num_ingreso`) = '';

UPDATE `factura_solicitudes`
   SET `estado` = 'PENDIENTE'
 WHERE `estado` IS NULL
    OR TRIM(`estado`) = '';

UPDATE `cotizacion_servicios_estandar`
   SET `activo` = 1
 WHERE `activo` IS NULL;

UPDATE `cotizacion_servicios_estandar`
   SET `orden_visual` = 1
 WHERE `orden_visual` IS NULL
    OR `orden_visual` = 0;

UPDATE `cotizacion_servicios_estandar`
   SET `valor_utm` = 1.00
 WHERE `valor_utm` IS NULL
    OR `valor_utm` <= 0;

/* -------------------------------------------------------------------------- */
/*  Indices (no unique)                                                        */
/* -------------------------------------------------------------------------- */

CALL ensure_index('registros', 'idx_nombre', 'ADD KEY `idx_nombre` (`nombre`)');
CALL ensure_index('registros', 'idx_rut', 'ADD KEY `idx_rut` (`rut`)');
CALL ensure_index('registros', 'idx_rol', 'ADD KEY `idx_rol` (`rol`)');

CALL ensure_index('sesiones', 'idx_sesiones_user', 'ADD KEY `idx_sesiones_user` (`user_id`)');
CALL ensure_index('sesiones', 'idx_sesiones_expires', 'ADD KEY `idx_sesiones_expires` (`expires_at`)');

CALL ensure_index('registro_historial', 'idx_historial_registro_fecha', 'ADD KEY `idx_historial_registro_fecha` (`registro_id`, `created_at`)');
CALL ensure_index('registro_historial', 'idx_historial_num_ingreso', 'ADD KEY `idx_historial_num_ingreso` (`num_ingreso`)');

CALL ensure_index('factura_solicitudes', 'idx_factura_solicitudes_estado', 'ADD KEY `idx_factura_solicitudes_estado` (`estado`)');
CALL ensure_index('factura_solicitudes', 'idx_factura_solicitudes_created_at', 'ADD KEY `idx_factura_solicitudes_created_at` (`created_at`)');
CALL ensure_index('factura_solicitudes', 'idx_factura_solicitudes_num_ingreso', 'ADD KEY `idx_factura_solicitudes_num_ingreso` (`num_ingreso`)');
CALL ensure_index('factura_solicitudes', 'idx_factura_solicitudes_created_sucursal', 'ADD KEY `idx_factura_solicitudes_created_sucursal` (`created_by_sucursal`)');

CALL ensure_index('utm_mensual', 'idx_utm_mensual_registrado_por_id', 'ADD KEY `idx_utm_mensual_registrado_por_id` (`registrado_por_id`)');
CALL ensure_index(
    'cotizacion_servicios_estandar',
    'idx_cotizacion_servicios_activo_orden',
    'ADD KEY `idx_cotizacion_servicios_activo_orden` (`activo`, `orden_visual`)'
);
CALL ensure_index(
    'correo_configuracion',
    'idx_correo_config_updated_by',
    'ADD KEY `idx_correo_config_updated_by` (`updated_by_id`)'
);

/* -------------------------------------------------------------------------- */
/*  Indices unique (con validacion previa para evitar perdida/bloqueo)         */
/* -------------------------------------------------------------------------- */

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'registros'
       AND INDEX_NAME = 'uq_num_ingreso'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `num_ingreso`
              FROM `registros`
             WHERE `num_ingreso` IS NOT NULL
               AND TRIM(`num_ingreso`) <> ''
             GROUP BY `num_ingreso`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_num_ingreso ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `registros` ADD UNIQUE KEY `uq_num_ingreso` (`num_ingreso`)',
        'SELECT ''Aviso: no se creo uq_num_ingreso porque hay num_ingreso duplicados.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'registros'
       AND INDEX_NAME = 'uq_anio_correlativo'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `anio`, `correlativo`
              FROM `registros`
             WHERE `anio` IS NOT NULL
               AND `correlativo` IS NOT NULL
             GROUP BY `anio`, `correlativo`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_anio_correlativo ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `registros` ADD UNIQUE KEY `uq_anio_correlativo` (`anio`, `correlativo`)',
        'SELECT ''Aviso: no se creo uq_anio_correlativo porque hay combinaciones anio/correlativo duplicadas.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'usuarios'
       AND INDEX_NAME = 'uq_usuarios_username'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `username`
              FROM `usuarios`
             WHERE `username` IS NOT NULL
               AND TRIM(`username`) <> ''
             GROUP BY `username`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_usuarios_username ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `usuarios` ADD UNIQUE KEY `uq_usuarios_username` (`username`)',
        'SELECT ''Aviso: no se creo uq_usuarios_username porque hay usernames duplicados.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'sesiones'
       AND INDEX_NAME = 'uq_sesiones_token_hash'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `token_hash`
              FROM `sesiones`
             WHERE `token_hash` IS NOT NULL
               AND TRIM(`token_hash`) <> ''
             GROUP BY `token_hash`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_sesiones_token_hash ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `sesiones` ADD UNIQUE KEY `uq_sesiones_token_hash` (`token_hash`)',
        'SELECT ''Aviso: no se creo uq_sesiones_token_hash porque hay token_hash duplicados.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'sucursales'
       AND INDEX_NAME = 'uq_sucursales_nombre'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT TRIM(`nombre`) AS `nombre_norm`
              FROM `sucursales`
             WHERE `nombre` IS NOT NULL
               AND TRIM(`nombre`) <> ''
             GROUP BY TRIM(`nombre`)
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_sucursales_nombre ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `sucursales` ADD UNIQUE KEY `uq_sucursales_nombre` (`nombre`)',
        'SELECT ''Aviso: no se creo uq_sucursales_nombre porque hay sucursales duplicadas.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'factura_solicitudes'
       AND INDEX_NAME = 'uq_factura_solicitudes_num_ingreso'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `num_ingreso`
              FROM `factura_solicitudes`
             WHERE `num_ingreso` IS NOT NULL
               AND TRIM(`num_ingreso`) <> ''
             GROUP BY `num_ingreso`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_factura_solicitudes_num_ingreso ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `factura_solicitudes` ADD UNIQUE KEY `uq_factura_solicitudes_num_ingreso` (`num_ingreso`)',
        'SELECT ''Aviso: no se creo uq_factura_solicitudes_num_ingreso porque hay num_ingreso duplicados en factura_solicitudes.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @idx_exists := (
    SELECT COUNT(*)
      FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE()
       AND TABLE_NAME = 'utm_mensual'
       AND INDEX_NAME = 'uq_utm_mensual_periodo'
);
SET @dup_count := (
    SELECT COUNT(*)
      FROM (
            SELECT `anio`, `mes`
              FROM `utm_mensual`
             GROUP BY `anio`, `mes`
            HAVING COUNT(*) > 1
           ) d
);
SET @sql := IF(
    @idx_exists > 0,
    'SELECT ''uq_utm_mensual_periodo ya existe.'' AS info',
    IF(
        @dup_count = 0,
        'ALTER TABLE `utm_mensual` ADD UNIQUE KEY `uq_utm_mensual_periodo` (`anio`, `mes`)',
        'SELECT ''Aviso: no se creo uq_utm_mensual_periodo porque hay periodos duplicados.'' AS aviso'
    )
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

/* -------------------------------------------------------------------------- */
/*  Foreign keys (se intentan crear sin abortar toda la migracion)             */
/* -------------------------------------------------------------------------- */

CALL ensure_fk_safe(
    'sesiones',
    'fk_sesiones_user_id',
    'ADD CONSTRAINT `fk_sesiones_user_id` FOREIGN KEY (`user_id`) REFERENCES `usuarios`(`id`) ON DELETE CASCADE'
);

CALL ensure_fk_safe(
    'registro_historial',
    'fk_historial_registro_id',
    'ADD CONSTRAINT `fk_historial_registro_id` FOREIGN KEY (`registro_id`) REFERENCES `registros`(`id`) ON DELETE CASCADE'
);

/* -------------------------------------------------------------------------- */
/*  Sincronizacion de sucursales desde usuarios                                */
/* -------------------------------------------------------------------------- */

INSERT IGNORE INTO `sucursales` (`nombre`)
SELECT DISTINCT TRIM(`u`.`sucursal`) AS `nombre`
  FROM `usuarios` u
 WHERE `u`.`sucursal` IS NOT NULL
   AND TRIM(`u`.`sucursal`) <> '';

/* -------------------------------------------------------------------------- */
/*  Seed inicial cotizacion (solo si tabla esta vacia)                        */
/* -------------------------------------------------------------------------- */

INSERT INTO `cotizacion_servicios_estandar` (`nombre_servicio`, `valor_utm`, `orden_visual`, `activo`)
SELECT 'Servicio estandar Geo Rural', 1.00, 1, 1
 WHERE NOT EXISTS (
    SELECT 1
      FROM `cotizacion_servicios_estandar`
 );

/* -------------------------------------------------------------------------- */
/*  Limpieza de helpers                                                        */
/* -------------------------------------------------------------------------- */

DROP PROCEDURE IF EXISTS `ensure_fk_safe`;
DROP PROCEDURE IF EXISTS `ensure_index`;
DROP PROCEDURE IF EXISTS `ensure_column`;
