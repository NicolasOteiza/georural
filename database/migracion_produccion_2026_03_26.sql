/*
  Migracion incremental de estructura para Geo Rural
  Fecha: 2026-03-26
  Objetivo:
    - asegurar columnas de edicion en registro_historial
    - normalizar roles incluyendo SUPERVISOR
    - crear tabla de alertas de documentos para OPERADOR/SUPER
  Nota: este script NO elimina datos de negocio.
*/

SET NAMES utf8mb4;

CREATE DATABASE IF NOT EXISTS `geo_rural` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `geo_rural`;

DELIMITER $$

DROP PROCEDURE IF EXISTS `ensure_column_safe`$$
CREATE PROCEDURE `ensure_column_safe`(
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

DROP PROCEDURE IF EXISTS `ensure_index_safe`$$
CREATE PROCEDURE `ensure_index_safe`(
    IN p_table VARCHAR(64),
    IN p_index VARCHAR(64),
    IN p_definition TEXT
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
        SET @sql = CONCAT('ALTER TABLE `', p_table, '` ADD ', p_definition);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$

DROP PROCEDURE IF EXISTS `ensure_fk_safe`$$
CREATE PROCEDURE `ensure_fk_safe`(
    IN p_table VARCHAR(64),
    IN p_fk VARCHAR(64),
    IN p_definition TEXT
)
BEGIN
    DECLARE v_exists INT DEFAULT 0;

    SELECT COUNT(*)
      INTO v_exists
      FROM information_schema.REFERENTIAL_CONSTRAINTS
     WHERE CONSTRAINT_SCHEMA = DATABASE()
       AND TABLE_NAME = p_table
       AND CONSTRAINT_NAME = p_fk;

    IF v_exists = 0 THEN
        SET @sql = CONCAT('ALTER TABLE `', p_table, '` ', p_definition);
        PREPARE stmt FROM @sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$

DELIMITER ;

/* -------------------------------------------------------------------------- */
/*  Historial: columnas de edicion                                             */
/* -------------------------------------------------------------------------- */

CALL ensure_column_safe('registro_historial', 'editado', '`editado` TINYINT(1) NOT NULL DEFAULT 0 AFTER `sucursal`');
CALL ensure_column_safe('registro_historial', 'editado_por', '`editado_por` VARCHAR(120) NULL AFTER `editado`');
CALL ensure_column_safe('registro_historial', 'editado_por_sucursal', '`editado_por_sucursal` VARCHAR(120) NULL AFTER `editado_por`');
CALL ensure_column_safe('registro_historial', 'editado_at', '`editado_at` DATETIME NULL AFTER `editado_por_sucursal`');

/* -------------------------------------------------------------------------- */
/*  Usuarios: normalizacion de roles                                           */
/* -------------------------------------------------------------------------- */

CALL ensure_column_safe('usuarios', 'role', "`role` VARCHAR(20) NOT NULL DEFAULT 'OPERADOR' AFTER `sucursal`");

UPDATE `usuarios`
   SET `role` = UPPER(TRIM(COALESCE(`role`, '')))
 WHERE `role` IS NOT NULL;

UPDATE `usuarios`
   SET `role` = 'OPERADOR'
 WHERE `role` IS NULL
    OR TRIM(`role`) = ''
    OR `role` NOT IN ('SUPER', 'ADMIN', 'SECRETARIA', 'OPERADOR', 'SUPERVISOR');

/* -------------------------------------------------------------------------- */
/*  Nueva tabla: alertas de documentos agregados                              */
/* -------------------------------------------------------------------------- */

CREATE TABLE IF NOT EXISTS `registro_documentos_alertas` (
    `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    `registro_id` INT UNSIGNED NOT NULL,
    `num_ingreso` VARCHAR(20) NOT NULL,
    `sucursal` VARCHAR(120) NULL,
    `documentos_agregados` JSON NOT NULL,
    `created_by` VARCHAR(120) NULL,
    `revisado` TINYINT(1) NOT NULL DEFAULT 0,
    `revisado_por` VARCHAR(120) NULL,
    `revisado_por_sucursal` VARCHAR(120) NULL,
    `revisado_at` DATETIME NULL,
    `created_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CALL ensure_index_safe(
    'registro_documentos_alertas',
    'idx_doc_alertas_revisado_fecha',
    'KEY `idx_doc_alertas_revisado_fecha` (`revisado`, `created_at`)'
);
CALL ensure_index_safe(
    'registro_documentos_alertas',
    'idx_doc_alertas_sucursal',
    'KEY `idx_doc_alertas_sucursal` (`sucursal`)'
);
CALL ensure_index_safe(
    'registro_documentos_alertas',
    'idx_doc_alertas_ingreso',
    'KEY `idx_doc_alertas_ingreso` (`num_ingreso`)'
);
CALL ensure_fk_safe(
    'registro_documentos_alertas',
    'fk_doc_alertas_registro_id',
    'ADD CONSTRAINT `fk_doc_alertas_registro_id` FOREIGN KEY (`registro_id`) REFERENCES `registros`(`id`) ON DELETE CASCADE'
);

/* -------------------------------------------------------------------------- */
/*  Limpieza                                                                   */
/* -------------------------------------------------------------------------- */

DROP PROCEDURE IF EXISTS `ensure_fk_safe`;
DROP PROCEDURE IF EXISTS `ensure_index_safe`;
DROP PROCEDURE IF EXISTS `ensure_column_safe`;
