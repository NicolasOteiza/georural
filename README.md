# Geo Rural - Registro con Node + MySQL

## Requisitos
- Node.js 18+
- MySQL/MariaDB (XAMPP)

## Configuracion
1. Edita `.env` si necesitas otro usuario/clave de MySQL.
2. Variables recomendadas:
   - `HOST=127.0.0.1` (evita exponer el servidor a la red).
   - `APP_URL=` (URL publica del sistema en hosting, por ejemplo `https://tu-dominio.com`).
   - `CORS_ORIGINS=http://localhost:3000,http://127.0.0.1:3000,http://localhost,http://127.0.0.1`
   - `API_WRITE_KEY=` (protege `POST/PUT` con header `X-API-Key`; es obligatorio si `HOST` no es localhost).
   - `AUTH_SESSION_HOURS=12` (duracion de la sesion).
   - `LOGIN_MAX_ATTEMPTS=5`, `LOGIN_WINDOW_MINUTES=15`, `LOGIN_LOCK_MINUTES=15` (limite de intentos fallidos de login).
   - `MAX_PASSWORD_LENGTH=200` (largo maximo permitido para claves).
   - `SESSION_CLEANUP_INTERVAL_MINUTES=30` (cada cuanto limpiar sesiones expiradas/revocadas).
   - `SMTP_HOST`, `SMTP_PORT`, `SMTP_SECURE`, `SMTP_USER`, `SMTP_PASSWORD`, `SMTP_FROM_EMAIL`, `SMTP_FROM_NAME` (respaldo SMTP por variables de entorno).
   - `SUPERUSER_USERNAME`, `SUPERUSER_PASSWORD`, `SUPERUSER_NAME`, `SUPERUSER_BRANCH` (cuenta SUPER principal).
   - `DEFAULT_USERS=...` (usuarios iniciales, formato `usuario|contrasena|sucursal|nombre|rol` separados por coma).
3. Instala dependencias:
   - `npm install`
4. Inicia servidor:
   - `npm start`
5. Verificacion rapida de sintaxis:
   - `npm run check`

El servidor corre por defecto en `http://127.0.0.1:3000`.

## Usuario MySQL recomendado (evitar root)
Si aparece `Access denied for user 'root'@'localhost' (using password: NO)`, crea un usuario dedicado para la app:

1. En phpMyAdmin (pestana SQL) ejecuta:
```sql
CREATE DATABASE IF NOT EXISTS `geo_rural` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER IF NOT EXISTS 'geo_rural_app'@'127.0.0.1' IDENTIFIED BY 'Cambia_Esta_Clave_2026!';
CREATE USER IF NOT EXISTS 'geo_rural_app'@'localhost' IDENTIFIED BY 'Cambia_Esta_Clave_2026!';
GRANT ALL PRIVILEGES ON `geo_rural`.* TO 'geo_rural_app'@'127.0.0.1';
GRANT ALL PRIVILEGES ON `geo_rural`.* TO 'geo_rural_app'@'localhost';
FLUSH PRIVILEGES;
```
2. Ajusta `.env`:
   - `DB_USER=geo_rural_app`
   - `DB_PASSWORD=Cambia_Esta_Clave_2026!`
3. Reinicia `iniciar_servidor_node.bat`.

## Inicio rapido en Windows (paso a paso)
1. Inicia Apache y MySQL desde XAMPP.
2. Abre la carpeta del proyecto: `C:\xampp\htdocs\geo_rural`.
3. Haz doble clic en `iniciar_servidor_node.bat`.
4. Espera el mensaje `Servidor activo en http://127.0.0.1:3000`.
5. Deja esa ventana abierta mientras usas el sistema.

Si prefieres terminal:
- En PowerShell usa `npm.cmd start` (no `npm start`, porque puede fallar por la politica de scripts).
- En CMD puedes usar `npm start`.

Si configuras `API_WRITE_KEY`, el frontend la recibe automaticamente al iniciar sesion.
Solo como respaldo/manual (por ejemplo para pruebas con Postman o consola), puedes definirla asi:
- `sessionStorage.setItem('geo_rural_api_key', 'TU_CLAVE')`

Si en el navegador quedo guardado un API base incorrecto, limpialo con:
- `localStorage.removeItem('geo_rural_api_base')`

## Inicio de sesion
- El frontend exige login para operar.
- El login permite entrar como `INVITADO` (modo solo lectura, sin edicion de registros).
- Si la tabla `usuarios` esta vacia al iniciar, el servidor crea usuarios desde `DEFAULT_USERS`.
- Si `DEFAULT_USERS` esta vacio, el servidor genera un `ADMIN` temporal con clave aleatoria y la muestra en logs de arranque.
- Cambia las claves iniciales o temporales inmediatamente en produccion.
- El `ADMIN` puede crear y eliminar usuarios desde la interfaz.
- Nuevo rol `SUPER`: acceso completo (ADMIN + SECRETARIA + OPERADOR), gestion de correo SMTP desde menu y solo una cuenta permitida.
- La cuenta `SUPER` no aparece en la lista de gestion cuando entra un `ADMIN`.
- Nuevo rol `SECRETARIA`: ve el panel de administracion con modos de ano (actual/historico), pero sin gestionar usuarios/sucursales/documentos.
- `SECRETARIA` recibe alerta diaria (primer ingreso del dia) con registros de 8+ dias desde su primer comentario y con comentario unico.
- Nuevo rol `SUPERVISOR`: similar a `INVITADO`, pero puede ver historial y guardar comentarios en registros existentes.
- `ADMIN` y `SUPER` pueden editar entradas de historial tipo `COMENTARIO` (fecha y texto), manteniendo autor original y marca de edicion.
- Solo `SUPER` puede crear o promocionar usuarios al rol `ADMIN`.
- `ADMIN` y `SECRETARIA` pueden cambiar `NRO INGRESO` manualmente (por ejemplo, para anos pasados).
- Si vas a publicar en subruta o dominio distinto, puedes definir la API con `<meta name="geo-rural-api-base" content="https://tu-dominio.com/api">`.

## API disponible
- `GET /api/health`
- `POST /api/auth/login`
- `POST /api/auth/guest`
- `GET /api/auth/me`
- `POST /api/auth/logout`
- `GET /api/admin/usuarios` (ADMIN)
- `POST /api/admin/usuarios` (ADMIN)
- `PUT /api/admin/usuarios/:userId` (ADMIN)
- `POST /api/admin/usuarios/:userId/reset-password` (ADMIN)
- `POST /api/admin/usuarios/:userId/guest-status` (SUPER, solo cuenta invitado)
- `DELETE /api/admin/usuarios/:userId` (ADMIN)
- `GET /api/admin/correo-config` (SUPER)
- `PUT /api/admin/correo-config` (SUPER)
- `GET /api/next-ingreso`
- `GET /api/utm/mes-actual` (SECRETARIA o SUPER)
- `POST /api/utm/mes-actual` (SECRETARIA o SUPER)
- `GET /api/cotizaciones/resumen` (SECRETARIA o SUPER)
- `POST /api/cotizaciones/enviar` (SECRETARIA o SUPER)
- `GET /api/facturas/solicitudes` (OPERADOR, SECRETARIA o SUPER)
- `GET /api/facturas/solicitudes/por-ingreso/:numIngreso` (OPERADOR, SECRETARIA o SUPER)
- `POST /api/facturas/solicitudes` (OPERADOR, SECRETARIA o SUPER)
- `POST /api/facturas/solicitudes/marcar-enviadas` (OPERADOR, SECRETARIA o SUPER)
- `POST /api/facturas/correo/enviar` (OPERADOR, SECRETARIA o SUPER)
- `POST /api/registros`
- `PUT /api/registros/:numIngreso`
- `POST /api/registros/:numIngreso/comentario`
- `PUT /api/registros/:numIngreso/historial/:historyId` (ADMIN o SUPER)
- `DELETE /api/registros/:numIngreso` (ADMIN o SUPER)
- `GET /api/registros/buscar?nombre=...&rut=...&rol=...&numIngreso=...`
- `GET /api/registros/sucursales-disponibles`
- `GET /api/registros/seguimiento-sin-movimiento?dias=8` (SECRETARIA o SUPER)
- `GET /api/registros/documentos-alertas` (OPERADOR o SUPER)
- `POST /api/registros/documentos-alertas/:alertId/revisar` (OPERADOR o SUPER)
- `GET /api/registros/buscar-rango-fechas?fechaDesde=YYYY-MM-DD&fechaHasta=YYYY-MM-DD&sucursal=...&region=...&comuna=...`
- `GET /api/registros/:numIngreso/historial`

## Notas
- Al iniciar, `server.js` crea automaticamente la base `geo_rural` y la tabla `registros` si no existen.
- Si prefieres crear el esquema manualmente, usa `database/schema.sql`.
- `NRO INGRESO` se reserva en transaccion al guardar, para evitar colisiones concurrentes.
- `NRO DE LOTES` ahora acepta solo enteros positivos.
- Cada alta/modificacion registra usuario + sucursal en `registro_historial`.
- La documentacion recibida de cada registro se persiste en `registros.documentos` (JSON).
- Al buscar un registro, el frontend muestra historial de comentarios con fecha y usuario.
- El envio de facturas al contador usa SMTP del sistema (igual que cotizaciones), ya no depende de `mailto`.
- Si una busqueda devuelve multiples coincidencias, el frontend solicita seleccionar `NRO INGRESO` exacto.
- La sesion se valida por cookie `HttpOnly` y ya no depende de token persistente en `localStorage`.
- Si el frontend corre en otro puerto local (ej: XAMPP en `localhost`), usa automaticamente `http://127.0.0.1:3000/api`.
