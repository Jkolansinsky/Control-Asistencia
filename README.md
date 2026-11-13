# Control de Asistencia v11 — pasos

## 1. Google Sheets / Apps Script
1. En tu proyecto de Apps Script reemplaza TODO el código por `Code.gs` (borra code2/code3: ya no se usan). Conserva tu `appsscript.json` (zona horaria, ejecutar como "Yo", acceso "Cualquier persona").
2. Ejecuta una vez `configuracionInicial()` y autoriza permisos. Crea columnas nuevas en Registros (I a R), `descriptor`/`actualizado` en BD (I:J) y `Hash`/`Salt` en admin (E:F).
3. Hoja **Sucursales**: A nombre, B "lat, lng". Opcional col. E = radio en metros (por defecto 200).
4. Hoja **Licencias**: actualiza las fechas de vencimiento (ahora SÍ se validan) y **vacía la columna G** para que cada dispositivo se enlace solo al activarse.
5. Implementar → Nueva implementación → Aplicación web. Copia la URL `/exec`.

## 2. GitHub Pages
1. Crea un repo, sube todos los archivos de esta carpeta (sin `Code.gs` si quieres).
2. En `core.js` pega la URL en `CFG.API`.
3. Settings → Pages → Branch `main` / root. Abre `https://TU_USUARIO.github.io/TU_REPO/`.
4. Opcional (más rápido y 100% estable): copia la carpeta `model` de `@vladmandic/face-api` a `models/` en tu repo y cambia `CFG.MODELS` a `'models/'`.

## 3. Primer uso
1. En cada celular/tablet abre la URL, escribe la licencia y activa. Usa "Agregar a pantalla de inicio".
2. Abre `admin.html` (usuario y contraseña de tu hoja admin; la contraseña en texto plano se convierte a hash en el primer ingreso).
3. Panel → **Migrar fotos antiguas** (una vez). Luego da de alta nuevos empleados con rostro.

## 4. Reportes PDF para gerentes (reporte.html)
1. En la hoja **admin** las columnas G y H son `Rol` y `Sucursal` (las crea `configuracionInicial()`). Los usuarios actuales sin rol son administradores (ven todas las sucursales).
2. Para cada gerente de franquicia: edita `U`, `CLAVE` y `SUC` en `crearGerente()` (Code.gs), ejecútala una vez y borra la clave del código.
3. Dales la URL `https://TU_USUARIO.github.io/TU_REPO/reporte.html`. Entran con su usuario y contraseña desde su celular.
4. Ejecuta cualquier función del editor para autorizar los permisos nuevos (Drive y conexión externa) y publica una **nueva versión** de la implementación.

## 5. Archivo mensual de registros (Archivo.gs)
1. Agrega `Archivo.gs` y reemplaza `Reporte.gs` en Apps Script.
2. Ejecuta `instalarArchivoMensual()` una vez (programa el día 3 de cada mes, 3:00 a.m.).
3. Para probarlo o hacer la primera limpieza: ejecuta `archivarRegistros()` y revisa el registro de ejecución.
4. Los archivos "Registros AAAA" se crean solos en la misma carpeta de Drive que tu hoja principal.
