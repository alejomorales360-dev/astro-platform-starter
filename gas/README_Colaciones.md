# Colaciones Semanales — Google Apps Script backend

App para controlar la inscripción semanal de los ~67 trabajadores al menú de
colaciones: evita que alguien quede sin anotarse y evita anotaciones dobles.
El frontend estático está en `public/colaciones/index.html` y se publica
junto con este sitio en Netlify (queda disponible en `/colaciones/`).

## Cómo funciona

- **Trabajadores** entran con su **RUT** (sin clave), ven el menú de la
  semana vigente y eligen **una opción por día**. Si eligen otra opción para
  el mismo día, se reemplaza la anterior (no se duplica). También pueden ver
  qué pidieron.
- **Administración** entra con una clave y puede:
  - Cargar, editar, desactivar o eliminar trabajadores.
  - Marcar trabajadores como **Spot** (visitas, clientes u otras personas
    que solo reciben colación durante un rango de fechas puntual) indicando
    fecha de inicio y fin de vigencia.
  - Definir el menú de cada día de la semana (varias opciones, ej. A/B/C).
  - Copiar el menú de la semana anterior como punto de partida.
  - Ver reportes: quién se anotó y quién no, resumen de opciones elegidas
    por día, exportar a CSV/Excel y exportar/imprimir en PDF (botón que usa
    la función de impresión del navegador con una vista lista para PDF).

## 1. Crear la planilla y las hojas

1. Crea una planilla de Google Sheets nueva (ej. "Colaciones 2026").
2. Ve a **Extensiones → Apps Script**.
3. Borra el contenido por defecto y pega el contenido de `gas/Colaciones.gs`
   de este repo.
4. Guarda (Ctrl+S / ⌘+S).
5. En el desplegable de funciones (arriba, junto a "Depurar"), elige
   **crearHojasIniciales** y presiona **Ejecutar**. La primera vez pedirá
   autorizar permisos: acéptalos.
   - Esto crea las hojas `Trabajadores`, `Menus`, `Pedidos` y `Config` con
     sus encabezados, y agrega una clave de admin por defecto
     (`admin_password = cambiar123`).
6. Abre la hoja **Config** y cambia el valor de `admin_password` por una
   clave real.
7. Carga a tus trabajadores en la hoja **Trabajadores**:
   `RUT | Nombre | Tipo | Activo | FechaInicio | FechaFin | Notas`
   - `Tipo`: `Fijo` o `Spot`.
   - `Activo`: `Si` o `No`.
   - `FechaInicio`/`FechaFin` solo aplican a `Spot` (formato `AAAA-MM-DD`).
     Una visita que viene solo un par de días se carga con esas dos fechas;
     fuera de ese rango no cuenta en los reportes de esa semana.

## 2. Publicar como Aplicación web

1. **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Ejecutar como: **Yo**.
4. Quién tiene acceso: **Cualquier persona** (para que los trabajadores
   puedan entrar sin cuenta de Google).
5. Presiona **Implementar** y copia la URL que termina en `/exec`.

## 3. Conectar el frontend

1. Abre `public/colaciones/index.html`.
2. Busca la línea:
   ```js
   const GAS_URL = 'PEGA_AQUI_TU_URL_DE_APPS_SCRIPT';
   ```
3. Reemplázala por la URL `/exec` copiada en el paso anterior.
4. Guarda y despliega el sitio (o pruébalo abriendo el archivo
   directamente en el navegador).

## Volver a desplegar tras editar `Colaciones.gs`

Igual que con el resto de los backends de este repo: editar el código no
actualiza la URL `/exec` ya publicada. Hay que ir a
**Implementar → Gestionar implementaciones**, editar (ícono de lápiz) la
implementación de tipo "Aplicación web", elegir **Nueva versión** en
"Versión" y presionar **Implementar**. La URL `/exec` no cambia, así que no
hay que tocar `GAS_URL` de nuevo.

## Diagnóstico rápido

Desde el editor de Apps Script, ejecuta manualmente:

- `listarHojasCol` — muestra en el Registro (`Ver → Registro de ejecución`)
  los nombres reales de las hojas de tu planilla.
- `testAPICol` — muestra en el Registro cuántos trabajadores, menús y
  pedidos hay cargados actualmente.
