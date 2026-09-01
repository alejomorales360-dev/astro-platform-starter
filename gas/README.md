# Tesorería 2° Básico A — Google Apps Script backend

Este backend vive en Google Apps Script, ligado a la planilla de Google
Sheets que contiene los datos (Alumnos, Conceptos, Pagos, Gastos,
Configuracion, Usuarios). El frontend estático está en
`public/tesoreria/index.html` y se publica junto con este sitio en Netlify.

## Bug encontrado: "no conecta"

El código original buscaba las hojas con nombres en mayúsculas
(`ALUMNOS`, `CONCEPTOS`, `PAGOS`, `GASTOS`, `CONFIGURACION`, `USUARIOS`),
pero la planilla real tiene los nombres capitalizados normalmente
(`Alumnos`, `Conceptos`, `Pagos`, `Gastos`, `Configuracion`, `Usuarios`).

`SpreadsheetApp.getSheetByName()` distingue mayúsculas de minúsculas, así
que **todas** las llamadas al backend (login, cargar datos, guardar pago,
etc.) fallaban con `Hoja no encontrada: ALUMNOS`. Eso es lo que se veía
como "no conecta".

`Code.gs` en esta carpeta ya trae el fix: `getHoja()` ahora también busca
la hoja ignorando mayúsculas/minúsculas, así que funciona sin importar
cómo esté nombrada la hoja en tu planilla.

## Cómo volver a desplegar el Apps Script

1. Abre tu Google Sheet → **Extensiones → Apps Script**.
2. Reemplaza el contenido de tu archivo `.gs` por el contenido de
   `gas/Code.gs` de este repo.
3. Guarda (Ctrl+S / ⌘+S).
4. **Implementar → Gestionar implementaciones** → edita (ícono de lápiz)
   la implementación de tipo "Aplicación web" que ya tienes.
5. En "Versión" elige **Nueva versión** y presiona **Implementar**.
   - Importante: solo editar el código no actualiza la URL `/exec` que ya
     usa el HTML — hay que crear una **nueva versión** de la implementación
     existente para que los cambios entren en vigor.
6. Confirma que sigue siendo "Ejecutar como: Yo" y "Quién tiene acceso:
   Cualquier persona" (o "Cualquier persona con una Cuenta de Google" si
   prefieres restringirlo, aunque eso puede bloquear a los apoderados sin
   cuenta de Google).
7. La URL `/exec` no cambia al crear una nueva versión, así que
   `GAS_URL` en `public/tesoreria/index.html` sigue siendo válida.

## Diagnóstico rápido

Si vuelve a fallar la conexión, desde el editor de Apps Script ejecuta la
función `listarHojas` (menú **Ejecutar**) y revisa el **Registro**
(`Ver → Registro de ejecución`, o Ctrl+Enter) para ver los nombres reales
de las hojas de tu planilla y compararlos con `HOJAS` en `Code.gs`.
