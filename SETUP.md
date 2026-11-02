# Instalación paso a paso

Son unos 10 minutos. Al final tendrás un link que tú y Fran abren en el celular, y cada gasto queda como una fila en una planilla de Google.

```
Celular (app web)  ──►  Google Apps Script  ──►  Google Sheets
GitHub Pages            (intermediario)          (Gastos + Resúmenes)
```

## 1. Crear la planilla

1. Entra a [sheets.new](https://sheets.new) con tu cuenta de Google.
2. Ponle de nombre **Gastos Familia**.

## 2. Pegar el script

1. En la planilla: **Extensiones → Apps Script**.
2. Borra todo lo que aparece en `Código.gs` y pega el contenido completo de [`apps-script/Code.gs`](apps-script/Code.gs).
3. Guarda (ícono del disquete).

## 3. Definir el PIN

1. En Apps Script, en el menú de la izquierda: **⚙️ Configuración del proyecto**.
2. Baja hasta **Propiedades del script → Agregar propiedad del script**.
3. Propiedad: `PIN` · Valor: el número que van a usar los dos (por ejemplo, 4 dígitos).
4. **Guardar propiedades del script**.

El PIN vive solo ahí, nunca en GitHub.

## 4. Crear las hojas

1. Vuelve al editor (`< >` en el menú izquierdo).
2. Arriba, en el selector de funciones, elige **setup** y presiona **Ejecutar**.
3. Google pedirá permisos: **Revisar permisos → tu cuenta → Configuración avanzada → Ir a … (no seguro) → Permitir**. Es normal: el script es tuyo y solo toca esta planilla.
4. Revisa la planilla: deberían aparecer las hojas **Gastos, Centros, Personas, Resumen mensual y Resumen anual**.

> En la hoja **Personas** vienen "Vicente" y "Fran". Cámbialos si quieres que aparezcan otros nombres.

## 5. Publicar el script como aplicación web

1. En Apps Script: **Implementar → Nueva implementación**.
2. En el engranaje de "Seleccionar tipo", elige **Aplicación web**.
3. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier usuario**.
4. **Implementar** y copia la **URL de la aplicación web** (termina en `/exec`).

"Cualquier usuario" es necesario para que la app pueda conectarse sin pedir login de Google; el PIN es lo que protege la escritura.

## 6. Pegar la URL en la app

1. En GitHub, abre el archivo [`config.js`](config.js) y presiona el lápiz (editar).
2. Reemplaza `PEGAR_AQUI_LA_URL_DE_APPS_SCRIPT` por la URL que copiaste (dejando las comillas).
3. **Commit changes**.

## 7. Activar GitHub Pages

1. Asegúrate de que estos archivos estén en la rama `main` (mergeando el pull request de esta rama).
2. En el repo: **Settings → General → Danger Zone → Change visibility → Public**. (Solo el código queda público; los gastos están en tu Google Drive.)
3. **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, Branch **main** y carpeta **/ (root)** → **Save**.
4. En 1–2 minutos la app queda en: **https://vossandon93.github.io/Finanzas/**

## 8. Instalar en los celulares

1. Abre el link, ingresa el PIN y elige quién eres.
2. Agrégala a la pantalla de inicio para que se vea como una app:
   - **iPhone (Safari):** botón Compartir → **Agregar a inicio**.
   - **Android (Chrome):** menú ⋮ → **Agregar a la pantalla principal**.
3. Mándale el link y el PIN a Fran y que haga lo mismo.

---

## Uso diario

- **Agregar un centro de costo:** en la app, en la lista de centros elige **+ Agregar centro de costo…** (o escríbelo en una fila nueva de la hoja **Centros**). No borres ni renombres centros que ya tienen gastos: el resumen los agrupa por nombre.
- **Cuentas dentro de un centro** (ej. Vicky: Sueldo, Contribuciones, Traslado): al elegir un centro aparecen sus cuentas para marcar una (es opcional); **+ Nueva cuenta** agrega otra a ese centro. Quedan en la hoja **Cuentas** (columna A el centro, B la cuenta) y cada gasto guarda la suya en la columna **Cuenta** de la hoja Gastos. La hoja **Resumen por cuenta** suma por mes, centro y cuenta.
- **Corregir un gasto:** bórralo desde la lista "Últimos gastos" y vuelve a ingresarlo, o edítalo directo en la hoja **Gastos**.
- **Métricas:** en la app, las pestañas **Por centro** (gráfico circular por centro de costo, por mes o año) y **Meses** (últimos 12 meses). En la planilla, las hojas **Resumen mensual** y **Resumen anual** las escribe el script cada vez que se guarda o borra un gasto desde la app. Si editas la hoja **Gastos** a mano, en script.google.com elige la función **actualizarResumenes** y presiona **Ejecutar**.
- **Descargar como Excel:** **Archivo → Descargar → Microsoft Excel (.xlsx)**.

## Pestañas personales (Rendición y Salud)

Además de las pestañas comunes, cada persona ve las suyas según quién eligió al entrar:

- **Vicente → Rendición:** sube la foto del comprobante (o un PDF), el monto y la descripción. La foto queda en la carpeta **Rendiciones - comprobantes** de tu Google Drive, la fila en la hoja **Rendiciones** y se envía un correo automático con la foto adjunta. Cada rendición se puede marcar como *Pagada* cuando te la reembolsen.
- **Fran → Salud:** contador de días sin azúcar (con récord y calendario de las últimas 5 semanas) y contador de proteína diaria con meta, botones de un toque y gráfico de los últimos 7 días. Los datos quedan en las hojas **Azúcar** y **Proteína**.

Para activarlas (una sola vez):

1. Pega el `Code.gs` nuevo en Apps Script y guarda.
2. En **Propiedades del script** agrega `CORREO_RENDICION` con el correo al que deben llegar las rendiciones (puedes poner varios separados por coma).
3. Elige la función **setup** y presiona **Ejecutar**. Google pedirá permisos nuevos (Drive para guardar las fotos y Gmail para enviar el correo): acéptalos igual que la primera vez.
4. **Implementar → Gestionar implementaciones → ✏️ → Nueva versión → Implementar.**

> Qué pestañas ve cada uno se define en `index.html`, en `PESTANAS_DE`. Ojo: es una separación de comodidad, no de seguridad; los dos usan el mismo PIN.

## Si cambias el script más adelante

Pega el nuevo código y ve a **Implementar → Gestionar implementaciones → ✏️ editar → Versión: Nueva versión → Implementar**. Así la URL se mantiene y no hay que tocar `config.js`.

## Problemas comunes

| Mensaje en la app | Solución |
|---|---|
| Falta configurar la URL de Apps Script en config.js | Paso 6. |
| PIN incorrecto | Revisa la propiedad `PIN` (paso 3). |
| El PIN no está configurado en el script | Falta el paso 3. |
| Falta la hoja Gastos: ejecuta setup() primero | Falta el paso 4. |
| No se pudo conectar | Revisa que la implementación tenga acceso "Cualquier usuario" (paso 5) y que la URL termine en `/exec`. |
