# Instalación paso a paso

Son unos 10 minutos. Al final tendrás un link que tú y Fran abren en el celular, y cada gasto queda como una fila en una planilla de Google.

```
Celular (app web)  ──►  Google Apps Script  ──►  Google Sheets
GitHub Pages            (intermediario)          (Gastos + Resúmenes)
```

## 1. Crear la planilla

1. Entra a [sheets.new](https://sheets.new) con tu cuenta de Google.
2. Ponle de nombre **Casa App**.

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

"Cualquier usuario" es necesario para que la app pueda conectarse; el acceso lo controla el propio script (PIN con límite de intentos, o cuentas de Google autorizadas: ver "Acceso privado").

## 6. Pegar la URL en la app

1. En GitHub, abre el archivo [`config.js`](config.js) y presiona el lápiz (editar).
2. Reemplaza `PEGAR_AQUI_LA_URL_DE_APPS_SCRIPT` por la URL que copiaste (dejando las comillas).
3. **Commit changes**.

## 7. Activar GitHub Pages

1. Asegúrate de que estos archivos estén en la rama `main` (mergeando el pull request de esta rama).
2. En el repo: **Settings → General → Danger Zone → Change visibility → Public**. (Solo el código queda público; los gastos están en tu Google Drive.)
3. **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, Branch **main** y carpeta **/ (root)** → **Save**.
4. En 1–2 minutos la app queda en: **https://vossandon93.github.io/Casa-app/**

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

## Ordenar centros y cuentas

La lista oficial está en `apps-script/Code.gs` (`CENTROS_INICIALES` y `CUENTAS_INICIALES`). Para dejar la planilla igual a esa lista y mover los gastos antiguos (por ejemplo "Luna" → Familia / Luna), basta con publicar la nueva versión: la primera vez que se abre la app lo hace sola (una sola vez). También se puede forzar eligiendo la función **reordenarCentrosYCuentas** y presionando **Ejecutar**. Cada gasto se cambia en su misma fila (no se duplica), antes se guarda un respaldo oculto de la hoja Gastos, y el registro de ejecución muestra qué se movió.

Cuando se agregan cuentas a la lista oficial (ej. Casa: Agua, Conectividad, Fumigación, Gas, Tintorería), la app las suma sola y también mueve los gastos antiguos a su cuenta: centros viejos con ese nombre o un sinónimo (Internet/Entel → Conectividad) y gastos de Casa sin cuenta cuya descripción lo deja claro (Metrogas → Gas). Si calza con más de una cuenta no se toca. Se cambia la misma fila (sin duplicar), con respaldo oculto, y se puede repetir a mano con **moverGastosAntiguos**.

Los centros y las cuentas siempre se muestran en orden alfabético, también los que se agreguen después desde la app (las hojas Centros y Cuentas se reordenan solas).

## Hoja Gastos

Columnas: ID · Fecha · **Centro de costo · Cuenta** · Monto · Descripción · Registró · Creado. El script busca cada columna por su encabezado, así que no cambies los títulos (sí puedes mover columnas). Las planillas antiguas, con la Cuenta al final, se ordenan solas la primera vez que se abre la app (o a mano con **ordenarHojaGastos**): se mueve la columna completa, sin copiar ni duplicar gastos.

**Editar a mano:** las columnas Centro de costo y Cuenta tienen listas desplegables (la de Cuenta muestra solo las cuentas del centro de esa fila; si cambias el centro y la cuenta no corresponde, se borra y te avisa). Al editar la hoja Gastos los resúmenes se recalculan solos. Para encontrar filas con errores (centro, cuenta, fecha o monto inválidos) ejecuta **revisarGastos** y mira el registro.

## Acceso privado: solo con cuentas de Google autorizadas

Sin esto la app se protege solo con el PIN (con límite: 5 intentos fallidos bloquean el ingreso 15 minutos). Con esto, **solo entran los correos que tú elijas**, con su cuenta de Google, y el PIN deja de servir. Los correos quedan en las propiedades del script, no en GitHub.

**A. Crear el ID de cliente de Google (una vez, ~5 min)**
1. Entra a https://console.cloud.google.com/ con tu cuenta y crea un proyecto (ej. "Casa App").
2. Menú → **APIs y servicios → Pantalla de consentimiento de OAuth** (o **Google Auth Platform**) → **Comenzar**: nombre "Casa App", tu correo de soporte, público **Externo**, acepta y crea.
3. En **Público / Audiencia → Usuarios de prueba** agrega los 2 correos que tendrán acceso (deja la app en modo *Prueba*: es otra capa de protección).
4. **Clientes → Crear cliente** → tipo **Aplicación web**:
   - Orígenes autorizados de JavaScript: `https://vossandon93.github.io`
   - URI de redireccionamiento autorizados: `https://vossandon93.github.io/Casa-app/` (con la barra final)
5. Copia el **ID de cliente** (termina en `.apps.googleusercontent.com`). No es secreto.

**B. Activarlo en Apps Script**
1. **Configuración del proyecto (⚙️) → Propiedades del script** → agrega:
   - `GOOGLE_CLIENT_ID` = el ID de cliente.
   - `CORREOS_PERMITIDOS` = los correos separados por coma, ej. `persona1@gmail.com, persona2@gmail.com`.
2. En el editor elige **activarAccesoGoogle** → **Ejecutar** → acepta el permiso nuevo ("conectarse a un servicio externo"). El registro debe decir "Listo".
3. **Implementar → Gestionar implementaciones → ✏️ → Nueva versión**.
4. Abre la app: aparece **Entrar con Google**. Después vuelve a activar Touch ID / Face ID en cada equipo.

**Para administrar**
- Quitar a alguien: bórralo de `CORREOS_PERMITIDOS` (sus sesiones dejan de servir al tiro).
- **verSesiones**: muestra qué cuentas tienen sesión abierta. **cerrarTodasLasSesiones**: obliga a todos a entrar de nuevo con Google (úsalo si pierdes un equipo).
- Para volver al PIN, borra `CORREOS_PERMITIDOS`.

## Pestañas personales (Rendición, Deporte y Salud)

Además de las pestañas comunes, cada persona ve las suyas según quién eligió al entrar:

- **Vicente → Rendición:** sube la foto del comprobante (o un PDF), el monto y la descripción. La foto queda en la carpeta **Rendiciones - comprobantes** de tu Google Drive, la fila en la hoja **Rendiciones** y se envía un correo automático con la foto adjunta. Cada rendición se puede marcar como *Pagada* cuando te la reembolsen.
- **Vicente → Deporte:** cuatro pestañas.
  - **Hoy:** sugiere qué zona toca según lo que llevas más días sin trabajar y arma una rutina de ~60 min (calentamiento, rodilla, fuerza, core, cardio, calma). Cada ejercicio se puede cambiar (↻) o dejar de proponer (✕). Los saltos e impacto están apagados por la rodilla.
  - **Registro:** qué hiciste cada día (Upper, Lower, Core, Cardio, Full body, Movilidad), minutos y ejercicios con series/peso. Queda en la hoja **Entrenamientos**.
  - **Progreso:** entrenamientos de la semana vs. la meta, calendario de 5 semanas, cada zona hace cuánto, último registro de cada ejercicio y peso vs. la meta.
  - **Diario:** simple: la misma tarjeta de proteína diaria de Salud (hoja **Proteína**, con la persona) y los tabacos del día, cada uno con su hora (hoja **Tabacos Vicente**; el total por día sigue en la columna Cigarros de **Diario Vicente**). Lo anotado antes de comida y peso se mantiene en la planilla.
  - Metas (se cambian desde la app): `DEPORTE_META_SEMANA` (4) y `DEPORTE_META_PESO` (76 kg) en Propiedades del script.
- **Fran → Salud:** contador de días sin azúcar (con récord y calendario de las últimas 5 semanas) y proteína diaria: un registro por día con el total de gramos y una nota opcional (base de huevos, comí azúcar, cómo se sintió…), meta, gráfico de los últimos 7 días y días anteriores con su nota. El mismo bloque de proteína se usa en Deporte de Vicente; cada persona tiene sus datos y su meta. Los datos quedan en las hojas **Azúcar** y **Proteína**.

Para activarlas (una sola vez):

1. Pega el `Code.gs` nuevo en Apps Script y guarda.
2. En **Propiedades del script** agrega `CORREO_RENDICION` con el correo al que deben llegar las rendiciones (puedes poner varios separados por coma).
3. Elige la función **setup** y presiona **Ejecutar**. Google pedirá permisos nuevos (Drive para guardar las fotos y Gmail para enviar el correo): acéptalos igual que la primera vez.
4. **Implementar → Gestionar implementaciones → ✏️ → Nueva versión → Implementar.**

> Qué pestañas ve cada uno se define en `index.html`, en `PESTANAS_DE`. Ojo: es una separación de comodidad, no de seguridad; los dos usan el mismo PIN.

## Formato de la planilla (que se vea limpia)

El diseño (encabezados azul oscuro, filas alternadas, anchos de columna, montos con `$`, ceros como "–" en los resúmenes, colores y orden de pestañas) está al final de `Code.gs`, en la sección **FORMATO VISUAL**. Solo cambia la apariencia, nunca los datos.

- Se aplica solo la primera vez que se abre la app con una versión nueva (`FORMATO_VERSION`), y los resúmenes quedan con formato cada vez que se reescriben.
- Para aplicarlo a mano (por ejemplo, después de agregar una hoja o columna): elige la función **darFormato** y presiona **Ejecutar**.
- Columnas nuevas: agrégalas en `COLUMNAS_ESTILO` (ancho, formato y alineación). Las que no estén ahí igual quedan con el estilo general.

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
