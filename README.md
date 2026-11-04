# Finanzas · Gastos Familia

App web simple para que dos personas anoten gastos desde el celular: se elige el centro de costo (Farmacia, Luna, Supermercado, Fran…), el monto y una descripción, y todo queda en una planilla de Google Sheets con resúmenes mensuales y anuales.

- **App:** `index.html` (HTML + JS sin dependencias), publicada con GitHub Pages.
- **Backend:** `apps-script/Code.gs`, un Google Apps Script pegado en la planilla.
- **Datos:** Google Sheets (hojas Gastos, Centros, Personas, Resumen mensual, Resumen anual). Se descarga como Excel cuando quieras.
- **Pestañas personales:** Rendición y Deporte (Vicente) y Salud (Fran: días sin azúcar y proteína diaria).
  - Rendición: gasto + foto con correo automático.
  - Deporte: rutina propuesta del día, registro de entrenamientos, progreso y diario de proteína, comida, cigarros y peso.

Instalación: ver [SETUP.md](SETUP.md).
