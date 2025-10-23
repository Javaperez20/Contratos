````markdown
# Tarifario dinámico · Generador de Contratos

Resumen
- Front-end que genera una interfaz de tarifario leyendo `data.xlsx` desde la raíz.
- `data.xlsx` necesita al menos la hoja `catalog` con columnas: `Código`, `Plan`, `Valor`, `Promo1`, `Meses1`, `Promo2`, `Meses2`, `Detalles`.
- Hoja `structure` opcional para controlar secciones/subsecciones y prefijos.

Principales características entregadas
- Interfaz (pestañas/subpestañas) generada automáticamente desde `data.xlsx`.
- Selecciones y toggles construidos desde la hoja `structure`.
- Móvil: tarjeta por línea (principal y adicionales) con segmented toggles, select de plan, checkbox "Porta" que muestra inputs de portabilidad solo si está marcado.
- Detalles: el campo `Detalles` del Excel se muestra preservando saltos de línea; además permite etiquetas HTML simples (p. ej. `<b>`) tal como fueron escritas en la celda.
- Precios: se muestran solo las promociones que tienen valor en Excel (si están vacías, no se muestran).
- Generación de contrato: usa un template `contrato_template.docx`. Las líneas móviles seleccionadas generan párrafos que se pasan a la plantilla en la clave `MOVIL` (cada línea → párrafo, separadas por doble salto).
- No hay controles para subir/descargar/generar Excel en la UI: la app carga siempre `data.xlsx` desde la raíz del proyecto (usa `fetch` por HTTP).

Recomendaciones de despliegue
- Sirve la carpeta con un servidor estático (por ejemplo `python -m http.server` o `npx http-server`) y coloca `data.xlsx` en la misma carpeta que `index.html`.
- Asegúrate que en `data.xlsx` los valores numéricos estén sin separadores de miles o que sean números (el parser intenta normalizar "13.490" → `13490`).

Seguridad y notas
- Actualmente permito que la celda `Detalles` contenga HTML (para respetar formateo). Si necesitas sanitizar (permitir solo un conjunto seguro de etiquetas) lo implemento.
- El contenido usado en el contrato viene del Excel y del formulario; revisa y valida datos en producción si múltiples usuarios editan el Excel.

Siguientes mejoras posibles
- Validación de número cuando Porta está marcado (impedir generar contrato si falta número).
- Sanitización más estricta de HTML en `Detalles` (permitir solo etiquetas seguras).
- Persistencia de última selección (localStorage).
- Exportación del `data.xlsx` de ejemplo al repo o PR automatizado.
