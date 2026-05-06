# Generador de Certificados SLB — Sullair Argentina

App para generar automáticamente los certificados de servicio para Schlumberger (SLB) a partir de la Planilla de Certificación mensual.

## Estructura del repositorio

```
├── app.py                  ← App principal Streamlit
├── requirements.txt        ← Dependencias Python
├── assets/
│   ├── logo_sullair.jpg    ← Logo Sullair Argentina
│   └── firma_mastracci.jpg ← Firma Antonella Mastracci
└── README.md
```

## Cómo usar la app

1. Ingresar el **N° de certificado inicial** del mes
2. El **período** se detecta automáticamente del Excel (o se puede escribir manualmente)
3. Subir la **Planilla de Certificación de Servicios Periféricos** (.xlsx)
4. Hacer clic en **Generar PDFs**
5. Descargar el **ZIP con los PDFs** y la **Planilla completada** (con N° de cert en columna O)

## Fuente de datos (columnas de la planilla)

| Campo | Columna |
|-------|---------|
| Segmento | G |
| CVU | D |
| Equipo (descripción) | L |
| Modelo / Interno | M |
| Desde | S |
| Hasta | T |
| N° Certificado (salida) | O |
| Período | B3 + B4 |

## Nomenclatura de archivos PDF

`SEGMENTO - INTERNO - N°CERT.pdf`

Ejemplo: `WCE DPM - E030286 - 684.pdf`
