# Tablero Venta B2B — Café Quindío · publicación automática

Este repositorio **arma y publica el tablero solo**. Tú solo subes el Excel del día;
GitHub Actions reprocesa todo y Vercel lo cuelga en la web en ~2 minutos.

## Cómo funciona
1. Los Excel fuente viven en `data/`.
2. Cuando cambia algo en `data/` (o en `pipeline/` / `web/`), **GitHub Actions** ejecuta
   `pipeline/build.py`, que reprocesa la venta (exclusiones B2B, filtros por
   vendedor/director/razón social/sucursal, lanzamiento, comparativo 2025, kilos) y
   genera el `index.html`.
2. El Action hace *commit* del `index.html` y **Vercel** publica esa versión automáticamente.

```
Proyecto_Auto_Vercel/
├── data/                         # Excel fuente (los que reemplazas)
│   ├── 8. INFORME FACTURACION Y AGOTADOS AGOSTO 2026.xlsx
│   ├── 7. INFORME FACTURACION Y AGOTADOS JULIO 2026..xlsx
│   ├── 6. INFORME FACTURACION Y AGOTADOS JUNIO 2026..xlsx
│   ├── INFORME FACTURACION Y AGOTADOS 2025.xlsx
│   └── 8. AVANCE PRESUPUESTAL AGOSTO 2026.xlsx
├── pipeline/
│   ├── build.py                  # genera index.html
│   └── config.json               # etiquetas, meses, presupuesto Julio, exclusiones
├── web/
│   └── plantilla.html            # tablero SIN datos (marcador /*__CQDATA__*/)
├── .github/workflows/build.yml   # la automatización (GitHub Actions)
├── index.html                    # SALIDA generada (la publica Vercel)
├── requirements.txt
├── vercel.json
└── README.md
```

## Montaje (una sola vez)
1. En el repo **Daniduenas/Ventas-diarias** → **Add file → Upload files**.
2. **Arrastra las carpetas y archivos de `Proyecto_Auto_Vercel`** (data, pipeline, web,
   .github, index.html, requirements.txt, vercel.json, README.md). GitHub conserva las subcarpetas.
3. **Commit changes.** Al terminar, entra a la pestaña **Actions** del repo: verás la
   ejecución "Construir y publicar tablero". Cuando termine (verde), Vercel publica.
4. Verifica que en **Vercel → el proyecto → Settings → Git** el *Root Directory* esté en la
   raíz (vacío) y que despliegue la rama principal.

> Si la pestaña Actions pide habilitar los workflows, dale "I understand… enable".

## Uso diario (lo único manual)
1. Entra al repo → carpeta **`data/`**.
2. **Add file → Upload files** y sube el **Informe de Agosto** nuevo
   (mismo nombre `8. INFORME FACTURACION Y AGOTADOS AGOSTO 2026.xlsx` → reemplaza).
3. **Commit changes.** Listo: el Action reprocesa y Vercel publica solo en ~2 min.

No necesitas volver a tocar el `index.html` nunca.

## Notas
- **Cambio de mes (p. ej. Septiembre):** avísame y actualizo `config.json`, `pipeline/build.py`
  y la plantilla para incluir el mes nuevo; también podrás subir el nuevo Informe/Avance a `data/`.
- **Presupuesto de Agosto** se lee del `8. AVANCE PRESUPUESTAL AGOSTO 2026.xlsx`; el de Julio
  está fijo en `config.json`.
- **Acceso:** el sitio es público. Para restringir: Vercel → Project → Settings →
  Deployment Protection → Password Protection.
- El build vuelve a leer el archivo de 2025 (grande) en cada corrida; tarda ~1–2 min en Actions.
