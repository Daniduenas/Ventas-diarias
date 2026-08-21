# Tablero Venta Diaria B2B — Café Quindío

Proyecto web (sitio estático) del tablero de seguimiento de venta diaria del canal B2B.
Es un único archivo autocontenido (`index.html`) con los datos ya incrustados: no necesita
servidor ni base de datos. Se abre en el navegador y se despliega en Vercel tal cual.

## Contenido del proyecto

```
Proyecto_Tablero_CQ/
├── index.html      # El tablero (datos + logo embebidos, Chart.js por CDN)
├── vercel.json     # Configuración de Vercel (sitio estático)
├── .gitignore
└── README.md
```

## Qué incluye el tablero
- **Presupuesto vs acumulado** (cumplimiento del mes).
- **Venta día anterior** (vs día previo, mismo día mes anterior y promedio).
- **Venta por producto** (clic en la barra de categoría para filtrar).
- **Detalle** por vendedor / director / línea / producto, con filtros globales.
- **Línea de tiempo** día a día (se filtra por director/vendedor).
- **Lanzamiento** (Café Cocora y Guadual 250 g): pesos, unidades y kilos.
- Filtro de mes (Agosto / Julio 2026) y descarga de la base a Excel.

## Desplegar en Vercel

### Opción A — Arrastrar y soltar (sin GitHub, ~30 s)
1. Entra a **https://vercel.com/new**.
2. Arrastra **esta carpeta** (`Proyecto_Tablero_CQ`) a la zona de despliegue.
3. Clic en **Deploy**. Vercel entrega una URL pública (p. ej. `https://tablero-cq.vercel.app`).

### Opción B — GitHub → Vercel (con actualización automática)
1. Crea un repositorio en GitHub y sube estos archivos (o esta carpeta como raíz).
2. En **https://vercel.com/new** elige *Import Git Repository* y selecciona el repo.
3. *Deploy*. Cada push al repo redepliega solo.

> No requiere *Build Command* ni *Output Directory*: Vercel detecta el sitio estático
> y sirve `index.html` en la raíz.

## Cómo actualizar los datos
Hoy los datos vienen "horneados" dentro de `index.html`. Para refrescarlos, se regenera
el `index.html` desde los Excel de facturación (Informe FAC. CONSOLIDADO) y el Avance
Presupuestal, y se vuelve a subir/hacer push. La lógica de reproceso (exclusión de
vendedores no B2B, mapa vendedor→director, kilos estimados del gramaje, referencias de
lanzamiento) está documentada en `ARQUITECTURA_Despliegue_GitHub_Vercel.md`.

## Acceso
Configurado como **público con enlace**. Como el tablero contiene venta por vendedor y
por cliente, si se quiere restringir: en Vercel → *Project → Settings → Deployment
Protection → Password Protection*.
