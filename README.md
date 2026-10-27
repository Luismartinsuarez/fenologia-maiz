# Fenología del maíz y grados día

Análisis de la fenología de cultivos reportada semanalmente en estaciones agrícolas del Perú, con énfasis en el
maíz: reconstrucción de campañas y fechas de fase, relación con estaciones climáticas del SENAMHI y estimación de
la temperatura base (Tb) y los grados día (GDD) requeridos por variedad.

Los datos no se incluyen en el repositorio.

## Estructura

| Archivo | Qué hace |
|---|---|
| `scripts/consolidar_fenologia.py` | Junta los libros semanales de fenología (una hoja por semana) en una sola tabla |
| `scripts/fases_maiz.py` | Arma las campañas de maíz (estación + siembra), fechas de cada fase y unifica nombres de variedades |
| `scripts/dias_fases_maiz.py` | Días desde la siembra hasta cada fase, por campaña y por estación–variedad |
| `notebooks/relacion_estaciones.ipynb` | Empareja los archivos climáticos del SENAMHI con las estaciones de fenología |
| `notebooks/campanias_estaciones_clima.ipynb` | Campañas y variedades por estación con su cobertura climática |
| `scripts/gdd_maiz.py` | GDD acumulados hasta cada fase para una o varias estaciones, con Tb fijas |
| `scripts/gdd_variedades.py` | Calibra Tb (0–14 °C, métodos medio y acotado) con validación cruzada y da los requerimientos por variedad |

## Datos esperados

```
datos/raw/fenologia/    libros Excel de monitoreo fenológico semanal
datos/raw/estaciones/   archivos diarios de estaciones SENAMHI (.XLS, tablas HTML)
datos/processed/        salidas de los scripts
```

## Uso

Requiere Python 3.10+, `pandas`, `numpy`, `openpyxl` y `jupyter`.

```bash
python scripts/consolidar_fenologia.py datos/raw/fenologia -o datos/processed/FENOLOGIA_consolidado.xlsx --csv
python scripts/fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv
python scripts/dias_fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv --csv
jupyter nbconvert --to notebook --execute --inplace notebooks/relacion_estaciones.ipynb
jupyter nbconvert --to notebook --execute --inplace notebooks/campanias_estaciones_clima.ipynb
python scripts/gdd_maiz.py datos/processed/FENOLOGIA_consolidado.csv datos/processed/clima_estaciones_diario.csv \
    --estacion todas --variedad "atlas 777" --tb 7 10
python scripts/gdd_variedades.py datos/processed/FENOLOGIA_consolidado.csv datos/processed/clima_estaciones_diario.csv
```

## Método de grados día

- **Medio:** GDD = max(min((Tmax + Tmin)/2, Tu) − Tb, 0), con Tu = 30 °C.
- **Acotado** (McMaster y Wilhelm, 1997): Tmax y Tmin se limitan a [Tb, Tu] antes de promediar.

Los GDD se acumulan desde el día siguiente a la siembra. La Tb de cada variedad se elige con el menor error al
predecir el inicio de panoja, espiga y maduración córnea dejando una campaña fuera.
