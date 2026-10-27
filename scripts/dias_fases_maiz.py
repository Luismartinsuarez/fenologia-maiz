#!/usr/bin/env python3
"""
Fecha de siembra y días transcurridos desde la siembra hasta el inicio de cada
fase del maíz, por estación y variedad.

Usa el mismo cálculo de campañas y fases que fases_maiz.py, pero entrega solo
una tabla limpia: los días que quedaron marcados para revisar (fase anterior a
la siembra, a más de un año, o fuera de orden) se dejan vacíos y se anotan en
la columna 'observacion'.

Uso:
    python dias_fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv
    python dias_fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv -o dias.xlsx --csv
    python dias_fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv --estacion "SANTO DOMINGO"
    python dias_fases_maiz.py datos/processed/FENOLOGIA_consolidado.csv --variedad "atlas 777"

Requiere: pandas, openpyxl
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from consolidar_fenologia import norm  # noqa: E402
from fases_maiz import FASES, calcular, escribir, leer_base, norm_variedad  # noqa: E402

FASES_DIAS = [f for f in FASES if f != "siembra"]
COLS_ID = ["ddrr", "departamento", "provincia", "distrito", "estacion", "variedad"]


def tabla_campanias(largo, ancho):
    """Una fila por estación + variedad + fecha de siembra con los días a cada fase."""
    ok = largo[largo["revisar"].isna() & largo["fase"].isin(FASES_DIAS)]
    dias = ok.pivot(index="ciclo_id", columns="fase", values="dias_desde_siembra")
    dias = dias.reindex(columns=FASES_DIAS).add_prefix("dias_")

    obs = (largo[largo["revisar"].notna() & largo["fase"].isin(FASES_DIAS)]
                .groupby("ciclo_id")
                .apply(lambda g: "; ".join(f"{f}: {r}" for f, r in zip(g["fase"], g["revisar"])),
                       include_groups=False)
                .rename("observacion"))

    t = (ancho[["ciclo_id"] + COLS_ID + ["fecha_siembra"]]
         .merge(dias, on="ciclo_id", how="left")
         .merge(obs, on="ciclo_id", how="left"))
    t = t[t["fecha_siembra"].notna()]
    t["anio_siembra"] = t["fecha_siembra"].dt.year
    t["fecha_siembra"] = t["fecha_siembra"].dt.date
    t = t.sort_values(["ddrr", "estacion", "variedad", "fecha_siembra"]).reset_index(drop=True)
    cols = COLS_ID + ["fecha_siembra", "anio_siembra"] + [f"dias_{f}" for f in FASES_DIAS] + ["observacion"]
    return t[cols]


def tabla_promedios(camp):
    """Promedio, mínimo y máximo de días por estación y variedad."""
    dcols = [f"dias_{f}" for f in FASES_DIAS]
    g = camp.groupby(COLS_ID, dropna=False)
    base = g.agg(campanias=("fecha_siembra", "size"),
                 fechas_siembra=("fecha_siembra",
                                 lambda s: ", ".join(d.strftime("%d/%m/%Y") for d in sorted(s))))
    prom = g[dcols].mean().round(0).add_prefix("prom_")
    mn = g[dcols].min().add_prefix("min_")
    mx = g[dcols].max().add_prefix("max_")
    # Columnas agrupadas por fase: prom, min, max
    orden = [c for f in FASES_DIAS for c in (f"prom_dias_{f}", f"min_dias_{f}", f"max_dias_{f}")]
    return base.join(prom).join(mn).join(mx)[["campanias", "fechas_siembra"] + orden].reset_index()


LEEME = [
    ("por_campania", "Una fila por estación + variedad + fecha de siembra (campaña)"),
    ("por_estacion_variedad", "Promedio, mínimo y máximo de días por estación y variedad, con la lista de fechas de siembra usadas"),
    ("dias_<fase>", "Días desde la siembra hasta el inicio de la fase. La fecha de inicio es la más repetida en los reportes semanales"),
    ("observacion", "Fases con fecha dudosa (anterior a la siembra, a más de 365 días o fuera de orden): su valor de días se deja vacío"),
    ("variedad", "Variedad más reportada en la campaña, en minúsculas y sin tildes"),
    ("fases", " → ".join(FASES_DIAS)),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("entrada", help="Base consolidada (.csv o .xlsx de consolidar_fenologia.py)")
    ap.add_argument("-o", "--salida", help="Excel de salida (por defecto dias_fases_maiz.xlsx junto a la entrada)")
    ap.add_argument("--estacion", action="append", help="Filtrar por estación (se puede repetir)")
    ap.add_argument("--variedad", action="append", help="Filtrar por variedad (se puede repetir)")
    ap.add_argument("--csv", action="store_true", help="Guardar también las dos tablas en CSV")
    a = ap.parse_args()

    entrada = Path(a.entrada)
    salida = Path(a.salida) if a.salida else entrada.with_name("dias_fases_maiz.xlsx")

    ancho, largo, *_ = calcular(leer_base(entrada))
    camp = tabla_campanias(largo, ancho)
    if a.estacion:
        claves = {norm(e) for e in a.estacion}
        camp = camp[camp["estacion"].map(norm).isin(claves)]
    if a.variedad:
        claves = {norm_variedad(v) for v in a.variedad}
        camp = camp[camp["variedad"].isin(claves)]
    if camp.empty:
        sys.exit("No hay campañas con esos filtros.")
    prom = tabla_promedios(camp)

    escribir(salida, {
        "por_campania": camp, "por_estacion_variedad": prom,
        "leeme": pd.DataFrame(LEEME, columns=["hoja / columna", "descripcion"]),
    })
    if a.csv:
        camp.to_csv(salida.with_name(salida.stem + "_por_campania.csv"), index=False, encoding="utf-8-sig")
        prom.to_csv(salida.with_name(salida.stem + "_por_estacion_variedad.csv"), index=False, encoding="utf-8-sig")

    print(f"{len(camp)} campañas, {len(prom)} combinaciones estación-variedad -> {salida}")


if __name__ == "__main__":
    main()
