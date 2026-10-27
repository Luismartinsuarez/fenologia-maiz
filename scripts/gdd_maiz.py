#!/usr/bin/env python3
"""
Grados-día (GDD) acumulados hasta cada fase del maíz, para una estación
fenológica y variedad, usando las temperaturas diarias de una estación
meteorológica. Acepta varias temperaturas base para compararlas.

Las fechas de siembra y de fases salen del mismo cálculo de fases_maiz.py.
Método (el mismo de E2/GDD/maiz_gdd.py):
    GDD = max(min((Tmax + Tmin) / 2, Tu) - Tb, 0)

Formatos de clima aceptados:
  - clima_estaciones_diario.csv (notebooks/relacion_estaciones.ipynb): columnas
    codigo, estacion_fenologia, fecha, tmax, tmin. Por defecto usa la estación
    climática relacionada con la estación fenológica (--met acepta un código).
  - mantaro.csv: columnas '<estación> - Temperatura Máxima/Minima - Val (ºC)'
    (--met es el nombre de la estación dentro del archivo).

Uso:
    python gdd_maiz.py datos/processed/FENOLOGIA_consolidado.csv datos/processed/clima_estaciones_diario.csv \\
        --estacion CO-CONTUMAZA --variedad amarillo --tb 7 10
    python gdd_maiz.py datos/processed/FENOLOGIA_consolidado.csv datos/processed/clima_estaciones_diario.csv \\
        --estacion todas --variedad amarillo --tb 7 10
    python gdd_maiz.py datos/processed/FENOLOGIA_consolidado.csv datos/raw/mantaro.csv \\
        --estacion "SANTA ANA" --variedad "blanco urubamba" --met "J2 - El Mantaro"

--estacion acepta varias estaciones, o 'todas' (todas las del archivo de clima
con campañas de la variedad; solo con clima_estaciones_diario.csv).

Requiere: pandas, openpyxl
"""
import argparse
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from consolidar_fenologia import norm  # noqa: E402
from fases_maiz import FASES, calcular, escribir, leer_base, norm_variedad  # noqa: E402

MET_MANTARO = "J2 - El Mantaro"


def _diaria(d, col_fecha):
    t = (d.assign(**{col_fecha: pd.to_datetime(d[col_fecha]).dt.normalize()})
          .groupby(col_fecha)[["tmax", "tmin"]].mean()
          .asfreq("D"))
    t.index.name = "fecha"
    malos = t["tmin"] > t["tmax"]
    t.loc[malos, ["tmin", "tmax"]] = None
    return t


def es_formato_diario(d):
    return {"estacion_fenologia", "fecha", "tmax", "tmin", "codigo"} <= set(d.columns)


def estaciones_con_clima(d):
    """Nombres de estación fenológica presentes en clima_estaciones_diario.csv."""
    return sorted({n for ns in d["estacion_fenologia"].dropna().str.split(r"\s*\|\s*") for n in ns})


def leer_clima(d, ruta, estacion_fen, met=None):
    """Devuelve (serie diaria tmax/tmin, descripción de la estación climática)."""
    if es_formato_diario(d):
        if met:
            sel = d[d["codigo"].str.zfill(6) == str(met).zfill(6)]
        else:
            nombres = d["estacion_fenologia"].fillna("").str.split(r"\s*\|\s*")
            sel = d[nombres.map(lambda ns: norm(estacion_fen) in {norm(n) for n in ns})]
        if sel.empty:
            disp = sorted(d["estacion_fenologia"].dropna().unique())
            sys.exit(f"No hay clima para '{met or estacion_fen}' en {Path(ruta).name}. "
                     f"Estaciones disponibles: {disp}")
        desc = f"código {sel['codigo'].iloc[0]} ({sel['archivo'].iloc[0] if 'archivo' in sel else ''})"
        return _diaria(sel, "fecha"), desc

    met = met or MET_MANTARO
    cols = {}
    for c in d.columns:
        m = re.fullmatch(rf"{re.escape(met)} - Temperatura (M[aá]xima|Minima|M[ií]nima) - Val \(ºC\)", c)
        if m:
            cols[c] = "tmax" if norm(m.group(1)) == "maxima" else "tmin"
    if set(cols.values()) != {"tmax", "tmin"}:
        disponibles = sorted({c.split(" - Temperatura")[0] for c in d.columns if " - Temperatura M" in c})
        sys.exit(f"No hay Tmax/Tmin para '{met}'. Estaciones en el archivo: {disponibles}")
    return _diaria(d.rename(columns=cols), "Fecha"), met


def gdd_diario(t, tb, tu):
    """GDD = max(min((Tmax+Tmin)/2, Tu) - Tb, 0); NaN si falta un dato."""
    tmed = (t["tmax"] + t["tmin"]) / 2
    return tmed, (tmed.clip(upper=tu) - tb).clip(lower=0)


def sfx(tb):
    return f"_tb{tb:g}"


def gdd_campania(clima, siembra, fases, tbs, tu):
    """Serie diaria desde la siembra hasta la última fase y GDD acumulado en cada fase, por Tb."""
    fin = fases["fecha_inicio"].max()
    t = clima.reindex(pd.date_range(siembra, fin, freq="D"))
    t["tmed"] = (t["tmax"] + t["tmin"]) / 2
    t["sin_dato"] = t["tmed"].isna()
    t["dias_sin_dato_acum"] = (t["sin_dato"] & (t.index > siembra)).cumsum()

    f = fases.sort_values("fecha_inicio").copy()
    f["dias_desde_fase_anterior"] = f["dias_desde_siembra"].diff().fillna(f["dias_desde_siembra"])
    f["dias_sin_dato"] = f["fecha_inicio"].map(t["dias_sin_dato_acum"])
    for tb in tbs:
        s = sfx(tb)
        _, t[f"gdd{s}"] = gdd_diario(t, tb, tu)
        # GDD del día de siembra no se cuenta: se acumula desde el día siguiente
        t[f"gdd_acum{s}"] = t[f"gdd{s}"].fillna(0).where(t.index > siembra, 0).cumsum()
        f[f"gdd_acumulado{s}"] = f["fecha_inicio"].map(t[f"gdd_acum{s}"])
        f[f"gdd_desde_fase_anterior{s}"] = f[f"gdd_acumulado{s}"].diff().fillna(f[f"gdd_acumulado{s}"])
        f[f"gdd_por_dia{s}"] = (f[f"gdd_desde_fase_anterior{s}"] / f["dias_desde_fase_anterior"]).round(2)

    # Fase vigente en cada día (la última que ya empezó)
    t["fase"] = None
    for _, r in f.iterrows():
        t.loc[t.index >= r["fecha_inicio"], "fase"] = r["fase"]
    t["fase"] = t["fase"].fillna("siembra")
    return t, f


def resumen(fases_df, tbs, por=()):
    """Por fase (y por las columnas de 'por'), entre campañas: media, mínimo, máximo y
    coeficiente de variación de días y GDD. La Tb con menor CV de GDD es la que hace
    más estable la fase entre años y lugares. Excluye las campañas con clima incompleto."""
    v = fases_df[(fases_df["fase"] != "siembra") & fases_df["clima_completo"]]
    medidas = [("dias", "dias_desde_siembra")] + [(f"gdd{sfx(tb)}", f"gdd_acumulado{sfx(tb)}") for tb in tbs]
    filas = []
    for clave, gv in (v.groupby(list(por)) if por else [((), v)]):
        clave = clave if isinstance(clave, tuple) else (clave,)
        for fase in [f for f in FASES if f in set(gv["fase"])]:
            g = gv[gv["fase"] == fase]
            fila = {**dict(zip(por, clave)), "fase": fase, "campanias": len(g)}
            for nombre, col in medidas:
                x = g[col].astype(float)
                fila[f"{nombre}_media"] = round(x.mean(), 1)
                fila[f"{nombre}_min"] = round(x.min(), 1)
                fila[f"{nombre}_max"] = round(x.max(), 1)
                fila[f"{nombre}_cv_pct"] = (round(100 * x.std() / x.mean(), 1)
                                            if len(x) > 1 and x.mean() else None)
            filas.append(fila)
    return pd.DataFrame(filas)


def gdd_estacion(dclima, ruta_clima, estacion, sel, largo, met, tbs, tu, max_sin_dato):
    """GDD de todas las campañas (filas de 'sel') de una estación."""
    clima, met_desc = leer_clima(dclima, ruta_clima, estacion, met)
    tabla_fases, diarios, avisos = [], [], []
    for _, c in sel.iterrows():
        siembra = c["fecha_siembra"]
        fases = largo[(largo["ciclo_id"] == c["ciclo_id"]) & (largo["fase"] != "siembra")]
        for _, r in fases[fases["revisar"].notna()].iterrows():
            avisos.append(f"Siembra {siembra:%d/%m/%Y}, {r['fase']}: {r['revisar']} (se excluye)")
        fases = fases[fases["revisar"].isna()][["fase", "fecha_inicio", "dias_desde_siembra"]]
        if fases.empty:
            avisos.append(f"Siembra {siembra:%d/%m/%Y}: sin fases con fecha válida")
            continue
        if siembra < clima.index.min() or fases["fecha_inicio"].max() > clima.index.max():
            avisos.append(f"Siembra {siembra:%d/%m/%Y}: la campaña sale del período del clima "
                          f"({clima.index.min():%d/%m/%Y} a {clima.index.max():%d/%m/%Y})")
        t, f = gdd_campania(clima, siembra, fases, tbs, tu)
        base = {"estacion": c["estacion"], "variedad": c["variedad"], "fecha_siembra": siembra.date()}
        fila_siembra = {**base, "fase": "siembra", "fecha_inicio": siembra.date(), "dias_desde_siembra": 0,
                        "dias_sin_dato": 0, **{f"gdd_acumulado{sfx(tb)}": 0.0 for tb in tbs}}
        f["fecha_inicio"] = f["fecha_inicio"].dt.date
        tabla_fases += [fila_siembra] + [{**base, **r} for r in f.to_dict("records")]
        diarios.append(t.reset_index(names="fecha").assign(**base, fecha=lambda x: x["fecha"].dt.date))
        n_sin = int((t["sin_dato"] & (t.index > siembra)).sum())
        pct_sin = 100 * n_sin / max(len(t) - 1, 1)
        completo = pct_sin <= max_sin_dato
        for fila in tabla_fases[-(len(f) + 1):]:
            fila["pct_dias_sin_dato"] = round(pct_sin, 1)
            fila["clima_completo"] = completo
        if n_sin:
            avisos.append(f"Siembra {siembra:%d/%m/%Y}: {n_sin} días sin temperatura ({pct_sin:.0f} %, cuentan como 0 GDD)"
                          + ("" if completo else f"; se excluye de los resúmenes (> {max_sin_dato:g} %)"))

    return tabla_fases, diarios, avisos, met_desc


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("fenologia", help="Base consolidada (.csv o .xlsx de consolidar_fenologia.py)")
    ap.add_argument("clima", help="CSV diario de clima (clima_estaciones_diario.csv o mantaro.csv)")
    ap.add_argument("--estacion", required=True, nargs="+",
                    help="Estación(es) fenológica(s), p. ej. CO-CONTUMAZA, o 'todas'")
    ap.add_argument("--variedad", required=True, help="Variedad de maíz, p. ej. 'amarillo'")
    ap.add_argument("--met", help="Estación climática: código SENAMHI (clima_estaciones_diario.csv) "
                                  f"o nombre dentro de mantaro.csv (por defecto '{MET_MANTARO}')")
    ap.add_argument("--tb", type=float, nargs="+", default=[10.0], help="Temperatura(s) base (°C), p. ej. --tb 7 10")
    ap.add_argument("--tu", type=float, default=30.0, help="Temperatura umbral superior (°C)")
    ap.add_argument("--max-sin-dato", type=float, default=10,
                    help="%% máximo de días sin temperatura para usar la campaña en los resúmenes (por defecto 10)")
    ap.add_argument("-o", "--salida", help="Excel de salida")
    a = ap.parse_args()
    tbs = sorted(set(a.tb))

    ancho, largo, *_ = calcular(leer_base(a.fenologia))
    sel_var = ancho[ancho["variedad"] == norm_variedad(a.variedad)]
    dclima = pd.read_csv(a.clima, dtype={"codigo": str}, low_memory=False)

    if a.estacion == ["todas"]:
        if not es_formato_diario(dclima):
            sys.exit("--estacion todas solo funciona con clima_estaciones_diario.csv")
        con_clima = {norm(e) for e in estaciones_con_clima(dclima)}
        estaciones = sorted(sel_var.loc[sel_var["estacion"].map(norm).isin(con_clima), "estacion"].unique())
        if not estaciones:
            sys.exit(f"Ninguna estación con clima tiene campañas de '{a.variedad}'")
    else:
        estaciones = a.estacion

    tabla_fases, diarios, avisos, mets = [], [], [], {}
    for est in estaciones:
        sel = sel_var[sel_var["estacion"].map(norm) == norm(est)]
        if sel.empty:
            vars_ = sorted(ancho.loc[ancho["estacion"].map(norm) == norm(est), "variedad"].dropna().unique())
            sys.exit(f"No hay campañas de '{a.variedad}' en '{est}'. Variedades de la estación: {vars_}")
        tf, di, av, mets[est] = gdd_estacion(dclima, a.clima, est, sel, largo, a.met, tbs, a.tu, a.max_sin_dato)
        tabla_fases += tf
        diarios += di
        avisos += [f"{est} · {x}" for x in av]

    if not tabla_fases:
        sys.exit("Ninguna campaña tiene fases con fecha válida.\n" + "\n".join(avisos))

    fases_df = pd.DataFrame(tabla_fases)
    fases_df["orden"] = fases_df["fase"].map({f: i for i, f in enumerate(FASES)})
    fases_df = fases_df.sort_values(["fecha_siembra", "orden"]).drop(columns="orden")
    cols_gdd = [f"{p}{sfx(tb)}" for tb in tbs for p in ("gdd_acumulado", "gdd_desde_fase_anterior", "gdd_por_dia")]
    fases_df = fases_df[["estacion", "variedad", "fecha_siembra", "fase", "fecha_inicio",
                         "dias_desde_siembra", "dias_desde_fase_anterior", *cols_gdd, "dias_sin_dato",
                         "pct_dias_sin_dato", "clima_completo"]]
    fases_df = fases_df.round({c: 1 for c in cols_gdd if not c.startswith("gdd_por_dia")})
    res = resumen(fases_df, tbs)
    res_est = resumen(fases_df, tbs, por=("estacion",))

    cols_d = [f"{p}{sfx(tb)}" for tb in tbs for p in ("gdd", "gdd_acum")]
    diario = pd.concat(diarios, ignore_index=True)[
        ["estacion", "variedad", "fecha_siembra", "fecha", "tmin", "tmax", "tmed", *cols_d, "fase", "sin_dato"]]
    diario = diario.round({"tmed": 2, **{c: (2 if c.startswith("gdd_tb") else 1) for c in cols_d}})

    tbs_txt = ", ".join(f"{tb:g}" for tb in tbs)
    leeme = pd.DataFrame([
        ("estación fenológica", ", ".join(estaciones)), ("variedad", a.variedad),
        *[("estación meteorológica", f"{e}: {m} · {Path(a.clima).name}") for e, m in mets.items()],
        ("método", f"GDD = max(min((Tmax + Tmin) / 2, Tu) - Tb, 0), con Tb = {tbs_txt} °C y Tu = {a.tu:g} °C"),
        ("columnas _tbX", "Cada columna de GDD se calcula con la temperatura base indicada en el sufijo"),
        ("acumulación", "Desde el día siguiente a la siembra hasta el día de inicio de cada fase (incluido)"),
        ("fechas de fases", "Fecha de inicio más repetida en los reportes semanales (ver fases_maiz.py). "
                            "Se excluyen las marcadas para revisar"),
        ("resumen", "Por fase, entre todas las campañas (de todas las estaciones): media, mínimo, máximo y "
                    "coeficiente de variación (CV) de días y GDD. La Tb que da menor CV de GDD es la que hace "
                    "más estable la fase entre años y lugares"),
        ("resumen_estacion", "Lo mismo, por estación"),
        ("dias_sin_dato", "Días sin temperatura entre la siembra y la fase (cuentan como 0 GDD)"),
        ("clima_completo", f"Campaña con {a.max_sin_dato:g} % o menos de días sin temperatura. "
                           "Solo estas entran en los resúmenes"),
        *[("aviso", x) for x in avisos],
    ], columns=["concepto", "valor"])

    nombre_est = "todas" if a.estacion == ["todas"] else "_".join(norm(e) for e in estaciones)
    salida = Path(a.salida) if a.salida else Path(a.fenologia).with_name(
        f"gdd_maiz_{nombre_est.replace(' ', '_')}_{norm_variedad(a.variedad).replace(' ', '_')}.xlsx")
    hojas = {"resumen": res}
    if len(estaciones) > 1:
        hojas["resumen_estacion"] = res_est
    escribir(salida, {**hojas, "fases": fases_df, "diario": diario, "leeme": leeme})

    print(f"{a.variedad} · Tb = {tbs_txt} °C · Tu = {a.tu:g} °C")
    ver = ["fase", "fecha_inicio", "dias_desde_siembra"] + [f"gdd_acumulado{sfx(tb)}" for tb in tbs]
    for (e, s), g in fases_df.groupby(["estacion", "fecha_siembra"]):
        print(f"\n{e} · siembra {s:%d/%m/%Y} · clima: {mets[e]}")
        print(g[ver].to_string(index=False))
    print("\nResumen entre campañas (CV = coeficiente de variación)")
    ver = ["fase", "campanias", "dias_media", "dias_cv_pct"] + [
        c for tb in tbs for c in (f"gdd{sfx(tb)}_media", f"gdd{sfx(tb)}_cv_pct")]
    print(res[ver].to_string(index=False))
    for x in avisos:
        print("AVISO:", x)
    print(f"\n-> {salida}")


if __name__ == "__main__":
    main()
