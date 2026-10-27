#!/usr/bin/env python3
"""
Requerimientos térmicos (GDD) por variedad de maíz: inventario de campañas con
clima, búsqueda de la temperatura base (Tb) y validación cruzada.

Para cada variedad con datos suficientes:
  1. Prueba Tb de --tb-min a --tb-max (cada --paso °C) con dos métodos de cálculo:
       media    GDD = max(min((Tmax + Tmin)/2, Tu) - Tb, 0)       (el de gdd_maiz.py)
       acotada  Tmax y Tmin se limitan a [Tb, Tu] antes de promediar
                (McMaster y Wilhelm 1997, método 2: descuenta el calor sobre Tu)
  2. Validación cruzada dejando una campaña fuera: con la media de GDD de las
     demás campañas se predice el día de inicio de la fase en la campaña dejada
     fuera. Se elige la Tb y el método con menor error (RMSE, en días) en las
     fases de calibración (panoja, espiga y maduración córnea).
  3. Compara ese error con el de predecir con la media de días calendario: si
     los GDD no lo mejoran, no aportan para esa variedad.

Solo usa campañas con clima completo (--max-sin-dato) y fases sin marca de
revisión (ver fases_maiz.py). Las variedades genéricas (VARIEDADES_GENERICAS en
fases_maiz.py) se excluyen salvo que se pidan con --variedad.

Uso:
    python gdd_variedades.py datos/processed/FENOLOGIA_consolidado.csv datos/processed/clima_estaciones_diario.csv
    python gdd_variedades.py ... --variedad "atlas 777" "blanco urubamba" --tb-min 4 --tb-max 12

Requiere: numpy, pandas, openpyxl
"""
import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from consolidar_fenologia import norm  # noqa: E402
from fases_maiz import FASES, VARIEDADES_GENERICAS, calcular, escribir, leer_base, norm_variedad  # noqa: E402
from gdd_maiz import es_formato_diario, estaciones_con_clima, leer_clima  # noqa: E402

FASES_CALIB = ["panoja", "espiga", "maduracion cornea"]   # floración y madurez fisiológica
FASES_TABLA = [f for f in FASES if f != "siembra"]
METODOS = ("media", "acotada")
MARGEN_DIAS = 150      # días de clima después de la última fase, para poder predecir más tarde
RANGO_TMED_MIN = 3.0   # °C de diferencia de Tmedia entre campañas para que la Tb sea identificable
N_CONFIABLE = 8        # campañas en maduración córnea para considerar el valor firme
TOL_DIAS = 0.5         # Tb con error hasta esto sobre el mínimo se consideran equivalentes
ANCHO_TB_MAX = 6.0     # intervalo de Tb equivalentes más ancho que esto -> Tb poco definida


# ---------------------------------------------------------------------------
# Datos
# ---------------------------------------------------------------------------

def campanias_con_clima(ancho, largo, dclima, ruta_clima, max_sin_dato):
    """Una entrada por campaña con estación climática: días a cada fase y Tmax/Tmin
    diarias desde la siembra (índice 0) hasta la última fase + MARGEN_DIAS."""
    con_clima = {norm(e) for e in estaciones_con_clima(dclima)}
    fases_ok = largo[largo["revisar"].isna() & (largo["fase"] != "siembra")]
    cache, out = {}, []
    for _, c in ancho[ancho["fecha_siembra"].notna()].iterrows():
        est = c["estacion"]
        if norm(est) not in con_clima:
            continue
        if est not in cache:
            cache[est] = leer_clima(dclima, ruta_clima, est)[0]
        f = fases_ok[fases_ok["ciclo_id"] == c["ciclo_id"]]
        f = f[f["dias_desde_siembra"] > 0]
        if f.empty:
            continue
        dias = dict(zip(f["fase"], f["dias_desde_siembra"].astype(int)))
        ult = max(dias.values())
        s = c["fecha_siembra"]
        t = cache[est].reindex(pd.date_range(s, s + pd.Timedelta(days=ult + MARGEN_DIAS), freq="D"))
        tx, tn = t["tmax"].to_numpy(float), t["tmin"].to_numpy(float)
        falta = np.isnan(tx[1:ult + 1]) | np.isnan(tn[1:ult + 1])
        tmed = (tx[1:ult + 1] + tn[1:ult + 1]) / 2
        out.append({
            "ciclo_id": c["ciclo_id"], "estacion": est, "variedad": c["variedad"],
            "fecha_siembra": s, "dias": dias, "tx": tx, "tn": tn,
            "pct_sin_dato": round(100 * falta.mean(), 1),
            "clima_completo": 100 * falta.mean() <= max_sin_dato,
            "tmed": round(float(np.nanmean(tmed)), 2) if (~falta).any() else None,
        })
    return out


def inventario(ancho, camps, variedades_pedidas, min_camp, min_est):
    """Campañas por variedad y si entran en el cálculo."""
    c = pd.DataFrame([{k: x[k] for k in ("ciclo_id", "estacion", "variedad", "clima_completo")} for x in camps])
    tot = ancho.groupby("variedad").agg(campanias=("ciclo_id", "size"), estaciones=("estacion", "nunique"))
    cc = c.groupby("variedad").agg(campanias_con_clima=("ciclo_id", "size"))
    ok = (c[c["clima_completo"]].groupby("variedad")
          .agg(campanias_clima_completo=("ciclo_id", "size"), estaciones_clima_completo=("estacion", "nunique")))
    inv = tot.join(cc).join(ok).fillna(0).astype(int).reset_index()
    inv["generica"] = inv["variedad"].isin(VARIEDADES_GENERICAS)

    def motivo(r):
        if variedades_pedidas is not None:
            return None if r["variedad"] in variedades_pedidas else "no pedida"
        if r["generica"]:
            return "nombre genérico"
        if r["campanias_clima_completo"] < min_camp:
            return f"menos de {min_camp} campañas con clima completo"
        if r["estaciones_clima_completo"] < min_est:
            return f"menos de {min_est} estaciones con clima completo"
        return None

    inv["motivo_exclusion"] = inv.apply(motivo, axis=1)
    inv["incluida"] = inv["motivo_exclusion"].isna() & (inv["campanias_clima_completo"] >= 2)
    return inv.sort_values(["incluida", "campanias_clima_completo", "campanias"], ascending=False)


# ---------------------------------------------------------------------------
# GDD y validación
# ---------------------------------------------------------------------------

def gdd_dia(tx, tn, tb, tu, metodo):
    if metodo == "media":
        g = np.clip(np.minimum((tx + tn) / 2, tu) - tb, 0, None)
    else:
        g = (np.clip(tx, tb, tu) + np.clip(tn, tb, tu)) / 2 - tb
    return np.nan_to_num(g)            # día sin dato = 0 GDD (igual que gdd_maiz.py)


def acumulado(c, tb, tu, metodo):
    """acum[d] = GDD desde el día siguiente a la siembra hasta el día d (incluido)."""
    g = gdd_dia(c["tx"], c["tn"], tb, tu, metodo)
    g[0] = 0
    return np.cumsum(g)


def evaluar(camps, tb, tu, metodo, fases, min_n):
    """Por fase: GDD de cada campaña y error de predicción dejando una fuera (con
    GDD y con días calendario). Devuelve (resumen por fase, errores por campaña)."""
    acums = [acumulado(c, tb, tu, metodo) for c in camps]
    res, errores = [], []
    for fase in fases:
        idx = [i for i, c in enumerate(camps) if fase in c["dias"]]
        if len(idx) < min_n:
            continue
        d = np.array([camps[i]["dias"][fase] for i in idx], float)
        g = np.array([acums[i][int(camps[i]["dias"][fase])] for i in idx])
        e_gdd, e_dias = [], []
        for k, i in enumerate(idx):
            otros = np.arange(len(idx)) != k
            meta = g[otros].mean()
            pred = int(np.searchsorted(acums[i], meta))     # primer día que alcanza la meta
            pred = pred if pred < len(acums[i]) else np.nan
            e_gdd.append(pred - d[k])
            e_dias.append(d[otros].mean() - d[k])
            errores.append({"ciclo_id": camps[i]["ciclo_id"], "estacion": camps[i]["estacion"],
                            "fecha_siembra": camps[i]["fecha_siembra"].date(), "tmed": camps[i]["tmed"],
                            "fase": fase, "dias_observados": int(d[k]), "gdd_observado": round(g[k], 1),
                            "gdd_meta": round(meta, 1), "dias_predichos_gdd": pred,
                            "error_gdd_dias": pred - d[k],
                            "dias_predichos_calendario": round(d[otros].mean(), 1),
                            "error_calendario_dias": round(d[otros].mean() - d[k], 1)})
        e_gdd, e_dias = np.array(e_gdd), np.array(e_dias)
        res.append({
            "fase": fase, "campanias": len(idx),
            "estaciones": len({camps[i]["estacion"] for i in idx}),
            "dias_media": round(d.mean(), 1), "dias_min": int(d.min()), "dias_max": int(d.max()),
            "dias_cv_pct": round(100 * d.std(ddof=1) / d.mean(), 1),
            "gdd_media": round(g.mean(), 1), "gdd_de": round(g.std(ddof=1), 1),
            "gdd_min": round(g.min(), 1), "gdd_max": round(g.max(), 1),
            "gdd_cv_pct": round(100 * g.std(ddof=1) / g.mean(), 1) if g.mean() else None,
            "rmse_gdd_dias": round(float(np.sqrt(np.nanmean(e_gdd ** 2))), 1),
            "rmse_calendario_dias": round(float(np.sqrt(np.mean(e_dias ** 2))), 1),
        })
    return pd.DataFrame(res), errores


def objetivo(res):
    """Error medio (días) de las fases de calibración evaluadas."""
    r = res[res["fase"].isin(FASES_CALIB)] if len(res) else res
    if r.empty:
        return None, None
    return r["rmse_gdd_dias"].mean(), r["rmse_calendario_dias"].mean()


def calibrar(camps, tbs, tu, min_n):
    """Búsqueda de Tb y método. Devuelve (tabla de la búsqueda, mejor (metodo, tb))."""
    filas = []
    for metodo in METODOS:
        for tb in tbs:
            res, _ = evaluar(camps, tb, tu, metodo, FASES_CALIB, min_n)
            err, err_cal = objetivo(res)
            fila = {"metodo": metodo, "tb": tb, "rmse_gdd_dias": None if err is None else round(err, 2),
                    "rmse_calendario_dias": None if err_cal is None else round(err_cal, 2)}
            for _, r in res.iterrows():
                fila[f"rmse_{r['fase']}"] = r["rmse_gdd_dias"]
                fila[f"cv_gdd_{r['fase']}"] = r["gdd_cv_pct"]
            filas.append(fila)
    b = pd.DataFrame(filas)
    if b["rmse_gdd_dias"].isna().all():
        return b, None
    mejor = b.loc[b["rmse_gdd_dias"].idxmin()]
    return b, (mejor["metodo"], float(mejor["tb"]))


def intervalo_tb(b, metodo, tol=TOL_DIAS):
    """Tb mínima y máxima cuyo error queda a menos de 'tol' días del mínimo (mismo método)."""
    m = b[b["metodo"] == metodo]
    ok = m[m["rmse_gdd_dias"] <= m["rmse_gdd_dias"].min() + tol]["tb"]
    return float(ok.min()), float(ok.max())


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

LEEME = [
    ("requerimientos", "Una fila por variedad: Tb y método elegidos, GDD y días hasta cada fase (la Tb se elige con panoja, espiga y maduración córnea), "
                       "error de predicción y confiabilidad"),
    ("por_fase", "Todas las fases de cada variedad con la Tb elegida: días y GDD (media, desviación, mínimo, "
                 "máximo, CV) y error de predicción"),
    ("por_estacion", "Lo mismo por estación (sin validación)"),
    ("busqueda_tb", "Error (RMSE en días) y CV de GDD para cada Tb y método probados"),
    ("validacion", "Predicción de cada campaña dejada fuera, con la Tb elegida y con días calendario"),
    ("campanias", "Campañas usadas, con su Tmedia (siembra a última fase) y GDD a cada fase con la Tb elegida"),
    ("inventario", "Todas las variedades: campañas, campañas con clima completo y motivo si no se calcularon"),
    ("método media", "GDD = max(min((Tmax + Tmin)/2, Tu) - Tb, 0). Es el mismo de gdd_maiz.py"),
    ("método acotada", "Tmax y Tmin se limitan a [Tb, Tu] antes de promediar (McMaster y Wilhelm 1997, método 2)"),
    ("acumulación", "Desde el día siguiente a la siembra hasta el día de inicio de la fase (incluido). "
                    "Días sin temperatura = 0 GDD; solo se usan campañas con clima completo"),
    ("validación", "Dejando una campaña fuera: la media de GDD de las demás predice el día de inicio de la fase. "
                   "RMSE = raíz del error cuadrático medio, en días. 'calendario' = predecir con la media de días"),
    ("elección de Tb", "Tb y método con menor RMSE medio en panoja, espiga y maduración córnea. La Tb se elige y se "
                       "valida con los mismos datos, así que el error real será algo mayor"),
    ("tb_intervalo", f"Tb con error a menos de {TOL_DIAS:g} días del mínimo: en ese intervalo cualquier Tb predice igual"),
    ("confiabilidad", f"alta: sin observaciones; media: una; baja: dos o más. Observaciones: rango de Tmedia "
                      f"< {RANGO_TMED_MIN:g} °C (Tb poco identificable), Tb en el borde de la búsqueda, "
                      f"intervalo de Tb de {ANCHO_TB_MAX:g} °C o más, "
                      f"GDD no mejoran a los días calendario, menos de {N_CONFIABLE} campañas en maduración córnea"),
    ("emergencia y hojas", "Con reportes semanales se fechan mal (CV alto): se muestran pero no se usan para elegir Tb"),
    ("cosecha", "Depende de la decisión del productor más que de la fenología: tomar con cautela"),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("fenologia", help="Base consolidada (.csv o .xlsx de consolidar_fenologia.py)")
    ap.add_argument("clima", help="clima_estaciones_diario.csv (notebooks/relacion_estaciones.ipynb)")
    ap.add_argument("--variedad", nargs="+", help="Variedades a calcular (por defecto: todas las que cumplen los mínimos)")
    ap.add_argument("--min-campanias", type=int, default=5, help="Campañas con clima completo (por defecto 5)")
    ap.add_argument("--min-estaciones", type=int, default=2, help="Estaciones con clima completo (por defecto 2)")
    ap.add_argument("--min-fase", type=int, default=4,
                    help="Campañas mínimas en una fase para validarla (por defecto 4)")
    ap.add_argument("--tb-min", type=float, default=0.0)
    ap.add_argument("--tb-max", type=float, default=14.0)
    ap.add_argument("--paso", type=float, default=0.5)
    ap.add_argument("--tu", type=float, default=30.0, help="Temperatura umbral superior (°C)")
    ap.add_argument("--max-sin-dato", type=float, default=10,
                    help="%% máximo de días sin temperatura en la campaña (por defecto 10)")
    ap.add_argument("-o", "--salida", help="Excel de salida (por defecto gdd_variedades.xlsx junto a la entrada)")
    a = ap.parse_args()

    dclima = pd.read_csv(a.clima, dtype={"codigo": str}, low_memory=False)
    if not es_formato_diario(dclima):
        sys.exit("El clima debe ser clima_estaciones_diario.csv")
    tbs = np.round(np.arange(a.tb_min, a.tb_max + a.paso / 2, a.paso), 2)

    ancho, largo, *_ = calcular(leer_base(a.fenologia))
    camps = campanias_con_clima(ancho, largo, dclima, a.clima, a.max_sin_dato)
    pedidas = {norm_variedad(v) for v in a.variedad} if a.variedad else None
    inv = inventario(ancho, camps, pedidas, a.min_campanias, a.min_estaciones)
    variedades = inv.loc[inv["incluida"], "variedad"].tolist()
    if not variedades:
        sys.exit("Ninguna variedad cumple los mínimos.")

    req, por_fase, por_est, busqueda, valid, tab_camp = [], [], [], [], [], []
    for var in variedades:
        cv = [c for c in camps if c["variedad"] == var and c["clima_completo"]]
        b, mejor = calibrar(cv, tbs, a.tu, a.min_fase)
        busqueda.append(b.assign(variedad=var))
        if mejor is None:
            req.append({"variedad": var, "campanias": len(cv), "observaciones":
                        f"ninguna fase de calibración con {a.min_fase}+ campañas"})
            continue
        metodo, tb = mejor
        tb_ini, tb_fin = intervalo_tb(b, metodo)
        res, err = evaluar(cv, tb, a.tu, metodo, FASES_TABLA, 2)
        por_fase.append(res.assign(variedad=var, metodo=metodo, tb=tb))
        valid += [{"variedad": var, **e} for e in err if e["fase"] in FASES_CALIB]
        res10, _ = evaluar(cv, 10.0, a.tu, "media", FASES_CALIB, a.min_fase)
        err_opt, err_cal = objetivo(res[res["campanias"] >= a.min_fase])
        err_10, _ = objetivo(res10)

        # GDD por campaña y por estación con la Tb elegida
        filas = []
        for c in cv:
            ac = acumulado(c, tb, a.tu, metodo)
            for fase in FASES_TABLA:
                if fase in c["dias"]:
                    filas.append({"variedad": var, "estacion": c["estacion"], "fecha_siembra": c["fecha_siembra"].date(),
                                  "tmed": c["tmed"], "pct_sin_dato": c["pct_sin_dato"], "fase": fase,
                                  "dias_desde_siembra": c["dias"][fase], "gdd_acumulado": round(ac[c["dias"][fase]], 1)})
        fc = pd.DataFrame(filas)
        tab_camp.append(fc)
        pe = (fc.groupby(["estacion", "fase"], sort=False)
                .agg(campanias=("gdd_acumulado", "size"), tmed=("tmed", "mean"),
                     dias_media=("dias_desde_siembra", "mean"), gdd_media=("gdd_acumulado", "mean"),
                     gdd_min=("gdd_acumulado", "min"), gdd_max=("gdd_acumulado", "max"),
                     gdd_cv_pct=("gdd_acumulado", lambda s: 100 * s.std() / s.mean() if len(s) > 1 else None))
                .round(1).reset_index())
        pe["orden"] = pe["fase"].map(FASES.index)
        por_est.append(pe.sort_values(["estacion", "orden"]).drop(columns="orden").assign(variedad=var, tb=tb))

        tmeds = [c["tmed"] for c in cv if c["tmed"] is not None]
        n_cornea = int(res.loc[res["fase"] == "maduracion cornea", "campanias"].sum())
        obs = []
        if max(tmeds) - min(tmeds) < RANGO_TMED_MIN:
            obs.append(f"rango de Tmedia {max(tmeds) - min(tmeds):.1f} °C: Tb poco identificable")
        if tb in (tbs[0], tbs[-1]):
            obs.append("Tb en el borde de la búsqueda")
        if tb_fin - tb_ini >= ANCHO_TB_MAX:
            obs.append(f"error casi igual entre Tb {tb_ini:g} y {tb_fin:g} °C: Tb poco definida")
        if err_opt is not None and err_opt >= err_cal:
            obs.append("los GDD no predicen mejor que los días calendario")
        if n_cornea < N_CONFIABLE:
            obs.append(f"{n_cornea} campañas en maduración córnea")
        fila = {"variedad": var, "campanias": len(cv), "estaciones": len({c["estacion"] for c in cv}),
                "tmed_min": round(min(tmeds), 1), "tmed_max": round(max(tmeds), 1),
                "metodo": metodo, "tb": tb, "tb_intervalo": f"{tb_ini:g}–{tb_fin:g}"}
        # Todas las fases, con columnas vacías si la variedad no la tiene (mismo orden en todas las filas)
        for fase in FASES_TABLA:
            r = res[res["fase"] == fase]
            r = r.iloc[0] if len(r) else pd.Series(dtype=object)
            fila |= {f"gdd_{fase}": r.get("gdd_media"), f"gdd_{fase}_min": r.get("gdd_min"),
                     f"gdd_{fase}_max": r.get("gdd_max"), f"cv_{fase}_pct": r.get("gdd_cv_pct"),
                     f"n_{fase}": r.get("campanias"), f"dias_{fase}": r.get("dias_media")}
        fila |= {"rmse_gdd_dias": round(err_opt, 1) if err_opt is not None else None,
                 "rmse_tb10_dias": round(err_10, 1) if err_10 is not None else None,
                 "rmse_calendario_dias": round(err_cal, 1) if err_cal is not None else None,
                 "mejora_vs_calendario_pct": (round(100 * (1 - err_opt / err_cal), 0)
                                              if err_opt is not None and err_cal else None),
                 "confiabilidad": ["alta", "media"][len(obs)] if len(obs) < 2 else "baja",
                 "observaciones": "; ".join(obs) or None}
        req.append(fila)

    req = pd.DataFrame(req)
    primero = lambda df, cols: df[cols + [c for c in df.columns if c not in cols]]
    hojas = {
        "requerimientos": req,
        "por_fase": primero(pd.concat(por_fase, ignore_index=True), ["variedad", "metodo", "tb"]),
        "por_estacion": primero(pd.concat(por_est, ignore_index=True), ["variedad", "tb"]),
        "busqueda_tb": primero(pd.concat(busqueda, ignore_index=True), ["variedad"]),
        "validacion": pd.DataFrame(valid),
        "campanias": pd.concat(tab_camp, ignore_index=True),
        "inventario": inv,
        "leeme": pd.DataFrame(LEEME + [("parámetros", f"Tb {a.tb_min:g} a {a.tb_max:g} cada {a.paso:g} °C; "
                                                      f"Tu {a.tu:g} °C; máx. {a.max_sin_dato:g} % días sin dato")],
                              columns=["hoja / concepto", "descripcion"]),
    }
    salida = Path(a.salida) if a.salida else Path(a.fenologia).with_name("gdd_variedades.xlsx")
    escribir(salida, hojas)

    ver = ["variedad", "campanias", "estaciones", "tmed_min", "tmed_max", "metodo", "tb", "tb_intervalo",
           "gdd_panoja", "gdd_maduracion cornea", "cv_maduracion cornea_pct",
           "rmse_gdd_dias", "rmse_tb10_dias", "rmse_calendario_dias", "confiabilidad"]
    pd.set_option("display.width", 250)
    print(req[[c for c in ver if c in req.columns]].to_string(index=False))
    for _, r in req[req["observaciones"].notna()].iterrows():
        print(f"  {r['variedad']}: {r['observaciones']}")
    print(f"\n{len(variedades)} variedades -> {salida}")


if __name__ == "__main__":
    main()
