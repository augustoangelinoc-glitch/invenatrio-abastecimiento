
import io
import re
import numpy as np
import pandas as pd


# ============================================================
# MAPA REAL DE LOS ARCHIVOS DEL PROYECTO
# ============================================================
# STOCK
# A Codigo
# B Descripción
# C U.Medida
# D Sistema              -> STOCK ACTUAL
# N Familia
# T Costo Cierre Mes     -> COSTO UNITARIO
#
# SALIDAS
# H F.Almac              -> FECHA DE SALIDA
# AX Unidades             -> CANTIDAD SALIDA
# Material                -> CÓDIGO/MATERIAL
#
# ORDENES DE COMPRA
# F. Docum.              -> FECHA DOCUMENTO
# Fecha Guia             -> FECHA RECEPCIÓN
# Material               -> CÓDIGO/MATERIAL
# Unidades               -> CANTIDAD COMPRADA
# Estado                 -> ESTADO DE LA OC
# ============================================================


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm(v):
    s = clean(v).upper()
    s = re.sub(r"\s+", " ", s)
    return s


def numeric(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)

    def cv(v):
        if pd.isna(v):
            return 0.0
        x = str(v).strip().replace("S/", "").replace(" ", "")
        if x == "":
            return 0.0
        if "," in x and "." in x:
            if x.rfind(",") > x.rfind("."):
                x = x.replace(".", "").replace(",", ".")
            else:
                x = x.replace(",", "")
        elif "," in x:
            x = x.replace(",", ".")
        try:
            return float(x)
        except Exception:
            return 0.0

    return s.map(cv)


def code_text(s):
    # Nunca convertir el código del material a número.
    return s.map(lambda x: "" if pd.isna(x) else str(x).strip())


def code_key(v):
    """
    Primera llave: texto exacto.
    Segunda llave: solo dígitos, para conectar casos donde Excel
    haya eliminado ceros iniciales en una de las fuentes.
    """
    s = clean(v)
    if not s:
        return ""
    s = re.sub(r"\.0+$", "", s)
    digits = re.sub(r"\D", "", s)
    return digits if digits else norm(s)


def find_col(df, names):
    by_norm = {norm(c): c for c in df.columns}
    for n in names:
        if norm(n) in by_norm:
            return by_norm[norm(n)]
    raise ValueError(
        f"No se encontró la columna requerida {names}. "
        f"Columnas disponibles: {list(df.columns)}"
    )


def read_excel(file):
    return pd.ExcelFile(file)


def choose_sheet(xls, words):
    for sheet in xls.sheet_names:
        n = norm(sheet)
        if any(w in n for w in words):
            return sheet
    return xls.sheet_names[0]


def load_stock(file):
    xls = read_excel(file)
    sheet = choose_sheet(xls, ["STOCK", "INVENTARIO", "SALDO"])
    d = pd.read_excel(file, sheet_name=sheet, dtype=object)

    c_codigo = find_col(d, ["Codigo", "Código"])
    c_desc = find_col(d, ["Descripción"])
    c_um = find_col(d, ["U.Medida"])
    c_stock = find_col(d, ["Sistema"])
    c_familia = find_col(d, ["Familia"])
    c_costo = find_col(d, ["Costo Cierre Mes"])

    out = pd.DataFrame({
        "codigo": code_text(d[c_codigo]),
        "descripcion": d[c_desc].map(clean),
        "unidad": d[c_um].map(clean),
        "stock_actual": numeric(d[c_stock]),
        "familia": d[c_familia].map(clean),
        "costo_unitario": numeric(d[c_costo]),
    })
    out["key"] = out["codigo"].map(code_key)
    out = out[out["codigo"] != ""].copy()

    # Si hubiera repetidos, conservar el último registro del stock.
    out = out.drop_duplicates("key", keep="last")
    return out.reset_index(drop=True), {
        "hoja": sheet,
        "Codigo": c_codigo,
        "Descripción": c_desc,
        "U.Medida": c_um,
        "Sistema": c_stock,
        "Familia": c_familia,
        "Costo Cierre Mes": c_costo,
    }


def load_salidas(file):
    xls = read_excel(file)
    sheet = choose_sheet(xls, ["SALIDA", "MOVIMIENTO", "KARDEX"])
    d = pd.read_excel(file, sheet_name=sheet, dtype=object)

    c_material = find_col(d, ["Material", "Cod. Origen", "Cod.Origen", "Codigo", "Código"])
    c_fecha = find_col(d, ["F.Almac", "Fecha Almacén", "Fecha Almacen", "F.Contab", "F.Docum."])
    c_qty = find_col(d, ["Unidades", "Cantidad", "Cantidad salida", "C. Kardex"])

    out = pd.DataFrame({
        "codigo": code_text(d[c_material]),
        "fecha": pd.to_datetime(d[c_fecha], errors="coerce", dayfirst=True),
        "cantidad": numeric(d[c_qty]).abs(),
    })
    out["key"] = out["codigo"].map(code_key)
    out = out[(out["key"] != "") & out["fecha"].notna()].copy()

    return out.reset_index(drop=True), {
        "hoja": sheet,
        "Material": c_material,
        "F.Almac": c_fecha,
        "Unidades": c_qty,
    }


def load_oc(file):
    xls = read_excel(file)
    sheet = choose_sheet(xls, ["ORDEN", "COMPRA", "OC"])
    d = pd.read_excel(file, sheet_name=sheet, dtype=object)

    c_material = find_col(d, ["Material", "Cod. Origen", "Cod.Origen", "Codigo", "Código"])
    c_doc = find_col(d, ["F. Docum.", "F.Docum.", "Fecha Documento", "F Docum"])
    c_guia = find_col(d, ["Fecha Guia", "Fecha Guía", "Fecha Guia ", "Fecha Recepcion", "Fecha Recepción"])
    c_estado = find_col(d, ["Estado", "Estado OC", "Situación"])
    c_qty = find_col(d, ["Unidades", "Cantidad", "Cantidad comprada"])

    out = pd.DataFrame({
        "codigo": code_text(d[c_material]),
        "f_docum": pd.to_datetime(d[c_doc], errors="coerce", dayfirst=True),
        "fecha_guia": pd.to_datetime(d[c_guia], errors="coerce", dayfirst=True),
        "estado": d[c_estado].map(norm),
        "cantidad": numeric(d[c_qty]).abs(),
    })
    out["key"] = out["codigo"].map(code_key)
    out["lead_time"] = (out["fecha_guia"] - out["f_docum"]).dt.days
    out.loc[(out["lead_time"] < 0) | (out["lead_time"] > 365), "lead_time"] = np.nan
    out = out[out["key"] != ""].copy()

    return out.reset_index(drop=True), {
        "hoja": sheet,
        "Material": c_material,
        "F. Docum.": c_doc,
        "Fecha Guia": c_guia,
        "Estado": c_estado,
        "Unidades": c_qty,
    }


def trend(values):
    y = np.asarray(values, dtype=float)
    if len(y) < 2 or np.allclose(y, y[0]):
        return "Estable"
    x = np.arange(len(y), dtype=float)
    slope = np.polyfit(x, y, 1)[0]
    base = max(float(np.mean(np.abs(y))), 1.0)
    if slope > 0.03 * base:
        return "Creciente"
    if slope < -0.03 * base:
        return "Decreciente"
    return "Estable"


def xyz(cv):
    if pd.isna(cv):
        return "Z"
    if cv <= 0.20:
        return "X"
    if cv <= 0.50:
        return "Y"
    return "Z"


def calculate(stock, sal, oc, horizon_months=3, z=1.65, as_of=None):
    # Fecha de análisis = última fecha real de salida, salvo que el usuario
    # quiera forzar una fecha posterior mediante as_of.
    if as_of is None:
        candidates = []
        if not sal.empty:
            candidates.append(sal["fecha"].max())
        if not oc.empty:
            candidates.append(oc["f_docum"].max())
        as_of = max(candidates) if candidates else pd.Timestamp.today()
    as_of = pd.Timestamp(as_of).normalize()

    # Meses mostrados: exactamente los meses que existen entre la primera
    # salida y el mes de análisis. Esto evita inventar meses.
    if sal.empty:
        periods = pd.period_range(as_of.to_period("M"), as_of.to_period("M"), freq="M")
    else:
        first = sal["fecha"].min().to_period("M")
        periods = pd.period_range(first, as_of.to_period("M"), freq="M")

    sal_by_key = {k: g for k, g in sal.groupby("key", sort=False)}
    oc_by_key = {k: g for k, g in oc.groupby("key", sort=False)}

    rows = []

    for _, r in stock.iterrows():
        key = r["key"]
        s = sal_by_key.get(key, pd.DataFrame(columns=sal.columns))
        o = oc_by_key.get(key, pd.DataFrame(columns=oc.columns))

        if not s.empty:
            s = s.copy()
            s["period"] = s["fecha"].dt.to_period("M")
            monthly = s.groupby("period")["cantidad"].sum()
        else:
            monthly = pd.Series(dtype=float)

        vals = np.array([float(monthly.get(p, 0.0)) for p in periods])
        total = float(vals.sum())
        months_with = int((vals > 0).sum())
        months_without = int(len(vals) - months_with)

        # REGLA IMPORTANTE:
        # promedio = consumo total / meses con consumo.
        # No se divide entre meses sin consumo.
        avg = total / months_with if months_with else 0.0
        daily = avg / 30.0 if avg > 0 else 0.0

        last_out = s["fecha"].max() if not s.empty else pd.NaT
        days_no_move = (
            int((as_of - last_out.normalize()).days)
            if pd.notna(last_out) else np.nan
        )

        coverage = (
            float(r["stock_actual"]) / daily
            if daily > 0 else np.inf
        )

        # Solo OC COMPRADO participa en última compra y Lead Time.
        comprado = o[o["estado"].str.contains("COMPRADO", na=False)].copy()

        if not comprado.empty:
            comprado = comprado.sort_values(["f_docum", "fecha_guia"])
            valid_lt = comprado["lead_time"].dropna()
            lt = float(valid_lt.median()) if not valid_lt.empty else np.nan

            last = comprado.iloc[-1]
            last_qty = float(last["cantidad"])
            last_buy_date = last["f_docum"]

            days_since_buy = (
                int((as_of - last_buy_date.normalize()).days)
                if pd.notna(last_buy_date) else np.nan
            )
            last_buy_coverage = (
                last_qty / avg if avg > 0 else np.nan
            )
        else:
            lt = np.nan
            last_qty = np.nan
            last_buy_date = pd.NaT
            days_since_buy = np.nan
            last_buy_coverage = np.nan

        # Variabilidad mensual.
        mean_all = float(vals.mean()) if len(vals) else 0.0
        std_all = float(vals.std(ddof=1)) if len(vals) > 1 else 0.0
        cv = std_all / mean_all if mean_all > 0 else np.nan

        if months_with == 0:
            tipo = "Sin consumo"
        elif months_with == 1:
            tipo = "Ocasional"
        elif months_with / len(vals) < 0.50:
            tipo = "Intermitente"
        else:
            tipo = "Frecuente"

        tr = trend(vals)
        variab = (
            "Baja" if pd.isna(cv) or cv <= 0.20
            else ("Media" if cv <= 0.50 else "Alta")
        )

        # Anomalía: regla conservadora; no marcar una sola salida como anomalía.
        anomaly = "No"
        if len(vals) >= 4 and std_all > 0:
            if np.any(np.abs(vals - mean_all) > 3 * std_all):
                anomaly = "Sí"

        # Seguridad:
        # demanda diaria * LT + componente estadístico de variabilidad.
        # Si LT no está disponible, NO inventamos un LT.
        if daily > 0 and pd.notna(lt):
            safety = z * daily * np.sqrt(max(lt, 0.0))
            reorder = daily * lt + safety
        else:
            safety = 0.0
            reorder = 0.0

        # Objetivo: cubrir el horizonte solicitado + seguridad.
        target = avg * max(float(horizon_months), 1.0) + safety
        replenish = max(target - float(r["stock_actual"]), 0.0)

        # Escenarios independientes para que el usuario vea cuánto compra
        # y cuántos días tendría en total después de comprar.
        def scenario(months):
            q = avg * months
            duration = (
                (float(r["stock_actual"]) + q) / daily
                if daily > 0 else np.inf
            )
            return q, duration

        q1, d1 = scenario(1)
        q2, d2 = scenario(2)
        q3, d3 = scenario(3)

        if total == 0:
            situation = "SIN CONSUMO"
            rupture = "FALSO"
        elif float(r["stock_actual"]) <= 0:
            situation = "ROTURA DE STOCK"
            rupture = "VERDADERO"
        elif pd.notna(lt) and coverage < lt:
            situation = "RIESGO DE ROTURA"
            rupture = "FALSO"
        elif daily > 0 and coverage < 30:
            situation = "STOCK BAJO"
            rupture = "FALSO"
        else:
            situation = "STOCK SUFICIENTE"
            rupture = "FALSO"

        if total == 0:
            when = "NO COMPRAR"
        elif float(r["stock_actual"]) <= reorder:
            when = "AHORA"
        else:
            when = "PROGRAMAR"

        priority = (
            "ALTO" if situation in ("ROTURA DE STOCK", "RIESGO DE ROTURA")
            else ("MEDIO" if situation == "STOCK BAJO" else "REVISAR")
        )

        if replenish > 0:
            recommendation = f"Abastecer {replenish:.2f} {r['unidad']}."
        else:
            recommendation = "No requiere abastecimiento inmediato."

        diagnosis = (
            f"{situation}. {tipo}. "
            f"{months_with} meses con consumo de {len(vals)}."
        )

        row = {
            "Código": r["codigo"],
            "Descripción": r["descripcion"],
            "Unidad de medida": r["unidad"],
            "Familia": r["familia"],
            "Stock actual": float(r["stock_actual"]),
            "Costo unitario (S/)": float(r["costo_unitario"]),
            "Valor del inventario (S/)": float(r["stock_actual"] * r["costo_unitario"]),
            "Consumo total": total,
            "Valor de salidas (S/)": total * float(r["costo_unitario"]),
            "Última salida": last_out,
        }

        for p, v in zip(periods, vals):
            row[str(p)] = float(v)

        row.update({
            "Meses con consumo": months_with,
            "Meses sin consumo": months_without,
            "Consumo mensual promedio": avg,
            "Consumo diario": daily,
            "Días sin movimiento": days_no_move,
            "Cobertura (días)": coverage,
            "Lead Time (días)": lt,
            "Tipo de consumo": tipo,
            "Situación de stock": situation,
            "Rotura de stock": rupture,
            "Tendencia del consumo": tr,
            "Coeficiente de variación": cv,
            "Nivel de variabilidad": variab,
            "Anomalía de consumo": anomaly,
            "Stock de seguridad": safety,
            "Punto de pedido": reorder,
            "Stock objetivo": target,
            "Cantidad a abastecer": replenish,
            "Última compra": last_qty,
            "Fecha última compra": last_buy_date,
            "Días desde última compra": days_since_buy,
            "Cobertura última compra (meses)": last_buy_coverage,
            "Compra para 1 mes": q1,
            "Duración después de comprar 1 mes (días)": d1,
            "Compra para 2 meses": q2,
            "Duración después de comprar 2 meses (días)": d2,
            "Compra para 3 meses": q3,
            "Duración después de comprar 3 meses (días)": d3,
            "Cuándo comprar": when,
            "Metodología utilizada": "Promedio de meses con consumo + variabilidad + Lead Time de OC COMPRADO",
            "Diagnóstico": diagnosis,
            "Prioridad": priority,
            "Recomendación": recommendation,
        })
        rows.append(row)

    df = pd.DataFrame(rows)

    # ABC: valor económico de consumo.
    economic = df["Consumo total"] * df["Costo unitario (S/)"]
    total_economic = float(economic.sum())
    if total_economic > 0:
        order = economic.sort_values(ascending=False).index
        cum = economic.loc[order].cumsum() / total_economic * 100
        abc = pd.Series(index=df.index, dtype=object)
        for i in order:
            p = cum.loc[i]
            abc.loc[i] = "A" if p <= 80 else ("B" if p <= 95 else "C")
        df["ABC acumulado (%)"] = cum.reindex(df.index)
        df["Clasificación ABC"] = abc
    else:
        df["ABC acumulado (%)"] = 100.0
        df["Clasificación ABC"] = "C"

    df["Clasificación XYZ"] = df["Coeficiente de variación"].map(xyz)

    fixed = [
        "Código","Descripción","Unidad de medida","Familia","Stock actual",
        "Costo unitario (S/)","Valor del inventario (S/)","Consumo total",
        "Valor de salidas (S/)","Última salida"
    ]
    months = [str(p) for p in periods]
    rest = [
        "Meses con consumo","Meses sin consumo","Consumo mensual promedio",
        "Consumo diario","Días sin movimiento","Cobertura (días)",
        "Lead Time (días)","Tipo de consumo","Situación de stock",
        "Rotura de stock","Tendencia del consumo","Coeficiente de variación",
        "Nivel de variabilidad","Anomalía de consumo","ABC acumulado (%)",
        "Clasificación ABC","Clasificación XYZ","Stock de seguridad",
        "Punto de pedido","Stock objetivo","Cantidad a abastecer",
        "Última compra","Fecha última compra","Días desde última compra",
        "Cobertura última compra (meses)","Compra para 1 mes",
        "Duración después de comprar 1 mes (días)","Compra para 2 meses",
        "Duración después de comprar 2 meses (días)","Compra para 3 meses",
        "Duración después de comprar 3 meses (días)","Cuándo comprar",
        "Metodología utilizada","Diagnóstico","Prioridad","Recomendación"
    ]
    df = df[fixed + months + rest]
    return df, periods


def excel_bytes(df):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analisis")
        ws = writer.sheets["Analisis"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        headers = {c.value: c.column for c in ws[1]}
        if "Código" in headers:
            col = headers["Código"]
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col).number_format = "@"

        for column in ws.columns:
            letter = column[0].column_letter
            max_len = max(
                len(str(c.value)) if c.value is not None else 0
                for c in column
            )
            ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 42)

    bio.seek(0)
    return bio.getvalue()
