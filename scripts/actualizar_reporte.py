#!/usr/bin/env python3
"""
Reescritura en Python de Modulo5.ActualizarAhora + ConstruirHtmlReporte
(VBA, Parking Switch 345.xlsm). Corre en GitHub Actions, sin depender de
Excel ni de ninguna maquina encendida.

Que hace, en orden:
  1. Login a ParkingApp con las credenciales de PARKINGAPP_EMAIL/PARKINGAPP_PASSWORD.
  2. Descarga la ventana de reconciliacion (ultimos RECONCILE_DAYS dias) +
     cualquier pago mas nuevo que el ultimo guardado.
  3. Reconstruye esa ventana en data/pagos.json (reemplaza, no solo agrega -
     atrapa reembolsos/ajustes retroactivos de ParkingApp) y agrega lo nuevo.
  4. Recalcula el ~20 valores dinamicos del reporte (agregacion mensual/diaria/
     horaria/heatmaps/ocupacion) sobre TODO data/pagos.json.
  5. Inserta esos valores en scripts/template.html -> index.html.
  6. git commit + push (usa el token que ya tenga configurado git en el runner).

Formulas y edge cases replicados 1:1 desde Modulo5_2026-09-10.bas
(ActualizarAhora, ConstruirHtmlReporte, TarifaMinuto, TopeDiario,
MarcarOcupacion) - ver spec en el historial de esta conversion.
"""
from __future__ import annotations

import json
import os
import re
import sys
from calendar import monthrange
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import requests

# ---------------------------------------------------------------------------
# Configuracion
# ---------------------------------------------------------------------------

TZ = ZoneInfo("America/Santiago")
RECONCILE_DAYS = 5
BASE_URL = "https://admin.parkingapp.cl"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/151.0.0.0",
    "Accept": "application/json",
    "X-Requested-With": "XMLHttpRequest",
}

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(REPO_ROOT, "data", "pagos.json")
TEMPLATE_PATH = os.path.join(REPO_ROOT, "scripts", "template.html")
OUTPUT_PATH = os.path.join(REPO_ROOT, "index.html")

MONTH_NAMES = [
    None, "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

BUCKET_LABELS = [
    None,
    "menos de $1.000",
    "$1.000 - $2.000",
    "$2.000 - $3.000",
    "$3.000 - $5.000",
    "$5.000 - $10.000",
    "$10.000 - $15.000",
    "$15.000 o mas (tope)",
]

WEEKDAY_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


# ---------------------------------------------------------------------------
# Reglas de negocio (identicas a TarifaMinuto/TopeDiario en VBA)
# ---------------------------------------------------------------------------

def tarifa_minuto(anio: int, mes: int) -> float:
    if anio > 2026:
        return 25.0
    if anio == 2026 and mes >= 6:
        return 25.0
    return 21.0


def tope_diario(anio: int) -> float:
    return 10000.0 if anio <= 2025 else 15000.0


def bucket_index(monto: float) -> int:
    if monto < 1000:
        return 1
    if monto < 2000:
        return 2
    if monto < 3000:
        return 3
    if monto < 5000:
        return 4
    if monto < 10000:
        return 5
    if monto < 15000:
        return 6
    return 7


def weekday_es(d: date) -> str:
    return WEEKDAY_ES[d.weekday()]  # Monday=0 .. Sunday=6, igual orden que WEEKDAY_ES


# ---------------------------------------------------------------------------
# ParkingApp: login + descarga paginada (equivalente a DoLogin + el loop de
# paginacion en ActualizarAhora)
# ---------------------------------------------------------------------------

def login(session: requests.Session, email: str, password: str) -> bool:
    r = session.get(f"{BASE_URL}/ingresar", headers=HEADERS, timeout=30)
    if r.status_code != 200:
        return False
    m = re.search(r'name="authenticity_token"\s+value="([^"]+)"', r.text)
    if not m:
        return False
    token = m.group(1)

    payload = {
        "utf8": "✓",
        "authenticity_token": token,
        "user_session[email]": email,
        "user_session[password]": password,
    }
    r2 = session.post(
        f"{BASE_URL}/sesiones",
        data=payload,
        headers={**HEADERS, "Referer": f"{BASE_URL}/ingresar"},
        timeout=30,
    )
    body = r2.text
    # Mismo criterio que el VBA: si la respuesta vuelve a traer el formulario
    # de sesion, el login fallo.
    return not ("user_session" in body and "authenticity_token" in body)


def clean_money(s) -> float:
    if s is None:
        return 0.0
    t = str(s).replace("$", "").replace(".", "").strip()
    if t == "":
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def fetch_range(session: requests.Session, start: date, end: date) -> list[dict]:
    records: list[dict] = []
    page = 1
    while True:
        params = {
            "page": page,
            "branch_id": "",
            "start_date": start.strftime("%d-%m-%Y"),
            "end_date": end.strftime("%d-%m-%Y"),
            "date_filter_type": "paid_at",
        }
        r = session.get(f"{BASE_URL}/pagos.json", params=params, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"ParkingApp respondio {r.status_code} en la pagina {page}")
        data = r.json()
        records.extend(data.get("collection") or [])
        if data.get("next_page") is None:
            break
        page += 1
        if page > 60:
            raise RuntimeError("Demasiadas paginas al descargar de ParkingApp (posible loop infinito)")
    return records


def parse_record(raw: dict) -> dict:
    """API -> {id, dt (datetime naive, hora local Chile), monto}."""
    date_at = raw.get("date_at", "")  # "dd-mm-yyyy"
    time_at = raw.get("time_at", "")  # "H:MM" o "HH:MM"
    dd, mm, yyyy = (int(p) for p in date_at.split("-"))
    hh, mi = (int(p) for p in time_at.split(":"))
    dt = datetime(yyyy, mm, dd, hh, mi)
    monto = clean_money(raw.get("amount")) + clean_money(raw.get("rounding"))
    return {"id": int(raw["id"]), "dt": dt, "monto": monto}


# ---------------------------------------------------------------------------
# Dataset persistido (data/pagos.json)
# ---------------------------------------------------------------------------

def load_dataset() -> list[dict]:
    with open(DATA_PATH, encoding="utf-8") as f:
        raw = json.load(f)
    return [{"id": r["id"], "dt": datetime.fromisoformat(r["dt"]), "monto": r["monto"]} for r in raw]


def save_dataset(records: list[dict]) -> None:
    records = sorted(records, key=lambda r: r["dt"])
    out = [{"id": r["id"], "dt": r["dt"].isoformat(timespec="minutes"), "monto": r["monto"]} for r in records]
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))


def reconciliar(session: requests.Session, dataset: list[dict], hoy: date) -> tuple[list[dict], bool]:
    """Replica la reconciliacion de ActualizarAhora. Devuelve (nuevo_dataset, hubo_cambios)."""
    if not dataset:
        raise RuntimeError("data/pagos.json esta vacio - no deberia pasar tras la migracion inicial")

    last_dt = max(r["dt"] for r in dataset)
    last_date = last_dt.date()
    reconcile_from = last_date - timedelta(days=RECONCILE_DAYS - 1)

    fetched_raw = fetch_range(session, reconcile_from, hoy)
    fetched = [parse_record(r) for r in fetched_raw]

    existing_ids = {r["id"] for r in dataset}

    window_block = [r for r in fetched if r["dt"].date() <= last_date]
    future_new = [r for r in fetched if r["dt"].date() > last_date and r["id"] not in existing_ids]

    if not window_block and not future_new:
        return dataset, False

    had_data_in_window = any(reconcile_from <= r["dt"].date() <= last_date for r in dataset)
    if not window_block and had_data_in_window:
        raise RuntimeError(
            "ParkingApp no devolvio transacciones para la ventana de reconciliacion "
            "pero el dataset SI tiene filas ahi - abortando por seguridad (no se borra nada)."
        )

    kept = [r for r in dataset if r["dt"].date() < reconcile_from]
    nuevo = kept + window_block + future_new
    return nuevo, True


# ---------------------------------------------------------------------------
# Agregacion (equivalente al loop principal de ConstruirHtmlReporte)
# ---------------------------------------------------------------------------

def marcar_ocupacion(dt_salida: datetime, duracion_min: float,
                     occupancy_hourly: list[int], occupancy_heatmap: list[list[int]],
                     daily_occup_hour: dict[str, int]) -> None:
    entrada = dt_salida - timedelta(minutes=duracion_min)
    cur = entrada
    safety = 0
    while True:
        hour_start = cur.replace(minute=0, second=0, microsecond=0)
        hh = hour_start.hour
        occupancy_hourly[hh] += 1
        wd = hour_start.weekday()  # 0=lunes .. 6=domingo, igual indexado que heatmapData
        occupancy_heatmap[wd][hh] += 1
        key = f"{hour_start.date().isoformat()}|{hh}"
        daily_occup_hour[key] = daily_occup_hour.get(key, 0) + 1
        cur = hour_start + timedelta(hours=1)
        safety += 1
        if not (cur <= dt_salida and safety < 48):
            break


def calcular_agregados(dataset: list[dict], hoy: date) -> dict:
    current_year = hoy.year
    doy_hoy = hoy.timetuple().tm_yday

    mon_cobrado = [0.0] * 13
    mon_trans = [0] * 13
    sum_estadia_mes = [0.0] * 13
    count_estadia_mes = [0] * 13

    daily_cobrado: dict[str, float] = {}
    daily_trans: dict[str, int] = {}
    daily_hour: dict[str, int] = {}
    daily_bucket: dict[str, int] = {}
    daily_occup_hour: dict[str, int] = {}

    hour_count = [0] * 24
    heatmap = [[0] * 24 for _ in range(7)]
    occupancy_hourly = [0] * 24
    occupancy_heatmap = [[0] * 24 for _ in range(7)]
    bucket_counts = [0] * 8

    sum_estadia = 0.0
    count_estadia = 0
    dia_completo_count = 0
    dia_completo_ocup_estim = 0

    total_cobrado_anio_anterior = 0.0
    total_trans_anio_anterior = 0

    for r in dataset:
        dt = r["dt"]
        monto = r["monto"]
        d = dt.date()

        if d.year == current_year - 1:
            if d.timetuple().tm_yday <= doy_hoy:
                total_cobrado_anio_anterior += monto
                total_trans_anio_anterior += 1

        if d.year != current_year:
            continue

        mes = d.month
        mon_cobrado[mes] += monto
        mon_trans[mes] += 1

        key_dia = d.isoformat()
        daily_cobrado[key_dia] = daily_cobrado.get(key_dia, 0.0) + monto
        daily_trans[key_dia] = daily_trans.get(key_dia, 0) + 1

        hora_val = dt.hour
        hour_count[hora_val] += 1
        wd = d.weekday()
        heatmap[wd][hora_val] += 1
        hkey = f"{key_dia}|{hora_val}"
        daily_hour[hkey] = daily_hour.get(hkey, 0) + 1

        b_idx = bucket_index(monto)
        bucket_counts[b_idx] += 1
        bkey = f"{key_dia}|{b_idx}"
        daily_bucket[bkey] = daily_bucket.get(bkey, 0) + 1

        tope = tope_diario(d.year)
        tarifa_min = tarifa_minuto(d.year, mes)

        if monto >= tope:
            dia_completo_count += 1
            if tarifa_min > 0:
                dia_completo_ocup_estim += 1
                marcar_ocupacion(dt, tope / tarifa_min, occupancy_hourly, occupancy_heatmap, daily_occup_hour)
        else:
            if tarifa_min > 0 and monto > 0:
                duracion_min = monto / tarifa_min
                sum_estadia += duracion_min
                count_estadia += 1
                sum_estadia_mes[mes] += duracion_min
                count_estadia_mes[mes] += 1
                marcar_ocupacion(dt, duracion_min, occupancy_hourly, occupancy_heatmap, daily_occup_hour)

    # --- monthly ---
    monthly = []
    prev_tdia = 0.0
    prev_mt = 0.0
    for mi in range(1, 13):
        if mi < hoy.month:
            dias_elapsed = monthrange(current_year, mi)[1]
        elif mi == hoy.month:
            dias_elapsed = hoy.day
        else:
            dias_elapsed = 0

        tdia = (mon_trans[mi] / dias_elapsed) if dias_elapsed > 0 else 0.0
        mt = (mon_cobrado[mi] / mon_trans[mi]) if mon_trans[mi] > 0 else 0.0
        delta_tdia = (tdia / prev_tdia - 1) if prev_tdia > 0 else 0.0
        delta_mt = (mt / prev_mt - 1) if prev_mt > 0 else 0.0
        estadia_prom_mes = (sum_estadia_mes[mi] / count_estadia_mes[mi]) if count_estadia_mes[mi] > 0 else 0.0
        cambio_tarifa = mi > 1 and tarifa_minuto(current_year, mi) != tarifa_minuto(current_year, mi - 1)

        monthly.append({
            "mes": mi,
            "nombre": MONTH_NAMES[mi],
            "cobrado": _num(mon_cobrado[mi]),
            "transacciones": mon_trans[mi],
            "tdia": round(tdia, 2),
            "montoTransac": round(mt, 2),
            "estadiaProm": round(estadia_prom_mes, 1),
            "deltaTdia": round(delta_tdia, 4),
            "deltaMontoTransac": round(delta_mt, 4),
            "cambioTarifa": cambio_tarifa,
        })
        prev_tdia = tdia
        prev_mt = mt

    # --- daily (orden DESCENDENTE, igual que el insertion sort del VBA) ---
    daily = []
    for key_dia in sorted(daily_cobrado.keys(), reverse=True):
        yyyy, mm, dd = (int(p) for p in key_dia.split("-"))
        cob = daily_cobrado[key_dia]
        trn = daily_trans[key_dia]
        mt_d = (cob / trn) if trn > 0 else 0.0
        horas = [daily_hour.get(f"{key_dia}|{h}", 0) for h in range(24)]
        horas_ocup = [daily_occup_hour.get(f"{key_dia}|{h}", 0) for h in range(24)]
        buckets_dia = [daily_bucket.get(f"{key_dia}|{b}", 0) for b in range(1, 8)]
        daily.append({
            "fecha": f"{dd:02d}/{mm:02d}/{yyyy}",
            "mes": mm,
            "monto": _num(cob),
            "transacciones": trn,
            "montoTransac": round(mt_d, 2),
            "dia": weekday_es(date(yyyy, mm, dd)),
            "horas": horas,
            "horasOcup": horas_ocup,
            "buckets": buckets_dia,
        })

    buckets = [
        {"rango": BUCKET_LABELS[b], "cantidad": bucket_counts[b]}
        for b in range(1, 8)
    ]

    estadia_prom = (sum_estadia / count_estadia) if count_estadia > 0 else 0.0
    total_trans_anio = sum(mon_trans[1:13])

    dias_transcurridos_mes_actual = hoy.day
    dias_totales_mes_actual = monthrange(current_year, hoy.month)[1]
    cobrado_mes_actual = mon_cobrado[hoy.month]
    proyeccion_mes_actual = (
        cobrado_mes_actual / dias_transcurridos_mes_actual * dias_totales_mes_actual
        if dias_transcurridos_mes_actual > 0 else 0.0
    )
    tarifa_min_actual = tarifa_minuto(current_year, hoy.month)

    return {
        "monthly": monthly,
        "daily": daily,
        "hourly": hour_count,
        "heatmapData": heatmap,
        "occupancyHourlyData": occupancy_hourly,
        "occupancyHeatmapData": occupancy_heatmap,
        "buckets": buckets,
        "estadiaProm": round(estadia_prom, 2),
        "diaCompletoCount": dia_completo_count,
        "diaCompletoOcupEstim": dia_completo_ocup_estim,
        "totalTransAnio": total_trans_anio,
        "cobradoMesActual": _num(cobrado_mes_actual),
        "proyeccionMesActual": round(proyeccion_mes_actual, 2),
        "diasTranscurridosMesActual": dias_transcurridos_mes_actual,
        "diasTotalesMesActual": dias_totales_mes_actual,
        "nombreMesActual": MONTH_NAMES[hoy.month],
        "currentYear": current_year,
        "totalCobradoAnioAnterior": round(total_cobrado_anio_anterior, 2),
        "totalTransAnioAnterior": total_trans_anio_anterior,
        "tarifaMinActual": tarifa_min_actual,
    }


def _num(x: float):
    """Como CStr de un Double en VBA: entero sin '.0' si no tiene decimales."""
    if float(x).is_integer():
        return int(x)
    return round(x, 2)


# ---------------------------------------------------------------------------
# Render: bloque de 20 lineas JS (equivalente a las lineas 1379-1398 del VBA)
# ---------------------------------------------------------------------------

def _fmt_monthly(m: dict) -> str:
    return (
        '{"mes":' + str(m["mes"]) +
        ',"nombre":"' + m["nombre"] + '"'
        ',"cobrado":' + json.dumps(m["cobrado"]) +
        ',"transacciones":' + str(m["transacciones"]) +
        ',"tdia":' + f'{m["tdia"]:.2f}' +
        ',"montoTransac":' + f'{m["montoTransac"]:.2f}' +
        ',"estadiaProm":' + f'{m["estadiaProm"]:.1f}' +
        ',"deltaTdia":' + f'{m["deltaTdia"]:.4f}' +
        ',"deltaMontoTransac":' + f'{m["deltaMontoTransac"]:.4f}' +
        ',"cambioTarifa":' + ("true" if m["cambioTarifa"] else "false") + "}"
    )


def _fmt_daily(d: dict) -> str:
    j = lambda obj: json.dumps(obj, separators=(",", ":"))
    return (
        '{"fecha":"' + d["fecha"] + '"'
        ',"mes":' + str(d["mes"]) +
        ',"monto":' + json.dumps(d["monto"]) +
        ',"transacciones":' + str(d["transacciones"]) +
        ',"montoTransac":' + f'{d["montoTransac"]:.2f}' +
        ',"dia":"' + d["dia"] + '"'
        ',"horas":' + j(d["horas"]) +
        ',"horasOcup":' + j(d["horasOcup"]) +
        ',"buckets":' + j(d["buckets"]) + "}"
    )


def render_js_block(v: dict) -> str:
    j = lambda obj: json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    monthly_json = "[" + ",".join(_fmt_monthly(m) for m in v["monthly"]) + "]"
    daily_json = "[" + ",".join(_fmt_daily(d) for d in v["daily"]) + "]"
    lines = [
        f"const monthly = {monthly_json};",
        f"const daily = {daily_json};",
        f"const hourly = {j(v['hourly'])};",
        f"const heatmapData = {j(v['heatmapData'])};",
        f"const occupancyHourlyData = {j(v['occupancyHourlyData'])};",
        f"const occupancyHeatmapData = {j(v['occupancyHeatmapData'])};",
        f"const buckets = {j(v['buckets'])};",
        f"const estadiaProm = {v['estadiaProm']:.2f};",
        f"const diaCompletoCount = {v['diaCompletoCount']};",
        f"const diaCompletoOcupEstim = {v['diaCompletoOcupEstim']};",
        f"const totalTransAnio = {v['totalTransAnio']};",
        f"const cobradoMesActual = {v['cobradoMesActual']};",
        f"const proyeccionMesActual = {v['proyeccionMesActual']:.2f};",
        f"const diasTranscurridosMesActual = {v['diasTranscurridosMesActual']};",
        f"const diasTotalesMesActual = {v['diasTotalesMesActual']};",
        f"const nombreMesActual = '{v['nombreMesActual']}';",
        f"const currentYear = {v['currentYear']};",
        f"const totalCobradoAnioAnterior = {v['totalCobradoAnioAnterior']:.2f};",
        f"const totalTransAnioAnterior = {v['totalTransAnioAnterior']};",
        f"const tarifaMinActual = {v['tarifaMinActual']:g};",
    ]
    return "\n".join(lines)


def render_report(dataset: list[dict], hoy: date) -> str:
    valores = calcular_agregados(dataset, hoy)
    bloque = render_js_block(valores)
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        template = f.read()
    if "{{DATOS_JS}}" not in template:
        raise RuntimeError("template.html no tiene el marcador {{DATOS_JS}} - revisar extraccion")
    return template.replace("{{DATOS_JS}}", bloque)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> int:
    email = os.environ.get("PARKINGAPP_EMAIL")
    password = os.environ.get("PARKINGAPP_PASSWORD")
    if not email or not password:
        print("Faltan PARKINGAPP_EMAIL / PARKINGAPP_PASSWORD en el entorno.", file=sys.stderr)
        return 1

    hoy = datetime.now(TZ).date()

    session = requests.Session()
    if not login(session, email, password):
        print("No se pudo iniciar sesion en ParkingApp.", file=sys.stderr)
        return 1

    dataset = load_dataset()
    nuevo_dataset, hubo_cambios = reconciliar(session, dataset, hoy)

    if not hubo_cambios:
        print("Nada nuevo que actualizar.")
        return 0

    save_dataset(nuevo_dataset)

    html = render_report(nuevo_dataset, hoy)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Reporte actualizado: {len(nuevo_dataset)} registros, {len(nuevo_dataset) - len(dataset)} nuevos/reconciliados.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
