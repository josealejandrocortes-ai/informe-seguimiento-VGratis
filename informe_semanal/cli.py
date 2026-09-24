"""Uso:  python -m informe_semanal.cli PLANTILLA.docx SOPORTE.xlsm -o INFORME.docx [--fecha 2026-09-18] [--sin-ia]"""
import argparse
import datetime as dt
import os
import sys

from .engine import (PROVEEDOR_DEFECTO, PROVEEDORES, Contexto, Libro, adivinar_fecha_corte, analizar,
                     construir, escanear)

GLOSARIO = {"Despacho del Ministro": "Despacho de la Ministra", "DESPACHO DEL MINISTRO": "DESPACHO DE LA MINISTRA"}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Diligencia el Informe Seguimiento Semanal desde el Excel de soporte")
    ap.add_argument("plantilla")
    ap.add_argument("excel")
    ap.add_argument("-o", "--salida", default="Informe_Seguimiento_Semanal_diligenciado.docx")
    ap.add_argument("--fecha", help="Fecha de corte AAAA-MM-DD (por defecto se deduce del Excel)")
    ap.add_argument("--sin-ia", action="store_true", help="No generar los análisis <…> (quedan los marcadores)")
    ap.add_argument("--proveedor", default=PROVEEDOR_DEFECTO, choices=list(PROVEEDORES))
    ap.add_argument("--modelo", default="", help="Por defecto, el del proveedor")
    ap.add_argument("--url", default="", help="URL base (solo proveedores compatibles con OpenAI)")
    a = ap.parse_args(argv)

    lib = Libro(a.excel)
    fecha = dt.date.fromisoformat(a.fecha) if a.fecha else adivinar_fecha_corte(lib)
    ctx = Contexto(lib, fecha, GLOSARIO)
    plantilla = open(a.plantilla, "rb").read()
    esc = escanear(plantilla, ctx)
    malos = [v for v in esc["valores"] if v["error"]]
    print(f"Fecha de corte: {fecha} | marcadores {{}}: {len(esc['valores'])} | análisis <>: {len(esc['analisis'])}")
    for v in malos:
        print("  ✗", v["marcador"], "→", v["error"])

    analisis = {}
    cfg = PROVEEDORES[a.proveedor]
    key = os.environ.get(cfg["env"], "") if cfg["env"] else ""
    if esc["analisis"] and not a.sin_ia:
        if cfg["requiere_key"] and not key:
            print(f"Sin {cfg['env']}: los análisis <…> quedan pendientes.")
        else:
            for it in esc["analisis"]:
                analisis[it["instruccion"]] = analizar(ctx, it["instruccion"], a.proveedor, key, a.modelo, a.url)
    datos, info = construir(plantilla, ctx, analisis)
    open(a.salida, "wb").write(datos)
    print(f"Listo: {a.salida} | reemplazos: {info['reemplazados']} | gráficos actualizados: {len(info['graficos'])}")
    for s in info["sin_resolver"]:
        print("  ✗ sin resolver:", s)
    for s in dict.fromkeys(info["analisis_pendientes"]):
        print("  … análisis pendiente:", s[:90])
    return 1 if malos or info["sin_resolver"] else 0


if __name__ == "__main__":
    sys.exit(main())
