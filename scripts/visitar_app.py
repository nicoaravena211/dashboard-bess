"""Abre la app de Streamlit con un navegador automático.

- Despierta la app si estaba dormida (botón "Yes, get this app back up").
- A las 8:00 y 12:00 (hora de Chile) la abre con ?actualizar=<token>: la app vacía su caché y vuelve a consultar las API.
- En las demás ejecuciones solo la visita, para que no pase 12 horas sin tráfico.

Variables de entorno: APP_URL (obligatoria), TOKEN_ACTUALIZACION (la misma de los Secrets de Streamlit), FORZAR ("true" para actualizar ahora).
"""
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urlencode, urlparse, urlunparse, parse_qsl
from zoneinfo import ZoneInfo

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

HORAS_ACTUALIZACION = (8, 12)  # hora local de Chile
TITULO = "Dashboard de operación"


def con_parametro(url, clave, valor):
    p = urlparse(url)
    q = dict(parse_qsl(p.query))
    q[clave] = valor
    return urlunparse(p._replace(query=urlencode(q)))


def main():
    url = os.environ["APP_URL"].strip()
    token = os.environ.get("TOKEN_ACTUALIZACION", "").strip()
    ahora = datetime.now(ZoneInfo("America/Santiago"))
    forzar = os.environ.get("FORZAR", "").lower() == "true"
    actualizar = bool(token) and (forzar or ahora.hour in HORAS_ACTUALIZACION)
    print(f"Hora de Chile: {ahora:%d-%m-%Y %H:%M} · {'ACTUALIZAR datos' if actualizar else 'solo mantener despierta la app'}")
    destino = con_parametro(url, "actualizar", token) if actualizar else url
    limite = 600 if actualizar else 240  # segundos de espera a que la app termine de cargar

    with sync_playwright() as pw:
        navegador = pw.chromium.launch()
        pagina = navegador.new_page(viewport={"width": 1400, "height": 900})
        pagina.goto(destino, wait_until="domcontentloaded", timeout=120_000)

        # ¿Está dormida? En Streamlit Community Cloud aparece un botón para despertarla.
        boton = pagina.get_by_role("button", name=re.compile("get this app back up", re.I))
        try:
            boton.wait_for(state="visible", timeout=15_000)
            print("La app estaba dormida: se despierta.")
            boton.click()
        except PWTimeout:
            pass

        # En Community Cloud la app va dentro de un iframe; en otros hostings, directo en la página.
        inicio = time.time()
        contexto = pagina
        while time.time() - inicio < 90:
            if pagina.locator('iframe[title="streamlitApp"]').count():
                contexto = pagina.frame_locator('iframe[title="streamlitApp"]')
                break
            if pagina.get_by_text(TITULO).count():
                break
            time.sleep(2)

        try:
            contexto.get_by_text(TITULO).first.wait_for(state="visible", timeout=limite * 1000)
        except PWTimeout:
            pagina.screenshot(path="fallo.png")
            print("ERROR: la app no terminó de cargar a tiempo (ver fallo.png).")
            navegador.close()
            return 1

        # Espera a que Streamlit deje de ejecutar el script (desaparece el indicador "Running").
        tranquilo = 0
        while tranquilo < 4 and time.time() - inicio < limite:
            corriendo = contexto.locator('[data-testid="stStatusWidget"]').count()
            tranquilo = 0 if corriendo else tranquilo + 1
            time.sleep(2.5)

        texto = contexto.locator("body").inner_text(timeout=30_000)
        m = re.search(r"Última consulta a las API:\s*([^\n]+)", texto)
        print("Última consulta a las API según la app:", m.group(1).strip() if m else "no se pudo leer")
        if "Traceback" in texto or "Error" in texto.split("Dashboard de operación")[0][:300]:
            print("AVISO: la página muestra un error. Revisa la app.")
        navegador.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
