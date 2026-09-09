# PlayGround / whatsapp: WhatsApp Web con Chrome persistente.
# Primera vez: escanea el QR. Siguientes: la misma carpeta chrome_whatsapp_perfil.
#   python main.py
# Grupo "Data Analytics" y texto de prueba. Número (lada país, sin +):
#   python main.py --para 5215512345678

import argparse
import os
import sys
import time

URL_WA = "https://web.whatsapp.com/"
PERFIL = "chrome_whatsapp_perfil"
GRUPO_DEFAULT = "Data Analytics"
TEXTO_DEFAULT = "Prueba de persistencia: la sesión sigue abierta; el QR solo se escanea una vez."


def _dir_script():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def _lanzar(pw, perfil):
    args = ["--disable-dev-shm-usage"]
    try:
        return pw.chromium.launch_persistent_context(
            perfil,
            headless=False,
            channel="chrome",
            args=args,
            locale="es-MX",
            no_viewport=True,
        )
    except Exception:
        return pw.chromium.launch_persistent_context(
            perfil,
            headless=False,
            args=args,
            locale="es-MX",
            no_viewport=True,
        )


def _hay_sesion(page):
    if page.locator('[data-testid="chat-list-search-container"]').count():
        return True
    if page.get_by_role("textbox", name="Search or start a new chat").count():
        return True
    if page.get_by_role("textbox", name="Buscar un chat o iniciar uno nuevo").count():
        return True
    if page.locator("#pane-side").count():
        return True
    return page.locator('[aria-label="Lista de chats"]').count() > 0


def _hay_qr(page):
    if page.locator("canvas[aria-label*='QR']").count():
        return True
    if page.locator("div[data-ref]").count():
        return True
    if page.get_by_text("Escanea el código QR", exact=False).count():
        return True
    return page.get_by_text("Scan the QR code", exact=False).count() > 0


def esperar_sesion(page, segundos=180):
    fin = time.time() + segundos
    aviso_qr = False
    while time.time() < fin:
        if _hay_sesion(page):
            print("Sesión lista (sin QR).")
            return True
        if _hay_qr(page) and not aviso_qr:
            print("Escanea el QR con el teléfono. Esta carpeta de Chrome se reusa la próxima vez.")
            aviso_qr = True
        page.wait_for_timeout(1000)
    print("No entró a WhatsApp a tiempo.")
    return False


def _caja_busqueda(page):
    loc = page.locator('[data-testid="chat-list-search-container"] input[data-tab="3"]')
    if loc.count():
        return loc.first
    loc = page.get_by_role("textbox", name="Search or start a new chat")
    if loc.count():
        return loc.first
    loc = page.get_by_role("textbox", name="Buscar un chat o iniciar uno nuevo")
    if loc.count():
        return loc.first
    loc = page.locator('input[aria-label="Search or start a new chat"]')
    if loc.count():
        return loc.first
    loc = page.locator('input[data-tab="3"]')
    if loc.count():
        return loc.first
    return None


def _caja_mensaje(page):
    nombres = (
        "Type a message",
        "Escribe un mensaje",
        "Mensaje",
    )
    for nombre in nombres:
        loc = page.locator(f'[contenteditable="true"][aria-label="{nombre}"]')
        if loc.count():
            return loc.last
        loc = page.get_by_role("textbox", name=nombre)
        if loc.count():
            return loc.last
    loc = page.locator("footer [contenteditable='true']")
    if loc.count():
        return loc.last
    loc = page.locator('#main footer [contenteditable="true"]')
    if loc.count():
        return loc.last
    return None


def abrir_grupo(page, nombre):
    caja = _caja_busqueda(page)
    if caja is None:
        print("No encontré la búsqueda de chats.")
        return False
    caja.click()
    page.wait_for_timeout(300)
    caja.fill(nombre)
    print(f'Busqué: {nombre}')
    page.wait_for_timeout(2000)

    chat = page.locator(f'span[title="{nombre}"]')
    if chat.count() == 0:
        chat = page.locator(f'div[title="{nombre}"]')
    if chat.count() == 0:
        chat = page.get_by_title(nombre, exact=True)
    try:
        chat.first.click(timeout=15000)
    except Exception:
        print(f'No encontré el grupo "{nombre}" en la lista filtrada.')
        return False

    try:
        page.get_by_role("banner").get_by_text(nombre, exact=False).first.wait_for(timeout=8000)
    except Exception:
        try:
            page.locator("#main header").get_by_text(nombre, exact=False).first.wait_for(timeout=8000)
        except Exception:
            page.wait_for_timeout(1500)
    print(f'Grupo abierto: {nombre}')
    return True


def escribir_y_enviar(page, texto):
    caja = _caja_mensaje(page)
    if caja is None:
        print("No encontré la caja de mensaje.")
        return False
    caja.click()
    page.wait_for_timeout(300)
    page.keyboard.insert_text(texto)
    page.wait_for_timeout(300)
    enviar_btn = page.locator('button[aria-label="Send"]')
    if enviar_btn.count() == 0:
        enviar_btn = page.locator('button[aria-label="Enviar"]')
    if enviar_btn.count() == 0:
        enviar_btn = page.locator('span[data-icon="send"]')
    if enviar_btn.count():
        try:
            enviar_btn.first.click(timeout=8000)
        except Exception:
            page.keyboard.press("Enter")
    else:
        page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    print("Mensaje enviado.")
    return True


def enviar_grupo(page, nombre, texto):
    if not abrir_grupo(page, nombre):
        return False
    return escribir_y_enviar(page, texto)


def enviar(page, telefono, texto):
    tel = "".join(c for c in telefono if c.isdigit())
    if not tel:
        print("Número vacío.")
        return False
    from urllib.parse import quote
    page.goto(
        f"https://web.whatsapp.com/send?phone={tel}&text={quote(texto)}",
        wait_until="domcontentloaded",
        timeout=120000,
    )
    if not esperar_sesion(page, 90):
        return False
    enviar_btn = page.locator('button[aria-label="Enviar"]')
    if enviar_btn.count() == 0:
        enviar_btn = page.locator('span[data-icon="send"]')
    try:
        enviar_btn.first.click(timeout=20000)
    except Exception:
        print("No encontré el botón Enviar. Revisa el número o si WhatsApp cambió la web.")
        return False
    page.wait_for_timeout(2000)
    print("Mensaje enviado (o al menos se pulsó Enviar).")
    return True


def main():
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Experimento WhatsApp Web (perfil persistente).")
    ap.add_argument("--grupo", default=GRUPO_DEFAULT, help="Nombre exacto del grupo")
    ap.add_argument("--para", default="", help="Número destino con lada país, sin + (opcional)")
    ap.add_argument("--texto", default=TEXTO_DEFAULT)
    ap.add_argument("--solo-abrir", action="store_true", help="No enviar, solo abrir sesión")
    args = ap.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Falta Playwright. En la carpeta del repo:")
        print("  python -m pip install -r requirements.txt")
        print("  python -m playwright install chromium")
        return 1

    perfil = os.path.join(_dir_script(), PERFIL)
    os.makedirs(perfil, exist_ok=True)
    print("Perfil Chrome (no es tu Chrome de diario):")
    print(" ", perfil)

    with sync_playwright() as pw:
        try:
            ctx = _lanzar(pw, perfil)
        except Exception as e:
            print("No se abrió Chrome.")
            print("  python -m pip install -r requirements.txt")
            print("  python -m playwright install chromium")
            print(" ", str(e).split("\n")[0])
            return 1
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(URL_WA, wait_until="domcontentloaded", timeout=120000)
        ok = esperar_sesion(page)
        if ok and not args.solo_abrir:
            if args.para:
                ok = enviar(page, args.para, args.texto)
            else:
                ok = enviar_grupo(page, args.grupo, args.texto)
        print("Cierra la ventana de Chrome cuando termines (o Enter aquí).")
        try:
            input()
        except EOFError:
            page.wait_for_timeout(8000)
        ctx.close()
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
