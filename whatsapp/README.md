# PlayGround / whatsapp

WhatsApp Web con un Chrome persistente: el QR se escanea **una vez** y la sesión queda en `chrome_whatsapp_perfil/` (junto a este `main.py`, no es tu Chrome de diario).

En cada máquina nueva hay que escanear el QR **en esa PC**. Esa carpeta no se sube a GitHub.

## Requisitos

- Python 3.9+
- Windows (laptop becario) o Mac (desarrollo)
- Chrome del sistema

```bat
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Configuración

1. Clona y entra a esta carpeta:
   ```bat
   git clone https://github.com/Becario33/PlayGround.git
   cd PlayGround
   cd whatsapp
   ```
2. No hay `.env`. Cierra el Chrome del experimento si quedó abierto.

## Uso

```bat
python main.py
```

1. Primera vez: escanea el QR de WhatsApp Web.
2. Busca el grupo **Data Analytics** y manda el texto de persistencia.
3. Enter en la consola (o cierra Chrome) para terminar.
4. Segunda vez: misma carpeta, sin QR.

Opcional:

```bat
python main.py --grupo "Data Analytics"
python main.py --texto "otro mensaje"
python main.py --para 5215512345678 --texto "hola"
python main.py --solo-abrir
```

## Qué hace cada parte

| Parte | Cómo |
|--------|------|
| Chrome | Playwright `launch_persistent_context` + Chrome del sistema |
| Perfil | Carpeta `chrome_whatsapp_perfil/` junto a `main.py` |
| WhatsApp | `web.whatsapp.com`: busca grupo y envía texto |

## Estructura

```
main.py                    Arranque
requirements.txt           Playwright
chrome_whatsapp_perfil/    Sesión local (no se sube)
```

## Versiones

- **v1.1.0** — Carpeta `whatsapp/` dentro de PlayGround
- **v1.0.0** — WhatsApp Web con perfil persistente

Autor: Becario33 <becario33@cinemex.net>
