# PlayGround / whatsapp

Top 10 películas por taquilla Comscore (mes pasado) en SQL Server: se arma un Excel, se captura y **la imagen** se manda al grupo **Data Analytics** por WhatsApp Web. Chrome persistente: QR **una vez** (carpeta `chrome_whatsapp_perfil/`, no tu Chrome de diario).

En cada PC nueva hay que escanear el QR. Esa carpeta y el `.env` no se suben a GitHub. La base `10.55.55.134` solo responde **en el corporativo**.

## Requisitos

- Python 3.9+
- Windows (laptop becario), red Cinemex
- Chrome del sistema
- ODBC Driver 17 (o 18) for SQL Server

```bat
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Configuración

1. Clona y entra:
   ```bat
   git clone https://github.com/Becario33/PlayGround.git
   cd PlayGround
   cd whatsapp
   ```
2. Credenciales SQL:
   ```bat
   copy .env.example .env
   ```
   Edita `.env`: `SQL_USER` y `SQL_PASSWORD` (no se sube).
3. Cierra el Chrome del experimento si quedó abierto.

## Uso

```bat
python main.py
```

1. Corre el TOP 10 por taquilla (mes pasado).
2. Arma el `.xlsx`, dibuja la tabla (Pillow) y abre WhatsApp.
3. En el grupo Data Analytics manda solo la **imagen**.
4. Enter y **cierra Chrome** antes de volver a correr.

Opcional:

```bat
python main.py --prueba
python main.py --texto "otro mensaje"
python main.py --grupo "Data Analytics"
python main.py --solo-abrir
```

## Qué hace cada parte

| Parte | Cómo |
|--------|------|
| SQL | TOP 10 por `Taquilla` del mes pasado (`Mes`) |
| Excel | openpyxl en `resultados/` |
| Captura | Pillow (tabla JPEG, estilo `v1.3.8`) |
| Chrome | Playwright `launch_persistent_context` + Chrome del sistema |
| Perfil | `chrome_whatsapp_perfil/` junto a `main.py` |
| WhatsApp | JPEG por Fotos y videos al grupo Data Analytics |

## Estructura

```
main.py                    Arranque
requirements.txt           Playwright + pyodbc
.env.example               Plantilla SQL (sin password)
.env                       Local, no se sube
resultados/                Excel y JPEG locales (no se suben)
chrome_whatsapp_perfil/    Sesión WhatsApp (no se sube)
```

## Versiones

- **v1.6.1** — Filtro mes pasado (columna Mes)
- **v1.6.0** — Top 10 por taquilla en vez de total asistencia
- **v1.5.1** — Captura sin bordes/lienzo 1280×720
- **v1.5.0** — Paste foto o Document (ya no clip image/* → sticker)
- **v1.4.16** — Cancel en Discard; ya no Escape con preview
- **v1.4.15** — Lienzo 1280×720 + clic Photos & videos (anti-sticker)
- **v1.4.14** — Input real `image/*` en footer (probado en WhatsApp Web)
- **v1.4.13** — Foto vía `input[type=file]` accept=video (no depende del clic)
- **v1.4.12** — Espera aria-expanded + clic JS a Photos & videos
- **v1.4.11** — Menú: clic a `menuitem` Photos & videos del HTML
- **v1.4.10** — Foto fija: `aria-label="Photos & videos"` + JPEG grande
- **v1.4.9** — Attach + expect_file_chooser en Photos & videos
- **v1.4.8** — Envío sin Explorador (clip + input video)
- **v1.4.7** — Restaura `main.py` de `v1.3.8` (foto estable)
- **v1.4.6** — Adjunto: hover + prueba inputs hasta preview
- **v1.4.5** — Anti-sticker: escala ≥1280×720 y solo Fotos/videos
- **v1.4.4** — Clip: adjunta por input oculto, sin Explorador
- **v1.4.3** — Foto por Fotos y videos (no sticker) + sin bordes
- **v1.4.2** — Por defecto Pillow estable; COM solo con `--excel-com`
- **v1.4.1** — Adjunto sin Explorador (input oculto / file chooser)
- **v1.4.0** — Captura literal Excel COM (fallback Pillow / `--pillow`)
- **v1.3.8** — Alto de filas al texto (sin lienzo blanco)
- **v1.3.7** — Fix recorte (import Image)
- **v1.3.6** — Recorte dinámico al tamaño de la tabla
- **v1.3.5** — Solo la tabla, sin marco blanco
- **v1.3.4** — Foto JPEG (no sticker) y sin texto extra
- **v1.3.3** — Enviar imagen: clic JS en Send 1 selected
- **v1.3.2** — Clic en Send 1 selected (preview)
- **v1.3.1** — Imagen nítida y clic en el avión verde
- **v1.3.0** — Captura Excel al grupo Data Analytics
- **v1.2.0** — Resultado SQL al grupo Data Analytics
- **v1.1.1** — Segunda corrida: cierra Chrome anterior; menos `--no-sandbox`
- **v1.1.0** — Carpeta `whatsapp/` dentro de PlayGround
- **v1.0.0** — WhatsApp Web con perfil persistente

Autor: Becario33 <becario33@cinemex.net>
