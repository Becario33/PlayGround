# PlayGround

Patio de experimentos. Cada prueba vive en **su carpeta**.

## Requisitos

- Python 3.9+
- Windows (laptop becario)

```bat
git clone https://github.com/Becario33/PlayGround.git
cd PlayGround
```

## Experimentos

| Carpeta | Qué es | Arranque |
|---------|--------|----------|
| `whatsapp/` | SQL → Excel → captura → WhatsApp | `cd whatsapp` → `python main.py` |
| `cinepolis/` | Dictamen: Cinépolis no es cazable como Cinemex | solo `DICTAMEN.md` |

Para uno nuevo: carpeta nueva + `main.py` (si hay código). No sueltes archivos en la raíz.

## Estructura

```
README.md
whatsapp/              WhatsApp Web
cinepolis/             Nota Cinépolis
```

## Versiones

- **v1.4.8** — Envío: clip + input foto, sin clic al Explorador
- **v1.4.7** — Restaura lógica estable de `v1.3.8` (foto, no sticker)
- **v1.4.6** — Clip: prueba inputs hasta preview; hover Fotos/videos
- **v1.4.5** — Foto real: JPEG ≥1280×720 + solo input Fotos/videos
- **v1.4.4** — Adjunto sin Explorador (input oculto tras el clip)
- **v1.4.3** — Foto (no sticker) + tabla sin bordes grises
- **v1.4.2** — Por defecto Pillow estable; COM opcional
- **v1.4.1** — Adjunto WhatsApp sin pelear con el Explorador
- **v1.4.0** — Captura literal Excel COM (fallback Pillow)
- **v1.3.8** — Alto de la tabla al texto, sin lienzo blanco
- **v1.3.7** — Fix import Image en el recorte
- **v1.3.6** — Recorte dinámico al contenido
- **v1.3.5** — Tabla sin marco blanco
- **v1.3.4** — JPEG por Fotos y videos; solo la imagen
- **v1.3.3** — Clic JS en Send 1 selected
- **v1.3.2** — Enviar imagen: botón Send 1 selected
- **v1.3.1** — Imagen nítida + botón enviar del preview
- **v1.3.0** — Captura Excel de la consulta al grupo
- **v1.2.0** — WhatsApp manda asistencia Comscore de ayer
- **v1.1.1** — WhatsApp: cierra Chrome antes de la segunda corrida
- **v1.1.0** — Experimentos por carpeta
- **v1.0.0** — WhatsApp Web con perfil persistente

Autor: Becario33 <becario33@cinemex.net>
