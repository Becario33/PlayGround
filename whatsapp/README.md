# PlayGround / whatsapp

Consulta de asistencia Comscore (ayer) en SQL Server: se arma un Excel, se captura y **la imagen** se manda al grupo **Data Analytics** por WhatsApp Web. Chrome persistente: QR **una vez** (carpeta `chrome_whatsapp_perfil/`, no tu Chrome de diario).

En cada PC nueva hay que escanear el QR. Esa carpeta y el `.env` no se suben a GitHub. La base `10.55.55.134` solo responde **en el corporativo**.

## Requisitos

- Python 3.9+
- Windows (laptop becario), red Cinemex
- Chrome del sistema
- Excel (para la captura real; si no, se manda una tabla en imagen)
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

1. Corre la consulta (asistencia industria de ayer).
2. Arma el `.xlsx`, captura el rango y abre WhatsApp.
3. En el grupo Data Analytics manda la **imagen** (el texto de la consulta va de pie).
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
| SQL | pyodbc a `Programacion.dbo.ComscoreMPAMexico` (ayer) |
| Excel | openpyxl en `resultados/` |
| Captura | Excel COM (Windows) o imagen de tabla |
| Chrome | Playwright `launch_persistent_context` + Chrome del sistema |
| Perfil | `chrome_whatsapp_perfil/` junto a `main.py` |
| WhatsApp | Adjunta el PNG al grupo Data Analytics |

## Estructura

```
main.py                    Arranque
requirements.txt           Playwright + pyodbc
.env.example               Plantilla SQL (sin password)
.env                       Local, no se sube
resultados/                Excel y PNG locales (no se suben)
chrome_whatsapp_perfil/    Sesión WhatsApp (no se sube)
```

## Versiones

- **v1.3.1** — Imagen nítida y clic en el avión verde
- **v1.3.0** — Captura Excel al grupo Data Analytics
- **v1.2.0** — Resultado SQL al grupo Data Analytics
- **v1.1.1** — Segunda corrida: cierra Chrome anterior; menos `--no-sandbox`
- **v1.1.0** — Carpeta `whatsapp/` dentro de PlayGround
- **v1.0.0** — WhatsApp Web con perfil persistente

Autor: Becario33 <becario33@cinemex.net>
