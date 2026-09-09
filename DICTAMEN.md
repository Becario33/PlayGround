# Dictamen Cinépolis — no cazar como Cinemex

Fecha: 25 ago 2026. Vistazo humano a [cinepolis.com/mx](https://cinepolis.com/mx), sin script y sin prototipo.

## Veredicto

**Tipo 2 (pase corto), con capas de tipo 3 en compra.** No es tipo 1 (JSON estable como Cinemex). **No-go:** no hay `main.py`, no se clona HORUS/HADES/IRIS, no se pelean tokens ni Chrome a escala cadena.

| Tipo | Qué sería | Qué se vio |
|------|-----------|------------|
| 1 Cinemex | REST + consumer key que aguanta semanas | No. No hay `X-API-Consumer-Key` ni `api.cinemex.com`-style |
| 2 Pase corto | Sesión / Authorization que nace al entrar y caduca | Sí. `sessionTtl` de México = **900 segundos (15 min)**; el bundle manda **Authorization** al GraphQL de cartelera |
| 3 Muro | Captcha, cola, app-only | Parcial: **reCAPTCHA** y **Queue-it** ya van en la home; no hace falta el muro 100 % del tiempo para ver posters |

El rumor de “a los 3 días se rompe” encaja, y el propio sitio es más agresivo: 15 minutos de TTL de sesión, no 3 días. Cazar un pase y pegarlo en Python es mantenimiento eterno, no un HADES.

## Qué se vio en la home (sin login)

- Cartelera de posters y títulos carga para cualquiera. “Elige tu cine” lista ciudades (CDMX: 74 cines). Guardar favoritos pide sesión.
- No es REST v2.37.2. El navegador habla **GraphQL** en un gateway (`api-g`, billboards / locations / config). CMS de posters aparte.
- Club Cinépolis usa un login tipo tenant (`club.cinepolis.com`, `clientId`, `audience`). Eso es cuenta de lealtad, no una llave de cartelera para scripts.
- Microfrontends (filtro de cines, carrito, banners), Dynatrace, cola Queue-it. Más superficie que Cinemex, más cosas que se pudren.

## Por qué no hay prototipo de un cine / un día

El plan solo autorizaba un HTTP chico **si** fuera tipo Cinemex. No lo es.

Un “un cine, un día” seguiría necesitando el pase corto (Authorization + TTL 15 min). Eso ya es el truco que el plan prohibió: refrescar tokens. Chrome contra la cadena no cabe en FIN-BECARIO-33-L (≤20 min, ≤3 ventanas).

## Qué sí usar para Cinépolis

Taquilla / asistencia / pantallas vs Cinemex: **Excel Rentrak** en Streamlit, como ahora.

HORUS (horarios), HADES (precio boleto) e IRIS (ficha de estreno) no tienen equivalente estable aquí. Precio y butacas estarían aún más atrás del pase y del recaptcha que la cartelera de posters.

## Cierre

Cinemex se pudo cazar porque el sitio deja un contrato JSON con clave de consumidor. Cinépolis deja un **pase**. El experimento terminó en el dictamen: no automatizar.
