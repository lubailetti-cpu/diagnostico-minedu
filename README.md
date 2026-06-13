# Diagnóstico de carga operativa — MINEDU

Cuestionario web para mapear la carga documental de los servidores del MINEDU y servir como línea de base del impacto de la plataforma **Moderniza**.

## Estructura

- `index.html` — Cuestionario web (HTML estático).
- `apps-script.gs` — Código del backend para correr dentro del Apps Script de la UNOME (no se ejecuta en este repo, es referencia para que el equipo lo pegue en su Google).
- `tutorial-javier.md` — Guía paso a paso para que la UNOME conecte el formulario a su Google Sheets.

## Despliegue

El HTML se sirve mediante GitHub Pages.

La data NO se guarda en este repositorio ni en ningún servidor de Tidú: viaja directo al Apps Script de la UNOME y se almacena en su Google Sheets del MINEDU.

## Privacidad

- Los resultados se comunican únicamente de forma agregada.
- La base de respuestas vive en el Google del MINEDU.
- Ningún tercero (incluida Tidú) tiene acceso a las respuestas individuales.

## Licencia

Uso interno MINEDU. Material desarrollado por Tidú como parte del piloto Moderniza.
