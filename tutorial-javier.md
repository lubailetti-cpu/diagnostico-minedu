# Cómo activar el diagnóstico — Guía paso a paso

Hola Javier,

Esta guía explica cómo conectar el formulario HTML de la encuesta con tu Google Sheets, en aproximadamente **15 minutos**. Todo el flujo queda dentro de tu cuenta Google del MINEDU: nadie más que vos tendrá acceso a las respuestas.

---

## Resumen del flujo

1. Creás una hoja de cálculo nueva en tu Google Drive.
2. Pegás un código corto en una herramienta de Google llamada **Apps Script**.
3. "Publicás" ese código como una aplicación web.
4. Te genera una URL que me pasás.
5. Yo la conecto al formulario y queda funcionando.

---

## Paso 1 — Crear la hoja de cálculo

1. Entrá a **drive.google.com** con tu cuenta del MINEDU.
2. Hacé clic en **Nuevo > Hoja de cálculo de Google**.
3. Cambialé el nombre arriba a la izquierda a algo claro, por ejemplo:
   **"Diagnóstico carga operativa — Respuestas"**
4. Dejala vacía. No necesitás escribir columnas: el código las crea solo la primera vez que llega una respuesta.

---

## Paso 2 — Abrir Apps Script

1. Dentro de la hoja, andá al menú **Extensiones > Apps Script**.
2. Se abre una pestaña nueva con un editor de código.
3. Borrá todo el código que aparece por defecto (suele ser algo como `function myFunction() {}`).

---

## Paso 3 — Pegar el código

1. Abrí el archivo **`apps-script.gs`** que te paso aparte (o copialo del repo de GitHub: `https://github.com/lubailetti-cpu/diagnostico-minedu`).
2. Pegá el contenido completo en el editor de Apps Script.
3. Guardá con **Ctrl + S** (o el ícono de disquete). Te va a pedir nombre del proyecto: ponele "Diagnóstico MINEDU".

---

## Paso 4 — Publicar como aplicación web

1. Arriba a la derecha, hacé clic en **Implementar > Nueva implementación**.
2. Hacé clic en el ícono de engranaje (⚙) al lado de "Seleccionar tipo".
3. Elegí **Aplicación web**.
4. Llenás los campos así:
   - **Descripción**: "Diagnóstico v1"
   - **Ejecutar como**: Yo (tu correo del MINEDU)
   - **Quién tiene acceso**: **Cualquier persona** ← importante para que el formulario pueda enviar datos
5. Hacé clic en **Implementar**.
6. **La primera vez** Google va a pedirte autorizar el script para que pueda escribir en tu hoja:
   - Hacé clic en **Autorizar acceso**.
   - Elegí tu cuenta del MINEDU.
   - Si te aparece "Google no verificó esta aplicación", hacé clic en **Configuración avanzada** y luego en **Ir a Diagnóstico MINEDU (no seguro)**. Es seguro: es tu propio código, en tu propia cuenta.
   - Aceptá los permisos.
7. Te muestra una **URL** que termina en **/exec**. Por ejemplo:
   `https://script.google.com/macros/s/AKfycb.../exec`
8. **Copiá esa URL y pasámela por mensaje.**

---

## Paso 5 — Yo conecto el formulario

Cuando recibo tu URL:
1. La pego en el HTML del formulario.
2. Subo el cambio.
3. Te aviso "ya está listo, podés probar".

---

## Paso 6 — Probar

1. Abrí la URL del formulario que te pase.
2. Llenalo vos mismo como si fueras un servidor cualquiera.
3. Apretá enviar.
4. Volvé a tu hoja de cálculo: debería aparecer una fila nueva con tus respuestas. La primera vez también aparecen los encabezados de las columnas (creados automáticamente).

Si funciona: listo, podés compartir el link del formulario con los servidores del MINEDU.

Si no funciona: avisame y revisamos juntos.

---

## Preguntas frecuentes

**¿Y si quiero cambiar algo del cuestionario después?**
Lo cambio en GitHub. Se actualiza solo. Tu hoja sigue recibiendo respuestas normalmente, pero las preguntas nuevas se agregan como columnas nuevas al final.

**¿Y si quiero parar de recibir respuestas?**
Volvés a Apps Script, **Implementar > Administrar implementaciones**, y archivás la implementación. El formulario va a seguir cargando pero no podrá enviar.

**¿Las respuestas son seguras?**
Sí. La hoja vive en tu Google Drive del MINEDU. Solo vos (y a quienes vos compartas la hoja) la pueden ver. Yo no tengo acceso, salvo que vos me agregues.

**¿Puedo limitar a que solo respondan correos @minedu.gob.pe?**
Sí pero requiere un cambio extra en el código. Si querés esa restricción, decime y te paso la versión modificada.

---

## Apoyo

Cualquier duda en cualquier paso, escribime. La parte del cuestionario está pensada para no requerir conocimientos técnicos: si algo no calza, lo más probable es que sea un paso mal explicado de mi lado, no que vos lo hagas mal.

Lucía
