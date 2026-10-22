# Test de Alma

Test de 16 preguntas (más 5 sobre pérdida) que asigna uno de 22 arquetipos, cada uno ligado a un arcano mayor del tarot. Se cobra con Mercado Pago y el resultado completo llega en PDF por email.

- `index.html` y `img/`: la web, publicada con GitHub Pages.
- `contenido/contenido-test-v2.json`: preguntas, puntuación, arquetipos y textos.
- `apps-script/`: script de Google que crea el pago en Mercado Pago y envía el PDF.
- `herramientas/`: plantilla y scripts que generan `index.html` a partir del contenido.

La dirección del script de pago va en la constante `PAGOS.endpoint` de `index.html`. Mientras esté vacía, los botones de compra avisan de que los pagos se habilitan pronto.
