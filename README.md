# Dúo Pizzería — pedidos online

| Archivo | Para quién | Qué hace |
|---|---|---|
| `index.html` | Clientes | Carta y pedidos para recoger o a domicilio (CP 31591). Con `?camarero=1` aparece la opción "Mesa" para el personal. |
| `panel-pedidos.html` | Cocina y dueño | Tablero de pedidos (avisa al cliente por WhatsApp al cambiar de estado), estadísticas, pausa rápida y configuración. |
| `Code.gs` | Google Apps Script | "Servidor" compartido: guarda pedidos y configuración en la Google Sheet y reparte los números de pedido. |

## Publicar cambios

- **Web y panel:** basta con subirlos a la rama `main`; GitHub Pages los publica en uno o dos minutos.
- **Code.gs:** no se publica desde GitHub. Pégalo en la Google Sheet (*Extensiones → Apps Script*), guarda y ve a
  *Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión → Implementar*,
  para mantener la misma URL que usan `index.html` y `panel-pedidos.html`.
