# Reporte de compras

Reporte local para consultar un archivo de seguimiento de recursos, filtrar pedidos y registrar ETA y almacén destino por orden y proveedor.

## Uso

1. Instala Node.js en Windows.
2. Ejecuta `iniciar-reporte.cmd` y mantén abierta la ventana mientras utilizas el reporte.
3. Abre `http://127.0.0.1:3783/` y selecciona el Excel de seguimiento.

El reporte guarda los datos de seguimiento en `seguimiento-datos.json` dentro de esta carpeta y los recupera al abrir un Excel actualizado. Este archivo y los Excel no se publican en Git. Conserva una copia de respaldo de la base de seguimiento.