# Nuevo Inventario Mejorado

Web app para inventariar a partir de un Excel, con:
- Carga de Excel (XLS/XLSX)
- Búsqueda difusa
- Dictado por voz (si el navegador lo soporta)
- Entrada de texto **multilínea** (compatible con apps como SuperWhisper)
- Reset y Exportar

## Uso
1. Sube tu archivo Excel (primera fila = cabeceras; debe existir columna **Producto** y opcionalmente **UMB**).
2. Las cantidades se inicializan a 0 para el recuento.
3. Añade con voz o pegando varias líneas tipo:
   - `3 Coca Cola`
   - `Naranjas 5`
   - `x3 Tónica Schweppes`
4. Exporta para guardar la columna `Stock` en el Excel.

## Despliegue en Netlify
- Arrastra la carpeta al dashboard de Netlify o conecta con GitHub.
- No requiere backend ni configuración extra.

## Notas
- Si tu Excel no tiene columna `Stock`, se añadirá automáticamente al exportar.
- La precisión de coincidencia de producto se puede ajustar en `threshold` del índice Fuse.