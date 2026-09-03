# Equivalencia de Códigos de Barra

App web simple para cruzar una planilla de pedido de un proveedor (columnas SKU y Color)
contra la planilla grande de "Equivalencia" (export tipo Crystal Reports) y devolver,
para cada línea del pedido, el/los código(s) de barra correspondientes.

## Cómo matchea

- De la **planilla grande** usa la columna A (Código de barra) y la columna B (Artículo,
  formato `SKU COLOR`, por ejemplo `28465 FEA`).
- De la **planilla de pedido** usa la columna A (SKU) y la columna B (Color).
- Arma la clave `SKU COLOR` de la planilla de pedido y busca coincidencia exacta contra
  la columna Artículo de la planilla grande.
- **No filtra por Talle.** Como un mismo SKU+Color tiene un código de barra distinto por
  cada talle, si hay varios talles vas a ver varios códigos candidatos en el resultado,
  cada uno con su talle entre paréntesis, para que elijas el correcto a mano. Si hay una
  sola coincidencia, se pone directo el código.

## Uso

1. Abrí la app en el navegador.
2. Subí la planilla grande de Equivalencia (`.xls` o `.xlsx`). Solo hace falta subirla la
   primera vez o cuando tengas una versión nueva — queda guardada en el servidor y se
   reutiliza en los próximos pedidos.
3. Subí la planilla de pedido del proveedor.
4. Tocá "Procesar y descargar resultado". Se descarga un Excel igual al de pedido, con dos
   columnas nuevas al final: `Código(s) de Barra` y `Cantidad de Coincidencias`.

## Generar etiqueta suelta (reimpresión)

Desde el botón "Generar etiqueta" en la página principal (o entrando directo a `/etiqueta`)
podés crear una etiqueta individual de **4 x 2 cm** ingresando a mano: descripción, color,
talle y código de barra. Genera un PNG a 300 DPI con el tamaño físico exacto de la etiqueta
(queda en los metadatos del archivo), con el código de barra real generado a partir del
número (EAN-13, UPC-A o Code128 según el largo del código). Al imprimir el PNG a "tamaño
real" / 100% (sin ajustar a página) sale del tamaño correcto.

## Deploy en Portainer (Docker Compose)

1. Subí este repo a GitHub.
2. En Portainer: **Stacks → Add stack → Repository**, apuntá al repo de GitHub y a
   `docker-compose.yml`.
3. Deploy. La app queda escuchando en el puerto `8020` del host (configurable en el
   `docker-compose.yml`).
4. Los datos (la planilla grande cacheada) se guardan en el volumen `equivalencia-data`,
   así que sobreviven a reinicios y updates del contenedor.

## Deploy manual con Docker

```bash
docker build -t equivalencia-app .
docker run -d --name equivalencia-app -p 8020:8000 -v equivalencia-data:/app/data equivalencia-app
```

## Desarrollo local

```bash
pip install -r requirements.txt
python app.py
```

La app queda en `http://localhost:8000`.
