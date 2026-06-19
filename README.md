# Previsualizador y editor de Excel

Monorepo con frontend en React Vite y backend en NestJS para cargar, previsualizar, editar y descargar archivos Excel.

## Stack

- `apps/web`: React + Vite + TypeScript.
- `apps/api`: NestJS + TypeScript + ExcelJS.

## Instalacion

```powershell
npm install
```

## Desarrollo

Levantar frontend y backend juntos:

```powershell
npm run dev
```

Servicios:

- Frontend: `http://localhost:5173`
- Backend: `http://localhost:3000/api`

## Scripts utiles

```powershell
npm run dev:web
npm run dev:api
npm run build
```

## API

- `GET /api/excel/health`: verifica que la API este disponible.
- `POST /api/excel/parse`: recibe un archivo `.xlsx` o `.csv` en `multipart/form-data` con el campo `file` y devuelve hojas editables.
- `POST /api/excel/export`: recibe el workbook editado en JSON y devuelve un `.xlsx`.

## Funciones

- Carga de `.xlsx` y `.csv`.
- Previsualizacion por hojas.
- Edicion de celdas.
- Agregar filas y columnas.
- Exportacion a `.xlsx` desde NestJS.
