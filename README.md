# PhishARG Add-in

Complemento de Outlook desarrollado con React, TypeScript y Office.js.

## Requisitos

- Node.js 18+
- npm 9+

## Instalación

```bash
npm install
```

## Scripts disponibles

En el directorio del proyecto podés ejecutar:

### `npm start`

Inicia el servidor de desarrollo local en `https://localhost:3000`.

### `npm run build`

Compila la aplicación para producción en la carpeta `dist/`.

### `npm test`

Ejecuta las pruebas del proyecto.

## Variables de entorno

Copiá el archivo `.env.example` a `.env` para configurar las variables necesarias:

- `API_BASE_URL`
- `PHISHARG_TOKEN_KEY`
- `PHISHARG_RATE_LIMIT_ID`

## Carga del complemento en Outlook

- **Local:** cargar `manifest-local.xml` (`https://localhost:3000`).
- **Producción:** cargar `manifest.xml`.
