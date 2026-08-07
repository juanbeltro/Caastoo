# Category Assortment Tool — V0.1 / Sprint 1

Primer módulo funcional del MVP acordado.

## Alcance implementado
- Lectura local de Excel (.xlsx/.xls) en el navegador.
- Selector de hoja.
- Preview de filas/columnas.
- Mapping manual al modelo canónico.
- 7 campos obligatorios y 3 recomendados.
- Validaciones iniciales: SKU/segmento faltante, SKU duplicado, métricas no numéricas/vacías.
- Construcción de modelo normalizado en memoria.
- No hay backend ni base de datos.

## Ejecutar localmente
```bash
npm install
npm run dev
```

## Build para Netlify
```bash
npm run build
```
Publicar el directorio `dist` o conectar el repositorio de GitHub a Netlify.

## Fuera de alcance de este sprint
- Motor de surtido / Pareto / 70-30.
- Recomendaciones.
- Construcción A/B/C.
- Exportación final.
- IA y transformación de Excel no tabulares.


## Sprint 2
Motor analítico: variaciones, desempeño 70/30, Pareto por segmento, casos especiales y recomendación explicable.
