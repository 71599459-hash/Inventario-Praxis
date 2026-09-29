# Inventario TIC — Colegio Praxis

Sistema web profesional para administrar el inventario tecnológico del **Colegio Praxis**.

## Estado actual

La aplicación funciona como sitio web estático y está preparada para importar el archivo oficial de inventario directamente desde el navegador.

> **Importante:** el repositorio es público. Por seguridad, el Excel oficial y sus datos reales (DNI, responsables, series, ubicaciones exactas, cámaras y redes) **no se publican en GitHub**.

## Funciones implementadas

- Dashboard con resumen del inventario.
- Búsqueda por:
  - Código TIC.
  - Código Padre TIC.
  - Equipo / material.
  - Marca y modelo.
  - Serie.
  - Colaborador.
  - DNI.
  - Sede.
  - Área / aula.
- Inventario maestro.
- General / equipos administrativos.
- Inventario separado por sede.
- Colaboradores y equipos a su cargo.
- Almacén TIC separado.
- Cámaras y redes.
- Insumos, materiales y herramientas.
- Movimientos e historial.
- Reportes por sede, equipo y estado.
- Exportación CSV.
- Copia de seguridad JSON.
- Acta de cargo imprimible por colaborador.
- Diseño adaptable a computadora y celular.

## Regla principal: Código TIC único

El sistema trabaja con la regla:

**1 Código TIC = 1 equipo = 1 registro maestro**

Si el mismo Código TIC aparece en varias hojas del Excel, la aplicación lo consolida en un solo registro y conserva las apariciones anteriores dentro del historial.

Al transferir un equipo:
- de un colaborador a otro,
- de un área a otra,
- de una sede a otra,
- o al Almacén TIC,

**no se crea otro Código TIC**. Solo cambia la ubicación/responsable actual y se registra un movimiento.

## Hojas procesadas del Excel

La importación reconoce:

- General
- TIC - Almacén
- Cámaras y Redes
- Insum, mat y herra
- Compras
- Ventas
- Préstamos

Además normaliza algunos datos antes de crear el inventario maestro, por ejemplo códigos vacíos o inválidos y variantes de nombres de sede.

## Cómo usarlo

1. Abrir la página web.
2. Pulsar **Importar Excel**.
3. Seleccionar el archivo oficial `.xlsx`.
4. Esperar a que finalice la consolidación.
5. Usar el buscador, filtros, sedes, colaboradores y almacén.

Los datos importados quedan almacenados en **IndexedDB del navegador** del equipo utilizado.

## Privacidad

El Excel real no está incluido en este repositorio.

Esto es intencional porque el inventario contiene información interna del Colegio Praxis.

Consulta: [docs/PRIVACIDAD.md](docs/PRIVACIDAD.md)

## Próxima etapa institucional

Para utilizar el sistema desde varios equipos y por varios usuarios, la arquitectura prevista es:

- PostgreSQL / Supabase.
- Inicio de sesión.
- Roles: Administrador TIC, Editor y Consulta.
- Row Level Security.
- Auditoría de movimientos.
- Copias de seguridad.
- Acceso HTTPS.

La estructura inicial está en:

`supabase/schema.sql`

## Estructura del proyecto

```
Inventario-Praxis/
├── index.html
├── styles.css
├── app.js
├── excel-import.js
├── storage.js
├── branding.js
├── manifest.webmanifest
├── docs/
│   └── PRIVACIDAD.md
└── supabase/
    └── schema.sql
```

## Colegio Praxis

Proyecto de gestión del inventario TIC — 2026.
