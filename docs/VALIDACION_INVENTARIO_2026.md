# Validación del Inventario Oficial 2026

Esta validación se realizó contra la estructura real del archivo **Inventario Oficial 2026 FUSIONADO OFICIAL** recibido para el proyecto.

No se publican nombres, DNI, series, ubicaciones exactas ni otros datos internos en este repositorio público.

## Resultado técnico de la revisión

| Indicador | Resultado |
|---|---:|
| Filas con información en las hojas procesadas | 4,321 |
| Registros de hojas activas | 2,054 |
| Códigos TIC únicos | 1,140 |
| Códigos TIC repetidos en hojas activas | 70 |
| Códigos con ubicaciones/responsables distintos | 58 |
| Registros útiles sin Código TIC | 591 |
| Registros maestros esperados después de consolidar | 1,731 |

## Hojas revisadas

- General
- TIC - Almacén
- Cámaras y Redes
- Insum, mat y herra
- Compras
- Ventas
- Préstamos

Las hojas auxiliares de resumen o detalle no se consideran fuente maestra del inventario.

## Hallazgo principal

El problema más importante no es únicamente que existan códigos repetidos, sino que un mismo Código TIC puede aparecer en más de una hoja con:

- otra sede;
- otra área;
- otro responsable;
- o una ubicación de almacén diferente.

El cruce más frecuente detectado fue **General ↔ TIC - Almacén**.

Por ello, el sistema aplica la regla:

> **1 Código TIC = 1 equipo = 1 registro maestro**

Las apariciones repetidas no generan copias del equipo. Se guardan como historial.

## Cómo determina la ubicación inicial

Al consolidar un Código TIC:

1. Se revisan las fechas disponibles en Fecha (I), Fecha (E) y Fecha (A).
2. Cuando existe una fecha claramente más reciente, esa aparición tiene prioridad como ubicación inicial.
3. Cuando no existe una fecha confiable, se utiliza una prioridad de fuente y calidad del registro.
4. Si las apariciones indican ubicaciones o responsables distintos, el registro queda marcado como **Revisar ubicación** aunque se haya seleccionado una ubicación inicial.

Esto evita asumir que una fila antigua sigue siendo la ubicación vigente.

## Módulo Revisión

La web incluye una pantalla de **Revisión y depuración** que muestra:

- Códigos TIC con ubicaciones conflictivas.
- Registros sin Código TIC.
- Cantidad de apariciones consolidadas.

Para un conflicto de ubicación, TIC puede:

- elegir una de las apariciones existentes del Excel;
- o definir manualmente la ubicación vigente.

Al confirmar:

- el Código TIC no cambia;
- no se crea un segundo registro;
- la ubicación/responsable actual se actualiza;
- y la decisión queda registrada en el historial.

## Traslado al almacén

Enviar un equipo al Almacén TIC significa actualizar su ubicación actual.

No se crea una nueva fila maestra ni un nuevo Código TIC.

El historial conserva:

- ubicación anterior;
- destino;
- responsable anterior y nuevo cuando corresponda;
- fecha;
- motivo u observación.

## Registros sin Código TIC

Los registros útiles sin Código TIC no se eliminan.

Se mantienen visibles en Revisión para poder:

- asignarles un Código TIC;
- completar datos;
- o determinar posteriormente si corresponden a componentes que dependen de un Código Padre TIC.

## Edición de fichas

La web permite editar datos como:

- equipo/material;
- código padre;
- marca;
- modelo;
- serie;
- sede;
- área;
- responsable;
- DNI;
- estado;
- condición;
- situación;
- observaciones.

Si se intenta asignar un Código TIC que ya pertenece a otro registro, el sistema bloquea el guardado.

## Privacidad

El archivo Excel real **no debe subirse al repositorio público** porque contiene información interna del Colegio Praxis.

La versión actual procesa el archivo dentro del navegador y guarda el estado en IndexedDB.

Para uso institucional desde varios equipos, la siguiente fase es una base de datos privada con autenticación y auditoría.
