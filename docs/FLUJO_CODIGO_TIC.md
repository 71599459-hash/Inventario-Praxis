# Flujo del Código TIC

## Regla obligatoria

**Un Código TIC representa un único equipo dentro del inventario.**

El sistema no debe crear un segundo registro con el mismo Código TIC cuando cambia:

- el colaborador responsable;
- el área o aula;
- la sede;
- el estado;
- o la ubicación hacia/desde Almacén TIC.

## Ejemplo de transferencia

1. El equipo está en **Sede Estrellas / Subdirección / Colaborador A**.
2. TIC selecciona **Mover / transferir**.
3. Se define el nuevo destino.
4. El registro maestro conserva el mismo Código TIC.
5. La ubicación actual se actualiza.
6. El movimiento anterior queda en el historial.

## Envío al almacén

Al seleccionar **Enviar al almacén**:

- se conserva el Código TIC;
- se limpia el responsable cuando corresponde;
- se cambia la ubicación a **TIC / ALMACÉN**;
- y se registra el movimiento.

## Duplicados del Excel

Durante la importación, si un Código TIC aparece varias veces:

- se crea **un solo registro maestro**;
- se conservan las apariciones como historial;
- si las apariciones indican sedes, áreas o responsables distintos, se marca **Revisar ubicación**;
- la alerta desaparece cuando TIC confirma un nuevo movimiento y define la ubicación vigente.

## Códigos inválidos

Valores como:

- vacío;
- `#REF!`;
- `-`;
- `0`;

no se consideran Códigos TIC válidos.

Los registros sin Código TIC pueden visualizarse para depuración, pero no se permite crear un nuevo equipo web sin un Código TIC válido.
