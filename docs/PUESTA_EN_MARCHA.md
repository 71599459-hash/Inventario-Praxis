# Puesta en marcha

## 1. Publicar la página

El proyecto incluye el workflow:

`.github/workflows/pages.yml`

Para la primera publicación, GitHub Pages debe habilitarse una vez desde el repositorio:

**Settings → Pages → Build and deployment → Source: GitHub Actions**

Después de habilitarlo, ejecutar nuevamente el workflow **Publicar Inventario Praxis**.

## 2. Cargar el inventario

Al abrir la página:

1. Pulsar **Importar Excel**.
2. Seleccionar el archivo oficial `.xlsx`.
3. Esperar el procesamiento.
4. Revisar el Dashboard.
5. Verificar los registros marcados **Revisar ubicación**.

## 3. Verificación inicial recomendada

Antes de utilizar el sistema como fuente oficial:

- comprobar los Códigos TIC consolidados;
- revisar equipos que aparecen en más de una ubicación;
- revisar registros sin Código TIC;
- revisar colaboradores sin DNI o área;
- validar equipos enviados a Almacén TIC.

## 4. Respaldo

Antes de cambios importantes utilizar:

**Inventario maestro → Copia JSON**

La copia se puede recuperar desde:

**Restaurar JSON**

## 5. Uso multiusuario

La versión actual almacena los datos importados en el navegador.

Para trabajar simultáneamente desde varios equipos se debe activar la siguiente etapa:

- Supabase/PostgreSQL;
- autenticación;
- roles;
- RLS;
- auditoría;
- copias de seguridad centralizadas.

Ver `supabase/schema.sql`.
