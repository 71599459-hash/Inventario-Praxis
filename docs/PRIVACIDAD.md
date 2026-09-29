# Privacidad y seguridad de los datos

El inventario del Colegio Praxis contiene información que no debe publicarse en un repositorio público, por ejemplo:

- DNI y nombres de colaboradores.
- Ubicación exacta de equipos.
- Series y códigos de activos.
- Inventario de cámaras, NVR, DVR y redes.
- Observaciones internas.

Por ello, el proyecto **no contiene el archivo Excel oficial ni una copia de sus datos reales**.

La aplicación importa el Excel directamente en el navegador y lo guarda en IndexedDB del equipo local. Esto permite probar y usar la interfaz sin publicar información sensible.

## Para uso institucional multiusuario

La siguiente etapa debe usar:

1. Base de datos privada PostgreSQL/Supabase.
2. Inicio de sesión.
3. Row Level Security.
4. Registro de auditoría.
5. HTTPS.
6. Copias de seguridad.
7. Acceso por roles.

El archivo `supabase/schema.sql` deja preparada la estructura principal.
