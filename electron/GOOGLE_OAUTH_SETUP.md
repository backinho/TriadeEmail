# Vinculación local de Gmail (Electron)

## Google Cloud Console

1. Selecciona el proyecto Google y habilita **Gmail API**.
2. Configura la pantalla de consentimiento OAuth. Si el estado de publicación es **Testing**, añade las cuentas que usarás como test users. Google puede limitar a siete días la duración de los refresh tokens de una app externa que continúa en Testing y solicita scopes de Gmail.
3. Crea credenciales OAuth de tipo **Desktop app** y copia el client ID.
4. No crees ni distribuyas un `client_secret`. El flujo de escritorio usa Authorization Code + PKCE y un callback loopback temporal en `127.0.0.1`.
5. Autoriza estos scopes:
   - `https://www.googleapis.com/auth/gmail.modify`: leer, destacar y mover correos entre Inbox, Spam y Papelera.
   - `https://www.googleapis.com/auth/gmail.send`: enviar correo.
   - `https://www.googleapis.com/auth/userinfo.email`: identificar la cuenta conectada.
   - Triade también pide `https://mail.google.com/` porque permite eliminación definitiva. Es restringido y requiere revisión de Google para distribuir públicamente. Si no necesitas eliminar definitivamente, quita ese scope de `GOOGLE_OAUTH_SCOPES` en `src/scripts/app.ts`.

## Configuración local

Añade el client ID público en el `.env` de la raíz:

```dotenv
PUBLIC_GOOGLE_CLIENT_ID=TU_CLIENT_ID.apps.googleusercontent.com
```

No añadas `client_secret`, claves de servicio de Supabase ni refresh tokens al `.env`.

Usa `npm run electron:dev` para probar el enlace Google. `npm run dev` abre Astro en un navegador normal y no tiene acceso a la API segura de Electron. El proceso principal abre el navegador del sistema, escucha en un puerto loopback temporal, valida `state`, completa PKCE y guarda el refresh token con `safeStorage` en el directorio de datos de Triade del sistema operativo. En Windows, Electron cifra el almacén con DPAPI. El renderer solo recibe access tokens temporales.

## Migración

Si usas Supabase para sincronizar los datos de la app, aplica `supabase/migrations/20261002000000_local_oauth_provider.sql` para guardar el proveedor y retirar los tokens Gmail antiguos de `user_accounts`. Google OAuth ya no necesita Edge Function, OAuth client secret ni secretos de cifrado de Supabase.

Desvincula y vuelve a vincular cada cuenta Gmail para crear su credencial local cifrada. Las credenciales antiguas no se migran automáticamente.# Vinculación local de Gmail (Electron)

## Google Cloud Console

1. Selecciona el proyecto de Google y habilita **Gmail API**.
2. Configura la pantalla de consentimiento OAuth. Mientras la app esté en modo **Testing**, añade las cuentas de Gmail que vas a usar como test users.
3. Crea credenciales OAuth con tipo de aplicación **Desktop app**. Copia su client ID.
4. No crees ni distribuyas un `client_secret` para la aplicación de escritorio. El flujo usa Authorization Code + PKCE y un callback loopback temporal en `127.0.0.1`; Google no necesita un callback web fijo para el cliente Desktop.
5. Autoriza los scopes que muestra la pantalla de consentimiento:
   - `https://www.googleapis.com/auth/gmail.modify`: leer, destacar y mover correos entre Inbox, Spam y Papelera.
   - `https://www.googleapis.com/auth/gmail.send`: enviar correo.
   - `https://www.googleapis.com/auth/userinfo.email`: identificar la cuenta conectada.
   - Triade actualmente también solicita `https://mail.google.com/` para eliminación definitiva. Es un scope restringido que puede exigir verificación y evaluación de seguridad antes de distribuir públicamente. Si no necesitas eliminar mensajes definitivamente, quítalo de `GOOGLE_OAUTH_SCOPES` en `src/scripts/app.ts`.

## Configuración local

En el `.env` de la raíz añade el client ID público:

```dotenv
PUBLIC_GOOGLE_CLIENT_ID=TU_CLIENT_ID.apps.googleusercontent.com
```

El ID no es secreto. No añadas `client_secret`, claves de servicio de Supabase ni refresh tokens al `.env`.

La vinculación OAuth de Gmail usa Electron y el navegador del sistema. Arranca el proyecto con `npm run electron:dev` (no con `npm run dev` en un navegador normal). El proceso principal crea un puerto local temporal, comprueba `state`, completa PKCE y conserva el refresh token cifrado mediante `safeStorage` en el directorio de datos de Triade del sistema operativo. El renderer recibe solo access tokens temporales. En Windows, Electron cifra ese almacén con DPAPI.

## Migración de datos existente

Aplica `supabase/migrations/20261002000000_local_oauth_provider.sql` a tu proyecto Supabase para guardar el proveedor y retirar refresh tokens Gmail de la tabla antigua `user_accounts`. La aplicación sigue usando Supabase para los datos ya sincronizados, pero Google OAuth no requiere Edge Function, OAuth client secret ni secretos de token en Supabase.

Después de migrar, desvincula y vuelve a vincular cada cuenta Gmail. Las credenciales implícitas antiguas no se convierten automáticamente. Si Google muestra `redirect_uri_mismatch`, confirma que la credencial es de tipo **Desktop app**, no **Web application**.