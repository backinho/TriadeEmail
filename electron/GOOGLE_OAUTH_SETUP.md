# Vinculacion local de Gmail (Electron)

## Google Cloud Console

1. Selecciona el proyecto Google y habilita **Gmail API**.
2. Configura la pantalla de consentimiento OAuth. Si la app esta en **Testing**, anade como test users las cuentas que la usaran.
3. Crea credenciales OAuth de tipo **Desktop app** y descarga el JSON de credenciales. Copia `client_id` y `client_secret` de ese archivo.
4. El flujo usa Authorization Code + PKCE y un callback loopback temporal en `127.0.0.1`. Google requiere enviar tambien `client_secret` al canjear el codigo y renovar el token.
5. Autoriza estos scopes:
   - `https://www.googleapis.com/auth/gmail.modify`: leer, destacar y mover correos entre Inbox, Spam y Papelera.
   - `https://www.googleapis.com/auth/gmail.send`: enviar correo.
   - `https://www.googleapis.com/auth/userinfo.email`: identificar la cuenta.
   - Triade tambien solicita `https://mail.google.com/` porque permite eliminar mensajes definitivamente. Es restringido y puede requerir revision de Google antes de distribuir la app publicamente. Si no necesitas eliminar definitivamente, quitale ese scope a `GOOGLE_OAUTH_SCOPES` en `src/scripts/app.ts`.

## Configuracion local

En el `.env` de la raiz configura ambos valores de la credencial **Desktop app**:

```dotenv
PUBLIC_GOOGLE_CLIENT_ID=TU_CLIENT_ID.apps.googleusercontent.com
PUBLIC_GOOGLE_CLIENT_SECRET=TU_CLIENT_SECRET_DEL_JSON_DESKTOP
```

El ID es publico. El `client_secret` de OAuth Desktop tampoco puede mantenerse confidencial en una app instalada; se incluye en el paquete y Google lo usa junto con PKCE. No uses el secreto de una credencial **Web application**, ni incluyas claves de servicio de Supabase o refresh tokens en `.env`.

Usa `npm run electron:dev` para probar Gmail; `npm run dev` abre Astro en un navegador normal, sin la API IPC de Electron. Al cambiar `.env`, cierra Triade, reconstruye y reinstala la app para incorporar los valores. Electron completa el callback, verifica `state`, canjea PKCE y guarda el refresh token con `safeStorage` en el almacenamiento del sistema operativo. En Windows ese almacenamiento usa DPAPI; el renderer solo recibe access tokens temporales.

## Supabase existente

Si Triade sincroniza sus datos con Supabase, aplica `supabase/migrations/20261002000000_local_oauth_provider.sql` para registrar el proveedor y retirar tokens Gmail heredados de `user_accounts`. El OAuth local no requiere Edge Function ni secretos de servidor en Supabase.

Despues de actualizar, desvincula Gmail en Triade y vuelve a autorizar la cuenta. Si la credencial no se renueva, comprueba que el client ID y el client secret provienen del mismo JSON Desktop y que la cuenta esta permitida en la pantalla de consentimiento.
