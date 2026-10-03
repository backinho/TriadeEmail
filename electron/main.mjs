import { app, BrowserWindow, ipcMain, safeStorage, shell } from 'electron';
import updaterPackage from 'electron-updater';
import { createServer } from 'node:http';
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto';
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const { autoUpdater } = updaterPackage;
const appDirectory = path.dirname(fileURLToPath(import.meta.url));
const webDirectory = path.join(appDirectory, '..', 'web-dist');
const googleTokenFile = path.join(app.getPath('userData'), 'oauth', 'google-tokens.enc');
const allowedGoogleScopes = new Set([
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://mail.google.com/',
]);
const mimeTypes = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.jpeg': 'image/jpeg',
  '.jpg': 'image/jpeg',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp',
  '.woff2': 'font/woff2',
};

function assertTrustedRenderer(event) {
  const url = new URL(event.senderFrame.url);
  if (url.protocol !== 'http:' || !['localhost', '127.0.0.1'].includes(url.hostname) || url.port !== '4321') {
    throw new Error('Solicitud OAuth rechazada desde un origen no confiable.');
  }
}

async function readGoogleTokens() {
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error('El almacén seguro del sistema no está disponible; no se guardó la cuenta de Google.');
  }
  try {
    const encrypted = await readFile(googleTokenFile);
    return JSON.parse(safeStorage.decryptString(encrypted));
  } catch (error) {
    if (error?.code === 'ENOENT') return {};
    throw error;
  }
}

async function writeGoogleTokens(tokens) {
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error('El almacén seguro del sistema no está disponible; no se guardó la cuenta de Google.');
  }
  await mkdir(path.dirname(googleTokenFile), { recursive: true });
  const temporaryFile = `${googleTokenFile}.${process.pid}.tmp`;
  await writeFile(temporaryFile, safeStorage.encryptString(JSON.stringify(tokens)), { mode: 0o600 });
  await rename(temporaryFile, googleTokenFile);
}

function makePkcePair() {
  const verifier = randomBytes(32).toString('base64url');
  const challenge = createHash('sha256').update(verifier).digest('base64url');
  return { verifier, challenge };
}

function receiveGoogleCallback(server, expectedState, timeout) {
  return new Promise((resolve, reject) => {
    const finish = (error, result) => {
      clearTimeout(timeout);
      server.close();
      if (error) reject(error);
      else resolve(result);
    };

    server.on('request', (request, response) => {
      const url = new URL(request.url, 'http://127.0.0.1');
      if (url.pathname !== '/') {
        response.writeHead(404).end('Not found');
        return;
      }

      const state = url.searchParams.get('state') || '';
      const expected = Buffer.from(expectedState);
      const received = Buffer.from(state);
      if (expected.length !== received.length || !timingSafeEqual(expected, received)) {
        response.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' }).end('OAuth state validation failed. You can close this window.');
        finish(new Error('Google OAuth state validation failed.'));
        return;
      }

      const error = url.searchParams.get('error');
      const code = url.searchParams.get('code');
      response.writeHead(error || !code ? 400 : 200, { 'Content-Type': 'text/html; charset=utf-8' });
      response.end('<!doctype html><meta charset="utf-8"><title>Triade Mail</title><p>Autorización recibida. Puedes volver a Triade Mail.</p>');
      if (error || !code) finish(new Error(`Google authorization failed: ${error || 'missing code'}`));
      else finish(null, code);
    });

    timeout = setTimeout(() => finish(new Error('Google OAuth timed out.')), 5 * 60 * 1000);
  });
}

ipcMain.handle('google-oauth:start', async (event, { clientId, clientSecret, scopes }) => {
  assertTrustedRenderer(event);
  if (!clientId || typeof clientId !== 'string' || !clientId.endsWith('.apps.googleusercontent.com')) {
    throw new Error('Configura un OAuth client ID de tipo Desktop app.');
  }
  if (typeof clientSecret !== 'string' || !clientSecret.trim()) {
    throw new Error('Falta PUBLIC_GOOGLE_CLIENT_SECRET. Cópialo del JSON de credenciales OAuth Desktop de Google Cloud.');
  }
  if (!Array.isArray(scopes) || scopes.some((scope) => typeof scope !== 'string' || !allowedGoogleScopes.has(scope))) {
    throw new Error('Scopes de Google inválidos.');
  }
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error('El almacén seguro del sistema no está disponible. No se puede guardar la cuenta.');
  }

  const { verifier, challenge } = makePkcePair();
  const state = randomBytes(32).toString('base64url');
  const server = createServer();
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('No se pudo abrir el callback local de Google.');
  const redirectUri = `http://127.0.0.1:${address.port}/`;
  let timeout;
  const callbackCode = receiveGoogleCallback(server, state, timeout);
  callbackCode.catch(() => {});
  const authorizationUrl = new URL('https://accounts.google.com/o/oauth2/v2/auth');
  authorizationUrl.search = new URLSearchParams({
    client_id: clientId,
    redirect_uri: redirectUri,
    response_type: 'code',
    scope: scopes.join(' '),
    access_type: 'offline',
    prompt: 'consent',
    code_challenge: challenge,
    code_challenge_method: 'S256',
    state,
  }).toString();

  try {
    await shell.openExternal(authorizationUrl.toString());
    const code = await callbackCode;
    const tokenResponse = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret.trim(),
        code,
        code_verifier: verifier,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
      }),
    });
    const tokenData = await tokenResponse.json();
    if (!tokenResponse.ok || !tokenData.access_token) {
      throw new Error(tokenData.error_description || tokenData.error || 'Google no pudo completar el canje OAuth.');
    }

    const profileResponse = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${tokenData.access_token}` },
    });
    const profile = profileResponse.ok ? await profileResponse.json() : null;
    const email = String(profile?.email || '').trim().toLowerCase();
    if (!email) throw new Error('No se pudo obtener el correo de Google.');

    const tokens = await readGoogleTokens();
    if (tokenData.refresh_token) tokens[email] = tokenData.refresh_token;
    if (!tokens[email]) throw new Error('Google no entregó refresh token. Revoca el acceso de Triade y vuelve a autorizar.');
    await writeGoogleTokens(tokens);

    return {
      account_email: email,
      access_token: tokenData.access_token,
      expires_in: Number(tokenData.expires_in) || 3600,
    };
  } catch (error) {
    server.close();
    throw error;
  }
});

ipcMain.handle('google-oauth:refresh', async (event, email, clientId, clientSecret) => {
  assertTrustedRenderer(event);
  const accountEmail = String(email || '').trim().toLowerCase();
  const tokens = await readGoogleTokens();
  const refreshToken = tokens[accountEmail];
  if (!refreshToken) throw new Error('La cuenta de Google debe volver a vincularse.');

  if (typeof clientId !== 'string' || !clientId.endsWith('.apps.googleusercontent.com')) {
    throw new Error('Falta PUBLIC_GOOGLE_CLIENT_ID en la configuración de la app.');
  }
  if (typeof clientSecret !== 'string' || !clientSecret.trim()) {
    throw new Error('Falta PUBLIC_GOOGLE_CLIENT_SECRET en la configuración de la app.');
  }
  const response = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret.trim(),
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
    }),
  });
  const data = await response.json();
  if (!response.ok || !data.access_token) throw new Error(data.error_description || 'No se pudo renovar el token de Google.');
  if (data.refresh_token) {
    tokens[accountEmail] = data.refresh_token;
    await writeGoogleTokens(tokens);
  }
  return { access_token: data.access_token, expires_in: Number(data.expires_in) || 3600 };
});

ipcMain.handle('google-oauth:unlink', async (event, email) => {
  assertTrustedRenderer(event);
  const accountEmail = String(email || '').trim().toLowerCase();
  const tokens = await readGoogleTokens();
  delete tokens[accountEmail];
  await writeGoogleTokens(tokens);
  return { ok: true };
});

function startWebServer() {
  const server = createServer(async (request, response) => {
    try {
      const url = new URL(request.url, 'http://localhost');
      const requestedPath = decodeURIComponent(url.pathname);
      let filePath = path.resolve(webDirectory, `.${requestedPath}`);

      if (!filePath.startsWith(`${path.resolve(webDirectory)}${path.sep}`) && filePath !== path.resolve(webDirectory)) {
        response.writeHead(403).end('Forbidden');
        return;
      }

      let fileInfo = await stat(filePath);
      if (fileInfo.isDirectory()) {
        filePath = path.join(filePath, 'index.html');
        fileInfo = await stat(filePath);
      }

      const content = await readFile(filePath);
      response.writeHead(200, {
        'Content-Length': content.length,
        'Content-Type': mimeTypes[path.extname(filePath)] || 'application/octet-stream',
      });
      response.end(content);
    } catch {
      response.writeHead(404).end('Not found');
    }
  });

  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(4321, '127.0.0.1', () => resolve(server));
  });
}

async function createWindow() {
  const server = await startWebServer();
  const window = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(appDirectory, 'preload.cjs'),
    },
  });

  window.on('closed', () => server.close());
  await window.loadURL('http://localhost:4321/');

  if (app.isPackaged) {
    autoUpdater.checkForUpdatesAndNotify();
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

autoUpdater.autoInstallOnAppQuit = true;