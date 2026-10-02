import { app, BrowserWindow } from 'electron';
import updaterPackage from 'electron-updater';
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const { autoUpdater } = updaterPackage;
const appDirectory = path.dirname(fileURLToPath(import.meta.url));
const webDirectory = path.join(appDirectory, '..', 'web-dist');
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