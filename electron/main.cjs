const { app, BrowserWindow, ipcMain, shell } = require("electron");
const path = require("path");

const PROTOCOL = "liteapp";
let mainWindow = null;
let pendingDeepLinkUrl = null;

if (process.defaultApp) {
  if (process.argv.length >= 2) {
    app.setAsDefaultProtocolClient(PROTOCOL, process.execPath, [path.resolve(process.argv[1])]);
  }
} else {
  app.setAsDefaultProtocolClient(PROTOCOL);
}

const gotSingleLock = app.requestSingleInstanceLock();
if (!gotSingleLock) {
  app.quit();
} else {
  app.on("second-instance", (_event, commandLine) => {
    const deepLink = commandLine.find((arg) => arg.startsWith(PROTOCOL + "://"));
    if (deepLink) {
      dispatchDeepLink(deepLink);
    }
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

function dispatchDeepLink(url) {
  if (!url) return;
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send("lite-auth-callback", url);
  } else {
    pendingDeepLinkUrl = url;
  }
}

app.on("open-url", (event, url) => {
  event.preventDefault();
  dispatchDeepLink(url);
});

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 960,
    minHeight: 640,
    title: "LITE — Không gian làm việc tài liệu",
    backgroundColor: "#141923",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith("http://") || url.startsWith("https://")) {
      void shell.openExternal(url);
    }
    return { action: "deny" };
  });

  const devUrl = process.env.VITE_DEV_SERVER_URL;
  if (devUrl) {
    void mainWindow.loadURL(devUrl);
  } else {
    void mainWindow.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  }

  mainWindow.webContents.on("did-finish-load", () => {
    const initialArg = process.argv.find((arg) => arg.startsWith(PROTOCOL + "://"));
    const urlToSend = pendingDeepLinkUrl || initialArg;
    if (urlToSend) {
      pendingDeepLinkUrl = null;
      mainWindow.webContents.send("lite-auth-callback", urlToSend);
    }
  });
}

ipcMain.handle("lite-open-external", async (_event, url) => {
  if (typeof url === "string" && (url.startsWith("https://") || url.startsWith("http://"))) {
    await shell.openExternal(url);
  }
});

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
