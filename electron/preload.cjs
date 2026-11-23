const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("liteDesktop", {
  isDesktop: true,
  openExternal: (url) => ipcRenderer.invoke("lite-open-external", url),
  saveFile: (payload) => ipcRenderer.invoke("lite-save-file", payload),
  showItemInFolder: (filePath) => ipcRenderer.invoke("lite-show-item-in-folder", filePath),
  onAuthCallback: (callback) => {
    const handler = (_event, url) => callback(url);
    ipcRenderer.on("lite-auth-callback", handler);
    return () => ipcRenderer.removeListener("lite-auth-callback", handler);
  },
});
