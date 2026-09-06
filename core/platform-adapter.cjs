"use strict";

class BrowserPlatformAdapter {
  constructor() {
    this.kind = "browser";
    this.contractVersion = "1.0";
  }

  capabilities() {
    return {
      runtime: "browser-pwa",
      filesystem_directory:
        typeof globalThis.showDirectoryPicker === "function",
      secure_keystore: false,
      native_scheduler: false,
      atomic_file_replace: false,
    };
  }

  async permission(handle, request = false) {
    if (!handle?.queryPermission) return false;
    let result = await handle.queryPermission({ mode: "readwrite" });
    if (result !== "granted" && request && handle.requestPermission)
      result = await handle.requestPermission({ mode: "readwrite" });
    return result === "granted";
  }

  async writeFile(handle, name, blob, requestPermission = false) {
    if (!(await this.permission(handle, requestPermission)))
      throw new Error("Chưa có quyền ghi vào thư mục sao lưu.");
    const fileHandle = await handle.getFileHandle(name, { create: true });
    const writable = await fileHandle.createWritable();
    try {
      await writable.write(blob);
      await writable.close();
    } catch (error) {
      await writable.abort?.();
      throw error;
    }
  }
}

module.exports = { BrowserPlatformAdapter };
