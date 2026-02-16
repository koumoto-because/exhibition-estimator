// src/main/preload.ts
import { contextBridge, ipcRenderer } from "electron";

type PromptGetResult = {
  stage: string;
  prompt: string;
  debug?: {
    promptPath?: string;
    promptsDir?: string;
    materialPath?: string;
    laborPath?: string;
    schemaPath?: string;
    hasMaterialTag?: boolean;
    hasLaborTag?: boolean;
    hasSchemaTag?: boolean;
    materialExists?: boolean;
    laborExists?: boolean;
    schemaExists?: boolean;
  };
  error?: string;
};

type EstimateExportResult = {
  ok: boolean;
  path?: string;
  error?: string;
};

type StateSaveResult = {
  ok: boolean;
  path?: string;
  error?: string;
};

type StateOpenResult = {
  ok: boolean;
  text?: string;
  path?: string;
  error?: string;
};

type StateSaveRequest = {
  saveAs?: boolean;
};

contextBridge.exposeInMainWorld("api", {
  // --- Prompt ---
  getPrompt: async (stage: string): Promise<PromptGetResult> => {
    // デバッグしやすいようにログは残す（不要なら削除OK）
    console.log("[preload] getPrompt called:", stage);
    const res = (await ipcRenderer.invoke("prompt:get", stage)) as PromptGetResult;
    return res;
  },

  exportEstimateXlsx: async (payload: unknown): Promise<EstimateExportResult> => {
    const res = (await ipcRenderer.invoke("estimate:exportXlsx", payload)) as EstimateExportResult;
    return res;
  },
  saveState: async (data: string, saveAs = false): Promise<StateSaveResult> => {
    const res = (await ipcRenderer.invoke("state:save", { data, saveAs })) as StateSaveResult;
    return res;
  },
  onStateOpen: (cb: (result: StateOpenResult) => void) => {
    ipcRenderer.on("state:openResult", (_e, result: StateOpenResult) => cb(result));
  },
  onStateSaveRequest: (cb: (result: StateSaveRequest) => void) => {
    ipcRenderer.on("state:saveRequest", (_e, result: StateSaveRequest) => cb(result));
  }
});

// TypeScript 側で window.api を認識させたい場合の型拡張（任意）
declare global {
  interface Window {
    api: {
      getPrompt: (stage: string) => Promise<PromptGetResult>;
      exportEstimateXlsx: (payload: unknown) => Promise<EstimateExportResult>;
      saveState: (data: string, saveAs?: boolean) => Promise<StateSaveResult>;
      onStateOpen: (cb: (result: StateOpenResult) => void) => void;
      onStateSaveRequest: (cb: (result: StateSaveRequest) => void) => void;
    };
  }
}
