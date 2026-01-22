// src/main/preload.ts
import { contextBridge, ipcRenderer } from "electron";

type AppState = {
  testMode: boolean;
};

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

contextBridge.exposeInMainWorld("api", {
  // --- App state ---
  getState: async (): Promise<AppState> => {
    return await ipcRenderer.invoke("app:getState");
  },

  setTestMode: async (value: boolean): Promise<AppState> => {
    return await ipcRenderer.invoke("app:setTestMode", value);
  },

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
  }
});

// TypeScript 側で window.api を認識させたい場合の型拡張（任意）
declare global {
  interface Window {
    api: {
      getState: () => Promise<AppState>;
      setTestMode: (value: boolean) => Promise<AppState>;
      getPrompt: (stage: string) => Promise<PromptGetResult>;
      exportEstimateXlsx: (payload: unknown) => Promise<EstimateExportResult>;
    };
  }
}
