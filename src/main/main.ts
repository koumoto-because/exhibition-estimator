import { app, BrowserWindow, ipcMain } from "electron";
import * as path from "path";
import * as fs from "fs";
import * as dotenv from "dotenv";

const APP_ROOT = app.getAppPath();
dotenv.config({ path: path.join(APP_ROOT, ".env") });

let mainWindow: BrowserWindow | null = null;
let isTestMode = process.env.APP_MODE === "test";

function getPromptFilename(stage: string) {
  switch (stage) {
    case "extract":
      return "01_extract.txt";
    case "update":
      return "02_update.txt";
    case "estimate":
      return "03_estimate.txt";
    default:
      return "01_extract.txt";
  }
}

function readTextSafe(p: string): string | null {
  try {
    return fs.readFileSync(p, "utf-8");
  } catch {
    return null;
  }
}

function loadPrompt(stage: string) {
  const promptsDir = path.join(APP_ROOT, "prompts");
  const promptFile = getPromptFilename(stage);
  const promptPath = path.join(promptsDir, promptFile);

  const raw = readTextSafe(promptPath);
  if (!raw) {
    return {
      stage,
      prompt: `【エラー】プロンプトファイルが読み込めませんでした: ${promptPath}`,
      debug: { promptPath }
    };
  }

  const materialPath = path.join(promptsDir, "materialInfo.json");
  const laborPath = path.join(promptsDir, "laborInfo.json");
  const schemaPath = path.join(promptsDir, "jsonSchema.json");

  const materialInfo = readTextSafe(materialPath);
  const laborInfo = readTextSafe(laborPath);
  const jsonSchema = readTextSafe(schemaPath);

  const replace = (
    text: string,
    tag: string,
    label: string,
    body: string | null,
    filePath: string
  ) => {
    if (!text.includes(tag)) return text;
    return text.split(tag).join(
      `${label}\n${body ?? `（読み込み失敗: ${filePath}）`}\n`
    );
  };

  let resolved = raw;
  resolved = replace(resolved, "<materialInfo>", "【材料詳細】", materialInfo, materialPath);
  resolved = replace(resolved, "<laborInfo>", "【人工単価・人工計算用係数】", laborInfo, laborPath);
  resolved = replace(resolved, "<jsonSchema>", "【Json Schema】", jsonSchema, schemaPath);

  return {
    stage,
    prompt: resolved,
    debug: {
      promptPath,
      promptsDir,
      materialPath,
      laborPath,
      schemaPath,
      hasMaterialTag: raw.includes("<materialInfo>"),
      hasLaborTag: raw.includes("<laborInfo>"),
      hasSchemaTag: raw.includes("<jsonSchema>"),
      materialExists: !!materialInfo,
      laborExists: !!laborInfo,
      schemaExists: !!jsonSchema
    }
  };
}

function createWindow() {
  const indexHtmlPath = path.join(APP_ROOT, "src", "renderer", "index.html");
  const preloadPath = path.join(APP_ROOT, "dist", "main", "preload.js");

  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: preloadPath,
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile(indexHtmlPath);
}

app.whenReady().then(() => {
  createWindow();
});

ipcMain.handle("app:getState", async () => ({ testMode: isTestMode }));

ipcMain.handle("app:setTestMode", async (_e, v) => {
  isTestMode = !!v;
  return { testMode: isTestMode };
});

ipcMain.handle("prompt:get", async (_e, stage) => {
  console.log("[ipc] prompt:get", stage);
  return loadPrompt(stage);
});
