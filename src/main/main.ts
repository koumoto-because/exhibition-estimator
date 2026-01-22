import { app, BrowserWindow, ipcMain, dialog } from "electron";
import * as path from "path";
import * as fs from "fs";
import * as dotenv from "dotenv";
import ExcelJS from "exceljs";

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
  const schemaFilename = stage === "estimate" ? "jsonSchema_estimate.json" : "jsonSchema_extract.json";
  const schemaPath = path.join(promptsDir, schemaFilename);

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

ipcMain.handle("estimate:exportXlsx", async (_e, payload) => {
  try {
    const templatePath = path.join(APP_ROOT, "refs", "見積もりシート.xlsx");
    if (!fs.existsSync(templatePath)) {
      return { ok: false, error: `テンプレートが見つかりません: ${templatePath}` };
    }

    const defaultName = "estimate.xlsx";
    const saveResult = await dialog.showSaveDialog({
      title: "見積書xlsxを保存",
      defaultPath: path.join(APP_ROOT, "refs", defaultName),
      filters: [{ name: "Excel", extensions: ["xlsx"] }]
    });
    if (saveResult.canceled || !saveResult.filePath) {
      return { ok: false, error: "保存がキャンセルされました。" };
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const sheet = workbook.getWorksheet(1) || workbook.worksheets[0];
    if (!sheet) {
      return { ok: false, error: "テンプレートのシートが見つかりません。" };
    }

    const estimate = (payload && typeof payload === "object" ? payload : {}) as any;
    const clientName = typeof estimate.client?.name === "string" ? estimate.client.name : "";
    const totalAmount = typeof estimate.total?.amount === "number" ? estimate.total.amount : 0;
    const breakdownTitle = typeof estimate.breakdown?.title === "string" ? estimate.breakdown.title : "";
    const taxIncluded = !!estimate.total?.tax?.included;

    sheet.getCell("A19").value = breakdownTitle || "工事内訳明細表";
    sheet.getCell("C11").value = typeof estimate.title === "string" ? estimate.title : "";
    sheet.getCell("A3").value = clientName;
    sheet.getCell("A7").value = totalAmount;
    sheet.getCell("E9").value = taxIncluded ? "※消費税は、含まれております。" : "※消費税は、含まれておりません。";

    const rows: Array<{
      groupNo: number | null;
      categoryName: string;
      description: string;
      quantity: number | null;
      unit: string;
      unitPrice: number | null;
      amount: number | string | null;
    }> = [];

    const groups = Array.isArray(estimate.breakdown?.groups) ? estimate.breakdown.groups : [];
    groups.forEach((group: any, groupIdx: number) => {
      const groupNo = typeof group.groupNo === "number" ? group.groupNo : null;
      const categoryName = typeof group.categoryName === "string" ? group.categoryName : "";
      const noteLines = Array.isArray(group.noteLines)
        ? group.noteLines.filter((line: any) => typeof line === "string" && line.trim())
        : [];
      const groupTotal =
        group.displayMode === "detailed" && Array.isArray(group.lineItems)
          ? group.lineItems.reduce((sum: number, line: any) => {
              const amt = typeof line.amount === "number" ? line.amount : 0;
              return sum + amt;
            }, 0)
          : typeof group.groupSummaryLine?.amount === "number"
            ? group.groupSummaryLine.amount
            : 0;
      const groupTotalDisplay = `<<${groupTotal}>>`;

      if (group.displayMode === "detailed" && Array.isArray(group.lineItems) && group.lineItems.length > 0) {
        rows.push({
          groupNo,
          categoryName,
          description: "",
          quantity: null,
          unit: "",
          unitPrice: null,
          amount: null
        });
        if (noteLines.length) {
          rows.push({
            groupNo: null,
            categoryName: `（${noteLines.join(" / ")}）`,
            description: "",
            quantity: null,
            unit: "",
            unitPrice: null,
            amount: null
          });
        }
        group.lineItems.forEach((line: any) => {
          const qty = typeof line.quantity === "number" ? line.quantity : null;
          const unit = typeof line.unit === "string" ? line.unit : "";
          const unitPrice =
            unit === "式" ? null : typeof line.unitPrice === "number" ? line.unitPrice : null;
          const amount =
            typeof line.amount === "number" ? line.amount : qty != null && unitPrice != null ? qty * unitPrice : null;
          const desc = typeof line.description === "string" ? line.description : "";
          rows.push({
            groupNo: null,
            categoryName: "",
            description: desc,
            quantity: qty,
            unit,
            unitPrice,
            amount
          });
        });
      } else {
        const summary = group.groupSummaryLine || {};
        const qty = typeof summary.quantity === "number" ? summary.quantity : null;
        const unit = typeof summary.unit === "string" ? summary.unit : "";
        const unitPrice = unit === "式" ? null : typeof summary.unitPrice === "number" ? summary.unitPrice : null;
        rows.push({
          groupNo,
          categoryName,
          description: "",
          quantity: qty,
          unit,
          unitPrice,
          amount: typeof summary.amount === "number" ? summary.amount : null
        });
        if (noteLines.length) {
          rows.push({
            groupNo: null,
            categoryName: `（${noteLines.join(" / ")}）`,
            description: "",
            quantity: null,
            unit: "",
            unitPrice: null,
            amount: null
          });
        }
      }
      if (groupIdx < groups.length - 1) {
        rows.push({
          groupNo: null,
          categoryName: "",
          description: "",
          quantity: null,
          unit: "",
          unitPrice: null,
          amount: null
        });
      }
    });

    const startRow = 21;
    const endRow = 51;
    for (let r = startRow; r <= endRow; r += 1) {
      sheet.getCell(`A${r}`).value = null;
      sheet.getCell(`B${r}`).value = null;
      sheet.getCell(`D${r}`).value = null;
      sheet.getCell(`G${r}`).value = null;
      sheet.getCell(`H${r}`).value = null;
      sheet.getCell(`I${r}`).value = null;
      sheet.getCell(`J${r}`).value = null;
    }

    rows.slice(0, endRow - startRow + 1).forEach((line, idx) => {
      const row = startRow + idx;
      if (line.groupNo != null) sheet.getCell(`A${row}`).value = line.groupNo;
      if (line.categoryName) sheet.getCell(`B${row}`).value = line.categoryName;
      sheet.getCell(`D${row}`).value = line.description || "";
      if (line.quantity != null) sheet.getCell(`G${row}`).value = line.quantity;
      if (line.unit) sheet.getCell(`H${row}`).value = line.unit;
      if (line.unitPrice != null) sheet.getCell(`I${row}`).value = line.unitPrice;
      if (line.amount != null) sheet.getCell(`J${row}`).value = line.amount;
    });

    await workbook.xlsx.writeFile(saveResult.filePath);
    return { ok: true, path: saveResult.filePath };
  } catch (err: any) {
    console.error(err);
    return { ok: false, error: "xlsx出力に失敗しました。" };
  }
});
