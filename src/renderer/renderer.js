// src/renderer/renderer.js

// ===== Schema compatibility =====
const SCHEMA_VERSION = "1.1.5";

const CATEGORIES = [
  { key: "woodworkItems", label: "木工造作物" },
  { key: "floorItems", label: "床" },
  { key: "finishingItems", label: "表装" },
  { key: "electricalItems", label: "電気" },
  { key: "leaseItems", label: "リース" },
  { key: "siteCosts", label: "現場費" }
];

let selectedCategoryKey = "woodworkItems";

let basePayload = null;
let currentPayload = null;
let selectedIndex = null;

let userChangedPathsCurrent = new Set();
let userChangedPathsPrevious = new Set();
let apiChangedPaths = new Set();

let stateHistory = [];
let redoHistory = [];

document.addEventListener("DOMContentLoaded", async () => {
  if (!window.api) {
    alert("window.api が見つかりません。preload の読み込みに失敗している可能性があります。");
    console.error("window.api is undefined. Check preload settings.");
    return;
  }

  const testModeToggle = document.getElementById("test-mode-toggle");
  const changeHighlightToggle = document.getElementById("change-highlight-toggle");
  const statusSpan = document.getElementById("status");


  function setStatus(msg) {
    try {
      if (statusSpan) statusSpan.textContent = msg || "";
    } catch (_) {}
    console.log("[status]", msg);
  }


  const stageSelect = document.getElementById("stage-select");
  const btnLoadPrompt = document.getElementById("btn-load-prompt");
  const promptTextarea = document.getElementById("prompt-text");

  const gptOutputTextarea = document.getElementById("gpt-output");
  const btnImportJson = document.getElementById("btn-import-json");
  const btnGenerateDiff = document.getElementById("btn-generate-diff");

  const btnRestoreState = document.getElementById("btn-restore-state");
  const btnRedoState = document.getElementById("btn-redo-state");

  const btnExportStateFile = document.getElementById("btn-export-state-file");
  const btnImportStateFile = document.getElementById("btn-import-state-file");
  const stateFileInput = document.getElementById("state-file-input");

  const categoryTabs = document.getElementById("category-tabs");

  const itemsView = document.getElementById("items-view");
  const itemEditorContainer = document.getElementById("item-editor-container");

  // nlCorrectionGlobal
  const nlGlobalTextarea = document.getElementById("nl-global");

  // siteCosts
  const siteCostsSection = document.getElementById("site-costs-section");
  const siteCostsLaborAmount = document.getElementById("sitecosts-labor-amount");
  const siteCostsLaborNotes = document.getElementById("sitecosts-labor-notes");
  const siteCostsTransportAmount = document.getElementById("sitecosts-transport-amount");
  const siteCostsTransportNotes = document.getElementById("sitecosts-transport-notes");
  const siteCostsWasteAmount = document.getElementById("sitecosts-waste-amount");
  const siteCostsWasteNotes = document.getElementById("sitecosts-waste-notes");

  const itemEditor = document.getElementById("item-editor");
  const itemEditorLabel = document.getElementById("item-editor-label");
  const btnApplyItemEdit = document.getElementById("btn-apply-item-edit");

  // ★ 手動追加（空行追加ボタン）
  const btnAddEmptyItem = document.getElementById("btn-add-empty-item");

  // フォーム
  const itemFormId = document.getElementById("item-form-id");
  const itemFormInclude = document.getElementById("item-form-include");
  const itemFormName = document.getElementById("item-form-name");
  const itemFormStructureType = document.getElementById("item-form-structure-type");
  const itemFormDimH = document.getElementById("item-form-dim-h");
  const itemFormDimW = document.getElementById("item-form-dim-w");
  const itemFormDimD = document.getElementById("item-form-dim-d");
  const itemFormQuantity = document.getElementById("item-form-quantity");
  const itemFormMaterials = document.getElementById("item-form-materials");
  const itemFormFinishes = document.getElementById("item-form-finishes");
  const itemFormPriceUnitPrice = document.getElementById("item-form-price-unitprice");
  const itemFormMaterialCostAmount = document.getElementById("item-form-materialcost-amount");
  const itemFormMaterialCostNotes = document.getElementById("item-form-materialcost-notes");
  const itemFormLaborPerUnitDays = document.getElementById("item-form-labor-perunit-days");
  const itemFormLaborDays = document.getElementById("item-form-labor-days");
  const itemFormCostAmount = document.getElementById("item-form-cost-amount");
  const itemFormCostNotes = document.getElementById("item-form-cost-notes");

  const itemFormIsBent = document.getElementById("item-form-isbent");
  const itemFormSpecialAngles = document.getElementById("item-form-special-angles");
  const itemFormSupportHeavy = document.getElementById("item-form-support-heavy");
  const itemFormFinishAreaValue = document.getElementById("item-form-finisharea-value");
  const itemFormFinishAreaNotes = document.getElementById("item-form-finisharea-notes");
  const itemFormLaborCoefValue = document.getElementById("item-form-laborcoef-value");
  const itemFormLaborCoefNotes = document.getElementById("item-form-laborcoef-notes");
  // ★ nlCorrection は詳細設定から外へ移動済み
  const itemFormNlCorrection = document.getElementById("item-form-nlcorrection");

  // タブ
  const tabForm = document.getElementById("tab-form");
  const tabJson = document.getElementById("tab-json");
  const tabContentForm = document.getElementById("tab-content-form");
  const tabContentJson = document.getElementById("tab-content-json");

  const woodworkOnlyEls = document.querySelectorAll("[data-woodwork-only]");
  const genericOnlyEls = document.querySelectorAll("[data-generic-only]");




// ===== Source filename banner (show at a glance) =====
function getSourceFilename(payload) {
  if (!payload) return "";
  const sd = payload.sourceDocument || {};
  return sd.filename || sd.fileName || sd.name || "";
}

function getSourceMetaText(payload) {
  if (!payload) return "図面ファイル:（未読み込み）";
  const sd = payload.sourceDocument || {};
  const fn = getSourceFilename(payload) || "（未設定）";
  const page = typeof sd.page === "number" ? ` / page:${sd.page}` : "";
  return `図面ファイル: ${fn}${page}`;
}

function ensureSourceInfoEl() {
  let el = document.getElementById("source-info");
  if (el) return el;

  el = document.createElement("div");
  el.id = "source-info";
  el.style.padding = "6px 10px";
  el.style.margin = "6px 0";
  el.style.border = "1px solid #ddd";
  el.style.borderRadius = "6px";
  el.style.fontWeight = "600";
  el.style.fontSize = "13px";
  el.style.background = "#f8f8f8";
  el.style.whiteSpace = "nowrap";
  el.style.overflow = "hidden";
  el.style.textOverflow = "ellipsis";

  // Prefer placing it near the status line if possible
  if (statusSpan && statusSpan.parentElement) {
    statusSpan.parentElement.insertBefore(el, statusSpan.parentElement.firstChild);
  } else if (itemsView && itemsView.parentElement) {
    itemsView.parentElement.insertBefore(el, itemsView);
  } else {
    document.body.insertBefore(el, document.body.firstChild);
  }
  return el;
}

function updateSourceInfoUI() {
  const el = ensureSourceInfoEl();
  const p = currentPayload || basePayload;
  el.textContent = getSourceMetaText(p);
  el.title = getSourceFilename(p) || "";
}

  function isHighlightEnabled() {
    return changeHighlightToggle ? changeHighlightToggle.checked : true;
  }

  function updateRestoreButtonEnabled() {
    if (btnRestoreState) btnRestoreState.disabled = stateHistory.length === 0;
  }
  function updateRedoButtonEnabled() {
    if (btnRedoState) btnRedoState.disabled = redoHistory.length === 0;
  }

  function setFormModeForCategory() {
    const isWoodwork = selectedCategoryKey === "woodworkItems";
    const isSiteCosts = selectedCategoryKey === "siteCosts";
    woodworkOnlyEls.forEach((el) => {
      el.hidden = !isWoodwork;
    });
    genericOnlyEls.forEach((el) => {
      el.hidden = isWoodwork;
    });
    if (siteCostsSection) siteCostsSection.hidden = !isSiteCosts;
    if (btnAddEmptyItem) btnAddEmptyItem.disabled = isSiteCosts;
    if (itemEditorContainer) itemEditorContainer.hidden = isSiteCosts;
  }

  function selectCategory(key) {
    if (!CATEGORIES.some((c) => c.key === key)) return;
    selectedCategoryKey = key;
    selectedIndex = null;
    if (itemEditor) itemEditor.value = "";
    resetForm();
    setFormModeForCategory();
    renderItemsTable();
    syncSiteCostsToUI();
    updateGlobalHighlight();
    updateFormHighlightsForSelected();
  }

  function renderCategoryTabs() {
    if (!categoryTabs) return;
    categoryTabs.innerHTML = "";
    CATEGORIES.forEach((cat) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "tab-btn" + (cat.key === selectedCategoryKey ? " active" : "");
      btn.textContent = cat.label;
      btn.addEventListener("click", () => {
        if (selectedCategoryKey === cat.key) return;
        selectCategory(cat.key);
        renderCategoryTabs();
      });
      categoryTabs.appendChild(btn);
    });
  }

  function buildStateSnapshot() {
    return {
      timestamp: new Date().toISOString(),
      basePayload: basePayload ? deepClone(basePayload) : null,
      currentPayload: currentPayload ? deepClone(currentPayload) : null,
      selectedIndex,
      selectedCategoryKey,
      userChangedPathsCurrent: Array.from(userChangedPathsCurrent),
      userChangedPathsPrevious: Array.from(userChangedPathsPrevious),
      apiChangedPaths: Array.from(apiChangedPaths),
      testMode: typeof testModeToggle?.checked === "boolean" ? testModeToggle.checked : null
    };
  }

  function pushStateHistory(clearRedo = true) {
    stateHistory.push(buildStateSnapshot());
    if (stateHistory.length > 5) stateHistory.shift();
    if (clearRedo) redoHistory = [];
    updateRestoreButtonEnabled();
    updateRedoButtonEnabled();
  }

  function applySnapshot(snap) {
    basePayload = snap.basePayload ? deepClone(snap.basePayload) : null;
    currentPayload = snap.currentPayload ? deepClone(snap.currentPayload) : null;
    selectedIndex = typeof snap.selectedIndex === "number" ? snap.selectedIndex : null;
    if (typeof snap.selectedCategoryKey === "string") selectedCategoryKey = snap.selectedCategoryKey;

    userChangedPathsCurrent = new Set(snap.userChangedPathsCurrent || []);
    userChangedPathsPrevious = new Set(snap.userChangedPathsPrevious || []);
    apiChangedPaths = new Set(snap.apiChangedPaths || []);

    if (typeof snap.testMode === "boolean" && testModeToggle) testModeToggle.checked = snap.testMode;

    // schema補正（v1.1.2）
    normalizePayloadForSchema(currentPayload);
    normalizePayloadForSchema(basePayload);

    // nlCorrectionGlobal の復元
    syncNlGlobalToUI();
    updateSourceInfoUI();
    renderCategoryTabs();

    renderItemsTable();
    resetForm();

    const items = getItems();
    if (selectedIndex != null && selectedIndex >= 0 && selectedIndex < items.length) {
      const item = items[selectedIndex];
      updateFormFromItem(item);
      if (itemEditor) itemEditor.value = JSON.stringify(item, null, 2);
      if (itemEditorLabel) itemEditorLabel.textContent = `ID: ${item.id || ""} / index: ${selectedIndex}`;
    } else {
      if (itemEditor) itemEditor.value = "";
      if (itemEditorLabel) itemEditorLabel.textContent = "（行をクリックするとここに表示されます）";
    }

    updateGlobalHighlight();
    updateFormHighlightsForSelected();
    setFormModeForCategory();
  }

  function restorePreviousState() {
    if (stateHistory.length === 0) {
      alert("復元できる以前の状態がありません。");
      return;
    }
    redoHistory.push(buildStateSnapshot());
    if (redoHistory.length > 5) redoHistory.shift();

    const snap = stateHistory.pop();
    applySnapshot(snap);

    updateRestoreButtonEnabled();
    updateRedoButtonEnabled();
    setStatus("以前の状態を復元しました（UNDO）");
  }

  function redoState() {
    if (redoHistory.length === 0) {
      alert("REDO できる状態がありません。");
      return;
    }
    stateHistory.push(buildStateSnapshot());
    if (stateHistory.length > 5) stateHistory.shift();

    const snap = redoHistory.pop();
    applySnapshot(snap);

    updateRestoreButtonEnabled();
    updateRedoButtonEnabled();
    setStatus("ロールバック前の状態に戻しました（REDO）");
  }

  // タブ
  if (tabForm && tabJson && tabContentForm && tabContentJson) {
    tabForm.addEventListener("click", () => {
      tabForm.classList.add("active");
      tabJson.classList.remove("active");
      tabContentForm.classList.remove("hidden");
      tabContentJson.classList.add("hidden");
    });
    tabJson.addEventListener("click", () => {
      tabJson.classList.add("active");
      tabForm.classList.remove("active");
      tabContentJson.classList.remove("hidden");
      tabContentForm.classList.add("hidden");
    });
  }

  // 初期状態
  try {
    const state = await window.api.getState();
    if (state && typeof state.testMode === "boolean" && testModeToggle) {
      testModeToggle.checked = state.testMode;
      setStatus(state.testMode ? "（テストモード）" : "（本番モード・API接続）");
    }
  } catch (err) {
    console.error("getState error:", err);
    setStatus("状態取得に失敗しました");
  }

  updateRestoreButtonEnabled();
  updateRedoButtonEnabled();

  // テストモード切り替え
  if (testModeToggle) {
    testModeToggle.addEventListener("change", async () => {
      const value = testModeToggle.checked;
      try {
        const state = await window.api.setTestMode(value);
        setStatus(state.testMode ? "（テストモード）" : "（本番モード・API接続）");
      } catch (err) {
        console.error("setTestMode error:", err);
        setStatus("テストモード切り替えに失敗しました");
      }
    });
  }

  // 変更ハイライト切り替え
  if (changeHighlightToggle) {
    changeHighlightToggle.addEventListener("change", () => {
      renderItemsTable();
      updateGlobalHighlight();
      updateFormHighlightsForSelected();
    });
  }

  // UNDO/REDO
  if (btnRestoreState) btnRestoreState.addEventListener("click", restorePreviousState);
  if (btnRedoState) btnRedoState.addEventListener("click", redoState);

  // プロンプト読込
  if (btnLoadPrompt && stageSelect && promptTextarea) {
    btnLoadPrompt.addEventListener("click", async () => {
      const stage = stageSelect.value || "extract";
      try {
        const result = await window.api.getPrompt(stage);
        console.log("getPrompt result:", result);       // ★追加
        console.log("getPrompt debug:", result.debug); // ★そのまま
        promptTextarea.value = result.prompt || "";
        setStatus("プロンプトを読み込みました: " + result.stage);
      } catch (err) {
        console.error("getPrompt error:", err);
        setStatus("プロンプトの読み込みに失敗しました");
      }
    });
  }

  // ===== nlCorrectionGlobal: UI -> JSON 即時反映 =====
  if (nlGlobalTextarea) {
    nlGlobalTextarea.addEventListener("input", () => {
      if (!currentPayload) return;
      normalizePayloadForSchema(currentPayload);
      currentPayload.nlCorrectionGlobal = nlGlobalTextarea.value || "";
      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }
      updateGlobalHighlight();
      renderItemsTable();
      setStatus("全体修正指示（nlCorrectionGlobal）を反映しました");
    });
  }

  // ===== siteCosts: UI -> JSON 即時反映 (schema v1.1.2) =====
  function syncSiteCostsToUI() {
    if (!siteCostsLaborAmount || !siteCostsTransportAmount || !siteCostsWasteAmount) return;
    if (!currentPayload) {
      siteCostsLaborAmount.value = "";
      if (siteCostsLaborNotes) siteCostsLaborNotes.value = "";
      siteCostsTransportAmount.value = "";
      if (siteCostsTransportNotes) siteCostsTransportNotes.value = "";
      siteCostsWasteAmount.value = "";
      if (siteCostsWasteNotes) siteCostsWasteNotes.value = "";
      return;
    }

    normalizePayloadForSchema(currentPayload);
    const sc = currentPayload.siteCosts || {};

    const labor = sc.laborCost || {};
    const transport = sc.transportCost || {};
    const waste = sc.wasteDisposalCost || {};

    siteCostsLaborAmount.value = typeof labor.amount === "number" ? String(labor.amount) : "";
    if (siteCostsLaborNotes) siteCostsLaborNotes.value = labor.notes || "";

    siteCostsTransportAmount.value = typeof transport.amount === "number" ? String(transport.amount) : "";
    if (siteCostsTransportNotes) siteCostsTransportNotes.value = transport.notes || "";

    siteCostsWasteAmount.value = typeof waste.amount === "number" ? String(waste.amount) : "";
    if (siteCostsWasteNotes) siteCostsWasteNotes.value = waste.notes || "";
  }

  function applySiteCostsFromUI() {
    if (!currentPayload) return;
    normalizePayloadForSchema(currentPayload);
    const toNum = (v) => {
      const n = parseFloat(String(v || ""));
      return Number.isFinite(n) ? n : null;
    };

    const laborAmount = siteCostsLaborAmount ? toNum(siteCostsLaborAmount.value) : null;
    const transportAmount = siteCostsTransportAmount ? toNum(siteCostsTransportAmount.value) : null;
    const wasteAmount = siteCostsWasteAmount ? toNum(siteCostsWasteAmount.value) : null;

    currentPayload.siteCosts.laborCost.amount = laborAmount ?? 0;
    currentPayload.siteCosts.laborCost.currency = "JPY";
    currentPayload.siteCosts.laborCost.notes = (siteCostsLaborNotes && siteCostsLaborNotes.value) || "";

    currentPayload.siteCosts.transportCost.amount = transportAmount ?? 0;
    currentPayload.siteCosts.transportCost.currency = "JPY";
    currentPayload.siteCosts.transportCost.notes = (siteCostsTransportNotes && siteCostsTransportNotes.value) || "";

    currentPayload.siteCosts.wasteDisposalCost.amount = wasteAmount ?? 0;
    currentPayload.siteCosts.wasteDisposalCost.currency = "JPY";
    currentPayload.siteCosts.wasteDisposalCost.notes = (siteCostsWasteNotes && siteCostsWasteNotes.value) || "";

    if (basePayload && currentPayload) {
      userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
    }

    updateGlobalHighlight();
    renderItemsTable();
    setStatus("現場費（siteCosts）を反映しました");
  }

  [
    siteCostsLaborAmount,
    siteCostsLaborNotes,
    siteCostsTransportAmount,
    siteCostsTransportNotes,
    siteCostsWasteAmount,
    siteCostsWasteNotes
  ].forEach((el) => {
    if (!el) return;
    const evt = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
    el.addEventListener(evt, () => {
      if (!currentPayload) return;
      applySiteCostsFromUI();
    });
  });

  function syncNlGlobalToUI() {
    if (!nlGlobalTextarea) return;
    if (!currentPayload) {
      nlGlobalTextarea.value = "";
      return;
    }
    normalizePayloadForSchema(currentPayload);
    nlGlobalTextarea.value = currentPayload.nlCorrectionGlobal || "";
  }

  function updateGlobalHighlight() {
    if (!nlGlobalTextarea) return;
    nlGlobalTextarea.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
    if (!isHighlightEnabled()) return;

    const prefix = "/nlCorrectionGlobal";
    if (hasChangeInSet(userChangedPathsCurrent, prefix)) nlGlobalTextarea.classList.add("user-changed-current");
    else if (hasChangeInSet(userChangedPathsPrevious, prefix)) nlGlobalTextarea.classList.add("user-changed-previous");
    else if (hasChangeInSet(apiChangedPaths, prefix)) nlGlobalTextarea.classList.add("api-changed");
  }

  // JSON取り込み
  if (btnImportJson && gptOutputTextarea) {
    btnImportJson.addEventListener("click", () => {
      const text = gptOutputTextarea.value.trim();
      if (!text) return alert("JSONが入力されていません。");

      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (err) {
        console.error("JSON parse error:", err);
        alert("JSONとして解析できませんでした。末尾カンマなどがないか確認してください。");
        return;
      }

      // schema v1.1.2（必須項目補正 + 旧形式の簡易移行）
      normalizePayloadForSchema(parsed);

      if (basePayload || currentPayload) pushStateHistory(true);

      if (basePayload) {
        apiChangedPaths = collectDiffPaths(basePayload, parsed);
        userChangedPathsPrevious = new Set(userChangedPathsCurrent);
      } else {
        apiChangedPaths = new Set();
        userChangedPathsPrevious = new Set();
      }

      basePayload = parsed;
      currentPayload = deepClone(parsed);

      userChangedPathsCurrent = new Set();
      selectedIndex = null;
      if (itemEditor) itemEditor.value = "";
      resetForm();

      syncNlGlobalToUI();
      renderCategoryTabs();
      setFormModeForCategory();
      syncSiteCostsToUI();
      updateGlobalHighlight();

      renderItemsTable();
      updateSourceInfoUI();
      setStatus(`JSONを取り込みました（${selectedCategoryKey}=${getItems().length}件）`);
      updateRestoreButtonEnabled();
      updateRedoButtonEnabled();
    });
  }

  function getItems() {
    if (!currentPayload) return [];
    if (selectedCategoryKey === "siteCosts") return [];
    const arr = currentPayload[selectedCategoryKey];
    return Array.isArray(arr) ? arr : [];
  }

  // ★手動追加（空の項目を追加）
  if (btnAddEmptyItem) {
    btnAddEmptyItem.addEventListener("click", () => {
      if (!currentPayload) {
        alert("先に ChatGPT 出力JSONを取り込んでください（カテゴリ配列に追加するため）。");
        return;
      }
      normalizePayloadForSchema(currentPayload);

      if (selectedCategoryKey === "siteCosts") {
        alert("現場費（siteCosts）は配列ではないため、『空の項目を追加』は使用できません。上の現場費フォームから編集してください。");
        return;
      }

      if (!Array.isArray(currentPayload[selectedCategoryKey])) currentPayload[selectedCategoryKey] = [];

      pushStateHistory(true);

      const newItem = createEmptyItem(currentPayload[selectedCategoryKey], selectedCategoryKey);
      currentPayload[selectedCategoryKey].push(newItem);

      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }

      selectedIndex = currentPayload[selectedCategoryKey].length - 1;

      renderItemsTable();
      syncNlGlobalToUI();
      updateGlobalHighlight();

      // 選択・フォームへ反映
      if (itemEditor) itemEditor.value = JSON.stringify(newItem, null, 2);
      updateFormFromItem(newItem);
      updateFormHighlightsForSelected();
      if (itemEditorLabel) itemEditorLabel.textContent = `ID: ${newItem.id || ""} / index: ${selectedIndex}`;

      setStatus(`空の項目を追加しました: ${newItem.id}（${selectedCategoryKey}）`);
      updateRestoreButtonEnabled();
      updateRedoButtonEnabled();
    });
  }

  // テーブル描画
  function renderItemsTable() {
    if (!itemsView) return;

    // siteCosts は配列ではないので、専用表示
    if (selectedCategoryKey === "siteCosts") {
      normalizePayloadForSchema(currentPayload);
      const sc = (currentPayload && currentPayload.siteCosts) || {};
      const labor = sc.laborCost || {};
      const transport = sc.transportCost || {};
      const waste = sc.wasteDisposalCost || {};

      const version = (currentPayload && currentPayload.version) || "";
      const stage = (currentPayload && currentPayload.stage) || "";
      const hasGlobal = currentPayload && typeof currentPayload.nlCorrectionGlobal === "string";

      let html = "";
      html += `<div style="margin-bottom:4px;">`;
      html += `<strong>version:</strong> ${escapeHtml(version)} / <strong>stage:</strong> ${escapeHtml(stage)} / `;
      html += `<strong>nlCorrectionGlobal:</strong> ${hasGlobal ? "OK" : "MISSING"} / `;
      html += `<strong>category:</strong> siteCosts</div>`;

      html += `<table>`;
      html += `<thead><tr><th>項目</th><th>金額(JPY)</th><th>メモ</th></tr></thead><tbody>`;
      html += `<tr><td>laborCost</td><td>${escapeHtml(typeof labor.amount === "number" ? labor.amount : "")}</td><td>${escapeHtml(labor.notes || "")}</td></tr>`;
      html += `<tr><td>transportCost</td><td>${escapeHtml(typeof transport.amount === "number" ? transport.amount : "")}</td><td>${escapeHtml(transport.notes || "")}</td></tr>`;
      html += `<tr><td>wasteDisposalCost</td><td>${escapeHtml(typeof waste.amount === "number" ? waste.amount : "")}</td><td>${escapeHtml(waste.notes || "")}</td></tr>`;
      html += `</tbody></table>`;
      itemsView.innerHTML = html;
      return;
    }

    const items = getItems();
    const version = (currentPayload && currentPayload.version) || "";
    const stage = (currentPayload && currentPayload.stage) || "";
    const hasGlobal = currentPayload && typeof currentPayload.nlCorrectionGlobal === "string";

    let html = "";
    html += `<div style="margin-bottom:4px;">`;
    html += `<strong>version:</strong> ${escapeHtml(version)} / <strong>stage:</strong> ${escapeHtml(stage)} / `;
    html += `<strong>nlCorrectionGlobal:</strong> ${hasGlobal ? "OK" : "MISSING"} / `;
    const catLabel = (CATEGORIES.find((c) => c.key === selectedCategoryKey) || {}).label || selectedCategoryKey;
    html += `<strong>category:</strong> ${escapeHtml(catLabel)} / <strong>count:</strong> ${items.length}件</div>`;

    if (items.length > 0) {
      const isWoodwork = selectedCategoryKey === "woodworkItems";
      html += `<table>`;
      html += `<thead><tr>
        <th>#</th>
        <th>積算</th>
        <th>名称</th>
        <th>種別</th>
        <th>寸法(H/W/D)</th>
        <th>材料(kind)</th>
        <th>仕上げ</th>
        <th>数量</th>
        <th>単価</th>
        <th>費用/材料費</th>
        <th>人工合計(人日)</th>
        <th>source</th>
      </tr></thead><tbody>`;

      items.forEach((item, index) => {
        const name = item.name || "";
        const structureType = isWoodwork ? (item.structureType || item.type || "") : "";

        const dims = isWoodwork ? item.dimensions || {} : {};
        const dimParts = [];
        if (typeof dims.height === "number") dimParts.push(`H${dims.height}`);
        if (typeof dims.width === "number") dimParts.push(`W${dims.width}`);
        if (typeof dims.depth === "number") dimParts.push(`D${dims.depth}`);
        const dimStr = dimParts.join(" ");

        const materials = isWoodwork && Array.isArray(item.materials) ? item.materials : [];
        const materialKinds = materials.map((m) => m && m.kind).filter((k) => !!k);
        const materialKindStr = materialKinds.join(", ");

        const finishes = isWoodwork && Array.isArray(item.finishes) ? item.finishes : [];
        const finishesStr = finishes.join(", ");

        const quantity = typeof item.quantity === "number" ? item.quantity : "";

        const unitPrice = item.price && typeof item.price.unitPrice === "number" ? item.price.unitPrice : "";

        const materialCostAmount = isWoodwork
          ? item.materialCost && typeof item.materialCost.amount === "number"
            ? item.materialCost.amount
            : ""
          : item.cost && typeof item.cost.amount === "number"
            ? item.cost.amount
            : "";

        const laborTotal = isWoodwork ? item.laborTotal || item.labor || null : null;
        let laborTotalDays = "";
        if (laborTotal && typeof laborTotal.amount === "number") {
          if (laborTotal.unit === "人日") laborTotalDays = laborTotal.amount.toString();
          else if (laborTotal.unit === "人時") laborTotalDays = (laborTotal.amount / 8).toFixed(2);
        }

        const includeInEstimateMark =
          typeof item.includeInEstimate === "boolean" ? (item.includeInEstimate ? "◯" : "×") : "";

        const src = item.source || "";

        const rowClass = index === selectedIndex ? "selected" : "";
        const highlightOn = isHighlightEnabled();

        const basePath = `/${selectedCategoryKey}/${index}`;
        const clsInclude = highlightOn ? classForPathPrefix(`${basePath}/includeInEstimate`) : "";
        const clsName = highlightOn ? classForPathPrefix(`${basePath}/name`) : "";
        const clsType = highlightOn && isWoodwork ? classForPathPrefix(`${basePath}/structureType`) : "";
        const clsDim = highlightOn && isWoodwork ? classForPathPrefix(`${basePath}/dimensions`) : "";
        const clsMat = highlightOn && isWoodwork ? classForPathPrefix(`${basePath}/materials`) : "";
        const clsFin = highlightOn && isWoodwork ? classForPathPrefix(`${basePath}/finishes`) : "";
        const clsQty = highlightOn ? classForPathPrefix(`${basePath}/quantity`) : "";
        const clsUnitPrice = highlightOn ? classForPathPrefix(`${basePath}/price/unitPrice`) : "";
        const clsMatCost = highlightOn ? classForPathPrefix(`${basePath}/${isWoodwork ? "materialCost" : "cost"}`) : "";
        const clsLabor = highlightOn && isWoodwork ? classForPathPrefix(`${basePath}/laborTotal`) : "";
        const clsSrc = highlightOn ? classForPathPrefix(`${basePath}/source`) : "";

        html += `<tr data-index="${index}" class="${rowClass}">
          <td>${index + 1}</td>
          <td class="${clsInclude}">${escapeHtml(includeInEstimateMark)}</td>
          <td class="${clsName}">${escapeHtml(name)}</td>
          <td class="${clsType}">${escapeHtml(structureType)}</td>
          <td class="${clsDim}">${escapeHtml(dimStr)}</td>
          <td class="${clsMat}">${escapeHtml(materialKindStr)}</td>
          <td class="${clsFin}">${escapeHtml(finishesStr)}</td>
          <td class="${clsQty}">${escapeHtml(quantity)}</td>
          <td class="${clsUnitPrice}">${escapeHtml(unitPrice)}</td>
          <td class="${clsMatCost}">${escapeHtml(materialCostAmount)}</td>
          <td class="${clsLabor}">${escapeHtml(laborTotalDays)}</td>
          <td class="${clsSrc}">${escapeHtml(src)}</td>
        </tr>`;
      });

      html += `</tbody></table>`;
    } else {
      html += `<div>items が見つかりませんでした。</div>`;
    }

    itemsView.innerHTML = html;
  }

  // 行クリック
  if (itemsView) {
    itemsView.addEventListener("click", (event) => {
      if (selectedCategoryKey === "siteCosts") return;
      const tr = event.target.closest("tr[data-index]");
      if (!tr) return;
      const index = parseInt(tr.getAttribute("data-index"), 10);
      if (isNaN(index)) return;

      selectedIndex = index;
      const items = getItems();
      const item = items[index];
      if (!item) return;

      if (itemEditor) itemEditor.value = JSON.stringify(item, null, 2);
      updateFormFromItem(item);
      updateFormHighlightsForSelected();

      if (itemEditorLabel) itemEditorLabel.textContent = `ID: ${item.id || ""} / index: ${index}`;

      renderItemsTable();
    });
  }

  // JSONタブ「反映」
  if (btnApplyItemEdit && itemEditor) {
    btnApplyItemEdit.addEventListener("click", () => {
      if (selectedCategoryKey === "siteCosts") return alert("現場費（siteCosts）はアイテム配列ではありません。上の現場費フォームから編集してください。");
      if (selectedIndex == null) return alert("まずテーブルからアイテムを選択してください。");
      if (!currentPayload || !Array.isArray(currentPayload[selectedCategoryKey]))
        return alert(`現在のJSONに ${selectedCategoryKey} がありません。`);

      let newItem;
      try {
        newItem = JSON.parse(itemEditor.value);
      } catch (err) {
        console.error("item JSON parse error:", err);
        alert("アイテムJSONとして解析できませんでした。構造やカンマを確認してください。");
        return;
      }

      currentPayload[selectedCategoryKey][selectedIndex] = newItem;

      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }

      updateFormFromItem(newItem);
      updateFormHighlightsForSelected();
      updateGlobalHighlight();
      renderItemsTable();
      setStatus("JSONの変更をアイテムに反映しました");
    });
  }

  // 差分JSON生成
  if (btnGenerateDiff) {
    btnGenerateDiff.addEventListener("click", async () => {
      if (!basePayload || !currentPayload) return alert("まず JSON を取り込んでください。");

      normalizePayloadForSchema(basePayload);
      normalizePayloadForSchema(currentPayload);

      const patch = generateJsonPatch(basePayload, currentPayload);

      const diff = {
        baseVersion: basePayload.version || null,
        baseStage: basePayload.stage || null,
        generatedAt: new Date().toISOString(),
        jsonPatch: patch
      };

      const diffJsonText = JSON.stringify(diff, null, 2);

      if (stageSelect) stageSelect.value = "update";
      if (!promptTextarea) return;

      try {
        const result = await window.api.getPrompt("update");
        let baseText = result.prompt || "";

        const basePayloadForPrompt = deepClone(basePayload);

        // ★ stage は review にする（要件維持）
        basePayloadForPrompt.stage = "review";

        // ★ version は v1.1.2 を保証（統一）
        basePayloadForPrompt.version = SCHEMA_VERSION;

        if (typeof basePayloadForPrompt.nlCorrectionGlobal !== "string") {
          basePayloadForPrompt.nlCorrectionGlobal = "";
        }

        const combined =
          baseText +
          "\n\n" +
          "===== 差分反映用の入力データ =====\n" +
          "【差分バッチ（jsonPatch フィールドが RFC6902 JSON Patch 規格に準拠）】\n" +
          diffJsonText +
          "\n\n" +
          "※ jsonPatch を前提に、差分が反映された最新の完全JSONを 1 つだけ出力してください。（元JSONはここには含めません）";

        promptTextarea.value = combined;

        setStatus("RFC6902 JSON Patch を生成し、更新用プロンプト欄にセットしました（operations=" + patch.length + "）");
      } catch (err) {
        console.error("getPrompt(update) error:", err);
        setStatus("更新用プロンプトの読み込みに失敗しました");
      }
    });
  }

  // 状態ファイル保存
  if (btnExportStateFile) {
    btnExportStateFile.addEventListener("click", () => {
      if (!basePayload && !currentPayload) {
        alert("エクスポートできる状態がまだありません。先にJSONを取り込んでください。");
        return;
      }
      const snap = buildStateSnapshot();
      const jsonText = JSON.stringify(snap, null, 2);
      const blob = new Blob([jsonText], { type: "application/json" });

      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      const filename = `takeoff-state-${timestamp}.json`;

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setStatus("現在の状態をファイルに保存しました");
    });
  }

  // 状態ファイル読込
  if (btnImportStateFile && stateFileInput) {
    btnImportStateFile.addEventListener("click", () => {
      stateFileInput.value = "";
      stateFileInput.click();
    });

    stateFileInput.addEventListener("change", () => {
      const file = stateFileInput.files && stateFileInput.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = () => {
        const text = String(reader.result || "");
        let snap;
        try {
          snap = JSON.parse(text);
        } catch (err) {
          console.error("state JSON parse error:", err);
          alert("状態JSONとして解析できませんでした。エクスポートしたものをそのまま読み込んでください。");
          return;
        }

        if (!("basePayload" in snap) || !("currentPayload" in snap)) {
          alert("状態JSONの形式が想定と異なります。エクスポートしたJSONファイルを使用してください。");
          return;
        }

        if (basePayload || currentPayload) pushStateHistory(true);

        applySnapshot(snap);

        updateRestoreButtonEnabled();
        updateRedoButtonEnabled();
        setStatus(`状態ファイル「${file.name}」を読み込み、状態を復元しました`);
      };

      reader.onerror = () => {
        console.error("state file read error:", reader.error);
        alert("ファイル読み込み中にエラーが発生しました。");
      };

      reader.readAsText(file);
    });
  }

  // ===== フォーム処理 =====

  function resetForm() {
    if (itemFormId) itemFormId.textContent = "-";
    if (itemFormInclude) itemFormInclude.checked = false;
    if (itemFormName) itemFormName.value = "";
    if (itemFormStructureType) itemFormStructureType.value = "";
    if (itemFormDimH) itemFormDimH.value = "";
    if (itemFormDimW) itemFormDimW.value = "";
    if (itemFormDimD) itemFormDimD.value = "";
    if (itemFormQuantity) itemFormQuantity.value = "";
    if (itemFormMaterials) itemFormMaterials.value = "";
    if (itemFormFinishes) itemFormFinishes.value = "";
    if (itemFormPriceUnitPrice) itemFormPriceUnitPrice.value = "";
    if (itemFormMaterialCostAmount) itemFormMaterialCostAmount.value = "";
    if (itemFormMaterialCostNotes) itemFormMaterialCostNotes.value = "";
    if (itemFormLaborPerUnitDays) itemFormLaborPerUnitDays.value = "";
    if (itemFormLaborDays) itemFormLaborDays.value = "";
    if (itemFormCostAmount) itemFormCostAmount.value = "";
    if (itemFormCostNotes) itemFormCostNotes.value = "";
    if (itemFormNlCorrection) itemFormNlCorrection.value = "";
    if (itemFormIsBent) itemFormIsBent.checked = false;
    if (itemFormSpecialAngles) itemFormSpecialAngles.checked = false;
    if (itemFormSupportHeavy) itemFormSupportHeavy.checked = false;
    if (itemFormFinishAreaValue) itemFormFinishAreaValue.value = "";
    if (itemFormFinishAreaNotes) itemFormFinishAreaNotes.value = "";
    if (itemFormLaborCoefValue) itemFormLaborCoefValue.value = "";
    if (itemFormLaborCoefNotes) itemFormLaborCoefNotes.value = "";
    clearFormHighlightClasses();
  }

  function updateFormFromItem(item) {
    if (!item) return resetForm();

    const isWoodwork = selectedCategoryKey === "woodworkItems";

    if (itemFormId) itemFormId.textContent = item.id || "-";
    if (itemFormInclude) itemFormInclude.checked = !!item.includeInEstimate;
    if (itemFormName) itemFormName.value = item.name || "";

    if (itemFormStructureType) {
      itemFormStructureType.value = isWoodwork ? item.structureType || "" : "";
      if (
        itemFormStructureType.value &&
        !Array.from(itemFormStructureType.options).some((opt) => opt.value === itemFormStructureType.value)
      ) {
        itemFormStructureType.value = "";
      }
    }

    const dims = isWoodwork ? item.dimensions || {} : {};
    if (itemFormDimH) itemFormDimH.value = typeof dims.height === "number" ? String(dims.height) : "";
    if (itemFormDimW) itemFormDimW.value = typeof dims.width === "number" ? String(dims.width) : "";
    if (itemFormDimD) itemFormDimD.value = typeof dims.depth === "number" ? String(dims.depth) : "";

    if (itemFormQuantity) itemFormQuantity.value = typeof item.quantity === "number" ? String(item.quantity) : "";

    if (itemFormMaterials) {
      const materials = isWoodwork && Array.isArray(item.materials) ? item.materials : [];
      const lines = materials
        .map((m) => {
          if (!m || !m.kind) return "";
          const q = typeof m.quantity === "number" && m.quantity !== 1 ? ` x ${m.quantity}` : "";
          return `${m.kind}${q}`;
        })
        .filter((s) => s.length > 0);
      itemFormMaterials.value = lines.join("\n");
    }

    if (itemFormFinishes) {
      const finishes = isWoodwork && Array.isArray(item.finishes) ? item.finishes : [];
      itemFormFinishes.value = finishes.join("\n");
    }

    if (itemFormPriceUnitPrice) {
      const up = item.price && typeof item.price.unitPrice === "number" ? item.price.unitPrice : "";
      itemFormPriceUnitPrice.value = up === "" ? "" : String(up);
    }

    const mc = isWoodwork ? item.materialCost || null : null;
    if (itemFormMaterialCostAmount) itemFormMaterialCostAmount.value = mc && typeof mc.amount === "number" ? String(mc.amount) : "";
    if (itemFormMaterialCostNotes) itemFormMaterialCostNotes.value = mc?.notes || "";

    const lp = isWoodwork ? item.laborPerUnit || null : null;
    if (itemFormLaborPerUnitDays) {
      let daysStr = "";
      if (lp && typeof lp.amount === "number") {
        if (lp.unit === "人日") daysStr = String(lp.amount);
        else if (lp.unit === "人時") daysStr = (lp.amount / 8).toFixed(2);
      }
      itemFormLaborPerUnitDays.value = daysStr;
    }

    const lt = isWoodwork ? item.laborTotal || item.labor || null : null;
    if (itemFormLaborDays) {
      let daysStr = "";
      if (lt && typeof lt.amount === "number") {
        if (lt.unit === "人日") daysStr = String(lt.amount);
        else if (lt.unit === "人時") daysStr = (lt.amount / 8).toFixed(2);
      }
      itemFormLaborDays.value = daysStr;
    }

    if (itemFormNlCorrection) itemFormNlCorrection.value = item.nlCorrection || "";

    if (itemFormIsBent) itemFormIsBent.checked = isWoodwork && !!item.isBent;
    if (itemFormSpecialAngles) itemFormSpecialAngles.checked = isWoodwork && !!item.hasSpecialAngles;
    if (itemFormSupportHeavy) itemFormSupportHeavy.checked = isWoodwork && !!item.supportsHeavyLoad;

    const fsa = isWoodwork ? item.finishSurfaceArea || null : null;
    if (itemFormFinishAreaValue) itemFormFinishAreaValue.value = fsa && typeof fsa.value === "number" ? String(fsa.value) : "";
    if (itemFormFinishAreaNotes) itemFormFinishAreaNotes.value = fsa?.notes || "";

    const lc = isWoodwork ? item.laborCoefficient || null : null;
    if (itemFormLaborCoefValue) itemFormLaborCoefValue.value = lc && typeof lc.value === "number" ? String(lc.value) : "";
    if (itemFormLaborCoefNotes) itemFormLaborCoefNotes.value = lc?.notes || "";

    const cost = !isWoodwork ? item.cost || null : null;
    if (itemFormCostAmount) itemFormCostAmount.value = cost && typeof cost.amount === "number" ? String(cost.amount) : "";
    if (itemFormCostNotes) itemFormCostNotes.value = cost?.notes || "";
  }

  function applyFormToItem(item, changedElement) {
    if (!item) return;
    const isChanged = (el) => !changedElement || changedElement === el;
    const isWoodwork = selectedCategoryKey === "woodworkItems";

    if (itemFormInclude && isChanged(itemFormInclude)) item.includeInEstimate = !!itemFormInclude.checked;
    if (itemFormName && isChanged(itemFormName)) if (itemFormName.value !== "") item.name = itemFormName.value;

    if (isWoodwork && itemFormStructureType && isChanged(itemFormStructureType)) {
      if (itemFormStructureType.value !== "") {
        const v = itemFormStructureType.value;
        item.structureType = v;
      }
    }

    if (
      isWoodwork &&
      itemFormDimH &&
      itemFormDimW &&
      itemFormDimD &&
      (isChanged(itemFormDimH) || isChanged(itemFormDimW) || isChanged(itemFormDimD))
    ) {
      const h = parseFloat(itemFormDimH.value);
      const w = parseFloat(itemFormDimW.value);
      const d = parseFloat(itemFormDimD.value);
      if (!isNaN(h) && !isNaN(w) && !isNaN(d)) {
        item.dimensions = { height: h, width: w, depth: d, unit: "mm" };
        item.boundingBoxVolume = (h / 1000) * (w / 1000) * (d / 1000);
      }
    }

    if (itemFormQuantity && isChanged(itemFormQuantity)) {
      if (itemFormQuantity.value !== "") {
        const q = isWoodwork ? parseInt(itemFormQuantity.value, 10) : parseFloat(itemFormQuantity.value);
        if (!isNaN(q) && q > 0) item.quantity = q;
      }
    }

    if (isWoodwork && itemFormMaterials && isChanged(itemFormMaterials)) {
      const lines = (itemFormMaterials.value || "")
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0);

      item.materials = lines.length
        ? lines.map((line) => {
            let kind = line;
            let qty = 1;
            const m = line.split(/[x×]/i);
            if (m.length >= 2) {
              kind = m[0].trim();
              const num = parseFloat(m[1].trim());
              if (!isNaN(num) && num > 0) qty = num;
            }
            return { kind, spec: "", quantity: qty, unit: "式" };
          })
        : [];
    }

    if (isWoodwork && itemFormFinishes && isChanged(itemFormFinishes)) {
      const lines = (itemFormFinishes.value || "")
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
      item.finishes = lines;
    }

    if (itemFormPriceUnitPrice && isChanged(itemFormPriceUnitPrice)) {
      if (itemFormPriceUnitPrice.value !== "") {
        const v = parseFloat(itemFormPriceUnitPrice.value);
        if (!isNaN(v)) {
          if (!item.price) item.price = { mode: "estimate_by_model", currency: "JPY", unitPrice: v, notes: "" };
          else item.price.unitPrice = v;
        }
      }
    }

    if (
      isWoodwork &&
      ((itemFormMaterialCostAmount && isChanged(itemFormMaterialCostAmount)) ||
        (itemFormMaterialCostNotes && isChanged(itemFormMaterialCostNotes)))
    ) {
      if (itemFormMaterialCostAmount && itemFormMaterialCostAmount.value !== "") {
        const amount = parseFloat(itemFormMaterialCostAmount.value);
        if (!isNaN(amount)) {
          const notes = (itemFormMaterialCostNotes && itemFormMaterialCostNotes.value) || "";
          item.materialCost = { amount, currency: "JPY", notes };
        }
      } else if (itemFormMaterialCostNotes && itemFormMaterialCostNotes.value !== "") {
        if (!item.materialCost) item.materialCost = { amount: 0, currency: "JPY", notes: itemFormMaterialCostNotes.value };
        else item.materialCost.notes = itemFormMaterialCostNotes.value;
      }
    }

    if (isWoodwork && itemFormLaborPerUnitDays && isChanged(itemFormLaborPerUnitDays)) {
      if (itemFormLaborPerUnitDays.value !== "") {
        const days = parseFloat(itemFormLaborPerUnitDays.value);
        if (!isNaN(days)) {
          const prevNotes = item.laborPerUnit?.notes || "";
          item.laborPerUnit = { amount: days, unit: "人日", notes: prevNotes };
        }
      }
    }

    if (isWoodwork && itemFormLaborDays && isChanged(itemFormLaborDays)) {
      if (itemFormLaborDays.value !== "") {
        const days = parseFloat(itemFormLaborDays.value);
        if (!isNaN(days)) {
          const prevNotes = item.laborTotal?.notes || item.labor?.notes || "";
          item.laborTotal = { amount: days, unit: "人日", notes: prevNotes };
        }
      }
    }

    // ★ nlCorrection（移動後）
    if (itemFormNlCorrection && isChanged(itemFormNlCorrection)) {
      item.nlCorrection = itemFormNlCorrection.value || "";
    }

    if (isWoodwork && itemFormIsBent && isChanged(itemFormIsBent)) item.isBent = !!itemFormIsBent.checked;
    if (isWoodwork && itemFormSpecialAngles && isChanged(itemFormSpecialAngles)) item.hasSpecialAngles = !!itemFormSpecialAngles.checked;
    if (isWoodwork && itemFormSupportHeavy && isChanged(itemFormSupportHeavy)) item.supportsHeavyLoad = !!itemFormSupportHeavy.checked;

    if (
      isWoodwork &&
      ((itemFormFinishAreaValue && isChanged(itemFormFinishAreaValue)) ||
        (itemFormFinishAreaNotes && isChanged(itemFormFinishAreaNotes)))
    ) {
      if (itemFormFinishAreaValue && itemFormFinishAreaValue.value !== "") {
        const v = parseFloat(itemFormFinishAreaValue.value);
        if (!isNaN(v)) {
          const notes = (itemFormFinishAreaNotes && itemFormFinishAreaNotes.value) || "";
          item.finishSurfaceArea = { value: v, notes };
        }
      } else if (itemFormFinishAreaNotes && itemFormFinishAreaNotes.value !== "") {
        if (!item.finishSurfaceArea) item.finishSurfaceArea = { value: 0, notes: itemFormFinishAreaNotes.value };
        else item.finishSurfaceArea.notes = itemFormFinishAreaNotes.value;
      }
    }

    if (
      isWoodwork &&
      ((itemFormLaborCoefValue && isChanged(itemFormLaborCoefValue)) ||
        (itemFormLaborCoefNotes && isChanged(itemFormLaborCoefNotes)))
    ) {
      if (itemFormLaborCoefValue && itemFormLaborCoefValue.value !== "") {
        const v = parseFloat(itemFormLaborCoefValue.value);
        if (!isNaN(v)) {
          const notes = (itemFormLaborCoefNotes && itemFormLaborCoefNotes.value) || "";
          item.laborCoefficient = { value: v, notes };
        }
      } else if (itemFormLaborCoefNotes && itemFormLaborCoefNotes.value !== "") {
        if (!item.laborCoefficient) item.laborCoefficient = { value: 0, notes: itemFormLaborCoefNotes.value };
        else item.laborCoefficient.notes = itemFormLaborCoefNotes.value;
      }
    }

    if (
      !isWoodwork &&
      ((itemFormCostAmount && isChanged(itemFormCostAmount)) || (itemFormCostNotes && isChanged(itemFormCostNotes)))
    ) {
      if (itemFormCostAmount && itemFormCostAmount.value !== "") {
        const amount = parseFloat(itemFormCostAmount.value);
        if (!isNaN(amount)) {
          const notes = (itemFormCostNotes && itemFormCostNotes.value) || "";
          item.cost = { amount, currency: "JPY", notes };
        }
      } else if (itemFormCostNotes && itemFormCostNotes.value !== "") {
        if (!item.cost) item.cost = { amount: 0, currency: "JPY", notes: itemFormCostNotes.value };
        else item.cost.notes = itemFormCostNotes.value;
      }
    }
  }

  setupFormAutoApply();
  function setupFormAutoApply() {
    const controls = [
      itemFormInclude,
      itemFormName,
      itemFormStructureType,
      itemFormDimH,
      itemFormDimW,
      itemFormDimD,
      itemFormQuantity,
      itemFormMaterials,
      itemFormFinishes,
      itemFormPriceUnitPrice,
      itemFormMaterialCostAmount,
      itemFormMaterialCostNotes,
      itemFormLaborPerUnitDays,
      itemFormLaborDays,
      itemFormCostAmount,
      itemFormCostNotes,
      itemFormNlCorrection, // ★移動後も自動反映
      itemFormIsBent,
      itemFormSpecialAngles,
      itemFormSupportHeavy,
      itemFormFinishAreaValue,
      itemFormFinishAreaNotes,
      itemFormLaborCoefValue,
      itemFormLaborCoefNotes
    ];

    controls.forEach((ctrl) => {
      if (!ctrl) return;
      const evt = ctrl.tagName === "SELECT" || ctrl.type === "checkbox" ? "change" : "input";
      ctrl.addEventListener(evt, (e) => {
        if (selectedIndex == null) return;
        if (selectedCategoryKey === "siteCosts") return;
        if (!currentPayload || !Array.isArray(currentPayload[selectedCategoryKey])) return;

        normalizePayloadForSchema(currentPayload);

        const item = currentPayload[selectedCategoryKey][selectedIndex];
        if (!item) return;

        applyFormToItem(item, e.target || ctrl);

        if (itemEditor) itemEditor.value = JSON.stringify(item, null, 2);

        if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);

        renderItemsTable();
        updateGlobalHighlight();
        updateFormHighlightsForSelected();
        setStatus("フォーム変更を反映しました");
      });
    });
  }

  function hasChangeInSet(set, prefix) {
    for (const p of set) {
      if (p === prefix || p.startsWith(prefix + "/")) return true;
    }
    return false;
  }

  function classForPathPrefix(prefix) {
    if (!isHighlightEnabled()) return "";
    if (hasChangeInSet(userChangedPathsCurrent, prefix)) return "user-changed-current";
    if (hasChangeInSet(userChangedPathsPrevious, prefix)) return "user-changed-previous";
    if (hasChangeInSet(apiChangedPaths, prefix)) return "api-changed";
    return "";
  }

  function clearFormHighlightClasses() {
    const formControls = [
      itemFormInclude,
      itemFormName,
      itemFormStructureType,
      itemFormDimH,
      itemFormDimW,
      itemFormDimD,
      itemFormQuantity,
      itemFormMaterials,
      itemFormFinishes,
      itemFormPriceUnitPrice,
      itemFormMaterialCostAmount,
      itemFormMaterialCostNotes,
      itemFormLaborPerUnitDays,
      itemFormLaborDays,
      itemFormCostAmount,
      itemFormCostNotes,
      itemFormNlCorrection,
      itemFormIsBent,
      itemFormSpecialAngles,
      itemFormSupportHeavy,
      itemFormFinishAreaValue,
      itemFormFinishAreaNotes,
      itemFormLaborCoefValue,
      itemFormLaborCoefNotes
    ];
    formControls.forEach((el) => {
      if (!el) return;
      el.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
    });
  }

  function updateFormHighlightsForSelected() {
    clearFormHighlightClasses();
    if (!isHighlightEnabled()) return;
    if (selectedIndex == null) return;

    const isWoodwork = selectedCategoryKey === "woodworkItems";
    const idx = selectedIndex;
    const basePath = `/${selectedCategoryKey}/${idx}`;
    const apply = (el, prefix) => {
      if (!el) return;
      el.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
      if (hasChangeInSet(userChangedPathsCurrent, prefix)) el.classList.add("user-changed-current");
      else if (hasChangeInSet(userChangedPathsPrevious, prefix)) el.classList.add("user-changed-previous");
      else if (hasChangeInSet(apiChangedPaths, prefix)) el.classList.add("api-changed");
    };

    apply(itemFormInclude, `${basePath}/includeInEstimate`);
    apply(itemFormName, `${basePath}/name`);
    // woodwork は structureType、それ以外は type を優先
    if (isWoodwork) apply(itemFormStructureType, `${basePath}/structureType`);
    if (isWoodwork) {
      apply(itemFormDimH, `${basePath}/dimensions/height`);
      apply(itemFormDimW, `${basePath}/dimensions/width`);
      apply(itemFormDimD, `${basePath}/dimensions/depth`);
    }
    apply(itemFormQuantity, `${basePath}/quantity`);
    if (isWoodwork) {
      apply(itemFormMaterials, `${basePath}/materials`);
      apply(itemFormFinishes, `${basePath}/finishes`);
    }
    apply(itemFormPriceUnitPrice, `${basePath}/price/unitPrice`);
    if (isWoodwork) {
      apply(itemFormMaterialCostAmount, `${basePath}/materialCost`);
      apply(itemFormMaterialCostNotes, `${basePath}/materialCost`);
      apply(itemFormLaborPerUnitDays, `${basePath}/laborPerUnit`);
      apply(itemFormLaborDays, `${basePath}/laborTotal`);
    } else {
      apply(itemFormCostAmount, `${basePath}/cost`);
      apply(itemFormCostNotes, `${basePath}/cost`);
    }
    apply(itemFormNlCorrection, `${basePath}/nlCorrection`);

    if (isWoodwork) {
      apply(itemFormIsBent, `${basePath}/isBent`);
      apply(itemFormSpecialAngles, `${basePath}/hasSpecialAngles`);
      apply(itemFormSupportHeavy, `${basePath}/supportsHeavyLoad`);
      apply(itemFormFinishAreaValue, `${basePath}/finishSurfaceArea`);
      apply(itemFormFinishAreaNotes, `${basePath}/finishSurfaceArea`);
      apply(itemFormLaborCoefValue, `${basePath}/laborCoefficient`);
      apply(itemFormLaborCoefNotes, `${basePath}/laborCoefficient`);
    }
  }

  // JSON Patch（深いpath・1op1パラメータ）
  function generateJsonPatch(base, curr) {
    const patch = [];
    // schema v1.1.2: items という固定配列は廃止。
    // ルート以下をすべて再帰差分にする（カテゴリ配列 + siteCosts を含む）。
    diffAny(base, curr, "", patch);
    return patch;
  }

  function diffAny(baseVal, currVal, path, patch) {
    const p = path === "" ? "/" : path;
    const baseUndef = typeof baseVal === "undefined";
    const currUndef = typeof currVal === "undefined";

    if (baseUndef && currUndef) return;
    if (baseUndef && !currUndef) return patch.push({ op: "add", path: p, value: currVal });
    if (!baseUndef && currUndef) return patch.push({ op: "remove", path: p });

    const baseIsArray = Array.isArray(baseVal);
    const currIsArray = Array.isArray(currVal);
    const baseIsObject = baseVal !== null && typeof baseVal === "object" && !baseIsArray;
    const currIsObject = currVal !== null && typeof currVal === "object" && !currIsArray;

    if (baseIsArray && currIsArray) {
      const maxLen = Math.max(baseVal.length, currVal.length);
      for (let i = 0; i < maxLen; i++) diffAny(baseVal[i], currVal[i], p + "/" + i, patch);
      return;
    }

    if (!baseIsObject || !currIsObject) {
      if (JSON.stringify(baseVal) !== JSON.stringify(currVal)) patch.push({ op: "replace", path: p, value: currVal });
      return;
    }

    const baseKeys = Object.keys(baseVal);
    const currKeys = Object.keys(currVal);
    const allKeys = new Set([...baseKeys, ...currKeys]);

    allKeys.forEach((key) => diffAny(baseVal[key], currVal[key], p + "/" + escapeJsonPointer(key), patch));
  }

  function escapeJsonPointer(str) {
    return String(str).replace(/~/g, "~0").replace(/\//g, "~1");
  }

  function collectDiffPaths(base, curr) {
    const paths = new Set();
    function walk(b, c, path) {
      const baseUndef = typeof b === "undefined";
      const currUndef = typeof c === "undefined";
      if (baseUndef && currUndef) return;
      if (baseUndef || currUndef) return paths.add(path);

      const bArr = Array.isArray(b);
      const cArr = Array.isArray(c);
      const bObj = b !== null && typeof b === "object" && !bArr;
      const cObj = c !== null && typeof c === "object" && !cArr;

      if (bArr && cArr) {
        const minLen = Math.min(b.length, c.length);

        // 共通部分
        for (let i = 0; i < minLen; i++) walk(b[i], c[i], path + "/" + i);

        // 追加（末尾）
        for (let i = minLen; i < c.length; i++) paths.add(path + "/-");

        // 削除（後ろから）
        for (let i = b.length - 1; i >= c.length; i--) paths.add(path + "/" + i);

        return;
      }

      if (!bObj || !cObj) {
        if (JSON.stringify(b) !== JSON.stringify(c)) paths.add(path);
        return;
      }

      const bKeys = Object.keys(b);
      const cKeys = Object.keys(c);
      const all = new Set([...bKeys, ...cKeys]);
      all.forEach((key) => walk(b[key], c[key], path + "/" + escapeJsonPointer(key)));
    }
    walk(base, curr, "");
    return paths;
  }

  function escapeHtml(str) {
    return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  // ===== v1.1.2: payload normalize (カテゴリ分類: items廃止) =====
  function normalizePayloadForSchema(payload) {
    if (!payload || typeof payload !== "object") return;

    payload.version = SCHEMA_VERSION;
    if (typeof payload.stage !== "string") payload.stage = "extraction";
    if (payload.stage === "extract") payload.stage = "extraction";
    if (payload.stage === "estimate") payload.stage = "ready_for_estimation";
    if (typeof payload.extractedAt !== "string") payload.extractedAt = new Date().toISOString();
    if (typeof payload.nlCorrectionGlobal !== "string") payload.nlCorrectionGlobal = "";

    // legacy: upholsteryItems -> finishingItems
    if (Array.isArray(payload.upholsteryItems) && !Array.isArray(payload.finishingItems)) {
      payload.finishingItems = payload.upholsteryItems;
    }
    if ("upholsteryItems" in payload) delete payload.upholsteryItems;

    const categoryKeys = ["woodworkItems", "floorItems", "finishingItems", "electricalItems", "leaseItems"];
    const hasAnyCategoryArray = categoryKeys.some((k) => Array.isArray(payload[k]));

    if (Array.isArray(payload.items)) {
      if (!hasAnyCategoryArray) {
        const legacyWoodwork = [];
        const legacyGeneric = [];
        payload.items.forEach((it) => {
          if (!it || typeof it !== "object") return;
          const looksWoodwork =
            (it.structureType && it.structureType !== "非木工造作物") ||
            it.dimensions ||
            it.materials ||
            it.finishes ||
            it.laborPerUnit ||
            it.laborTotal ||
            it.materialCost ||
            it.laborCost;
          if (looksWoodwork) legacyWoodwork.push(it);
          else legacyGeneric.push(it);
        });
        payload.woodworkItems = legacyWoodwork;
        payload.finishingItems = legacyGeneric;
      }
      delete payload.items;
    }

    if (Array.isArray(payload.laborItems)) {
      payload.laborItems.forEach((it) => {
        if (!it || typeof it !== "object") return;
        const rawCat = it.category || it.targetCategory || it.target || "";
        const normalized = String(rawCat).toLowerCase();
        let dest = "finishingItems";
        if (normalized.includes("floor")) dest = "floorItems";
        else if (normalized.includes("finish") || normalized.includes("upholstery")) dest = "finishingItems";
        else if (normalized.includes("electrical")) dest = "electricalItems";
        else if (normalized.includes("lease")) dest = "leaseItems";
        if (!Array.isArray(payload[dest])) payload[dest] = [];
        payload[dest].push(it);
      });
      delete payload.laborItems;
    }

    categoryKeys.forEach((k) => {
      if (!Array.isArray(payload[k])) payload[k] = [];
    });

    if (!payload.siteCosts || typeof payload.siteCosts !== "object" || Array.isArray(payload.siteCosts)) {
      payload.siteCosts = {};
    }

    const ensureCost = (obj, key) => {
      if (!obj[key] || typeof obj[key] !== "object" || Array.isArray(obj[key])) {
        obj[key] = {};
      }
      if (typeof obj[key].amount !== "number") obj[key].amount = 0;
      if (typeof obj[key].currency !== "string") obj[key].currency = "JPY";
      if (typeof obj[key].notes !== "string") obj[key].notes = "";
    };
    ensureCost(payload.siteCosts, "laborCost");
    ensureCost(payload.siteCosts, "transportCost");
    ensureCost(payload.siteCosts, "wasteDisposalCost");
    if (typeof payload.siteCosts.notes !== "string") payload.siteCosts.notes = "";

    const normalizeBaseFields = (it) => {
      if (!it || typeof it !== "object") return;
      if (typeof it.id !== "string") it.id = "itm-000";
      if (typeof it.name !== "string") it.name = "";
      if (typeof it.quantity !== "number" || it.quantity <= 0) it.quantity = 1;
      if (typeof it.includeInEstimate !== "boolean") it.includeInEstimate = true;
      if (typeof it.nlCorrection !== "string") it.nlCorrection = "";
      if (typeof it.source === "string") {
        const s = it.source.toLowerCase();
        if (s === "chatgpt") it.source = "chatGPT";
        else if (s === "user") it.source = "user";
      }
      if (it.source !== "chatGPT" && it.source !== "user") it.source = "chatGPT";
      if (!it.price || typeof it.price !== "object") it.price = { mode: "estimate_by_model", currency: "JPY", unitPrice: null, notes: "" };
      if (typeof it.price.mode !== "string") it.price.mode = "estimate_by_model";
      if (typeof it.price.currency !== "string") it.price.currency = "JPY";
      if (!("unitPrice" in it.price)) it.price.unitPrice = null;
      if (typeof it.price.notes !== "string") it.price.notes = "";
      if (typeof it.notes !== "string") it.notes = "";

      if (typeof it.unitPrice === "number" && (it.price.unitPrice == null || isNaN(it.price.unitPrice))) {
        it.price.unitPrice = it.unitPrice;
      }
      if ("unitPrice" in it) delete it.unitPrice;
    };

    const normalizeWoodworkItem = (it) => {
      if (!it || typeof it !== "object") return;
      normalizeBaseFields(it);

      if (typeof it.structureType !== "string") {
        it.structureType = typeof it.type === "string" ? it.type : "単純什器";
      }
      if ("type" in it) delete it.type;
      if ("classification" in it) delete it.classification;
      if (!Number.isInteger(it.quantity) || it.quantity < 1) it.quantity = 1;

      const toPositive = (v) => (typeof v === "number" && v > 0 ? v : 1);
      if (!it.dimensions || typeof it.dimensions !== "object") it.dimensions = {};
      it.dimensions.height = toPositive(it.dimensions.height);
      it.dimensions.width = toPositive(it.dimensions.width);
      it.dimensions.depth = toPositive(it.dimensions.depth);
      if (typeof it.dimensions.unit !== "string") it.dimensions.unit = "mm";

      if (!Array.isArray(it.materials)) it.materials = [];
      if (!Array.isArray(it.finishes)) it.finishes = [];

      if (!it.finishSurfaceArea || typeof it.finishSurfaceArea !== "object") it.finishSurfaceArea = {};
      if (typeof it.finishSurfaceArea.value !== "number") it.finishSurfaceArea.value = 0.01;
      if (typeof it.finishSurfaceArea.notes !== "string") it.finishSurfaceArea.notes = "";

      if (typeof it.isBent !== "boolean") it.isBent = false;
      if (typeof it.hasSpecialAngles !== "boolean") it.hasSpecialAngles = false;
      if (typeof it.supportsHeavyLoad !== "boolean") it.supportsHeavyLoad = false;
      if (typeof it.outlineDescription !== "string") it.outlineDescription = "";

      const ensureLabor = (key) => {
        if (!it[key] || typeof it[key] !== "object") it[key] = {};
        if (typeof it[key].amount !== "number") it[key].amount = 0.01;
        if (typeof it[key].unit !== "string") it[key].unit = "人日";
        if (typeof it[key].notes !== "string") it[key].notes = "";
      };
      ensureLabor("laborPerUnit");
      ensureLabor("laborTotal");

      if (!it.laborCoefficient || typeof it.laborCoefficient !== "object") it.laborCoefficient = {};
      if (typeof it.laborCoefficient.value !== "number") it.laborCoefficient.value = 0.01;
      if (typeof it.laborCoefficient.notes !== "string") it.laborCoefficient.notes = "";

      if (typeof it.boundingBoxVolume !== "number") {
        const h = it.dimensions.height / 1000;
        const w = it.dimensions.width / 1000;
        const d = it.dimensions.depth / 1000;
        it.boundingBoxVolume = h * w * d;
      }

      ensureCost(it, "materialCost");
      ensureCost(it, "laborCost");

      if (typeof it.recognitionConfidence !== "number") it.recognitionConfidence = 1;
      if ("cost" in it) delete it.cost;
    };

    const normalizeQuantifiedItem = (it) => {
      if (!it || typeof it !== "object") return;
      normalizeBaseFields(it);

      if (!it.cost || typeof it.cost !== "object") {
        const mat = it.materialCost?.amount || 0;
        const labor = it.laborCost?.amount || 0;
        it.cost = { amount: mat + labor, currency: "JPY", notes: "" };
      }
      ensureCost(it, "cost");
      if (typeof it.unit !== "string") it.unit = "式";

      const removeKeys = [
        "structureType",
        "dimensions",
        "materials",
        "finishes",
        "finishSurfaceArea",
        "isBent",
        "hasSpecialAngles",
        "supportsHeavyLoad",
        "outlineDescription",
        "laborPerUnit",
        "laborTotal",
        "laborCoefficient",
        "boundingBoxVolume",
        "materialCost",
        "laborCost",
        "recognitionConfidence",
        "classification",
        "type"
      ];
      removeKeys.forEach((k) => {
        if (k in it) delete it[k];
      });
    };

    const normalizeElectricalItem = (it) => {
      if (!it || typeof it !== "object") return;
      normalizeQuantifiedItem(it);
      if (typeof it.electricalType !== "string") it.electricalType = "other";
      if (typeof it.isHighPlace !== "boolean") it.isHighPlace = false;
      if (!("spec" in it)) it.spec = null;
    };

    payload.woodworkItems.forEach(normalizeWoodworkItem);
    ["floorItems", "finishingItems", "leaseItems"].forEach((k) => {
      payload[k].forEach(normalizeQuantifiedItem);
    });
    payload.electricalItems.forEach(normalizeElectricalItem);
  }

  // ===== 空 item 生成 =====
  function nextItmId(items) {
    let max = 0;
    for (const it of items) {
      const m = String(it?.id || "").match(/^itm-(\d{3,})$/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (!isNaN(n) && n > max) max = n;
      }
    }
    const next = max + 1;
    return `itm-${next.toString().padStart(3, "0")}`;
  }

  function createEmptyItem(items, categoryKey) {
    const id = nextItmId(items);
    const h = 1, w = 1, d = 1;

    // スキーマ詳細はカテゴリごとに異なるため、
    // ここでは "壊れにくい最小セット" を基本として、木工だけ拡張する。
    const base = {
      id,
      name: "",
      quantity: 1,
      includeInEstimate: true,
      price: { mode: "estimate_by_model", currency: "JPY", unitPrice: null, notes: "" },
      nlCorrection: "",
      source: "user",
      notes: ""
    };

    if (categoryKey === "woodworkItems") {
      return {
        ...base,
        structureType: "単純什器",
        dimensions: { height: h, width: w, depth: d, unit: "mm" },
        materials: [],
        finishes: [],
        finishSurfaceArea: { value: 0.01, notes: "" },
        isBent: false,
        hasSpecialAngles: false,
        supportsHeavyLoad: false,
        outlineDescription: "",
        recognitionConfidence: 1,
        laborPerUnit: { amount: 0.01, unit: "人日", notes: "" },
        laborTotal: { amount: 0.01, unit: "人日", notes: "" },
        laborCoefficient: { value: 0.01, notes: "" },
        boundingBoxVolume: (h / 1000) * (w / 1000) * (d / 1000),
        materialCost: { amount: 0, currency: "JPY", notes: "" },
        laborCost: { amount: 0, currency: "JPY", notes: "" }
      };
    }

    if (categoryKey === "electricalItems") {
      return {
        ...base,
        unit: "式",
        electricalType: "other",
        isHighPlace: false,
        spec: null,
        cost: { amount: 0, currency: "JPY", notes: "" }
      };
    }

    return {
      ...base,
      unit: "式",
      cost: { amount: 0, currency: "JPY", notes: "" }
    };
  }

  // 初期表示
  syncNlGlobalToUI();
  updateSourceInfoUI();
  updateGlobalHighlight();
  renderCategoryTabs();
  setFormModeForCategory();
  syncSiteCostsToUI();
  renderItemsTable();
});
