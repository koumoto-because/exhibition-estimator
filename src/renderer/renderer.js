// src/renderer/renderer.js

// ===== Schema compatibility =====
const SCHEMA_VERSION = "1.1.13";

const CATEGORIES = [
  { key: "woodworkItems", label: "木工造作物" },
  { key: "floorItems", label: "床" },
  { key: "finishingItems", label: "表装" },
  { key: "signItems", label: "サイン" },
  { key: "electricalItems", label: "電気" },
  { key: "leaseItems", label: "リース" },
  { key: "siteCosts", label: "現場費" }
];

let selectedCategoryKey = "woodworkItems";

let basePayload = null;
let currentPayload = null;
let selectedIndex = null;

let estimatePayload = null;
let estimateSelectedGroupIndex = null;

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

  // page tabs
  const pageTabExtract = document.getElementById("page-tab-extract");
  const pageTabEstimate = document.getElementById("page-tab-estimate");
  const extractPage = document.getElementById("extract-page");
  const estimatePage = document.getElementById("estimate-page");

  // nlCorrection per category
  const nlWoodworkTextarea = document.getElementById("nl-woodwork");
  const nlFloorTextarea = document.getElementById("nl-floor");
  const nlFinishingTextarea = document.getElementById("nl-finishing");
  const nlSignTextarea = document.getElementById("nl-sign");
  const nlElectricalTextarea = document.getElementById("nl-electrical");
  const nlLeaseTextarea = document.getElementById("nl-lease");
  const nlSiteCostsTextarea = document.getElementById("nl-sitecosts");

  // siteCosts
  const siteCostsSection = document.getElementById("site-costs-section");
  const siteCostsLaborUnitPrice = document.getElementById("sitecosts-labor-unitprice");
  const siteCostsLaborList = document.getElementById("sitecosts-labor-list");
  const siteCostsLaborAdd = document.getElementById("sitecosts-labor-add");
  const siteCostsTransportList = document.getElementById("sitecosts-transport-list");
  const siteCostsTransportAdd = document.getElementById("sitecosts-transport-add");
  const siteCostsWasteVehicles = document.getElementById("sitecosts-waste-vehicles");
  const siteCostsWasteUnitPrice = document.getElementById("sitecosts-waste-unitprice");
  const siteCostsWasteCurrency = document.getElementById("sitecosts-waste-currency");
  const siteCostsWasteNotes = document.getElementById("sitecosts-waste-notes");
  const siteCostsFormLabor = document.getElementById("sitecosts-form-labor");
  const siteCostsFormTransport = document.getElementById("sitecosts-form-transport");
  const siteCostsFormWaste = document.getElementById("sitecosts-form-waste");

  const itemEditor = document.getElementById("item-editor");
  const itemEditorLabel = document.getElementById("item-editor-label");
  const btnApplyItemEdit = document.getElementById("btn-apply-item-edit");

  // estimate editor
  const btnEstimateGenerate = document.getElementById("btn-estimate-generate");
  const estimateTitleInput = document.getElementById("estimate-title");
  const estimateClientInput = document.getElementById("estimate-client");
  const estimateTotalAmountInput = document.getElementById("estimate-total-amount");
  const estimateTaxIncludedInput = document.getElementById("estimate-tax-included");
  const estimateTaxRateInput = document.getElementById("estimate-tax-rate");
  const estimateBreakdownTitleInput = document.getElementById("estimate-breakdown-title");
  const estimateGroupsView = document.getElementById("estimate-groups-view");
  const btnEstimateGroupAdd = document.getElementById("btn-estimate-group-add");
  const btnEstimateGroupRemove = document.getElementById("btn-estimate-group-remove");
  const estimateGroupLabel = document.getElementById("estimate-group-label");
  const estimateGroupNoInput = document.getElementById("estimate-group-no");
  const estimateGroupCategoryInput = document.getElementById("estimate-group-category");
  const estimateGroupDisplaySelect = document.getElementById("estimate-group-display");
  const estimateGroupNotesInput = document.getElementById("estimate-group-notes");
  const estimateGroupMarginInput = document.getElementById("estimate-group-margin");
  const estimateGroupHiddenInput = document.getElementById("estimate-group-hidden");
  const estimateGroupPricingNotesInput = document.getElementById("estimate-group-pricing-notes");
  const estimateSummarySection = document.getElementById("estimate-group-summary");
  const estimateLinesSection = document.getElementById("estimate-group-lines");
  const estimateSummaryQtyInput = document.getElementById("estimate-summary-quantity");
  const estimateSummaryUnitInput = document.getElementById("estimate-summary-unit");
  const estimateSummaryUnitPriceInput = document.getElementById("estimate-summary-unitprice");
  const estimateSummaryAmountInput = document.getElementById("estimate-summary-amount");
  const estimateLineItemsList = document.getElementById("estimate-lineitems-list");
  const btnEstimateLineAdd = document.getElementById("btn-estimate-line-add");
  const estimateJsonTextarea = document.getElementById("estimate-json");
  const btnEstimateApplyJson = document.getElementById("btn-estimate-apply-json");
  const btnEstimateExportXlsx = document.getElementById("btn-estimate-export-xlsx");

  // ★ 手動追加（空行追加ボタン）
  const btnItemAdd = document.getElementById("btn-item-add");
  const btnItemRemove = document.getElementById("btn-item-remove");

  // フォーム
  const itemFormId = document.getElementById("item-form-id");
  const itemFormInclude = document.getElementById("item-form-include");
  const itemFormName = document.getElementById("item-form-name");
  const itemFormStructureType = document.getElementById("item-form-structure-type");
  const itemFormDimH = document.getElementById("item-form-dim-h");
  const itemFormDimW = document.getElementById("item-form-dim-w");
  const itemFormDimD = document.getElementById("item-form-dim-d");
  const itemFormQuantity = document.getElementById("item-form-quantity");
  const itemFormUnit = document.getElementById("item-form-unit");
  const itemFormMaterialsList = document.getElementById("item-form-materials-list");
  const itemFormMaterialsAdd = document.getElementById("item-form-materials-add");
  const itemFormFinishesList = document.getElementById("item-form-finishes-list");
  const itemFormFinishesAdd = document.getElementById("item-form-finishes-add");
  const itemFormSignType = document.getElementById("item-form-sign-type");
  const itemFormPriceUnitPrice = document.getElementById("item-form-price-unitprice");
  const itemFormMaterialCostAmount = document.getElementById("item-form-materialcost-amount");
  const itemFormMaterialCostNotes = document.getElementById("item-form-materialcost-notes");
  const itemFormLaborPerUnitDays = document.getElementById("item-form-labor-perunit-days");
  const itemFormLaborDays = document.getElementById("item-form-labor-days");
  const itemFormCostAmount = document.getElementById("item-form-cost-amount");
  const itemFormCostNotes = document.getElementById("item-form-cost-notes");
  const itemFormElectricalType = document.getElementById("item-form-electrical-type");
  const itemFormIsHighPlace = document.getElementById("item-form-is-high-place");
  const itemFormSpecVoltage = document.getElementById("item-form-spec-voltage");
  const itemFormSpecWatt = document.getElementById("item-form-spec-watt");
  const itemFormSpecPhase = document.getElementById("item-form-spec-phase");
  const itemFormSpecBreaker = document.getElementById("item-form-spec-breaker");
  const itemFormSpecNotes = document.getElementById("item-form-spec-notes");

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
  const electricalOnlyEls = document.querySelectorAll("[data-electrical-only]");
  const signOnlyEls = document.querySelectorAll("[data-sign-only]");

  const SITE_COST_ITEMS = [
    { key: "laborCost", label: "人工費", formEl: siteCostsFormLabor },
    { key: "transportCost", label: "運搬費", formEl: siteCostsFormTransport },
    { key: "wasteDisposalCost", label: "残材処理費", formEl: siteCostsFormWaste }
  ];

  const MATERIAL_KIND_OPTIONS = ["小割", "垂木", "ベニヤ", "曲げベニヤ", "ポリ板", "メラミン", "ラワンランバ"];
  const MATERIAL_UNIT_OPTIONS = ["枚", "本", "m", "m²", "m³", "式", "個"];
  const FINISH_KIND_OPTIONS = ["メラミン", "ポリ板", "経師", "出力シート", "塗装"];
  const FINISH_UNIT_OPTIONS = ["m²", "枚", "ロール", "m"];




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
    const isElectrical = selectedCategoryKey === "electricalItems";
    const isSign = selectedCategoryKey === "signItems";
    if (document.body) document.body.dataset.category = selectedCategoryKey;
    woodworkOnlyEls.forEach((el) => {
      el.hidden = !isWoodwork;
    });
    genericOnlyEls.forEach((el) => {
      el.hidden = isWoodwork;
    });
    electricalOnlyEls.forEach((el) => {
      el.hidden = !isElectrical;
    });
    signOnlyEls.forEach((el) => {
      el.hidden = !isSign;
    });
    if (siteCostsSection) siteCostsSection.hidden = !isSiteCosts;
    if (btnItemAdd) btnItemAdd.disabled = isSiteCosts;
    if (btnItemRemove) btnItemRemove.disabled = isSiteCosts;
    if (itemEditorContainer) itemEditorContainer.hidden = false;
    if (isSiteCosts && tabForm && tabJson && tabContentForm && tabContentJson) {
      tabForm.classList.add("active");
      tabJson.classList.remove("active");
      tabContentForm.classList.remove("hidden");
      tabContentJson.classList.add("hidden");
    }
    syncSiteCostsFormVisibility();
  }

  function selectCategory(key) {
    if (!CATEGORIES.some((c) => c.key === key)) return;
    selectedCategoryKey = key;
    selectedIndex = null;
    if (selectedCategoryKey === "siteCosts") selectedIndex = 0;
    if (itemEditor) itemEditor.value = "";
    resetForm();
    setFormModeForCategory();
    renderItemsTable();
    syncSiteCostsToUI();
    syncSiteCostsFormVisibility();
    syncSiteCostEditorForSelection();
    updateCategoryHighlight();
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

    // nlCorrection per category の復元
    syncNlCategoryToUI();
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

    updateCategoryHighlight();
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
      updateCategoryHighlight();
      updateFormHighlightsForSelected();
    });
  }

  // UNDO/REDO
  if (btnRestoreState) btnRestoreState.addEventListener("click", restorePreviousState);
  if (btnRedoState) btnRedoState.addEventListener("click", redoState);

  // キーボードショートカット: Undo/Redo
  document.addEventListener("keydown", (event) => {
    const isMac = navigator.platform && navigator.platform.toLowerCase().includes("mac");
    const isUndoKey = (isMac ? event.metaKey : event.ctrlKey) && event.key.toLowerCase() === "z";
    if (!isUndoKey) return;

    const target = event.target;
    const isEditable =
      target &&
      ((target.tagName === "INPUT" && target.type !== "checkbox") ||
        target.tagName === "TEXTAREA" ||
        target.isContentEditable);
    if (isEditable) return;

    event.preventDefault();
    if (event.shiftKey) redoState();
    else restorePreviousState();
  });

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

  const nlCategoryTextareas = {
    woodworkItems: nlWoodworkTextarea,
    floorItems: nlFloorTextarea,
    finishingItems: nlFinishingTextarea,
    signItems: nlSignTextarea,
    electricalItems: nlElectricalTextarea,
    leaseItems: nlLeaseTextarea,
    siteCosts: nlSiteCostsTextarea
  };

  const nlCategoryKeys = {
    woodworkItems: "nlCorrectionWoodworkItems",
    floorItems: "nlCorrectionFloorItems",
    finishingItems: "nlCorrectionFinishingItems",
    signItems: "nlCorrectionSignItems",
    electricalItems: "nlCorrectionElectricalItems",
    leaseItems: "nlCorrectionLeaseItems",
    siteCosts: "nlCorrectionSiteCosts"
  };

  // ===== nlCorrection per category: UI -> JSON 即時反映 =====
  Object.entries(nlCategoryTextareas).forEach(([category, el]) => {
    if (!el) return;
    el.addEventListener("input", () => {
      if (!currentPayload) return;
      normalizePayloadForSchema(currentPayload);
      const key = nlCategoryKeys[category];
      if (key) currentPayload[key] = el.value || "";
      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }
      updateCategoryHighlight();
      renderItemsTable();
      setStatus(`カテゴリ修正指示（${category}）を反映しました`);
    });
  });

  // ===== siteCosts: UI -> JSON 即時反映 (schema v1.1.12) =====
  function syncSiteCostsToUI() {
    if (!currentPayload) {
      if (siteCostsLaborUnitPrice) siteCostsLaborUnitPrice.value = "";
      if (siteCostsWasteVehicles) siteCostsWasteVehicles.value = "";
      if (siteCostsWasteUnitPrice) siteCostsWasteUnitPrice.value = "";
      if (siteCostsWasteCurrency) siteCostsWasteCurrency.value = "JPY";
      if (siteCostsWasteNotes) siteCostsWasteNotes.value = "";
      if (siteCostsTransportList) siteCostsTransportList.innerHTML = "";
      if (siteCostsLaborList) siteCostsLaborList.innerHTML = "";
      return;
    }

    normalizePayloadForSchema(currentPayload);
    const sc = currentPayload.siteCosts || {};

    const labor = sc.laborCost || {};
    const transport = sc.transportCost || {};
    const waste = sc.wasteDisposalCost || {};
    const transportCost = transport.cost || transport;
    const wasteCost = waste.cost || waste;

    if (siteCostsLaborUnitPrice) siteCostsLaborUnitPrice.value = typeof labor.unitPrice === "number" ? String(labor.unitPrice) : "";

    if (siteCostsWasteVehicles) siteCostsWasteVehicles.value = typeof waste.vehicles === "number" ? String(waste.vehicles) : "";
    if (siteCostsWasteUnitPrice) siteCostsWasteUnitPrice.value = typeof waste.unitPrice === "number" ? String(waste.unitPrice) : "";
    if (siteCostsWasteCurrency) siteCostsWasteCurrency.value = wasteCost.currency || "JPY";
    if (siteCostsWasteNotes) siteCostsWasteNotes.value = waste.notes || "";

    renderTransportList(siteCostsTransportList, transport.entries || []);
    renderLaborList(siteCostsLaborList, labor.entries || []);
  }

  function syncSiteCostsFormVisibility() {
    const isSiteCosts = selectedCategoryKey === "siteCosts";
    SITE_COST_ITEMS.forEach((item, index) => {
      if (!item.formEl) return;
      item.formEl.hidden = !(isSiteCosts && selectedIndex === index);
    });
  }

  function buildSelect(options, value) {
    const select = document.createElement("select");
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "（未設定）";
    select.appendChild(empty);
    options.forEach((opt) => {
      const el = document.createElement("option");
      el.value = opt;
      el.textContent = opt;
      select.appendChild(el);
    });
    if (typeof value === "string") select.value = value;
    return select;
  }

  function createMaterialRow(entry, index) {
    const row = document.createElement("div");
    row.className = "repeat-row";
    row.dataset.materialRow = "1";
    if (Number.isInteger(index)) row.dataset.index = String(index);

    const kindSelect = buildSelect(MATERIAL_KIND_OPTIONS, entry.kind || "");
    kindSelect.dataset.field = "kind";
    kindSelect.className = "repeat-kind";

    const specInput = document.createElement("input");
    specInput.type = "text";
    specInput.placeholder = "仕様";
    specInput.value = entry.spec || "";
    specInput.dataset.field = "spec";
    specInput.className = "repeat-spec";

    const qtyInput = document.createElement("input");
    qtyInput.type = "number";
    qtyInput.step = "0.01";
    qtyInput.placeholder = "数量";
    qtyInput.value = typeof entry.quantity === "number" ? String(entry.quantity) : "";
    qtyInput.dataset.field = "quantity";
    qtyInput.className = "repeat-small repeat-qty";

    const unitSelect = buildSelect(MATERIAL_UNIT_OPTIONS, entry.unit || "");
    unitSelect.dataset.field = "unit";
    unitSelect.className = "repeat-small repeat-unit";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "-";
    removeBtn.dataset.action = "remove-material";

    row.appendChild(kindSelect);
    row.appendChild(specInput);
    row.appendChild(qtyInput);
    row.appendChild(unitSelect);
    row.appendChild(removeBtn);
    return row;
  }

  function createFinishRow(entry, index) {
    const row = document.createElement("div");
    row.className = "repeat-row";
    row.dataset.finishRow = "1";
    if (Number.isInteger(index)) row.dataset.index = String(index);

    const kindSelect = buildSelect(FINISH_KIND_OPTIONS, entry.kind || "");
    kindSelect.dataset.field = "kind";
    kindSelect.className = "repeat-kind";

    const specInput = document.createElement("input");
    specInput.type = "text";
    specInput.placeholder = "仕様";
    specInput.value = entry.spec || "";
    specInput.dataset.field = "spec";
    specInput.className = "repeat-spec";

    const areaInput = document.createElement("input");
    areaInput.type = "number";
    areaInput.step = "0.01";
    areaInput.placeholder = "面積";
    areaInput.value = typeof entry.surfaceAreaValue === "number" ? String(entry.surfaceAreaValue) : "";
    areaInput.dataset.field = "surfaceArea/value";
    areaInput.className = "repeat-small repeat-qty";

    const unitSelect = buildSelect(FINISH_UNIT_OPTIONS, entry.surfaceAreaUnit || "");
    unitSelect.dataset.field = "surfaceArea/unit";
    unitSelect.className = "repeat-small repeat-unit";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "-";
    removeBtn.dataset.action = "remove-finish";

    row.appendChild(kindSelect);
    row.appendChild(specInput);
    row.appendChild(areaInput);
    row.appendChild(unitSelect);
    row.appendChild(removeBtn);
    return row;
  }

  function renderMaterialsList(list) {
    if (!itemFormMaterialsList) return;
    itemFormMaterialsList.innerHTML = "";
    const items = Array.isArray(list) ? list : [];
    items.forEach((entry, index) => {
      const normalized = typeof entry === "string" ? { kind: entry } : entry || {};
      itemFormMaterialsList.appendChild(
        createMaterialRow({
          kind: normalized.kind || "",
          spec: normalized.spec || "",
          quantity: typeof normalized.quantity === "number" ? normalized.quantity : 1,
          unit: normalized.unit || ""
        }, index)
      );
    });
  }

  function renderFinishesList(list) {
    if (!itemFormFinishesList) return;
    itemFormFinishesList.innerHTML = "";
    const items = Array.isArray(list) ? list : [];
    items.forEach((entry, index) => {
      const normalized = typeof entry === "string" ? { kind: entry } : entry || {};
      const surfaceArea = normalized.surfaceArea || {};
      itemFormFinishesList.appendChild(
        createFinishRow({
          kind: normalized.kind || "",
          spec: normalized.spec || "",
          surfaceAreaValue: typeof surfaceArea.value === "number" ? surfaceArea.value : null,
          surfaceAreaUnit: typeof surfaceArea.unit === "string" ? surfaceArea.unit : ""
        }, index)
      );
    });
  }

  function readMaterialsList() {
    if (!itemFormMaterialsList) return [];
    const rows = itemFormMaterialsList.querySelectorAll("[data-material-row]");
    const items = [];
    rows.forEach((row) => {
      const get = (field) => row.querySelector(`[data-field="${field}"]`);
      const kind = (get("kind") && get("kind").value) || "";
      if (!kind) return;
      const spec = (get("spec") && get("spec").value) || "";
      const qtyRaw = get("quantity") && get("quantity").value;
      const qtyVal = parseFloat(String(qtyRaw || ""));
      const quantity = Number.isFinite(qtyVal) && qtyVal > 0 ? qtyVal : 1;
      const unit = (get("unit") && get("unit").value) || "式";
      items.push({ kind, spec, quantity, unit });
    });
    return items;
  }

  function readFinishesList() {
    if (!itemFormFinishesList) return [];
    const rows = itemFormFinishesList.querySelectorAll("[data-finish-row]");
    const items = [];
    rows.forEach((row) => {
      const get = (field) => row.querySelector(`[data-field="${field}"]`);
      const kind = (get("kind") && get("kind").value) || "";
      if (!kind) return;
      const spec = (get("spec") && get("spec").value) || "";
      const areaRaw = get("surfaceArea/value") && get("surfaceArea/value").value;
      const areaVal = parseFloat(String(areaRaw || ""));
      const area = Number.isFinite(areaVal) && areaVal > 0 ? areaVal : null;
      const unit = (get("surfaceArea/unit") && get("surfaceArea/unit").value) || "";
      const entry = { kind, spec };
      if (area != null || unit) {
        entry.surfaceArea = { value: area != null ? area : 0.01, unit: unit || "m²", notes: "" };
      }
      items.push(entry);
    });
    return items;
  }

  function applySiteCostsFromUI() {
    if (!currentPayload) return;
    normalizePayloadForSchema(currentPayload);
    const toNum = (v) => {
      const n = parseFloat(String(v || ""));
      return Number.isFinite(n) ? n : null;
    };

    const laborUnitPrice = siteCostsLaborUnitPrice ? toNum(siteCostsLaborUnitPrice.value) : null;
    const laborCostNotes = currentPayload.siteCosts.laborCost.cost?.notes || "";
    const laborEntries = readLaborEntries(siteCostsLaborList);
    if (laborEntries.length === 0) {
      laborEntries.push({ label: "作業", people: 0, days: 0, notes: "" });
    }
    const transportCostNotes = currentPayload.siteCosts.transportCost.cost?.notes || "";
    const transportEntries = readTransportEntries(siteCostsTransportList);
    const wasteVehicles = siteCostsWasteVehicles ? toNum(siteCostsWasteVehicles.value) : null;
    const wasteUnitPrice = siteCostsWasteUnitPrice ? toNum(siteCostsWasteUnitPrice.value) : null;
    const wasteCostNotes = currentPayload.siteCosts.wasteDisposalCost.cost?.notes || "";

    const laborUnit = laborUnitPrice ?? 0;
    const laborTotal = laborEntries.reduce((sum, entry) => sum + (entry.people || 0) * (entry.days || 0), 0);
    const laborAmount = laborUnit * laborTotal;

    if (transportEntries.length === 0) {
      transportEntries.push({
        label: "運搬",
        vehicleType: "",
        vehicles: 0,
        days: 0,
        unitPrice: 0,
        cost: { amount: 0, currency: "JPY", notes: "" },
        notes: ""
      });
    }

    const transportAmount = transportEntries.reduce((sum, entry) => {
      const amount = entry && entry.cost && typeof entry.cost.amount === "number" ? entry.cost.amount : 0;
      return sum + amount;
    }, 0);

    const wasteAmount = (wasteVehicles ?? 0) * (wasteUnitPrice ?? 0);

    currentPayload.siteCosts.laborCost.unitPrice = laborUnitPrice ?? 0;
    currentPayload.siteCosts.laborCost.entries = laborEntries;
    currentPayload.siteCosts.laborCost.cost = {
      amount: laborAmount ?? 0,
      currency: "JPY",
      notes: laborCostNotes
    };

    currentPayload.siteCosts.transportCost.entries = transportEntries;
    currentPayload.siteCosts.transportCost.cost = {
      amount: transportAmount ?? 0,
      currency: "JPY",
      notes: transportCostNotes
    };

    currentPayload.siteCosts.wasteDisposalCost.vehicles = wasteVehicles ?? 0;
    currentPayload.siteCosts.wasteDisposalCost.unitPrice = wasteUnitPrice ?? 0;
    currentPayload.siteCosts.wasteDisposalCost.cost = {
      amount: wasteAmount ?? 0,
      currency: (siteCostsWasteCurrency && siteCostsWasteCurrency.value) || "JPY",
      notes: wasteCostNotes
    };
    currentPayload.siteCosts.wasteDisposalCost.notes = (siteCostsWasteNotes && siteCostsWasteNotes.value) || "";

    if (basePayload && currentPayload) {
      userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
    }

    updateCategoryHighlight();
    syncSiteCostEditorForSelection();
    renderItemsTable();
    setStatus("現場費（siteCosts）を反映しました");
  }

  [
    siteCostsLaborUnitPrice,
    siteCostsWasteVehicles,
    siteCostsWasteUnitPrice,
    siteCostsWasteNotes
  ].forEach((el) => {
    if (!el) return;
    const evt = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
    el.addEventListener(evt, () => {
      if (!currentPayload) return;
      applySiteCostsFromUI();
    });
  });

  if (siteCostsLaborList) {
    siteCostsLaborList.addEventListener("input", () => {
      if (!currentPayload) return;
      applySiteCostsFromUI();
    });
    siteCostsLaborList.addEventListener("change", () => {
      if (!currentPayload) return;
      applySiteCostsFromUI();
    });
  }

  function renderTransportList(container, entries) {
    if (!container) return;
    container.innerHTML = "";
    const list = Array.isArray(entries) ? entries : [];
    list.forEach((entry, index) => {
      const row = document.createElement("div");
      row.className = "sitecosts-transport-grid";
      row.dataset.transportRow = "1";
      row.dataset.transportIndex = String(index);

      const labelInput = document.createElement("input");
      labelInput.type = "text";
      labelInput.placeholder = "ラベル";
      labelInput.value = typeof entry.label === "string" ? entry.label : "";

      const typeInput = document.createElement("input");
      typeInput.type = "text";
      typeInput.placeholder = "車種";
      typeInput.value = typeof entry.vehicleType === "string" ? entry.vehicleType : "";

      const vehiclesInput = document.createElement("input");
      vehiclesInput.type = "number";
      vehiclesInput.step = "1";
      vehiclesInput.placeholder = "台数";
      vehiclesInput.value = typeof entry.vehicles === "number" ? String(entry.vehicles) : "";

      const daysInput = document.createElement("input");
      daysInput.type = "number";
      daysInput.step = "0.5";
      daysInput.placeholder = "日数";
      daysInput.value = typeof entry.days === "number" ? String(entry.days) : "";

      const unitPriceInput = document.createElement("input");
      unitPriceInput.type = "number";
      unitPriceInput.step = "1";
      unitPriceInput.placeholder = "単価(JPY)";
      unitPriceInput.value = typeof entry.unitPrice === "number" ? String(entry.unitPrice) : "";

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "削除";
      removeBtn.addEventListener("click", () => {
        if (!currentPayload) return;
        normalizePayloadForSchema(currentPayload);
        const target = currentPayload.siteCosts.transportCost.entries;
        if (Array.isArray(target)) target.splice(index, 1);
        if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
        syncSiteCostsToUI();
        updateCategoryHighlight();
        renderItemsTable();
      });

      row.appendChild(labelInput);
      row.appendChild(typeInput);
      row.appendChild(vehiclesInput);
      row.appendChild(daysInput);
      row.appendChild(unitPriceInput);
      row.appendChild(removeBtn);
      container.appendChild(row);
    });
  }

  function renderLaborList(container, entries) {
    if (!container) return;
    container.innerHTML = "";
    const list = Array.isArray(entries) ? entries : [];
    list.forEach((entry, index) => {
      const row = document.createElement("div");
      row.className = "sitecosts-labor-grid";
      row.dataset.laborRow = "1";
      row.dataset.laborIndex = String(index);

      const labelInput = document.createElement("input");
      labelInput.type = "text";
      labelInput.placeholder = "ラベル";
      labelInput.value = typeof entry.label === "string" ? entry.label : "";

      const peopleInput = document.createElement("input");
      peopleInput.type = "number";
      peopleInput.step = "1";
      peopleInput.placeholder = "人数";
      peopleInput.value = typeof entry.people === "number" ? String(entry.people) : "";

      const daysInput = document.createElement("input");
      daysInput.type = "number";
      daysInput.step = "0.5";
      daysInput.placeholder = "日数";
      daysInput.value = typeof entry.days === "number" ? String(entry.days) : "";

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "削除";
      removeBtn.addEventListener("click", () => {
        if (!currentPayload) return;
        normalizePayloadForSchema(currentPayload);
        const target = currentPayload.siteCosts.laborCost.entries;
        if (Array.isArray(target)) target.splice(index, 1);
        if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
        syncSiteCostsToUI();
        updateCategoryHighlight();
        renderItemsTable();
      });

      row.appendChild(labelInput);
      row.appendChild(peopleInput);
      row.appendChild(daysInput);
      row.appendChild(removeBtn);
      container.appendChild(row);
    });
  }

  function readLaborEntries(container) {
    if (!container) return [];
    const rows = container.querySelectorAll("[data-labor-row]");
    const entries = [];
    rows.forEach((row) => {
      const inputs = row.querySelectorAll("input");
      const getVal = (i) => (inputs[i] ? inputs[i].value : "");
      const toNum = (v) => {
        const n = parseFloat(String(v || ""));
        return Number.isFinite(n) ? n : 0;
      };
      const label = getVal(0);
      const people = toNum(getVal(1));
      const days = toNum(getVal(2));
      if (!label) return;
      entries.push({ label, people, days, notes: "" });
    });
    return entries;
  }

  function readTransportEntries(container) {
    if (!container) return [];
    const rows = container.querySelectorAll("[data-transport-row]");
    const entries = [];
    rows.forEach((row) => {
      const inputs = row.querySelectorAll("input");
      const getVal = (i) => (inputs[i] ? inputs[i].value : "");
      const toNum = (v) => {
        const n = parseFloat(String(v || ""));
        return Number.isFinite(n) ? n : 0;
      };
      const label = getVal(0);
      const vehicleType = getVal(1);
      const vehicles = toNum(getVal(2));
      const days = toNum(getVal(3));
      const unitPrice = toNum(getVal(4));
      const amount = vehicles * days * unitPrice;
      if (!label) return;
      entries.push({
        label,
        vehicleType: vehicleType || "",
        vehicles,
        days,
        unitPrice,
        cost: { amount, currency: "JPY", notes: "" },
        notes: ""
      });
    });
    return entries;
  }

  if (siteCostsTransportList) {
    siteCostsTransportList.addEventListener("input", () => {
      if (!currentPayload) return;
      applySiteCostsFromUI();
    });
  }

  if (siteCostsTransportAdd) {
    siteCostsTransportAdd.addEventListener("click", () => {
      if (!currentPayload) return;
      normalizePayloadForSchema(currentPayload);
      if (!Array.isArray(currentPayload.siteCosts.transportCost.entries)) {
        currentPayload.siteCosts.transportCost.entries = [];
      }
      currentPayload.siteCosts.transportCost.entries.push({
        label: "",
        vehicleType: "",
        vehicles: 0,
        days: 0,
        unitPrice: 0,
        cost: { amount: 0, currency: "JPY", notes: "" },
        notes: ""
      });
      if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      syncSiteCostsToUI();
      updateCategoryHighlight();
      renderItemsTable();
    });
  }

  if (siteCostsLaborAdd) {
    siteCostsLaborAdd.addEventListener("click", () => {
      if (!currentPayload) return;
      normalizePayloadForSchema(currentPayload);
      if (!Array.isArray(currentPayload.siteCosts.laborCost.entries)) {
        currentPayload.siteCosts.laborCost.entries = [];
      }
      currentPayload.siteCosts.laborCost.entries.push({
        label: "",
        people: 0,
        days: 0,
        notes: ""
      });
      if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      syncSiteCostsToUI();
      updateCategoryHighlight();
      renderItemsTable();
    });
  }

  if (pageTabExtract) {
    pageTabExtract.addEventListener("click", () => setPageMode("extract"));
  }
  if (pageTabEstimate) {
    pageTabEstimate.addEventListener("click", () => setPageMode("estimate"));
  }

  if (btnEstimateGenerate) {
    btnEstimateGenerate.addEventListener("click", () => {
      if (!currentPayload) {
        alert("先に抽出JSONを読み込んでください。");
        return;
      }
      estimatePayload = buildEstimateFromExtraction(currentPayload);
      estimateSelectedGroupIndex = 0;
      syncEstimateToUI();
      setPageMode("estimate");
    });
  }

  [
    estimateTitleInput,
    estimateClientInput,
    estimateTaxRateInput,
    estimateTaxIncludedInput,
    estimateBreakdownTitleInput
  ].forEach((el) => {
    if (!el) return;
    const evt = el.type === "checkbox" ? "change" : "input";
    el.addEventListener(evt, () => {
      if (!estimatePayload) ensureEstimatePayload();
      applyEstimateHeaderFromUI();
    });
  });

  [
    estimateGroupNoInput,
    estimateGroupCategoryInput,
    estimateGroupDisplaySelect,
    estimateGroupNotesInput,
    estimateGroupMarginInput,
    estimateGroupHiddenInput,
    estimateGroupPricingNotesInput,
    estimateSummaryQtyInput,
    estimateSummaryUnitInput,
    estimateSummaryUnitPriceInput,
    estimateSummaryAmountInput
  ].forEach((el) => {
    if (!el) return;
    const evt = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
    el.addEventListener(evt, () => {
      applyEstimateGroupFromUI();
      if (estimateGroupDisplaySelect === el) {
        syncEstimateGroupEditor();
      }
    });
  });

  if (estimateLineItemsList) {
    estimateLineItemsList.addEventListener("input", () => {
      applyEstimateGroupFromUI();
    });
  }

  if (btnEstimateLineAdd) {
    btnEstimateLineAdd.addEventListener("click", () => {
      ensureEstimatePayload();
      const group = estimatePayload.breakdown.groups?.[estimateSelectedGroupIndex];
      if (!group) return;
      if (!Array.isArray(group.lineItems)) group.lineItems = [];
      group.lineItems.push(createDefaultLineItem());
      renderEstimateLineItems(estimateLineItemsList, group.lineItems);
      applyEstimateGroupFromUI();
    });
  }

  if (btnEstimateGroupAdd) {
    btnEstimateGroupAdd.addEventListener("click", () => {
      ensureEstimatePayload();
      const groupNo = estimatePayload.breakdown.groups.length + 1;
      estimatePayload.breakdown.groups.push(createDefaultEstimateGroup(groupNo, ""));
      estimateSelectedGroupIndex = estimatePayload.breakdown.groups.length - 1;
      renderEstimateGroups();
      syncEstimateGroupEditor();
      recalcEstimateTotals();
      syncEstimateJson();
    });
  }

  if (btnEstimateGroupRemove) {
    btnEstimateGroupRemove.addEventListener("click", () => {
      ensureEstimatePayload();
      if (!Number.isInteger(estimateSelectedGroupIndex)) return;
      estimatePayload.breakdown.groups.splice(estimateSelectedGroupIndex, 1);
      if (estimatePayload.breakdown.groups.length === 0) {
        estimatePayload.breakdown.groups.push(createDefaultEstimateGroup(1, ""));
        estimateSelectedGroupIndex = 0;
      } else if (estimateSelectedGroupIndex >= estimatePayload.breakdown.groups.length) {
        estimateSelectedGroupIndex = estimatePayload.breakdown.groups.length - 1;
      }
      renderEstimateGroups();
      syncEstimateGroupEditor();
      recalcEstimateTotals();
      syncEstimateJson();
    });
  }

  if (btnEstimateApplyJson) {
    btnEstimateApplyJson.addEventListener("click", () => {
      if (!estimateJsonTextarea) return;
      try {
        const parsed = JSON.parse(estimateJsonTextarea.value || "{}");
        estimatePayload = parsed;
        estimateSelectedGroupIndex = 0;
        syncEstimateToUI();
      } catch (err) {
        alert("JSONの解析に失敗しました。構文を確認してください。");
        console.error(err);
      }
    });
  }

  if (btnEstimateExportXlsx) {
    btnEstimateExportXlsx.addEventListener("click", async () => {
      if (!estimatePayload) {
        alert("見積書JSONがありません。先に下書き作成またはJSON反映を行ってください。");
        return;
      }
      if (!window.api || !window.api.exportEstimateXlsx) {
        alert("xlsx出力APIが利用できません。");
        return;
      }
      try {
        const res = await window.api.exportEstimateXlsx(estimatePayload);
        if (!res || !res.ok) {
          alert(res?.error || "xlsx出力に失敗しました。");
          return;
        }
        setStatus(`xlsxを出力しました: ${res.path}`);
      } catch (err) {
        console.error(err);
        alert("xlsx出力中にエラーが発生しました。");
      }
    });
  }

  const siteCostsSummaries = document.querySelectorAll(".sitecosts-drawer summary");
  siteCostsSummaries.forEach((summary) => {
    summary.addEventListener("click", (e) => {
      const tag = e.target && e.target.tagName;
      if (tag === "INPUT" || tag === "LABEL" || tag === "BUTTON") {
        e.preventDefault();
      }
    });
  });

  function syncNlCategoryToUI() {
    Object.values(nlCategoryTextareas).forEach((el) => {
      if (el) el.value = "";
    });
    if (!currentPayload) return;
    normalizePayloadForSchema(currentPayload);
    Object.entries(nlCategoryTextareas).forEach(([category, el]) => {
      if (!el) return;
      const key = nlCategoryKeys[category];
      el.value = typeof currentPayload[key] === "string" ? currentPayload[key] : "";
    });
  }

  function updateCategoryHighlight() {
    Object.values(nlCategoryTextareas).forEach((el) => {
      if (!el) return;
      el.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
    });
    if (!isHighlightEnabled()) return;
    const key = nlCategoryKeys[selectedCategoryKey];
    const el = nlCategoryTextareas[selectedCategoryKey];
    if (!el || !key) return;
    const prefix = `/${key}`;
    if (hasChangeInSet(userChangedPathsCurrent, prefix)) el.classList.add("user-changed-current");
    else if (hasChangeInSet(userChangedPathsPrevious, prefix)) el.classList.add("user-changed-previous");
    else if (hasChangeInSet(apiChangedPaths, prefix)) el.classList.add("api-changed");
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
      if (selectedCategoryKey === "siteCosts") selectedIndex = 0;
      if (itemEditor) itemEditor.value = "";
      resetForm();

      syncNlCategoryToUI();
      renderCategoryTabs();
      setFormModeForCategory();
      syncSiteCostsToUI();
      syncSiteCostsFormVisibility();
      syncSiteCostEditorForSelection();
      updateCategoryHighlight();

      renderItemsTable();
      updateSourceInfoUI();
      setStatus(`JSONを取り込みました（${selectedCategoryKey}=${getItems().length}件）`);
      updateRestoreButtonEnabled();
      updateRedoButtonEnabled();
    });
  }

  function getItems() {
    if (!currentPayload) return [];
    if (selectedCategoryKey === "siteCosts") return getSiteCostItems();
    const arr = currentPayload[selectedCategoryKey];
    return Array.isArray(arr) ? arr : [];
  }

  function getSiteCostItems() {
    if (!currentPayload) return [];
    normalizePayloadForSchema(currentPayload);
    return SITE_COST_ITEMS.map((item) => {
      const entry = currentPayload.siteCosts?.[item.key] || {};
      return {
        key: item.key,
        label: item.label,
        cost: entry.cost || entry
      };
    });
  }

  function getSiteCostEntry(index) {
    if (!currentPayload) return null;
    normalizePayloadForSchema(currentPayload);
    const meta = SITE_COST_ITEMS[index];
    if (!meta) return null;
    return currentPayload.siteCosts?.[meta.key] || null;
  }

  function syncSiteCostEditorForSelection() {
    if (selectedCategoryKey !== "siteCosts") return;
    const meta = SITE_COST_ITEMS[selectedIndex] || null;
    const entry = getSiteCostEntry(selectedIndex);
    if (itemEditor) itemEditor.value = entry ? JSON.stringify(entry, null, 2) : "";
    if (itemEditorLabel) itemEditorLabel.textContent = meta ? `項目: ${meta.label}` : "項目: -";
  }

  // ★手動追加（空の項目を追加）
  if (btnItemAdd) {
    btnItemAdd.addEventListener("click", () => {
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
      syncNlCategoryToUI();
      updateCategoryHighlight();

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

  // ★削除（選択項目）
  if (btnItemRemove) {
    btnItemRemove.addEventListener("click", () => {
      if (!currentPayload) {
        alert("先に ChatGPT 出力JSONを取り込んでください。");
        return;
      }
      if (selectedCategoryKey === "siteCosts") {
        alert("現場費（siteCosts）は削除できません。");
        return;
      }
      if (selectedIndex == null) {
        alert("削除する項目をリストから選択してください。");
        return;
      }
      if (!Array.isArray(currentPayload[selectedCategoryKey])) {
        alert(`現在のJSONに ${selectedCategoryKey} がありません。`);
        return;
      }

      pushStateHistory(true);
      currentPayload[selectedCategoryKey].splice(selectedIndex, 1);

      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }

      selectedIndex = null;
      if (itemEditor) itemEditor.value = "";
      resetForm();
      renderItemsTable();
      updateCategoryHighlight();
      updateFormHighlightsForSelected();
      setStatus("選択中の項目を削除しました");
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
      const items = getSiteCostItems();
      const version = (currentPayload && currentPayload.version) || "";
      const stage = (currentPayload && currentPayload.stage) || "";
      const nlKey = nlCategoryKeys.siteCosts;
      const hasCategoryCorrection = nlKey && currentPayload && typeof currentPayload[nlKey] === "string";

      let html = "";
      html += `<div style="margin-bottom:4px;">`;
      html += `<strong>version:</strong> ${escapeHtml(version)} / <strong>stage:</strong> ${escapeHtml(stage)} / `;
      html += `<strong>nlCorrection:</strong> ${hasCategoryCorrection ? "OK" : "MISSING"} / `;
      html += `<strong>category:</strong> siteCosts / <strong>count:</strong> ${items.length}件</div>`;

      html += `<table>`;
      html += `<thead><tr><th>#</th><th>項目</th><th>合計</th></tr></thead><tbody>`;

      const highlightOn = isHighlightEnabled();
      items.forEach((item, index) => {
        const rowClass = index === selectedIndex ? "selected" : "";
        const amount = typeof item.cost?.amount === "number" ? item.cost.amount : "";
        const basePath = `/siteCosts/${item.key}`;
        const clsAmount = highlightOn ? classForPathPrefix(`${basePath}/cost`) : "";
        html += `<tr data-index="${index}" class="${rowClass}">
          <td>${index + 1}</td>
          <td>${escapeHtml(item.label)}</td>
          <td class="${clsAmount}">${escapeHtml(amount)}</td>
        </tr>`;
      });

      html += `</tbody></table>`;
      itemsView.innerHTML = html;
      return;
    }

    const items = getItems();
    const version = (currentPayload && currentPayload.version) || "";
    const stage = (currentPayload && currentPayload.stage) || "";
    const nlKey = nlCategoryKeys[selectedCategoryKey];
    const hasCategoryCorrection = nlKey && currentPayload && typeof currentPayload[nlKey] === "string";

    let html = "";
    html += `<div style="margin-bottom:4px;">`;
    html += `<strong>version:</strong> ${escapeHtml(version)} / <strong>stage:</strong> ${escapeHtml(stage)} / `;
    html += `<strong>nlCorrection:</strong> ${hasCategoryCorrection ? "OK" : "MISSING"} / `;
    const catLabel = (CATEGORIES.find((c) => c.key === selectedCategoryKey) || {}).label || selectedCategoryKey;
    html += `<strong>category:</strong> ${escapeHtml(catLabel)} / <strong>count:</strong> ${items.length}件</div>`;

    {
      const isWoodwork = selectedCategoryKey === "woodworkItems";
      const isElectrical = selectedCategoryKey === "electricalItems";
      html += `<table>`;
      let emptyColspan = 0;
      if (isWoodwork) {
        emptyColspan = 13;
        html += `<thead><tr>
          <th>#</th>
          <th>積算</th>
          <th>名称</th>
          <th>種別</th>
          <th>寸法(H/W/D)</th>
          <th>材料</th>
          <th>仕上げ</th>
          <th>数量</th>
          <th>単価</th>
          <th>材料費</th>
          <th>人工(人日)</th>
          <th>人工費</th>
          <th>合計</th>
        </tr></thead><tbody>`;
      } else if (isElectrical) {
        emptyColspan = 9;
        html += `<thead><tr>
          <th>#</th>
          <th>積算</th>
          <th>名称</th>
          <th>電気種別</th>
          <th>電気仕様</th>
          <th>数量</th>
          <th>単価</th>
          <th>合計</th>
          <th>source</th>
        </tr></thead><tbody>`;
      } else {
        emptyColspan = 7;
        html += `<thead><tr>
          <th>#</th>
          <th>積算</th>
          <th>名称</th>
          <th>数量</th>
          <th>単価</th>
          <th>合計</th>
          <th>source</th>
        </tr></thead><tbody>`;
      }

      if (items.length === 0) {
        html += `<tr><td colspan="${emptyColspan}">items が見つかりませんでした。</td></tr>`;
      }

      items.forEach((item, index) => {
        const name = item.name || "";
        const structureType = isWoodwork ? (item.structureType || item.type || "") : "";
        const electricalType = isElectrical ? (item.electricalType || "") : "";

        const dims = isWoodwork ? item.dimensions || {} : {};
        const dimParts = [];
        if (typeof dims.height === "number") dimParts.push(`H${dims.height}`);
        if (typeof dims.width === "number") dimParts.push(`W${dims.width}`);
        if (typeof dims.depth === "number") dimParts.push(`D${dims.depth}`);
        const dimStr = dimParts.join(" ");

        const materials = isWoodwork && Array.isArray(item.materials) ? item.materials : [];
        const materialStr = materials
          .map((m) => {
            if (!m) return "";
            if (typeof m === "string") return m;
            return m.kind || "";
          })
          .filter((s) => s.length > 0)
          .join(", ");

        const finishes = isWoodwork && Array.isArray(item.finishes) ? item.finishes : [];
        const finishesStr = finishes
          .map((f) => {
            if (!f) return "";
            if (typeof f === "string") return f;
            return f.kind || "";
          })
          .filter((s) => s.length > 0)
          .join(", ");

        const quantity = typeof item.quantity === "number" ? item.quantity : "";
        const unit = !isWoodwork && typeof item.unit === "string" ? item.unit : "";
        const quantityDisplay = !isWoodwork && unit ? `${quantity}${unit}` : quantity;

        const unitPriceValue = item.price && typeof item.price.unitPrice === "number" ? item.price.unitPrice : null;
        const unitPriceDisplay = typeof unitPriceValue === "number" ? unitPriceValue : "";

        const materialCostValue = isWoodwork
          ? item.materialCost && typeof item.materialCost.amount === "number"
            ? item.materialCost.amount
            : null
          : item.cost && typeof item.cost.amount === "number"
            ? item.cost.amount
            : null;
        const materialCostDisplay = typeof materialCostValue === "number" ? materialCostValue : "";

        const laborCostValue =
          isWoodwork && item.laborCost && typeof item.laborCost.amount === "number" ? item.laborCost.amount : null;
        const laborCostDisplay = typeof laborCostValue === "number" ? laborCostValue : "";

        let totalValue = null;
        if (typeof unitPriceValue === "number" && typeof item.quantity === "number") {
          totalValue = unitPriceValue * item.quantity;
        } else if (isWoodwork) {
          const material = typeof materialCostValue === "number" ? materialCostValue : 0;
          const labor = typeof laborCostValue === "number" ? laborCostValue : 0;
          totalValue = material + labor;
        } else if (typeof materialCostValue === "number") {
          totalValue = materialCostValue;
        }
        const totalDisplay = typeof totalValue === "number" ? totalValue : "";

        const spec = isElectrical ? item.spec || {} : {};
        let specDisplay = "";
        if (isElectrical) {
          const specParts = [];
          if (typeof spec.voltageV === "number") specParts.push(`${spec.voltageV}V`);
          if (typeof spec.wattW === "number") specParts.push(`${spec.wattW}W`);
          if (typeof spec.phase === "string" && spec.phase)
            specParts.push(spec.phase === "single" ? "単相" : spec.phase === "three" ? "三相" : spec.phase);
          if (typeof spec.breakerA === "number") specParts.push(`${spec.breakerA}A`);
          specDisplay = specParts.join(" / ");
        }

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
        let highlightOnRow = highlightOn;
        if (highlightOn && basePayload && Array.isArray(basePayload[selectedCategoryKey]) && item && item.id) {
          const baseItems = basePayload[selectedCategoryKey];
          const baseIndex = baseItems.findIndex((b) => b && b.id === item.id);
          if (baseIndex !== -1 && baseIndex !== index) {
            highlightOnRow = false;
          }
        }

        const basePath = `/${selectedCategoryKey}/${index}`;
        const clsInclude = highlightOnRow ? classForPathPrefix(`${basePath}/includeInEstimate`) : "";
        const clsName = highlightOnRow ? classForPathPrefix(`${basePath}/name`) : "";
        const clsType = highlightOnRow
          ? isWoodwork
            ? classForPathPrefix(`${basePath}/structureType`)
            : isElectrical
              ? classForPathPrefix(`${basePath}/electricalType`)
              : ""
          : "";
        const clsDim = highlightOnRow
          ? isWoodwork
            ? classForPathPrefix(`${basePath}/dimensions`)
            : isElectrical
              ? classForPathPrefix(`${basePath}/spec`)
              : ""
          : "";
        const clsMat = highlightOnRow && isWoodwork ? classForPathPrefix(`${basePath}/materials`) : "";
        const clsFin = highlightOnRow && isWoodwork ? classForPathPrefix(`${basePath}/finishes`) : "";
        const clsQty = highlightOnRow ? classForPathPrefix(`${basePath}/quantity`) : "";
        const clsUnitPrice = highlightOnRow ? classForPathPrefix(`${basePath}/price/unitPrice`) : "";
        const clsMatCost = highlightOnRow ? classForPathPrefix(`${basePath}/${isWoodwork ? "materialCost" : "cost"}`) : "";
        const clsLaborCost = highlightOnRow && isWoodwork ? classForPathPrefix(`${basePath}/laborCost`) : "";
        const clsTotal = highlightOnRow ? classForPathPrefix(`${basePath}/price`) : "";
        const clsLabor = highlightOnRow && isWoodwork ? classForPathPrefix(`${basePath}/laborTotal`) : "";
        const clsSrc = highlightOnRow ? classForPathPrefix(`${basePath}/source`) : "";

        if (isWoodwork) {
          html += `<tr data-index="${index}" class="${rowClass}">
            <td>${index + 1}</td>
            <td class="${clsInclude}">${escapeHtml(includeInEstimateMark)}</td>
            <td class="${clsName}">${escapeHtml(name)}</td>
            <td class="${clsType}">${escapeHtml(structureType)}</td>
            <td class="${clsDim}">${escapeHtml(dimStr)}</td>
            <td class="${clsMat}">${escapeHtml(materialStr)}</td>
            <td class="${clsFin}">${escapeHtml(finishesStr)}</td>
            <td class="${clsQty}">${escapeHtml(quantity)}</td>
            <td class="${clsUnitPrice}">${escapeHtml(unitPriceDisplay)}</td>
            <td class="${clsMatCost}">${escapeHtml(materialCostDisplay)}</td>
            <td class="${clsLabor}">${escapeHtml(laborTotalDays)}</td>
            <td class="${clsLaborCost}">${escapeHtml(laborCostDisplay)}</td>
            <td class="${clsTotal}">${escapeHtml(totalDisplay)}</td>
          </tr>`;
        } else if (isElectrical) {
          html += `<tr data-index="${index}" class="${rowClass}">
            <td>${index + 1}</td>
            <td class="${clsInclude}">${escapeHtml(includeInEstimateMark)}</td>
            <td class="${clsName}">${escapeHtml(name)}</td>
            <td class="${clsType}">${escapeHtml(electricalType)}</td>
            <td class="${clsDim}">${escapeHtml(specDisplay)}</td>
            <td class="${clsQty}">${escapeHtml(quantityDisplay)}</td>
            <td class="${clsUnitPrice}">${escapeHtml(unitPriceDisplay)}</td>
            <td class="${clsMatCost}">${escapeHtml(totalDisplay)}</td>
            <td class="${clsSrc}">${escapeHtml(src)}</td>
          </tr>`;
        } else {
          html += `<tr data-index="${index}" class="${rowClass}">
            <td>${index + 1}</td>
            <td class="${clsInclude}">${escapeHtml(includeInEstimateMark)}</td>
            <td class="${clsName}">${escapeHtml(name)}</td>
            <td class="${clsQty}">${escapeHtml(quantityDisplay)}</td>
            <td class="${clsUnitPrice}">${escapeHtml(unitPriceDisplay)}</td>
            <td class="${clsMatCost}">${escapeHtml(totalDisplay)}</td>
            <td class="${clsSrc}">${escapeHtml(src)}</td>
          </tr>`;
        }
      });

      html += `</tbody></table>`;
    }

    itemsView.innerHTML = html;
  }

  // 行クリック
  if (itemsView) {
    itemsView.addEventListener("click", (event) => {
      const tr = event.target.closest("tr[data-index]");
      if (!tr) return;
      const index = parseInt(tr.getAttribute("data-index"), 10);
      if (isNaN(index)) return;

      selectedIndex = index;
      if (selectedCategoryKey === "siteCosts") {
        syncSiteCostsToUI();
        syncSiteCostsFormVisibility();
        syncSiteCostEditorForSelection();
        renderItemsTable();
        return;
      }

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
      if (selectedIndex == null) return alert("まずテーブルからアイテムを選択してください。");
      if (!currentPayload) return alert("現在のJSONがありません。");

      let newItem;
      try {
        newItem = JSON.parse(itemEditor.value);
      } catch (err) {
        console.error("item JSON parse error:", err);
        alert("アイテムJSONとして解析できませんでした。構造やカンマを確認してください。");
        return;
      }

      if (selectedCategoryKey === "siteCosts") {
        const meta = SITE_COST_ITEMS[selectedIndex];
        if (!meta) return alert("現場費の選択項目が見つかりません。");
        if (!currentPayload.siteCosts || typeof currentPayload.siteCosts !== "object") currentPayload.siteCosts = {};
        currentPayload.siteCosts[meta.key] = newItem;
        normalizePayloadForSchema(currentPayload);
        syncSiteCostsToUI();
        syncSiteCostsFormVisibility();
        syncSiteCostEditorForSelection();
      } else {
        if (!Array.isArray(currentPayload[selectedCategoryKey]))
          return alert(`現在のJSONに ${selectedCategoryKey} がありません。`);
        currentPayload[selectedCategoryKey][selectedIndex] = newItem;
        updateFormFromItem(newItem);
        updateFormHighlightsForSelected();
      }

      if (basePayload && currentPayload) {
        userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
      }

      updateCategoryHighlight();
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
        if (typeof basePayloadForPrompt.nlCorrectionWoodworkItems !== "string") basePayloadForPrompt.nlCorrectionWoodworkItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionFloorItems !== "string") basePayloadForPrompt.nlCorrectionFloorItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionFinishingItems !== "string") basePayloadForPrompt.nlCorrectionFinishingItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionSignItems !== "string") basePayloadForPrompt.nlCorrectionSignItems = "";
        if (typeof basePayloadForPrompt.nlCollectionSignItems !== "string") basePayloadForPrompt.nlCollectionSignItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionElectricalItems !== "string") basePayloadForPrompt.nlCorrectionElectricalItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionLeaseItems !== "string") basePayloadForPrompt.nlCorrectionLeaseItems = "";
        if (typeof basePayloadForPrompt.nlCorrectionSiteCosts !== "string") basePayloadForPrompt.nlCorrectionSiteCosts = "";

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
    if (itemFormUnit) itemFormUnit.value = "";
    renderMaterialsList([]);
    renderFinishesList([]);
    if (itemFormSignType) itemFormSignType.value = "";
    if (itemFormPriceUnitPrice) itemFormPriceUnitPrice.value = "";
    if (itemFormMaterialCostAmount) itemFormMaterialCostAmount.value = "";
    if (itemFormMaterialCostNotes) itemFormMaterialCostNotes.value = "";
    if (itemFormLaborPerUnitDays) itemFormLaborPerUnitDays.value = "";
    if (itemFormLaborDays) itemFormLaborDays.value = "";
    if (itemFormCostAmount) itemFormCostAmount.value = "";
    if (itemFormCostNotes) itemFormCostNotes.value = "";
    if (itemFormElectricalType) itemFormElectricalType.value = "";
    if (itemFormIsHighPlace) itemFormIsHighPlace.checked = false;
    if (itemFormSpecVoltage) itemFormSpecVoltage.value = "";
    if (itemFormSpecWatt) itemFormSpecWatt.value = "";
    if (itemFormSpecPhase) itemFormSpecPhase.value = "";
    if (itemFormSpecBreaker) itemFormSpecBreaker.value = "";
    if (itemFormSpecNotes) itemFormSpecNotes.value = "";
    if (itemFormNlCorrection) itemFormNlCorrection.value = "";
    if (itemFormIsBent) itemFormIsBent.checked = false;
    if (itemFormSpecialAngles) itemFormSpecialAngles.checked = false;
    if (itemFormSupportHeavy) itemFormSupportHeavy.checked = false;
    if (itemFormFinishAreaValue) itemFormFinishAreaValue.value = "";
    if (itemFormFinishAreaNotes) itemFormFinishAreaNotes.value = "";
    if (itemFormLaborCoefValue) itemFormLaborCoefValue.value = "";
      if (itemFormLaborCoefNotes) itemFormLaborCoefNotes.value = "";
      if (itemFormSignType) itemFormSignType.value = "";
    clearFormHighlightClasses();
  }

  function updateFormFromItem(item) {
    if (!item) return resetForm();

    const isWoodwork = selectedCategoryKey === "woodworkItems";
    const isSign = selectedCategoryKey === "signItems";

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
    if (itemFormUnit) itemFormUnit.value = !isWoodwork && typeof item.unit === "string" ? item.unit : "";
    if (itemFormSignType) itemFormSignType.value = isSign && typeof item.signType === "string" ? item.signType : "";

    if (isWoodwork) {
      renderMaterialsList(item.materials || []);
      renderFinishesList(item.finishes || []);
    } else {
      renderMaterialsList([]);
      renderFinishesList([]);
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

    if (!isWoodwork && itemFormElectricalType) {
      itemFormElectricalType.value = typeof item.electricalType === "string" ? item.electricalType : "";
    }
    if (!isWoodwork && itemFormIsHighPlace) {
      itemFormIsHighPlace.checked = !!item.isHighPlace;
    }
    if (!isWoodwork) {
      const spec = item.spec || {};
      if (itemFormSpecVoltage) itemFormSpecVoltage.value = typeof spec.voltageV === "number" ? String(spec.voltageV) : "";
      if (itemFormSpecWatt) itemFormSpecWatt.value = typeof spec.wattW === "number" ? String(spec.wattW) : "";
      if (itemFormSpecPhase) itemFormSpecPhase.value = typeof spec.phase === "string" ? spec.phase : "";
      if (itemFormSpecBreaker) itemFormSpecBreaker.value = typeof spec.breakerA === "number" ? String(spec.breakerA) : "";
      if (itemFormSpecNotes) itemFormSpecNotes.value = typeof spec.notes === "string" ? spec.notes : "";
    }
  }

  function applyFormToItem(item, changedElement) {
    if (!item) return;
    const isChanged = (el) => !changedElement || changedElement === el;
    const isWoodwork = selectedCategoryKey === "woodworkItems";
    const isSign = selectedCategoryKey === "signItems";

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

    if (!isWoodwork && itemFormUnit && isChanged(itemFormUnit)) {
      if (itemFormUnit.value !== "") item.unit = itemFormUnit.value;
    }
    if (isSign && itemFormSignType && isChanged(itemFormSignType)) {
      if (itemFormSignType.value !== "") item.signType = itemFormSignType.value;
    }


    if (itemFormPriceUnitPrice && isChanged(itemFormPriceUnitPrice)) {
      if (itemFormPriceUnitPrice.value !== "") {
        const v = parseFloat(itemFormPriceUnitPrice.value);
        if (!isNaN(v)) {
          if (!item.price) {
            item.price = { mode: isSign ? "override_fixed" : "estimate_by_model", currency: "JPY", unitPrice: v, notes: "" };
          }
          else item.price.unitPrice = v;
          if (isSign) item.price.mode = "override_fixed";
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

    if (!isWoodwork && itemFormElectricalType && isChanged(itemFormElectricalType)) {
      if (itemFormElectricalType.value !== "") item.electricalType = itemFormElectricalType.value;
    }

    if (!isWoodwork && itemFormIsHighPlace && isChanged(itemFormIsHighPlace)) {
      item.isHighPlace = !!itemFormIsHighPlace.checked;
    }

    if (
      !isWoodwork &&
      (isChanged(itemFormSpecVoltage) ||
        isChanged(itemFormSpecWatt) ||
        isChanged(itemFormSpecPhase) ||
        isChanged(itemFormSpecBreaker) ||
        isChanged(itemFormSpecNotes))
    ) {
      const v = itemFormSpecVoltage && itemFormSpecVoltage.value !== "" ? parseFloat(itemFormSpecVoltage.value) : null;
      const w = itemFormSpecWatt && itemFormSpecWatt.value !== "" ? parseFloat(itemFormSpecWatt.value) : null;
      const b = itemFormSpecBreaker && itemFormSpecBreaker.value !== "" ? parseFloat(itemFormSpecBreaker.value) : null;
      const phase = itemFormSpecPhase && itemFormSpecPhase.value !== "" ? itemFormSpecPhase.value : null;
      const notes = (itemFormSpecNotes && itemFormSpecNotes.value) || "";
      item.spec = {
        voltageV: Number.isFinite(v) ? v : null,
        wattW: Number.isFinite(w) ? w : null,
        breakerA: Number.isFinite(b) ? b : null,
        phase,
        notes
      };
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
      itemFormUnit,
      itemFormSignType,
      itemFormPriceUnitPrice,
      itemFormMaterialCostAmount,
      itemFormMaterialCostNotes,
      itemFormLaborPerUnitDays,
      itemFormLaborDays,
      itemFormCostAmount,
      itemFormCostNotes,
      itemFormElectricalType,
      itemFormIsHighPlace,
      itemFormSpecVoltage,
      itemFormSpecWatt,
      itemFormSpecPhase,
      itemFormSpecBreaker,
      itemFormSpecNotes,
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
        updateCategoryHighlight();
        updateFormHighlightsForSelected();
        setStatus("フォーム変更を反映しました");
      });
    });
  }

  function applyMaterialsFromUI() {
    if (selectedIndex == null) return;
    if (selectedCategoryKey !== "woodworkItems") return;
    if (!currentPayload || !Array.isArray(currentPayload[selectedCategoryKey])) return;

    normalizePayloadForSchema(currentPayload);
    const item = currentPayload[selectedCategoryKey][selectedIndex];
    if (!item) return;

    item.materials = readMaterialsList();

    if (itemEditor) itemEditor.value = JSON.stringify(item, null, 2);
    if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
    renderItemsTable();
    updateCategoryHighlight();
    updateFormHighlightsForSelected();
    setStatus("材料を反映しました");
  }

  function applyFinishesFromUI() {
    if (selectedIndex == null) return;
    if (selectedCategoryKey !== "woodworkItems") return;
    if (!currentPayload || !Array.isArray(currentPayload[selectedCategoryKey])) return;

    normalizePayloadForSchema(currentPayload);
    const item = currentPayload[selectedCategoryKey][selectedIndex];
    if (!item) return;

    const existing = Array.isArray(item.finishes) ? item.finishes : [];
    const next = readFinishesList();
    item.finishes = next.map((entry, idx) => {
      const prev = existing[idx] && typeof existing[idx] === "object" ? existing[idx] : {};
      if (!entry.surfaceArea && prev.surfaceArea) entry.surfaceArea = prev.surfaceArea;
      if (!entry.notes && typeof prev.notes === "string") entry.notes = prev.notes;
      if (entry.surfaceArea && prev.surfaceArea && typeof prev.surfaceArea.notes === "string") {
        entry.surfaceArea.notes = entry.surfaceArea.notes || prev.surfaceArea.notes;
      }
      return entry;
    });

    if (itemEditor) itemEditor.value = JSON.stringify(item, null, 2);
    if (basePayload && currentPayload) userChangedPathsCurrent = collectDiffPaths(basePayload, currentPayload);
    renderItemsTable();
    updateCategoryHighlight();
    updateFormHighlightsForSelected();
    setStatus("仕上げを反映しました");
  }

  if (itemFormMaterialsAdd) {
    itemFormMaterialsAdd.addEventListener("click", () => {
      if (!itemFormMaterialsList) return;
      if (currentPayload && selectedCategoryKey === "woodworkItems") pushStateHistory(true);
      itemFormMaterialsList.appendChild(createMaterialRow({ kind: "", spec: "", quantity: 1, unit: "" }));
      applyMaterialsFromUI();
    });
  }

  if (itemFormFinishesAdd) {
    itemFormFinishesAdd.addEventListener("click", () => {
      if (!itemFormFinishesList) return;
      if (currentPayload && selectedCategoryKey === "woodworkItems") pushStateHistory(true);
      itemFormFinishesList.appendChild(createFinishRow({ kind: "", spec: "", surfaceAreaValue: null, surfaceAreaUnit: "" }));
      applyFinishesFromUI();
    });
  }

  if (itemFormMaterialsList) {
    itemFormMaterialsList.addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-action=\"remove-material\"]");
      if (!btn) return;
      if (currentPayload && selectedCategoryKey === "woodworkItems") pushStateHistory(true);
      const row = btn.closest("[data-material-row]");
      if (row) row.remove();
      applyMaterialsFromUI();
    });
    itemFormMaterialsList.addEventListener("input", applyMaterialsFromUI);
    itemFormMaterialsList.addEventListener("change", applyMaterialsFromUI);
  }

  if (itemFormFinishesList) {
    itemFormFinishesList.addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-action=\"remove-finish\"]");
      if (!btn) return;
      if (currentPayload && selectedCategoryKey === "woodworkItems") pushStateHistory(true);
      const row = btn.closest("[data-finish-row]");
      if (row) row.remove();
      applyFinishesFromUI();
    });
    itemFormFinishesList.addEventListener("input", applyFinishesFromUI);
    itemFormFinishesList.addEventListener("change", applyFinishesFromUI);
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
      itemFormUnit,
      itemFormMaterialsList,
      itemFormFinishesList,
      itemFormSignType,
      itemFormPriceUnitPrice,
      itemFormMaterialCostAmount,
      itemFormMaterialCostNotes,
      itemFormLaborPerUnitDays,
      itemFormLaborDays,
      itemFormCostAmount,
      itemFormCostNotes,
      itemFormElectricalType,
      itemFormIsHighPlace,
      itemFormSpecVoltage,
      itemFormSpecWatt,
      itemFormSpecPhase,
      itemFormSpecBreaker,
      itemFormSpecNotes,
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
    if (itemFormMaterialsList) {
      itemFormMaterialsList.querySelectorAll("[data-field]").forEach((el) => {
        el.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
      });
    }
    if (itemFormFinishesList) {
      itemFormFinishesList.querySelectorAll("[data-field]").forEach((el) => {
        el.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
      });
    }
  }

  function updateFormHighlightsForSelected() {
    clearFormHighlightClasses();
    if (!isHighlightEnabled()) return;
    if (selectedIndex == null) return;
    if (selectedCategoryKey === "siteCosts") return;

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
    if (!isWoodwork) apply(itemFormUnit, `${basePath}/unit`);
    if (isWoodwork) {
      const baseItem = basePayload && basePayload[selectedCategoryKey] && basePayload[selectedCategoryKey][idx];
      const baseMaterials = Array.isArray(baseItem?.materials) ? baseItem.materials : [];
      const baseFinishes = Array.isArray(baseItem?.finishes) ? baseItem.finishes : [];
      const currItem = currentPayload && currentPayload[selectedCategoryKey] && currentPayload[selectedCategoryKey][idx];
      const currMaterials = Array.isArray(currItem?.materials) ? currItem.materials : [];
      const currFinishes = Array.isArray(currItem?.finishes) ? currItem.finishes : [];

      const materialEquals = (a, b) => {
        if (!a || !b) return false;
        if (typeof a === "string" || typeof b === "string") return String(a) === String(b);
        return (
          (a.kind || "") === (b.kind || "") &&
          (a.spec || "") === (b.spec || "") &&
          (typeof a.quantity === "number" ? a.quantity : null) === (typeof b.quantity === "number" ? b.quantity : null) &&
          (a.unit || "") === (b.unit || "")
        );
      };

      const finishEquals = (a, b) => {
        if (!a || !b) return false;
        if (typeof a === "string" || typeof b === "string") return String(a) === String(b);
        const aArea = a.surfaceArea || {};
        const bArea = b.surfaceArea || {};
        return (
          (a.kind || "") === (b.kind || "") &&
          (a.spec || "") === (b.spec || "") &&
          (typeof aArea.value === "number" ? aArea.value : null) === (typeof bArea.value === "number" ? bArea.value : null) &&
          (aArea.unit || "") === (bArea.unit || "")
        );
      };

      const findMatchIndex = (list, entry, eqFn) => {
        for (let i = 0; i < list.length; i++) {
          if (eqFn(list[i], entry)) return i;
        }
        return -1;
      };

      if (itemFormMaterialsList) {
        itemFormMaterialsList.querySelectorAll("[data-material-row]").forEach((row) => {
          const idx = row.dataset.index;
          if (typeof idx !== "string") return;
          const rowIndex = parseInt(idx, 10);
          if (!Number.isFinite(rowIndex)) return;
          const currEntry = currMaterials[rowIndex];
          const matchIndex = currEntry ? findMatchIndex(baseMaterials, currEntry, materialEquals) : -1;
          const isShifted = matchIndex !== -1 && matchIndex !== rowIndex;
          const isAdded = rowIndex >= baseMaterials.length;
          row.querySelectorAll("[data-field]").forEach((fieldEl) => {
            fieldEl.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
            const field = fieldEl.dataset.field;
            if (!field) return;
            const prefix = `${basePath}/materials/${idx}/${field}`;
            if (isShifted) return;
            if (hasChangeInSet(userChangedPathsCurrent, prefix) || (isAdded && hasChangeInSet(userChangedPathsCurrent, `${basePath}/materials/-`))) {
              fieldEl.classList.add("user-changed-current");
            } else if (hasChangeInSet(userChangedPathsPrevious, prefix)) {
              fieldEl.classList.add("user-changed-previous");
            } else if (hasChangeInSet(apiChangedPaths, prefix) || (isAdded && hasChangeInSet(apiChangedPaths, `${basePath}/materials/-`))) {
              fieldEl.classList.add("api-changed");
            }
          });
        });
      }
      if (itemFormFinishesList) {
        itemFormFinishesList.querySelectorAll("[data-finish-row]").forEach((row) => {
          const idx = row.dataset.index;
          if (typeof idx !== "string") return;
          const rowIndex = parseInt(idx, 10);
          if (!Number.isFinite(rowIndex)) return;
          const currEntry = currFinishes[rowIndex];
          const matchIndex = currEntry ? findMatchIndex(baseFinishes, currEntry, finishEquals) : -1;
          const isShifted = matchIndex !== -1 && matchIndex !== rowIndex;
          const isAdded = rowIndex >= baseFinishes.length;
          row.querySelectorAll("[data-field]").forEach((fieldEl) => {
            fieldEl.classList.remove("user-changed-current", "user-changed-previous", "api-changed");
            const field = fieldEl.dataset.field;
            if (!field) return;
            const prefix = `${basePath}/finishes/${idx}/${field}`;
            if (isShifted) return;
            if (hasChangeInSet(userChangedPathsCurrent, prefix) || (isAdded && hasChangeInSet(userChangedPathsCurrent, `${basePath}/finishes/-`))) {
              fieldEl.classList.add("user-changed-current");
            } else if (hasChangeInSet(userChangedPathsPrevious, prefix)) {
              fieldEl.classList.add("user-changed-previous");
            } else if (hasChangeInSet(apiChangedPaths, prefix) || (isAdded && hasChangeInSet(apiChangedPaths, `${basePath}/finishes/-`))) {
              fieldEl.classList.add("api-changed");
            }
          });
        });
      }
    }
    if (selectedCategoryKey === "signItems") apply(itemFormSignType, `${basePath}/signType`);
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
    } else if (selectedCategoryKey === "electricalItems") {
      apply(itemFormElectricalType, `${basePath}/electricalType`);
      apply(itemFormIsHighPlace, `${basePath}/isHighPlace`);
      apply(itemFormSpecVoltage, `${basePath}/spec/voltageV`);
      apply(itemFormSpecWatt, `${basePath}/spec/wattW`);
      apply(itemFormSpecPhase, `${basePath}/spec/phase`);
      apply(itemFormSpecBreaker, `${basePath}/spec/breakerA`);
      apply(itemFormSpecNotes, `${basePath}/spec/notes`);
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

  function joinPath(base, segment) {
    return base === "" ? "/" + segment : base + "/" + segment;
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
      if (path.endsWith("materials") || path.endsWith("finishes")) {
        const eqMaterial = (a, d) => {
          if (!a || !d) return false;
          if (typeof a === "string" || typeof d === "string") return String(a) === String(d);
          return (
            (a.kind || "") === (d.kind || "") &&
            (a.spec || "") === (d.spec || "") &&
            (typeof a.quantity === "number" ? a.quantity : null) === (typeof d.quantity === "number" ? d.quantity : null) &&
            (a.unit || "") === (d.unit || "")
          );
        };
        const eqFinish = (a, d) => {
          if (!a || !d) return false;
          if (typeof a === "string" || typeof d === "string") return String(a) === String(d);
          const aArea = a.surfaceArea || {};
          const dArea = d.surfaceArea || {};
          return (
            (a.kind || "") === (d.kind || "") &&
            (a.spec || "") === (d.spec || "") &&
            (typeof aArea.value === "number" ? aArea.value : null) === (typeof dArea.value === "number" ? dArea.value : null) &&
            (aArea.unit || "") === (dArea.unit || "")
          );
        };
        const eqFn = path.endsWith("materials") ? eqMaterial : eqFinish;
        const used = new Set();
        const replaceOps = [];
        const addOps = [];
        const removeOps = [];

        for (let i = 0; i < currVal.length; i++) {
          const currEntry = currVal[i];
          let match = -1;
          for (let j = 0; j < baseVal.length; j++) {
            if (used.has(j)) continue;
            if (eqFn(baseVal[j], currEntry)) {
              match = j;
              used.add(j);
              break;
            }
          }
          if (match === -1) {
            addOps.push({ op: "add", path: joinPath(path, "-"), value: currEntry });
          } else {
            diffAny(baseVal[match], currEntry, joinPath(path, String(match)), replaceOps);
          }
        }

        for (let j = baseVal.length - 1; j >= 0; j--) {
          if (!used.has(j)) removeOps.push({ op: "remove", path: joinPath(path, String(j)) });
        }
        patch.push(...replaceOps);
        patch.push(...removeOps);
        patch.push(...addOps);
        return;
      }

      if (path.endsWith("Items")) {
        const baseById = new Map();
        baseVal.forEach((item, idx) => {
          if (item && typeof item.id === "string") baseById.set(item.id, { item, idx });
        });
        const usedIds = new Set();
        const replaceOps = [];
        const addOps = [];
        const removeOps = [];

        for (let i = 0; i < currVal.length; i++) {
          const currItem = currVal[i];
          const id = currItem && currItem.id;
          if (typeof id === "string" && baseById.has(id)) {
            const baseEntry = baseById.get(id);
            usedIds.add(id);
            diffAny(baseEntry.item, currItem, joinPath(path, String(baseEntry.idx)), replaceOps);
          } else {
            addOps.push({ op: "add", path: joinPath(path, "-"), value: currItem });
          }
        }

        for (let j = baseVal.length - 1; j >= 0; j--) {
          const baseItem = baseVal[j];
          const id = baseItem && baseItem.id;
          if (!id || !usedIds.has(id)) removeOps.push({ op: "remove", path: joinPath(path, String(j)) });
        }
        patch.push(...replaceOps);
        patch.push(...removeOps);
        patch.push(...addOps);
        return;
      }

      const maxLen = Math.max(baseVal.length, currVal.length);
      for (let i = 0; i < maxLen; i++) diffAny(baseVal[i], currVal[i], joinPath(path, String(i)), patch);
      return;
    }

    if (!baseIsObject || !currIsObject) {
      if (JSON.stringify(baseVal) !== JSON.stringify(currVal)) patch.push({ op: "replace", path: p, value: currVal });
      return;
    }

    const baseKeys = Object.keys(baseVal);
    const currKeys = Object.keys(currVal);
    const allKeys = new Set([...baseKeys, ...currKeys]);

    allKeys.forEach((key) => diffAny(baseVal[key], currVal[key], joinPath(path, escapeJsonPointer(key)), patch));
  }

  function escapeJsonPointer(str) {
    return String(str).replace(/~/g, "~0").replace(/\//g, "~1");
  }

  function collectDiffPaths(base, curr) {
    const paths = new Set();
    if (!base || !curr) return paths;
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
        if (path.endsWith("materials") || path.endsWith("finishes")) {
          const used = new Set();
          const eqMaterial = (a, d) => {
            if (!a || !d) return false;
            if (typeof a === "string" || typeof d === "string") return String(a) === String(d);
            return (
              (a.kind || "") === (d.kind || "") &&
              (a.spec || "") === (d.spec || "") &&
              (typeof a.quantity === "number" ? a.quantity : null) === (typeof d.quantity === "number" ? d.quantity : null) &&
              (a.unit || "") === (d.unit || "")
            );
          };
          const eqFinish = (a, d) => {
            if (!a || !d) return false;
            if (typeof a === "string" || typeof d === "string") return String(a) === String(d);
            const aArea = a.surfaceArea || {};
            const dArea = d.surfaceArea || {};
            return (
              (a.kind || "") === (d.kind || "") &&
              (a.spec || "") === (d.spec || "") &&
              (typeof aArea.value === "number" ? aArea.value : null) === (typeof dArea.value === "number" ? dArea.value : null) &&
              (aArea.unit || "") === (dArea.unit || "")
            );
          };
          const eqFn = path.endsWith("materials") ? eqMaterial : eqFinish;

          for (let i = 0; i < c.length; i++) {
            const currEntry = c[i];
            let match = -1;
            for (let j = 0; j < b.length; j++) {
              if (used.has(j)) continue;
              if (eqFn(b[j], currEntry)) {
                match = j;
                used.add(j);
                break;
              }
            }
            if (match === -1) {
              paths.add(path + "/-");
            } else {
              walk(b[match], currEntry, path + "/" + i);
            }
          }

          for (let j = 0; j < b.length; j++) {
            if (!used.has(j)) paths.add(path + "/" + j);
          }
          return;
        }

        if (path.endsWith("Items")) {
          const baseById = new Map();
          b.forEach((item, idx) => {
            if (item && typeof item.id === "string") baseById.set(item.id, { item, idx });
          });
          const usedIds = new Set();

          for (let i = 0; i < c.length; i++) {
            const currItem = c[i];
            const id = currItem && currItem.id;
            if (typeof id === "string" && baseById.has(id)) {
              usedIds.add(id);
              walk(baseById.get(id).item, currItem, path + "/" + i);
            } else {
              paths.add(path + "/-");
            }
          }

          b.forEach((item, idx) => {
            const id = item && item.id;
            if (!id || !usedIds.has(id)) paths.add(path + "/" + idx);
          });
          return;
        }

        const minLen = Math.min(b.length, c.length);

        for (let i = 0; i < minLen; i++) walk(b[i], c[i], path + "/" + i);
        for (let i = minLen; i < c.length; i++) paths.add(path + "/-");
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
    if (typeof payload.nlCorrectionWoodworkItems !== "string") payload.nlCorrectionWoodworkItems = "";
    if (typeof payload.nlCorrectionFloorItems !== "string") payload.nlCorrectionFloorItems = "";
    if (typeof payload.nlCorrectionFinishingItems !== "string") payload.nlCorrectionFinishingItems = "";
    if (typeof payload.nlCorrectionSignItems !== "string") payload.nlCorrectionSignItems = "";
    if (typeof payload.nlCollectionSignItems !== "string") payload.nlCollectionSignItems = "";
    if (typeof payload.nlCorrectionElectricalItems !== "string") payload.nlCorrectionElectricalItems = "";
    if (typeof payload.nlCorrectionLeaseItems !== "string") payload.nlCorrectionLeaseItems = "";
    if (typeof payload.nlCorrectionSiteCosts !== "string") payload.nlCorrectionSiteCosts = "";

    // legacy: upholsteryItems -> finishingItems
    if (Array.isArray(payload.upholsteryItems) && !Array.isArray(payload.finishingItems)) {
      payload.finishingItems = payload.upholsteryItems;
    }
    if ("upholsteryItems" in payload) delete payload.upholsteryItems;

    const categoryKeys = ["woodworkItems", "floorItems", "finishingItems", "signItems", "electricalItems", "leaseItems"];
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

    const normalizeCostObject = (entry) => {
      const base = entry && typeof entry === "object" && !Array.isArray(entry) ? entry : {};
      const nested = base.cost && typeof base.cost === "object" && !Array.isArray(base.cost) ? base.cost : null;
      const amount = typeof base.amount === "number" ? base.amount : nested && typeof nested.amount === "number" ? nested.amount : 0;
      const currency =
        typeof base.currency === "string" ? base.currency : nested && typeof nested.currency === "string" ? nested.currency : "JPY";
      const notes =
        nested && typeof nested.notes === "string" ? nested.notes : typeof base.notes === "string" ? base.notes : "";
      return { amount, currency, notes };
    };

    const normalizeLaborEntry = (entry) => {
      const base = entry && typeof entry === "object" && !Array.isArray(entry) ? entry : {};
      return {
        label: typeof base.label === "string" ? base.label : "",
        people: typeof base.people === "number" ? base.people : 0,
        days: typeof base.days === "number" ? base.days : 0,
        notes: typeof base.notes === "string" ? base.notes : ""
      };
    };

    const normalizeTransportEntry = (entry) => {
      const base = entry && typeof entry === "object" && !Array.isArray(entry) ? entry : {};
      return {
        label: typeof base.label === "string" ? base.label : "",
        vehicleType: typeof base.vehicleType === "string" ? base.vehicleType : "",
        vehicles: typeof base.vehicles === "number" ? base.vehicles : 0,
        days: typeof base.days === "number" ? base.days : 0,
        unitPrice: typeof base.unitPrice === "number" ? base.unitPrice : 0,
        cost: normalizeCostObject(base),
        notes: typeof base.notes === "string" ? base.notes : ""
      };
    };

    const laborBase = payload.siteCosts.laborCost && typeof payload.siteCosts.laborCost === "object" ? payload.siteCosts.laborCost : {};
    const legacyEntries = [];
    if (laborBase.setup && typeof laborBase.setup === "object") {
      legacyEntries.push({
        label: "設営",
        people: typeof laborBase.setup.people === "number" ? laborBase.setup.people : 0,
        days: typeof laborBase.setup.days === "number" ? laborBase.setup.days : 0,
        notes: typeof laborBase.setup.notes === "string" ? laborBase.setup.notes : ""
      });
    }
    if (laborBase.teardown && typeof laborBase.teardown === "object") {
      legacyEntries.push({
        label: "撤去",
        people: typeof laborBase.teardown.people === "number" ? laborBase.teardown.people : 0,
        days: typeof laborBase.teardown.days === "number" ? laborBase.teardown.days : 0,
        notes: typeof laborBase.teardown.notes === "string" ? laborBase.teardown.notes : ""
      });
    }
    const entries = Array.isArray(laborBase.entries) ? laborBase.entries.map(normalizeLaborEntry) : legacyEntries;
    payload.siteCosts.laborCost = {
      unitPrice: typeof laborBase.unitPrice === "number" ? laborBase.unitPrice : 0,
      entries: entries.length ? entries : [{ label: "作業", people: 0, days: 0, notes: "" }],
      cost: normalizeCostObject(laborBase),
      notes: typeof laborBase.notes === "string" ? laborBase.notes : ""
    };

    const transportBase =
      payload.siteCosts.transportCost && typeof payload.siteCosts.transportCost === "object" ? payload.siteCosts.transportCost : {};
    const legacyTransport = [];
    if (Array.isArray(transportBase.setup)) {
      transportBase.setup.forEach((entry) => {
        legacyTransport.push({ label: "搬入", ...entry });
      });
    }
    if (Array.isArray(transportBase.teardown)) {
      transportBase.teardown.forEach((entry) => {
        legacyTransport.push({ label: "搬出", ...entry });
      });
    }
    const transportEntries = Array.isArray(transportBase.entries)
      ? transportBase.entries.map(normalizeTransportEntry)
      : legacyTransport.map(normalizeTransportEntry);
    payload.siteCosts.transportCost = {
      entries: transportEntries.length
        ? transportEntries
        : [
            {
              label: "運搬",
              vehicleType: "",
              vehicles: 0,
              days: 0,
              unitPrice: 0,
              cost: { amount: 0, currency: "JPY", notes: "" },
              notes: ""
            }
          ],
      cost: normalizeCostObject(transportBase),
      notes: typeof transportBase.notes === "string" ? transportBase.notes : ""
    };

    const wasteBase =
      payload.siteCosts.wasteDisposalCost && typeof payload.siteCosts.wasteDisposalCost === "object"
        ? payload.siteCosts.wasteDisposalCost
        : {};
    payload.siteCosts.wasteDisposalCost = {
      vehicles: typeof wasteBase.vehicles === "number" ? wasteBase.vehicles : 0,
      unitPrice: typeof wasteBase.unitPrice === "number" ? wasteBase.unitPrice : 0,
      cost: normalizeCostObject(wasteBase),
      notes: typeof wasteBase.notes === "string" ? wasteBase.notes : ""
    };
    if (typeof payload.siteCosts.notes !== "string") payload.siteCosts.notes = "";

    const ensureCost = (obj, key) => {
      if (!obj[key] || typeof obj[key] !== "object" || Array.isArray(obj[key])) {
        obj[key] = {};
      }
      if (typeof obj[key].amount !== "number") obj[key].amount = 0;
      if (typeof obj[key].currency !== "string") obj[key].currency = "JPY";
      if (typeof obj[key].notes !== "string") obj[key].notes = "";
    };

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

      const normalizeMaterialSpec = (entry) => {
        if (typeof entry === "string") {
          return { kind: entry, spec: "", quantity: 1, unit: "式" };
        }
        const base = entry && typeof entry === "object" ? entry : {};
        return {
          kind: typeof base.kind === "string" ? base.kind : "",
          spec: typeof base.spec === "string" ? base.spec : "",
          quantity: typeof base.quantity === "number" && base.quantity > 0 ? base.quantity : 1,
          unit: typeof base.unit === "string" ? base.unit : "式"
        };
      };

      const normalizeFinishSpec = (entry) => {
        if (typeof entry === "string") {
          return { kind: entry, spec: "", surfaceArea: { value: 0.01, unit: "m²", notes: "" }, notes: "" };
        }
        const base = entry && typeof entry === "object" ? entry : {};
        const surfaceAreaBase = base.surfaceArea && typeof base.surfaceArea === "object" ? base.surfaceArea : {};
        const surfaceArea =
          typeof base.quantity === "number"
            ? { value: base.quantity, unit: "m²", notes: "" }
            : {
                value: typeof surfaceAreaBase.value === "number" ? surfaceAreaBase.value : 0.01,
                unit: typeof surfaceAreaBase.unit === "string" ? surfaceAreaBase.unit : "m²",
                notes: typeof surfaceAreaBase.notes === "string" ? surfaceAreaBase.notes : ""
              };
        return {
          kind: typeof base.kind === "string" ? base.kind : "",
          spec: typeof base.spec === "string" ? base.spec : "",
          surfaceArea,
          notes: typeof base.notes === "string" ? base.notes : ""
        };
      };

      it.materials = it.materials.map(normalizeMaterialSpec).filter((m) => m.kind);
      it.finishes = it.finishes.map(normalizeFinishSpec).filter((f) => f.kind);

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

    const normalizeSignItem = (it) => {
      if (!it || typeof it !== "object") return;
      normalizeQuantifiedItem(it);
      if (typeof it.signType !== "string") it.signType = "その他";
      if (!it.price || typeof it.price !== "object") {
        it.price = { mode: "override_fixed", currency: "JPY", unitPrice: 0, notes: "" };
      }
      if (typeof it.price.mode !== "string" || it.price.mode !== "override_fixed") it.price.mode = "override_fixed";
      if (typeof it.price.currency !== "string") it.price.currency = "JPY";
      if (typeof it.price.unitPrice !== "number") it.price.unitPrice = 0;
      if (typeof it.price.notes !== "string") it.price.notes = "";
    };

    payload.woodworkItems.forEach(normalizeWoodworkItem);
    ["floorItems", "finishingItems", "leaseItems"].forEach((k) => {
      payload[k].forEach(normalizeQuantifiedItem);
    });
    payload.signItems.forEach(normalizeSignItem);
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

    if (categoryKey === "signItems") {
      return {
        ...base,
        unit: "式",
        cost: { amount: 0, currency: "JPY", notes: "" },
        signType: "その他",
        price: { mode: "override_fixed", currency: "JPY", unitPrice: 0, notes: "" }
      };
    }

    return {
      ...base,
      unit: "式",
      cost: { amount: 0, currency: "JPY", notes: "" }
    };
  }

  function createDefaultEstimateGroup(groupNo, categoryName) {
    return {
      groupNo,
      categoryName: categoryName || "",
      noteLines: [],
      displayMode: "summary",
      pricingRule: {
        marginRateOnCost: 0,
        isHiddenFromDocument: true,
        notes: ""
      },
      groupSummaryLine: {
        quantity: 1,
        unit: "式",
        unitPrice: 0,
        amount: 0
      },
      lineItems: []
    };
  }

  function createDefaultEstimatePayload() {
    return {
      title: "",
      client: { name: "" },
      total: { amount: 0, tax: { included: true, rate: 0.1 } },
      breakdown: { title: "工事内訳明細表", groups: [createDefaultEstimateGroup(1, "")] }
    };
  }

  function ensureEstimatePayload() {
    if (!estimatePayload || typeof estimatePayload !== "object") {
      estimatePayload = createDefaultEstimatePayload();
    }
    if (!estimatePayload.client || typeof estimatePayload.client !== "object") {
      estimatePayload.client = { name: "" };
    }
    if (!estimatePayload.total || typeof estimatePayload.total !== "object") {
      estimatePayload.total = { amount: 0, tax: { included: true, rate: 0.1 } };
    }
    if (!estimatePayload.total.tax || typeof estimatePayload.total.tax !== "object") {
      estimatePayload.total.tax = { included: true, rate: 0.1 };
    }
    if (!estimatePayload.breakdown || typeof estimatePayload.breakdown !== "object") {
      estimatePayload.breakdown = { title: "工事内訳明細表", groups: [] };
    }
    if (!Array.isArray(estimatePayload.breakdown.groups)) {
      estimatePayload.breakdown.groups = [];
    }
  }

  function setPageMode(mode) {
    const isEstimate = mode === "estimate";
    if (extractPage) extractPage.hidden = isEstimate;
    if (estimatePage) estimatePage.hidden = !isEstimate;
    if (pageTabExtract) pageTabExtract.classList.toggle("active", !isEstimate);
    if (pageTabEstimate) pageTabEstimate.classList.toggle("active", isEstimate);
    if (isEstimate) syncEstimateToUI();
  }

  function calcEstimateLineAmount(line) {
    const amount = typeof line.amount === "number" ? line.amount : null;
    if (typeof amount === "number") return amount;
    const qty = typeof line.quantity === "number" ? line.quantity : 0;
    const unitPrice = typeof line.unitPrice === "number" ? line.unitPrice : 0;
    return qty * unitPrice;
  }

  function calcEstimateGroupAmount(group) {
    if (!group || typeof group !== "object") return 0;
    if (group.displayMode === "summary") {
      const line = group.groupSummaryLine || {};
      if (typeof line.amount === "number") return line.amount;
      if (typeof line.quantity === "number" && typeof line.unitPrice === "number") {
        return line.quantity * line.unitPrice;
      }
      return 0;
    }
    const items = Array.isArray(group.lineItems) ? group.lineItems : [];
    return items.reduce((sum, line) => sum + calcEstimateLineAmount(line), 0);
  }

  function recalcEstimateTotals() {
    if (!estimatePayload) return;
    const groups = estimatePayload.breakdown?.groups || [];
    const total = groups.reduce((sum, group) => sum + calcEstimateGroupAmount(group), 0);
    estimatePayload.total.amount = total;
    if (estimateTotalAmountInput) estimateTotalAmountInput.value = total ? String(total) : "0";
  }

  function renderEstimateGroups() {
    if (!estimateGroupsView) return;
    ensureEstimatePayload();
    const groups = estimatePayload.breakdown.groups;
    if (estimateSelectedGroupIndex == null && groups.length) {
      estimateSelectedGroupIndex = 0;
    }

    const rows = groups
      .map((group, index) => {
        const cost = typeof group.groupSummaryLine?.amount === "number" ? group.groupSummaryLine.amount : 0;
        const marginRate = typeof group.pricingRule?.marginRateOnCost === "number" ? group.pricingRule.marginRateOnCost : 0;
        const subtotal = Math.round(cost * (1 + marginRate));
        const marginDisplay = marginRate ? `${Math.round(marginRate * 100)}%` : "0%";
        const rowClass = index === estimateSelectedGroupIndex ? "selected" : "";
        return `
          <tr class="${rowClass}" data-group-index="${index}">
            <td>${group.groupNo ?? ""}</td>
            <td>${escapeHtml(group.categoryName || "")}</td>
            <td>${group.displayMode || ""}</td>
            <td>${cost || 0}</td>
            <td>${marginDisplay}</td>
            <td>${subtotal || 0}</td>
          </tr>
        `;
      })
      .join("");

    estimateGroupsView.innerHTML = `
      <table class="estimate-table">
        <thead>
          <tr>
            <th>#</th>
            <th>種別</th>
            <th>表示</th>
            <th>原価</th>
            <th>利益率</th>
            <th>小計</th>
          </tr>
        </thead>
        <tbody>
          ${rows || `<tr><td colspan="6" style="color:#777;">グループがありません</td></tr>`}
        </tbody>
      </table>
    `;

    const tbody = estimateGroupsView.querySelector("tbody");
    if (tbody) {
      tbody.querySelectorAll("tr[data-group-index]").forEach((row) => {
        row.addEventListener("click", () => {
          const idx = Number(row.getAttribute("data-group-index"));
          if (Number.isFinite(idx)) {
            estimateSelectedGroupIndex = idx;
            renderEstimateGroups();
            syncEstimateGroupEditor();
          }
        });
      });
    }
  }

  function syncEstimateGroupEditor() {
    ensureEstimatePayload();
    const groups = estimatePayload.breakdown.groups;
    const group = Number.isInteger(estimateSelectedGroupIndex) ? groups[estimateSelectedGroupIndex] : null;
    const hasGroup = !!group;

    if (estimateGroupLabel) {
      estimateGroupLabel.textContent = hasGroup
        ? `#${group.groupNo ?? ""} ${group.categoryName || ""}`
        : "（行をクリックするとここに表示されます）";
    }

    const setDisabled = (el, disabled) => {
      if (!el) return;
      el.disabled = disabled;
    };

    [
      estimateGroupNoInput,
      estimateGroupCategoryInput,
      estimateGroupDisplaySelect,
      estimateGroupNotesInput,
      estimateGroupMarginInput,
      estimateGroupHiddenInput,
      estimateGroupPricingNotesInput,
      estimateSummaryQtyInput,
      estimateSummaryUnitInput,
      estimateSummaryUnitPriceInput,
      estimateSummaryAmountInput,
      btnEstimateLineAdd
    ].forEach((el) => setDisabled(el, !hasGroup));

    if (!hasGroup) {
      if (estimateGroupNoInput) estimateGroupNoInput.value = "";
      if (estimateGroupCategoryInput) estimateGroupCategoryInput.value = "";
      if (estimateGroupDisplaySelect) estimateGroupDisplaySelect.value = "summary";
      if (estimateGroupNotesInput) estimateGroupNotesInput.value = "";
      if (estimateGroupMarginInput) estimateGroupMarginInput.value = "";
      if (estimateGroupHiddenInput) estimateGroupHiddenInput.checked = false;
      if (estimateGroupPricingNotesInput) estimateGroupPricingNotesInput.value = "";
      if (estimateSummaryQtyInput) estimateSummaryQtyInput.value = "";
      if (estimateSummaryUnitInput) estimateSummaryUnitInput.value = "";
      if (estimateSummaryUnitPriceInput) estimateSummaryUnitPriceInput.value = "";
      if (estimateSummaryAmountInput) estimateSummaryAmountInput.value = "";
      if (estimateLineItemsList) estimateLineItemsList.innerHTML = "";
      return;
    }

    if (estimateGroupNoInput) estimateGroupNoInput.value = group.groupNo ?? "";
    if (estimateGroupCategoryInput) estimateGroupCategoryInput.value = group.categoryName || "";
    if (estimateGroupDisplaySelect) estimateGroupDisplaySelect.value = group.displayMode || "summary";
    if (estimateGroupNotesInput) estimateGroupNotesInput.value = Array.isArray(group.noteLines) ? group.noteLines.join("\n") : "";

    const pricing = group.pricingRule || {};
    if (estimateGroupMarginInput) estimateGroupMarginInput.value = pricing.marginRateOnCost ?? "";
    if (estimateGroupHiddenInput) estimateGroupHiddenInput.checked = pricing.isHiddenFromDocument !== false;
    if (estimateGroupPricingNotesInput) estimateGroupPricingNotesInput.value = pricing.notes || "";

    const summary = group.groupSummaryLine || {};
    if (estimateSummaryQtyInput) estimateSummaryQtyInput.value = summary.quantity ?? "";
    if (estimateSummaryUnitInput) estimateSummaryUnitInput.value = summary.unit || "";
    if (estimateSummaryUnitPriceInput) estimateSummaryUnitPriceInput.value = summary.unitPrice ?? "";
    if (estimateSummaryAmountInput) estimateSummaryAmountInput.value = summary.amount ?? "";

    if (estimateSummarySection) estimateSummarySection.hidden = group.displayMode !== "summary";
    if (estimateLinesSection) estimateLinesSection.hidden = group.displayMode !== "detailed";

    if (group.displayMode === "detailed") {
      if (!Array.isArray(group.lineItems) || group.lineItems.length === 0) {
        group.lineItems = [createDefaultLineItem()];
      }
      renderEstimateLineItems(estimateLineItemsList, group.lineItems);
    } else {
      if (estimateLineItemsList) estimateLineItemsList.innerHTML = "";
    }
  }

  function syncEstimateHeaderToUI() {
    ensureEstimatePayload();
    if (estimateTitleInput) estimateTitleInput.value = estimatePayload.title || "";
    if (estimateClientInput) estimateClientInput.value = estimatePayload.client?.name || "";
    if (estimateTotalAmountInput) estimateTotalAmountInput.value = estimatePayload.total?.amount ?? "";
    if (estimateTaxIncludedInput) estimateTaxIncludedInput.checked = !!estimatePayload.total?.tax?.included;
    if (estimateTaxRateInput) estimateTaxRateInput.value = estimatePayload.total?.tax?.rate ?? "";
    if (estimateBreakdownTitleInput) estimateBreakdownTitleInput.value = estimatePayload.breakdown?.title || "";
  }

  function syncEstimateJson() {
    if (!estimateJsonTextarea) return;
    ensureEstimatePayload();
    estimateJsonTextarea.value = JSON.stringify(estimatePayload, null, 2);
  }

  function syncEstimateToUI() {
    ensureEstimatePayload();
    syncEstimateHeaderToUI();
    renderEstimateGroups();
    syncEstimateGroupEditor();
    recalcEstimateTotals();
    syncEstimateJson();
  }

  function applyEstimateHeaderFromUI() {
    if (!estimatePayload) return;
    if (estimateTitleInput) estimatePayload.title = estimateTitleInput.value || "";
    if (estimateClientInput) estimatePayload.client = { name: estimateClientInput.value || "" };
    if (estimateBreakdownTitleInput) estimatePayload.breakdown.title = estimateBreakdownTitleInput.value || "";
    if (estimateTaxIncludedInput) estimatePayload.total.tax.included = estimateTaxIncludedInput.checked;
    if (estimateTaxRateInput) {
      const rate = parseFloat(String(estimateTaxRateInput.value || ""));
      estimatePayload.total.tax.rate = Number.isFinite(rate) ? rate : 0;
    }
    recalcEstimateTotals();
    syncEstimateJson();
  }

  function applyEstimateGroupFromUI() {
    if (!estimatePayload) return;
    const groups = estimatePayload.breakdown.groups;
    const group = Number.isInteger(estimateSelectedGroupIndex) ? groups[estimateSelectedGroupIndex] : null;
    if (!group) return;

    const num = parseInt(String(estimateGroupNoInput?.value || ""), 10);
    group.groupNo = Number.isFinite(num) ? num : group.groupNo || 1;
    group.categoryName = estimateGroupCategoryInput?.value || "";
    group.displayMode = estimateGroupDisplaySelect?.value === "detailed" ? "detailed" : "summary";
    if (estimateGroupNotesInput) {
      const raw = estimateGroupNotesInput.value || "";
      group.noteLines = raw
        .split("\n")
        .map((line) => line.trim())
        .filter((line) => line.length > 0);
    }
    group.pricingRule = {
      marginRateOnCost: parseFloat(String(estimateGroupMarginInput?.value || "")) || 0,
      isHiddenFromDocument: estimateGroupHiddenInput?.checked !== false,
      notes: estimateGroupPricingNotesInput?.value || ""
    };

    if (estimateGroupLabel) {
      estimateGroupLabel.textContent = `#${group.groupNo ?? ""} ${group.categoryName || ""}`;
    }

    if (group.displayMode === "summary") {
      if (!group.groupSummaryLine || typeof group.groupSummaryLine !== "object") {
        group.groupSummaryLine = { quantity: 0, unit: "", unitPrice: 0, amount: 0 };
      }
      const qty = parseFloat(String(estimateSummaryQtyInput?.value || ""));
      const unitPrice = parseFloat(String(estimateSummaryUnitPriceInput?.value || ""));
      const amount = parseFloat(String(estimateSummaryAmountInput?.value || ""));
      group.groupSummaryLine.quantity = Number.isFinite(qty) ? qty : 0;
      group.groupSummaryLine.unit = estimateSummaryUnitInput?.value || "";
      group.groupSummaryLine.unitPrice = Number.isFinite(unitPrice) ? unitPrice : 0;
      group.groupSummaryLine.amount = Number.isFinite(amount) ? amount : group.groupSummaryLine.quantity * group.groupSummaryLine.unitPrice;
    } else {
      if (!Array.isArray(group.lineItems)) group.lineItems = [];
      const nextLines = readEstimateLineItems(estimateLineItemsList, group.lineItems);
      if (nextLines.length > 0) {
        group.lineItems = nextLines;
      } else if (!Array.isArray(group.lineItems) || group.lineItems.length === 0) {
        group.lineItems = [createDefaultLineItem()];
      }
    }

    recalcEstimateTotals();
    renderEstimateGroups();
    syncEstimateJson();
  }

  function createDefaultLineItem() {
    return {
      id: "",
      description: "",
      quantity: 1,
      unit: "式",
      unitPrice: 0,
      amount: 0,
      note: ""
    };
  }


  function renderEstimateLineItems(container, items) {
    if (!container) return;
    container.innerHTML = "";
    const list = Array.isArray(items) ? items : [];
    list.forEach((item, index) => {
      const row = document.createElement("div");
      row.className = "estimate-lines-row";
      row.dataset.lineRow = "1";
      row.dataset.lineIndex = String(index);

      const makeInput = (field, value, type = "text", step) => {
        const input = document.createElement("input");
        input.type = type;
        if (step) input.step = step;
        input.value = value != null ? String(value) : "";
        input.dataset.field = field;
        return input;
      };

      const descInput = makeInput("description", item.description || "");
      descInput.placeholder = "内容";
      const qtyInput = makeInput("quantity", item.quantity ?? "", "number", "0.01");
      const unitInput = makeInput("unit", item.unit || "");
      const unitPriceInput = makeInput("unitPrice", item.unitPrice ?? "", "number", "1");
      const amountInput = makeInput("amount", item.amount ?? "", "number", "1");
      const noteInput = makeInput("note", item.note || "");

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "削除";
      removeBtn.addEventListener("click", () => {
        if (!estimatePayload) return;
        const group = estimatePayload.breakdown.groups?.[estimateSelectedGroupIndex];
        if (!group || !Array.isArray(group.lineItems)) return;
        group.lineItems.splice(index, 1);
        if (group.lineItems.length === 0) {
          group.lineItems.push(createDefaultLineItem());
        }
        renderEstimateLineItems(container, group.lineItems);
        applyEstimateGroupFromUI();
      });

      row.appendChild(descInput);
      row.appendChild(qtyInput);
      row.appendChild(unitInput);
      row.appendChild(unitPriceInput);
      row.appendChild(amountInput);
      row.appendChild(noteInput);
      row.appendChild(removeBtn);
      container.appendChild(row);
    });
  }

  function readEstimateLineItems(container, existingItems) {
    if (!container) return [];
    const rows = container.querySelectorAll("[data-line-row]");
    const items = [];
    const existing = Array.isArray(existingItems) ? existingItems : [];
    rows.forEach((row) => {
      const index = parseInt(row.dataset.lineIndex || "", 10);
      const get = (field) => row.querySelector(`[data-field="${field}"]`);
      const toNum = (v) => {
        const n = parseFloat(String(v || ""));
        return Number.isFinite(n) ? n : 0;
      };
      const description = (get("description") && get("description").value) || "";
      const quantity = toNum(get("quantity") && get("quantity").value);
      const unit = (get("unit") && get("unit").value) || "";
      const unitPrice = toNum(get("unitPrice") && get("unitPrice").value);
      const amountRaw = get("amount") && get("amount").value;
      const amount = Number.isFinite(parseFloat(String(amountRaw || "")))
        ? parseFloat(String(amountRaw || ""))
        : quantity * unitPrice;
      const note = (get("note") && get("note").value) || "";
      const base = Number.isFinite(index) && existing[index] ? existing[index] : createDefaultLineItem();
      if (!base.id && Number.isFinite(index)) base.id = `line-${index}-${Date.now()}`;
      base.description = description || "未設定";
      base.quantity = quantity;
      base.unit = unit;
      base.unitPrice = unitPrice;
      base.amount = amount;
      base.note = note;
      items.push(base);
    });
    return items;
  }

  function buildEstimateFromExtraction(payload) {
    if (!payload || typeof payload !== "object") return createDefaultEstimatePayload();
    const categories = [
      { key: "woodworkItems", label: "木工造作" },
      { key: "floorItems", label: "床" },
      { key: "finishingItems", label: "表装" },
      { key: "signItems", label: "サイン" },
      { key: "electricalItems", label: "電気" },
      { key: "leaseItems", label: "リース" }
    ];

    const calcItemAmount = (item, isWoodwork) => {
      if (!item || typeof item !== "object") return 0;
      if (item.includeInEstimate === false) return 0;
      const unitPrice = item.price && typeof item.price.unitPrice === "number" ? item.price.unitPrice : null;
      const qty = typeof item.quantity === "number" ? item.quantity : null;
      if (typeof unitPrice === "number" && typeof qty === "number") return unitPrice * qty;
      if (isWoodwork) {
        const mat = item.materialCost?.amount || 0;
        const labor = item.laborCost?.amount || 0;
        return mat + labor;
      }
      const cost = item.cost?.amount;
      return typeof cost === "number" ? cost : 0;
    };

    const groups = [];
    categories.forEach((cat, index) => {
      const items = Array.isArray(payload[cat.key]) ? payload[cat.key] : [];
      const total = items.reduce((sum, item) => sum + calcItemAmount(item, cat.key === "woodworkItems"), 0);
      if (items.length === 0 && total === 0) return;
      const group = createDefaultEstimateGroup(groups.length + 1, cat.label);
      if (cat.key === "leaseItems") {
        group.displayMode = "detailed";
      }
      group.lineItems = items.map((item) => {
        const unitPrice = item?.price && typeof item.price.unitPrice === "number" ? item.price.unitPrice : null;
        const qty = typeof item?.quantity === "number" ? item.quantity : 1;
        const unit = typeof item?.unit === "string" ? item.unit : "式";
        const amount = calcItemAmount(item, cat.key === "woodworkItems");
        const lineId = item?.id || `line-${cat.key}-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
        const lineItem = {
          id: lineId,
          description: item?.name || "未設定",
          quantity: qty,
          unit,
          unitPrice: unitPrice ?? 0,
          amount,
          note: item?.notes || ""
        };
        return lineItem;
      });
      group.groupSummaryLine.amount = total;
      group.groupSummaryLine.unitPrice = total;
      groups.push(group);
    });

    const site = payload.siteCosts || {};
    const laborAmount = site.laborCost?.cost?.amount || 0;
    const laborUnitPrice = site.laborCost?.unitPrice || 0;
    const laborQuantity = Array.isArray(site.laborCost?.entries)
      ? site.laborCost.entries.reduce((sum, entry) => sum + (entry.people || 0) * (entry.days || 0), 0)
      : 0;
    const transportAmount = site.transportCost?.cost?.amount || 0;
    const wasteAmount = site.wasteDisposalCost?.cost?.amount || 0;
    const siteTotal = laborAmount + transportAmount + wasteAmount;
    const group = createDefaultEstimateGroup(groups.length + 1, "現場費");
    group.displayMode = "detailed";
    group.lineItems = [
      {
        id: "sitecosts-labor",
        description: "人工費",
        quantity: laborQuantity || 0,
        unit: "人工",
        unitPrice: laborUnitPrice,
        amount: laborAmount,
        note: ""
      },
      {
        id: "sitecosts-transport",
        description: "運搬費",
        quantity: 1,
        unit: "式",
        unitPrice: transportAmount,
        amount: transportAmount,
        note: ""
      },
      {
        id: "sitecosts-waste",
        description: "残材処理費",
        quantity: 1,
        unit: "式",
        unitPrice: wasteAmount,
        amount: wasteAmount,
        note: ""
      }
    ];
    group.groupSummaryLine.amount = siteTotal;
    group.groupSummaryLine.unitPrice = siteTotal;
    groups.push(group);

    const totalAmount = groups.reduce((sum, group) => sum + calcEstimateGroupAmount(group), 0);
    return {
      title: "御 見 積 書",
      client: { name: "" },
      total: { amount: totalAmount, tax: { included: true, rate: 0.1 } },
      breakdown: { title: "工事内訳明細表", groups: groups.length ? groups : [createDefaultEstimateGroup(1, "")] }
    };
  }


  // 初期表示
  syncNlCategoryToUI();
  updateSourceInfoUI();
  updateCategoryHighlight();
  renderCategoryTabs();
  setFormModeForCategory();
  syncSiteCostsToUI();
  syncSiteCostsFormVisibility();
  syncSiteCostEditorForSelection();
  renderItemsTable();
  syncEstimateToUI();
  setPageMode("extract");
});
