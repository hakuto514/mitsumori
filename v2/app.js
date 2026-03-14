/* ===== V2 御見積書ジェネレーター - app.js ===== */

/* ===== STATE ===== */
let parts = [],
  labors = [],
  discounts = [];
let staffList = [
  "小椋 政継",
  "牧田 寛之",
  "山澤 亮太",
  "髙橋 大輔",
  "重親 亮介",
  "磯江 亮輔",
];
let selectedStaff = 0;
const STORAGE_KEY = "mitsumori_v2_data";
const TEMPLATE_KEY = "mitsumori_v2_templates";

/* ===== THEME ===== */
function initTheme() {
  const saved = localStorage.getItem("mitsumori_theme");
  if (saved === "dark")
    document.documentElement.setAttribute("data-theme", "dark");
}
function toggleTheme() {
  const isDark = document.documentElement.getAttribute("data-theme") === "dark";
  document.documentElement.setAttribute(
    "data-theme",
    isDark ? "light" : "dark",
  );
  localStorage.setItem("mitsumori_theme", isDark ? "light" : "dark");
}

/* ===== TAB ===== */
function switchTab(t) {
  document
    .querySelectorAll(".tab-content")
    .forEach((e) => e.classList.remove("active"));
  document
    .querySelectorAll(".tab-btn")
    .forEach((e) => e.classList.remove("active"));
  document.getElementById("tab-" + t).classList.add("active");
  document
    .querySelector('.tab-btn[data-tab="' + t + '"]')
    .classList.add("active");
}

/* ===== TOAST ===== */
function showToast(msg) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2500);
}

/* ===== UTILS ===== */
const floor100 = (v) => Math.floor(v / 100) * 100;
const fmt = (n) => "¥" + Number(n || 0).toLocaleString();
const fmtN = (n) => Number(n || 0).toLocaleString();

/* ===== VALIDATION ===== */
function validateField(id, label) {
  const el = document.getElementById(id);
  if (!el) return true;
  const val = el.value.trim();
  const grp = el.closest(".form-group");
  if (!val) {
    el.classList.add("invalid");
    if (grp) {
      grp.classList.add("has-error");
      let m = grp.querySelector(".validation-msg");
      if (!m) {
        m = document.createElement("div");
        m.className = "validation-msg";
        grp.appendChild(m);
      }
      m.textContent = label + "を入力してください";
      m.style.display = "block";
    }
    return false;
  }
  el.classList.remove("invalid");
  if (grp) {
    grp.classList.remove("has-error");
    const m = grp.querySelector(".validation-msg");
    if (m) m.style.display = "none";
  }
  return true;
}

function validateAll() {
  let ok = true;
  [
    ["estNo", "見積書番号"],
    ["issueDate", "発行日"],
    ["custName", "顧客名"],
    ["compName", "会社名"],
  ].forEach(([id, lbl]) => {
    if (!validateField(id, lbl)) ok = false;
  });
  return ok;
}

function updateTabBadges() {
  const basicFields = [
    "estNo",
    "issueDate",
    "custName",
    "compName",
    "compAddr",
    "compDept",
  ];
  const filled = basicFields.filter((id) => {
    const e = document.getElementById(id);
    return e && e.value.trim();
  }).length;
  const pct = Math.round((filled / basicFields.length) * 100);
  setBadge("badge-basic", pct);
  setBadge("badge-parts", parts.length > 0 ? 100 : 0);
  setBadge("badge-labor", labors.length > 0 ? 100 : 0);
  setBadge("badge-output", 100);
}

function setBadge(id, pct) {
  const el = document.getElementById(id);
  if (!el) return;
  if (pct >= 100) {
    el.className = "tab-badge complete";
    el.textContent = "✓";
  } else if (pct > 0) {
    el.className = "tab-badge partial";
    el.textContent = pct + "%";
  } else {
    el.className = "tab-badge empty";
    el.textContent = "○";
  }
}

/* ===== PARTS ===== */
function addPart(p) {
  if (!p && parts.length >= 15) {
    showToast("部品は最大15行までです");
    return;
  }
  parts.push(
    p || {
      name: "",
      code: "",
      qty: 1,
      costPrice: 0,
      sellInput: 0,
      note: "",
      atCost: false,
      source: "other",
    },
  );
  renderParts();
}
function removePart(i) {
  parts.splice(i, 1);
  renderParts();
}
function updatePart(i, k, v) {
  if (k === "atCost") {
    parts[i][k] = v;
    renderParts();
    return;
  }
  if (k === "source") {
    parts[i][k] = v;
    renderParts();
    return;
  }
  parts[i][k] =
    k === "qty" || k === "costPrice" || k === "sellInput" ? Number(v) : v;
  recalc();
}

function adjustQty(i, delta) {
  const newQty = (parts[i].qty || 1) + delta;
  if (newQty >= 1) {
    parts[i].qty = newQty;
    // Re-render the specific qty input to avoid re-rendering entire list and losing focus if typing
    const el = document.getElementById("qtyInput" + i);
    if (el) el.value = newQty;
    recalc();
  }
}

function renderParts() {
  const c = document.getElementById("partsContainer");
  c.innerHTML = "";
  parts.forEach((p, i) => {
    const sell = floor100(p.sellInput);
    const isSelf = p.source === "self";
    const cost = isSelf ? Math.round(sell * 0.62) : p.costPrice;
    const margin =
      sell > 0
        ? isSelf
          ? 38
          : p.atCost
            ? 0
            : Math.round(((sell - cost) / sell) * 100)
        : 0;
    const bc =
      margin >= 60
        ? "badge-green"
        : margin >= 40
          ? "badge-yellow"
          : "badge-red";
    const badgeText = isSelf ? "自社38%" : p.atCost ? "原価" : margin + "%";
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <div class="card-header">
        <span style="font-weight:700">部品 #${i + 1}</span>
        <div><span class="badge ${p.atCost && !isSelf ? "badge-yellow" : bc}">${badgeText}</span>
        <button class="del-btn" onclick="removePart(${i})">✕</button></div>
      </div>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <button class="btn btn-sm ${isSelf ? "btn-primary" : "btn-outline"}" onclick="updatePart(${i},'source','self')">自社</button>
        <button class="btn btn-sm ${!isSelf ? "btn-primary" : "btn-outline"}" onclick="updatePart(${i},'source','other')">他社</button>
      </div>
      <div class="form-row c2">
        <div class="form-group"><label>部品名</label><input value="${p.name}" onchange="updatePart(${i},'name',this.value)"></div>
        <div class="form-group"><label>部品コード</label><input value="${p.code}" onchange="updatePart(${i},'code',this.value)"></div>
      </div>
      <div class="form-row ${isSelf ? "c2" : "c3"}">
        <div class="form-group"><label>数量</label>
          <div class="qty-input-group">
            <button class="qty-btn" onclick="adjustQty(${i}, -1)">－</button>
            <input type="number" id="qtyInput${i}" value="${p.qty}" min="1" onchange="updatePart(${i},'qty',this.value)">
            <button class="qty-btn" onclick="adjustQty(${i}, 1)">＋</button>
          </div>
        </div>
        ${isSelf ? "" : `<div class="form-group"><label>仕入単価</label><input type="number" value="${p.costPrice}" onchange="updatePart(${i},'costPrice',this.value)"></div>`}
        <div class="form-group"><label>${isSelf ? "定価" : "売価単価"}</label><input type="number" value="${p.sellInput}" onchange="updatePart(${i},'sellInput',this.value)"></div>
      </div>
      <div class="form-group"><label>備考</label><input value="${p.note}" onchange="updatePart(${i},'note',this.value)"></div>
      ${isSelf ? "" : `<div class="check-row"><input type="checkbox" ${p.atCost ? "checked" : ""} onchange="updatePart(${i},'atCost',this.checked)"><label style="font-size:12px">原価のまま（粗利計算から除外）</label></div>`}`;
    c.appendChild(card);
  });
  recalc();
}

/* ===== LABOR ===== */
function addLabor(l) {
  if (!l && labors.length >= 8) {
    showToast("作業費は最大8行までです");
    return;
  }
  labors.push(l || { name: "", amount: 0, note: "" });
  renderLabor();
}
function removeLabor(i) {
  labors.splice(i, 1);
  renderLabor();
}
function updateLabor(i, k, v) {
  labors[i][k] = k === "amount" ? Number(v) : v;
  recalc();
}
function renderLabor() {
  const c = document.getElementById("laborContainer");
  c.innerHTML = "";
  labors.forEach((l, i) => {
    const isExp = l.name === "諸経費";
    const d = document.createElement("div");
    d.className = "card";
    d.innerHTML = `
      <div class="card-header"><span style="font-weight:600">${l.name || "作業費項目"}</span><button class="del-btn" onclick="removeLabor(${i})">✕</button></div>
      <div class="form-row c2">
        <div class="form-group"><label>項目名</label><input value="${l.name}" onchange="updateLabor(${i},'name',this.value)" ${isExp ? "readonly" : ""}></div>
        <div class="form-group"><label>金額${isExp ? " (自動計算)" : ""}</label><input type="number" value="${l.amount}" id="labor-amt-${i}" onchange="updateLabor(${i},'amount',this.value)" ${isExp ? "readonly" : ""}></div>
      </div>
      <div class="form-group"><label>備考</label><input value="${l.note}" onchange="updateLabor(${i},'note',this.value)"></div>`;
    c.appendChild(d);
  });
  recalc();
}

/* ===== DISCOUNT ===== */
function addDiscount() {
  if (discounts.length >= 3) {
    showToast("値引きは最大3行までです");
    return;
  }
  discounts.push({ name: "値引き", amount: 0, note: "" });
  renderDiscount();
}
function removeDiscount(i) {
  discounts.splice(i, 1);
  renderDiscount();
}
function updateDiscount(i, k, v) {
  discounts[i][k] = k === "amount" ? Number(v) : v;
  recalc();
}
function renderDiscount() {
  const c = document.getElementById("discountContainer");
  c.innerHTML = "";
  discounts.forEach((d, i) => {
    const el = document.createElement("div");
    el.className = "card";
    el.innerHTML = `
      <div class="card-header"><span style="font-weight:600">${d.name}</span><button class="del-btn" onclick="removeDiscount(${i})">✕</button></div>
      <div class="form-row c2">
        <div class="form-group"><label>項目名</label><input value="${d.name}" onchange="updateDiscount(${i},'name',this.value)"></div>
        <div class="form-group"><label>金額</label><input type="number" value="${d.amount}" onchange="updateDiscount(${i},'amount',this.value)"></div>
      </div>
      <div class="form-group"><label>備考</label><input value="${d.note}" onchange="updateDiscount(${i},'note',this.value)"></div>`;
    c.appendChild(el);
  });
  recalc();
}

/* ===== STAFF ===== */
function renderStaff() {
  const p = document.getElementById("staffPills");
  p.innerHTML = staffList
    .map((s, i) => {
      const active = i === selectedStaff;
      return `<div class="pill ${active ? "" : "inactive"}" onclick="selectedStaff=${i};renderStaff()">${s}<button class="pill-del" onclick="event.stopPropagation();removeStaffFromList(${i})">×</button></div>`;
    })
    .join("");
}
function openStaffModal() {
  document.getElementById("staffModal").classList.add("show");
  renderStaffList();
}
function closeStaffModal() {
  document.getElementById("staffModal").classList.remove("show");
  renderStaff();
}
function renderStaffList() {
  document.getElementById("staffList").innerHTML = staffList
    .map(
      (s, i) =>
        `<div class="card" style="padding:10px;display:flex;justify-content:space-between;align-items:center"><span>${s}</span><button class="del-btn" onclick="staffList.splice(${i},1);if(selectedStaff>=staffList.length)selectedStaff=Math.max(0,staffList.length-1);renderStaffList()">✕</button></div>`,
    )
    .join("");
}
function addStaff() {
  const n = document.getElementById("newStaff");
  if (n.value.trim()) {
    staffList.push(n.value.trim());
    n.value = "";
    renderStaffList();
    showToast("担当者を追加しました");
  }
}
function removeStaffFromList(i) {
  staffList.splice(i, 1);
  if (selectedStaff >= staffList.length)
    selectedStaff = Math.max(0, staffList.length - 1);
  renderStaff();
}

/* ===== CALC ===== */
function recalc() {
  const aux = Number(document.getElementById("auxMaterial").value) || 0;
  let costTotal = 0,
    sellTotal = 0,
    costForMargin = 0,
    sellForMargin = 0;
  parts.forEach((p) => {
    const sell = floor100(p.sellInput);
    const isSelf = p.source === "self";
    const lineTotal = sell * p.qty;
    const lineCost = isSelf
      ? Math.round(sell * 0.62) * p.qty
      : p.costPrice * p.qty;
    sellTotal += lineTotal;
    costTotal += lineCost;
    if (isSelf || !p.atCost) {
      sellForMargin += lineTotal;
      costForMargin += lineCost;
    }
  });
  const partMargin =
    sellForMargin > 0
      ? Math.round(((sellForMargin - costForMargin) / sellForMargin) * 100)
      : 0;
  document.getElementById("sumCost").textContent = fmt(costTotal);
  document.getElementById("sumSell").textContent = fmt(sellTotal);
  document.getElementById("sumProfit").textContent = partMargin + "%";
  document.getElementById("sumMargin").textContent = fmt(
    sellForMargin - costForMargin,
  );

  let laborSumExMisc = 0;
  labors.forEach((l) => {
    if (l.name !== "諸経費") laborSumExMisc += l.amount;
  });
  const miscBase = sellTotal + aux + laborSumExMisc;
  const miscExp = Math.floor((miscBase * 0.1) / 100) * 100;
  labors.forEach((l, i) => {
    if (l.name === "諸経費") {
      l.amount = miscExp;
      const el = document.getElementById("labor-amt-" + i);
      if (el) el.value = miscExp;
    }
  });

  let laborSum = 0;
  labors.forEach((l) => (laborSum += l.amount));
  let discSum = 0;
  discounts.forEach((d) => (discSum += d.amount));
  const subtotal = sellTotal + aux + laborSum - discSum;
  const tax = Math.floor(subtotal * 0.1);
  const total = subtotal + tax;
  const grossMargin =
    subtotal > 0 ? ((subtotal - costTotal) / subtotal) * 100 : 0;

  document.getElementById("monSubtotal").textContent = fmt(subtotal);
  document.getElementById("monTotal").textContent = fmt(total);
  const gm = Math.round(grossMargin);
  document.getElementById("monMargin").textContent = gm + "%";
  document.getElementById("monMargin").style.color =
    gm >= 60 ? "var(--green)" : gm >= 40 ? "var(--yellow)" : "var(--red)";
  const bar = document.getElementById("monBar");
  bar.style.width = Math.min(100, Math.max(0, gm)) + "%";
  bar.style.background =
    gm >= 60 ? "var(--green)" : gm >= 40 ? "var(--yellow)" : "var(--red)";

  autoSave();
  updateTabBadges();
}

/* ===== STAMP (Canvas) ===== */
function createStampCanvas(staffName, dateStr, size) {
  const cv = document.createElement("canvas");
  cv.width = size;
  cv.height = size;
  const ctx = cv.getContext("2d");
  const cx = size / 2,
    cy = size / 2,
    r = size / 2 - 2;
  // Outer circle
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.strokeStyle = "#c83c2c";
  ctx.lineWidth = Math.max(1.5, size * 0.03);
  ctx.stroke();
  // Extract surname
  const surname = staffName.split(/[\s　]/)[0] || staffName;
  let top, bottom;
  if (surname.length <= 1) {
    top = surname;
    bottom = "";
  } else if (surname.length === 2) {
    top = surname[0];
    bottom = surname[1];
  } else if (surname.length === 3) {
    top = surname[0];
    bottom = surname.slice(1);
  } else {
    const h = Math.ceil(surname.length / 2);
    top = surname.slice(0, h);
    bottom = surname.slice(h);
  }
  const dd = dateStr || "";
  const dateLabel = dd.replace(
    /(\d{4})-(\d{2})-(\d{2})/,
    (m, y, mo, d) => y.slice(2) + "." + mo + "." + d,
  );
  ctx.fillStyle = "#c83c2c";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  // Horizontal date-band lines
  const ly1 = cy - r * 0.19,
    ly2 = cy + r * 0.19,
    lx = r * 0.82;
  ctx.strokeStyle = "#c83c2c";
  ctx.lineWidth = Math.max(1, size * 0.012);
  ctx.beginPath();
  ctx.moveTo(cx - lx, ly1);
  ctx.lineTo(cx + lx, ly1);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(cx - lx, ly2);
  ctx.lineTo(cx + lx, ly2);
  ctx.stroke();
  // Top kanji — balanced size
  const topFs = (r * 0.58) / Math.max(top.length, 1);
  ctx.font = "bold " + topFs + 'px "Hiragino Mincho ProN","Yu Mincho",serif';
  ctx.fillText(top, cx, cy - r * 0.52);
  // Date text — slightly larger for readability
  const dateFs = r * 0.30;
  ctx.font = "bold " + dateFs + 'px "Helvetica Neue",sans-serif';
  ctx.fillText(dateLabel, cx, cy);
  // Bottom kanji — balanced size
  const botFs = (r * 0.58) / Math.max(bottom.length, 1);
  ctx.font = "bold " + botFs + 'px "Hiragino Mincho ProN","Yu Mincho",serif';
  if (bottom) ctx.fillText(bottom, cx, cy + r * 0.54);
  return cv;
}

/* ===== LOCALSTORAGE ===== */
function getFormData() {
  const fields = [
    "estNo",
    "issueDate",
    "validity",
    "payTerms",
    "custName",
    "delivPlace",
    "delivDate",
    "prodName",
    "modelNo",
    "installDate",
    "serialNo",
    "callContent",
    "causeAction",
    "compName",
    "compAddr",
    "compDept",
    "compBranch",
    "compTel",
    "compFax",
    "auxMaterial",
    "remarks",
  ];
  const data = {};
  fields.forEach((id) => {
    const el = document.getElementById(id);
    if (el) data[id] = el.value;
  });
  data.parts = parts;
  data.labors = labors;
  data.discounts = discounts;
  data.staffList = staffList;
  data.selectedStaff = selectedStaff;
  return data;
}

function setFormData(data) {
  const fields = [
    "estNo",
    "issueDate",
    "validity",
    "payTerms",
    "custName",
    "delivPlace",
    "delivDate",
    "prodName",
    "modelNo",
    "installDate",
    "serialNo",
    "callContent",
    "causeAction",
    "compName",
    "compAddr",
    "compDept",
    "compBranch",
    "compTel",
    "compFax",
    "auxMaterial",
    "remarks",
  ];
  fields.forEach((id) => {
    const el = document.getElementById(id);
    if (el && data[id] !== undefined) {
      if (id === "installDate" && !/^\d{4}-\d{2}-\d{2}$/.test(data[id])) {
        // It's in the old text format (e.g. 1996年6月14日), ignore it so the type="date" input doesn't break
        el.value = "";
      } else {
        el.value = data[id];
      }
    }
  });
  if (data.parts) parts = data.parts;
  if (data.labors) labors = data.labors;
  if (data.discounts) discounts = data.discounts;
  if (data.staffList) staffList = data.staffList;
  if (data.selectedStaff !== undefined) selectedStaff = data.selectedStaff;
  renderParts();
  renderLabor();
  renderDiscount();
  renderStaff();
  recalc();
}

function autoSave() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(getFormData()));
  } catch (e) {}
}

function autoLoad() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      setFormData(JSON.parse(raw));
      return true;
    }
  } catch (e) {}
  return false;
}

/* ===== TEMPLATES ===== */
function getTemplates() {
  try {
    return JSON.parse(localStorage.getItem(TEMPLATE_KEY) || "[]");
  } catch (e) {
    return [];
  }
}

function saveTemplate() {
  const name = prompt("テンプレート名を入力してください:");
  if (!name) return;
  const templates = getTemplates();
  templates.push({
    name,
    date: new Date().toISOString().slice(0, 10),
    data: getFormData(),
  });
  localStorage.setItem(TEMPLATE_KEY, JSON.stringify(templates));
  renderTemplates();
  showToast("テンプレートを保存しました");
}

function loadTemplate(idx) {
  const templates = getTemplates();
  if (templates[idx]) {
    setFormData(templates[idx].data);
    showToast("テンプレートを読み込みました");
  }
}

function deleteTemplate(idx) {
  if (!confirm("このテンプレートを削除しますか？")) return;
  const templates = getTemplates();
  templates.splice(idx, 1);
  localStorage.setItem(TEMPLATE_KEY, JSON.stringify(templates));
  renderTemplates();
  showToast("テンプレートを削除しました");
}

function renderTemplates() {
  const c = document.getElementById("templateList");
  if (!c) return;
  const templates = getTemplates();
  if (templates.length === 0) {
    c.innerHTML =
      '<div style="text-align:center;color:var(--text-muted);padding:20px;font-size:13px">保存済みテンプレートはありません</div>';
    return;
  }
  c.innerHTML = templates
    .map(
      (t, i) => `
    <div class="template-item">
      <div><div class="template-item-name">${t.name}</div><div class="template-item-date">${t.date}</div></div>
      <div class="template-actions">
        <button class="btn btn-sm btn-primary" onclick="loadTemplate(${i})">読込</button>
        <button class="btn btn-sm btn-danger" onclick="deleteTemplate(${i})">削除</button>
      </div>
    </div>`,
    )
    .join("");
}

/* ===== PREVIEW & PDF ===== */
function openPreview() {
  if (!validateAll()) {
    showToast("必須項目を入力してください");
    switchTab("basic");
    return;
  }
  const m = document.getElementById("previewModal");
  m.classList.add("show");
  const body = document.getElementById("previewBody");
  const aux = Number(document.getElementById("auxMaterial").value) || 0;
  let sellTotal = 0,
    costTotal = 0;
  parts.forEach((p) => {
    const s = floor100(p.sellInput);
    sellTotal += s * p.qty;
    costTotal +=
      (p.source === "self" ? Math.round(s * 0.62) : p.costPrice) * p.qty;
  });
  let laborSum = 0;
  labors.forEach((l) => (laborSum += l.amount));
  let discSum = 0;
  discounts.forEach((d) => (discSum += d.amount));
  const subtotal = sellTotal + aux + laborSum - discSum;
  const tax = Math.floor(subtotal * 0.1);
  const total = subtotal + tax;
  const staff = staffList[selectedStaff] || "";
  const dateVal = document.getElementById("issueDate").value;

  let partsRows = "";
  for (let i = 0; i < 15; i++) {
    const p = parts[i];
    if (p) {
      const sell = floor100(p.sellInput);
      partsRows += `<tr><td style="text-align:center">${i + 1}</td><td>${p.name}</td><td>${p.code}</td><td style="text-align:center">${p.qty}</td><td style="text-align:left">${fmtN(sell)}</td><td style="text-align:left">${fmtN(sell * p.qty)}</td><td>${p.note}</td></tr>`;
    } else {
      partsRows += `<tr><td style="text-align:center">${i + 1}</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`;
    }
  }
  const allLaborsAndDiscounts = [];
  labors.forEach(l => allLaborsAndDiscounts.push({type: 'labor', ...l}));
  discounts.forEach(d => allLaborsAndDiscounts.push({type: 'disc', ...d}));

  let laborDiscRowsHtml = "";
  if (allLaborsAndDiscounts.length === 0) {
     laborDiscRowsHtml = `<tr><td style="width:90px;text-align:center;padding:1px 2px;letter-spacing:1px"><span style="font-size:8pt;font-weight:700">作業名<br>及び値引き</span></td><td style="padding:1px 2px;border-top:none"></td><td style="width:55px;padding:1px 2px;border-top:none"></td><td style="width:145px;padding:1px 2px;border-top:none"></td></tr>`;
  } else {
     allLaborsAndDiscounts.forEach((item, index) => {
         let rowHtml = '<tr>';
         if (index === 0) {
             rowHtml += `<td rowspan="${allLaborsAndDiscounts.length}" style="width:90px;text-align:center;padding:1px 2px;letter-spacing:1px;border-top:none"><span style="font-size:8pt;font-weight:700">作業名<br>及び値引き</span></td>`;
         }
         const qtyText = item.type === 'labor' ? (item.amount ? "一式" : "") : "";
         rowHtml += `<td style="padding:1px 2px;border-top:${index===0?'none':'1px solid #333'}">${item.name}</td>`;
         rowHtml += `<td style="width:55px;text-align:center;padding:1px 2px;border-top:${index===0?'none':'1px solid #333'}">${qtyText}</td>`;
         rowHtml += `<td style="width:145px;text-align:left;padding:1px 2px;border-top:${index===0?'none':'1px solid #333'}">${item.amount ? fmtN(item.amount) : ""}</td>`;
         rowHtml += '</tr>';
         laborDiscRowsHtml += rowHtml;
     });
  }
  const dFmt = dateVal.replace(/(\d{4})-(\d{2})-(\d{2})/, "$1年$2月$3日");

  // Calculate scale so it fits entirely without scroll (overlay padding + modal body padding = ~80px width, ~140px height)
  const aw = window.innerWidth - 90;
  const ah = window.innerHeight - 160;
  const scale = Math.min(1, aw / 794, ah / 1123);
  const wrapperStyle = scale < 1 ? `width:${794*scale}px;height:${1123*scale}px;margin:0 auto;overflow:hidden;` : '';
  const scaleStyle = scale < 1 ? `transform:scale(${scale});transform-origin:top left;` : '';

  let html = "";
  if (scale < 1) html += `<div id="pdfWrapper" style="${wrapperStyle}">`;
  html += `
    <div id="pdfPrintArea" style="width:794px;height:1123px;padding:14px 28px;box-sizing:border-box;background:#fff;color:#000;font-family:'Hiragino Kaku Gothic ProN','Yu Gothic',sans-serif;margin:0 auto;overflow:hidden;position:relative;${scaleStyle}">
      <div style="text-align:center;font-size:18pt;font-weight:700;letter-spacing:14px;margin-bottom:4px;">御　見　積　書</div>
      <div style="display:flex;justify-content:flex-end;gap:24px;font-size:8pt;margin-bottom:4px">
        <span>見積書番号：${document.getElementById("estNo").value}</span><span>${dFmt}</span>
      </div>
      <div style="display:flex;gap:12px;margin-bottom:6px">
        <div style="flex:1">
          <div style="font-size:13pt;font-weight:700;border-bottom:2px solid #000;padding-bottom:1px;margin-bottom:4px">${document.getElementById("custName").value}　<span style="font-size:10pt">御中</span></div>
          <table class="pv-table" style="font-size:8pt;width:100%">
            <tr><td style="width:80px;background:#f5f5f5;padding:1px 3px">納入場所</td><td style="padding:1px 3px">${document.getElementById("delivPlace").value}</td></tr>
            <tr><td style="background:#f5f5f5;padding:1px 3px">納期</td><td style="padding:1px 3px">${document.getElementById("delivDate").value}</td></tr>
            <tr><td style="background:#f5f5f5;padding:1px 3px">御支払条件</td><td style="padding:1px 3px">${document.getElementById("payTerms").value}</td></tr>
            <tr><td style="background:#f5f5f5;padding:1px 3px">見積有効期限</td><td style="padding:1px 3px">${document.getElementById("validity").value}</td></tr>
          </table>
        </div>
        <div style="flex:1;font-size:7.5pt;text-align:right">
          <div style="font-weight:700;font-size:11pt;margin-bottom:1px">${document.getElementById("compName").value}</div>
          <div style="margin-bottom:0">${document.getElementById("compAddr").value}</div>
          <div style="margin-bottom:0">TEL:${document.getElementById("compTel").value} FAX:${document.getElementById("compFax").value}</div>
          <div style="margin-top:1px;margin-bottom:0">${document.getElementById("compDept").value}</div>
          <div>${document.getElementById("compBranch").value}</div>
          <div style="margin-top:2px;border-bottom:1px solid #000;display:inline-block;padding-bottom:0;min-width:110px;text-align:left">担当者：${staff}</div>
          <div style="margin-top:4px;display:flex;justify-content:flex-end;gap:6px">
            <div><div style="text-align:center;font-size:6pt;margin-bottom:0">上長</div><div class="stamp-box" style="width:48px;height:48px"></div></div>
            <div><div style="text-align:center;font-size:6pt;margin-bottom:0">担当者</div><div class="stamp-box" id="pdfStampBox" style="width:48px;height:48px"></div></div>
          </div>
        </div>
      </div>
      <table class="pv-table" style="margin-bottom:3px;font-size:8pt">
        <tr><td style="width:20%;background:#f5f5f5;font-weight:600;padding:1px 3px">製品名</td><td style="width:30%;padding:1px 3px">${document.getElementById("prodName").value}</td><td style="width:20%;background:#f5f5f5;font-weight:600;padding:1px 3px">機種</td><td style="width:30%;padding:1px 3px">${document.getElementById("modelNo").value}</td></tr>
        <tr><td style="background:#f5f5f5;font-weight:600;padding:1px 3px">設置日</td><td style="padding:1px 3px">${document.getElementById("installDate").value ? document.getElementById("installDate").value.replace(/(\d{4})-(\d{2})-(\d{2})/, "$1年$2月$3日") : ""}</td><td style="background:#f5f5f5;font-weight:600;padding:1px 3px">機番</td><td style="padding:1px 3px">${document.getElementById("serialNo").value}</td></tr>
      </table>
      <div style="font-size:7.5pt;margin-bottom:1px"><b>【コール内容】</b>${document.getElementById("callContent").value}</div>
      <div style="font-size:7.5pt;margin-bottom:4px"><b>【故障原因・処置】</b>${document.getElementById("causeAction").value}</div>
      <div style="font-size:8pt;font-weight:700;margin-bottom:2px;border:1px solid #333;padding:1px 6px;background:#f9f9f9">故障箇所の部品のみの修繕内容です。</div>
      <table class="pv-table" style="font-size:8pt">
        <tr style="background:#f5f5f5;font-weight:700"><td style="width:24px;text-align:center;padding:1px 2px">No</td><td style="padding:1px 2px">部品名</td><td style="width:80px;padding:1px 2px">部品コード</td><td style="width:24px;text-align:center;padding:1px 2px">数量</td><td style="width:55px;text-align:left;padding:1px 2px">単価</td><td style="width:65px;text-align:left;padding:1px 2px">金額</td><td style="width:80px;padding:1px 2px">備考</td></tr>
        ${partsRows.replace(/<td/g, '<td style="padding:1px 2px"')}
        <tr><td colspan="4" style="text-align:center;font-weight:600;padding:1px 2px">補修・補助材料費</td><td style="text-align:center;padding:1px 2px">一式</td><td style="text-align:left;padding:1px 2px">${fmtN(aux)}</td><td></td></tr>
        <tr style="font-weight:700"><td colspan="5" style="text-align:center;padding:1px 2px">使用部品費合計</td><td style="text-align:left;padding:1px 2px">${fmtN(sellTotal + aux)}</td><td></td></tr>
      </table>
      <table class="pv-table" style="margin-top:-1px;font-size:8pt">
        ${laborDiscRowsHtml}
      </table>
      <table class="pv-table" style="margin-top:2px;font-size:8pt">
        <tr><td style="border:none"></td><td style="text-align:center;font-weight:600;width:55px;padding:1px 2px">小計</td><td style="text-align:right;width:145px;padding:1px 2px">${fmtN(subtotal)}</td></tr>
        <tr><td style="border:none"></td><td style="text-align:center;font-weight:600;width:55px;padding:1px 2px">消費税</td><td style="text-align:right;width:145px;padding:1px 2px">${fmtN(tax)}</td></tr>
        <tr class="pv-total-row"><td style="border:none"></td><td style="text-align:center;background:#1a365d;color:#fff;font-size:11pt;font-weight:700;padding:2px">合計</td><td style="text-align:right;background:#1a365d;color:#fff;font-size:13pt;padding:2px">¥${fmtN(total)}</td></tr>
      </table>
      <table class="pv-table" style="margin-top:4px;font-size:8pt">
        <tr><td rowspan="3" style="width:24px;padding:1px 2px" class="pv-vert"><span style="font-size:7pt">御返答欄</span></td><td colspan="6" style="text-align:center;padding:1px 2px">上記御見積書の御返答を御願い致します。</td></tr>
        <tr><td colspan="3" style="padding:1px 2px">１．修理する</td><td style="width:60px;padding:1px 2px">御返答日</td><td colspan="2" style="text-align:right;padding:1px 2px">　　年　　月　　日</td></tr>
        <tr><td colspan="3" style="padding:1px 2px">２．修理しない</td><td style="padding:1px 2px">発注者</td><td colspan="2" style="text-align:right;padding:1px 2px">様　　印</td></tr>
      </table>
      <div style="font-size:6.5pt;margin-top:4px;white-space:pre-wrap;line-height:1.2">${document.getElementById("remarks").value}</div>
    </div>`;
  if (scale < 1) html += `</div>`;
  body.innerHTML = html;

  const sb = document.getElementById("pdfStampBox");
  if (sb) {
    // Canvas at 88px (2x for retina), display at 44px to fit inside 48px box (2px border)
    const sc = createStampCanvas(staff, dateVal, 88);
    sc.style.width = "44px";
    sc.style.height = "44px";
    sb.appendChild(sc);
  }
}

function closePreview() {
  document.getElementById("previewModal").classList.remove("show");
}

async function downloadPDF() {
  const element = document.getElementById("pdfPrintArea");
  if (!element) {
    showToast("先にプレビューを開いてください。");
    return;
  }
  const btn = document.querySelector(
    "#previewModal .modal-header .btn-primary",
  );
  const originalText = btn.innerHTML;
  btn.innerHTML = "⏳ 生成中...";
  btn.disabled = true;
  const modal = element.closest(".modal"),
    modalOverlay = element.closest(".modal-overlay"),
    modalBody = document.getElementById("previewBody");
  const saved = {
    mh: modal.style.maxHeight,
    mo: modal.style.overflowY,
    oo: modalOverlay.style.overflowY,
  };
  modal.style.maxHeight = "none";
  modal.style.overflowY = "visible";
  modalOverlay.style.overflowY = "visible";
  if (modalBody) modalBody.scrollTop = 0;

  const wrapper = document.getElementById("pdfWrapper");
  const savedTransform = element.style.transform;
  const savedOrigin = element.style.transformOrigin;
  let savedWH = null;
  if (wrapper) {
      savedWH = { w: wrapper.style.width, h: wrapper.style.height };
      wrapper.style.width = "794px";
      wrapper.style.height = "1123px";
  }
  element.style.transform = "none";
  element.style.transformOrigin = "initial";

  setTimeout(async () => {
    try {
      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
        scrollY: -window.scrollY,
        scrollX: 0,
      });
      
      element.style.transform = savedTransform;
      element.style.transformOrigin = savedOrigin;
      if (wrapper && savedWH) {
          wrapper.style.width = savedWH.w;
          wrapper.style.height = savedWH.h;
      }
      modal.style.maxHeight = saved.mh;
      modal.style.overflowY = saved.mo;
      modalOverlay.style.overflowY = saved.oo;
      const imgData = canvas.toDataURL("image/jpeg", 0.98);
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        unit: "px",
        format: [794, 1123],
        orientation: "portrait",
        hotfixes: ["px_scaling"],
      });
      pdf.addImage(imgData, "JPEG", 0, 0, 794, 1123);
      const pdfFilename =
        "御見積書_" +
        (document.getElementById("estNo").value || "draft") +
        ".pdf";
      const pdfBlob = pdf.output("blob");
      if (window.showSaveFilePicker) {
        try {
          const handle = await showSaveFilePicker({
            suggestedName: pdfFilename,
            types: [
              { description: "PDF", accept: { "application/pdf": [".pdf"] } },
            ],
          });
          const w = await handle.createWritable();
          await w.write(pdfBlob);
          await w.close();
        } catch (e) {
          if (e.name !== "AbortError") alert("保存エラー: " + e.message);
        }
      } else {
        pdf.save(pdfFilename);
      }
      btn.innerHTML = originalText;
      btn.disabled = false;
      showToast("PDFを出力しました");
    } catch (err) {
      element.style.transform = savedTransform;
      element.style.transformOrigin = savedOrigin;
      if (wrapper && savedWH) {
          wrapper.style.width = savedWH.w;
          wrapper.style.height = savedWH.h;
      }
      modal.style.maxHeight = saved.mh;
      modal.style.overflowY = saved.mo;
      modalOverlay.style.overflowY = saved.oo;
      btn.disabled = false;
      alert("PDF生成エラー: " + err.message);
    }
  }, 300);
}

async function sharePDF() {
  const element = document.getElementById("pdfPrintArea");
  if (!element) {
    showToast("先にプレビューを開いてください。");
    return;
  }
  
  if (!navigator.share || !navigator.canShare) {
    alert("お使いのブラウザ/端末は共有機能に対応していません。\nPDFダウンロードをご利用ください。");
    return;
  }

  const btn = document.getElementById("shareBtn");
  const originalText = btn.innerHTML;
  btn.innerHTML = "⏳ 準備中...";
  btn.disabled = true;

  // Temporarily remove constraints for a full capture
  const modal = element.closest(".modal"),
    modalOverlay = element.closest(".modal-overlay"),
    modalBody = document.getElementById("previewBody");
  const saved = {
    mh: modal.style.maxHeight,
    mo: modal.style.overflowY,
    oo: modalOverlay.style.overflowY,
  };
  modal.style.maxHeight = "none";
  modal.style.overflowY = "visible";
  modalOverlay.style.overflowY = "visible";
  if (modalBody) modalBody.scrollTop = 0;

  const wrapper = document.getElementById("pdfWrapper");
  const savedTransform = element.style.transform;
  const savedOrigin = element.style.transformOrigin;
  let savedWH = null;
  if (wrapper) {
      savedWH = { w: wrapper.style.width, h: wrapper.style.height };
      wrapper.style.width = "794px";
      wrapper.style.height = "1123px";
  }
  element.style.transform = "none";
  element.style.transformOrigin = "initial";

  setTimeout(async () => {
    try {
      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
        scrollY: -window.scrollY,
        scrollX: 0,
      });

      // Restore constraints immediately
      element.style.transform = savedTransform;
      element.style.transformOrigin = savedOrigin;
      if (wrapper && savedWH) {
          wrapper.style.width = savedWH.w;
          wrapper.style.height = savedWH.h;
      }
      modal.style.maxHeight = saved.mh;
      modal.style.overflowY = saved.mo;
      modalOverlay.style.overflowY = saved.oo;

      const imgData = canvas.toDataURL("image/jpeg", 0.98);
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        unit: "px",
        format: [794, 1123],
        orientation: "portrait",
        hotfixes: ["px_scaling"],
      });
      pdf.addImage(imgData, "JPEG", 0, 0, 794, 1123);
      
      const pdfFilename = "御見積書_" + (document.getElementById("estNo").value || "draft") + ".pdf";
      const pdfBlob = pdf.output("blob");
      const pdfFile = new File([pdfBlob], pdfFilename, { type: "application/pdf" });

      if (navigator.canShare({ files: [pdfFile] })) {
        await navigator.share({
          files: [pdfFile],
          title: "御見積書",
          text: document.getElementById("custName").value + "宛の御見積書です。",
        });
        showToast("共有メニューを開きました");
      } else {
        alert("お使いのブラウザはPDFファイルの直接共有に対応していません。ダウンロードを使用してください。");
      }
    } catch (err) {
      if (err.name !== "AbortError") {
        alert("共有エラー: " + err.message);
      }
    } finally {
      btn.innerHTML = originalText;
      btn.disabled = false;
    }
  }, 300);
}
