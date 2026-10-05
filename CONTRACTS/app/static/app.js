const $ = (id) => document.getElementById(id);

/* Lớp BACKEND do backend-server.js (chạy trên máy) hoặc backend-pyodide.js
   (chạy thẳng trong trình duyệt) cung cấp, nên phần giao diện dưới đây dùng
   chung cho cả hai bản. */

const state = {
  people: [],
  selected: new Set(),
  filter: "all",
  query: "",
  drawerRow: null,
};

const TYPE_CLASS = { HDLD: "hdld", HDKS: "hdks" };

/* ───────────────────────── helpers ───────────────────────── */

function toast(msg, kind = "") {
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.textContent = msg;
  $("toasts").appendChild(el);
  setTimeout(() => el.remove(), kind === "err" ? 7000 : 4000);
}

function busy(on, text = "Đang xử lý…") {
  $("loadingText").textContent = text;
  $("loading").hidden = !on;
}

function download(blob, ten) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = ten;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/* ───────────────────────── tải file lên ───────────────────────── */

async function handleFile(file) {
  if (!file) return;
  busy(true, "Đang chuẩn bị…");
  try {
    await BACKEND.khoiDong((buoc) => busy(true, buoc));
    busy(true, "Đang đọc file Excel…");
    const data = await BACKEND.napExcel(file);
    state.people = data.people;
    state.selected = new Set(data.people.filter((p) => p.marked && p.type).map((p) => p.row));

    $("uploadView").hidden = true;
    $("listView").hidden = false;
    $("btnReset").hidden = false;
    $("fileChip").hidden = false;
    $("fileChip").textContent = file.name;

    renderStats();
    render();
    toast(`Đã nạp ${data.people.length} nhân sự từ ${file.name}`, "ok");
  } catch (e) {
    toast(e.message, "err");
  } finally {
    busy(false);
  }
}

const dz = $("dropzone");
dz.addEventListener("click", () => $("fileInput").click());
dz.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("fileInput").click(); }
});
$("fileInput").addEventListener("change", (e) => handleFile(e.target.files[0]));
["dragenter", "dragover"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
["dragleave", "drop"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
dz.addEventListener("drop", (e) => handleFile(e.dataTransfer.files[0]));

$("btnReset").addEventListener("click", () => {
  Object.assign(state, { people: [], selected: new Set(), filter: "all", query: "" });
  $("uploadView").hidden = false;
  $("listView").hidden = true;
  $("btnReset").hidden = true;
  $("fileChip").hidden = true;
  $("actionbar").hidden = true;
  $("fileInput").value = "";
  $("search").value = "";
});

/* ───────────────────────── hiển thị ───────────────────────── */

function renderStats() {
  const p = state.people;
  $("stTotal").textContent = p.length;
  $("stHdld").textContent = p.filter((x) => x.type === "HDLD").length;
  $("stHdks").textContent = p.filter((x) => x.type === "HDKS").length;
  $("stReady").textContent = p.filter((x) => x.ready).length;
  $("stMissing").textContent = p.filter((x) => !x.ready).length;
}

function visible() {
  const q = state.query.trim().toLowerCase();
  return state.people.filter((p) => {
    if (state.filter === "HDLD" && p.type !== "HDLD") return false;
    if (state.filter === "HDKS" && p.type !== "HDKS") return false;
    if (state.filter === "ready" && !p.ready) return false;
    if (state.filter === "missing" && p.ready) return false;
    if (!q) return true;
    return [p.name, p.role, p.department, p.employment].join(" ").toLowerCase().includes(q);
  });
}

function statusCell(p) {
  if (!p.type) return `<span class="status bad">Chưa chọn loại HĐ</span>`;
  if (p.ready) return `<span class="status ok">Đủ thông tin</span>`;
  return `<span class="status warn">Thiếu ${p.missing.length} mục</span>`;
}

function render() {
  const rows = visible();
  $("tbody").innerHTML = rows.map((p) => `
    <tr data-row="${p.row}" class="${state.selected.has(p.row) ? "selected" : ""}">
      <td class="col-check"><input type="checkbox" data-check="${p.row}" ${state.selected.has(p.row) ? "checked" : ""}
          ${p.type ? "" : "disabled"} aria-label="Chọn ${p.name}"></td>
      <td class="col-stt">${p.stt}</td>
      <td>
        <div class="person"><b>${p.name}</b><span>${[p.role, p.department].filter(Boolean).join(" · ") || "—"}</span></div>
      </td>
      <td><span class="badge ${TYPE_CLASS[p.type] || "none"}">${p.type_label}</span></td>
      <td class="mono">${p.contract_no
        ? (p.auto_no ? `${p.contract_no} <span class="dim">(tự cấp)</span>` : p.contract_no)
        : '<span class="dim">—</span>'}</td>
      <td class="mono">${p.start ? `${p.start} – ${p.end}` : '<span class="dim">—</span>'}</td>
      <td class="num">${p.salary ? `${p.salary} ₫` : '<span class="dim">—</span>'}</td>
      <td>${statusCell(p)}</td>
      <td><button class="link-btn" data-detail="${p.row}">Chi tiết</button></td>
    </tr>`).join("");

  $("emptyMsg").hidden = rows.length > 0;
  const selectable = rows.filter((p) => p.type);
  $("checkAll").checked = selectable.length > 0 && selectable.every((p) => state.selected.has(p.row));
  $("checkAll").indeterminate = !$("checkAll").checked && selectable.some((p) => state.selected.has(p.row));
  renderActionbar();
}

function renderActionbar() {
  const n = state.selected.size;
  $("actionbar").hidden = n === 0;
  $("selCount").textContent = n;
  $("mergeWrap").hidden = n < 2;

  const incomplete = state.people.filter((p) => state.selected.has(p.row) && !p.ready).length;
  $("selWarn").hidden = incomplete === 0;
  $("selWarn").textContent = incomplete
    ? `${incomplete} hợp đồng còn thiếu thông tin, chỗ trống sẽ để dấu chấm lửng`
    : "";
}

$("tbody").addEventListener("click", (e) => {
  const detail = e.target.closest("[data-detail]");
  if (detail) { openDrawer(Number(detail.dataset.detail)); return; }
  const box = e.target.closest("[data-check]");
  if (box) {
    const row = Number(box.dataset.check);
    box.checked ? state.selected.add(row) : state.selected.delete(row);
    e.target.closest("tr").classList.toggle("selected", box.checked);
    renderActionbar();
    render();
    return;
  }
  const tr = e.target.closest("tr[data-row]");
  if (tr) {
    const cb = tr.querySelector("[data-check]");
    if (cb && !cb.disabled) { cb.checked = !cb.checked; cb.dispatchEvent(new Event("click", { bubbles: true })); }
  }
});

$("checkAll").addEventListener("change", (e) => {
  visible().filter((p) => p.type).forEach((p) =>
    e.target.checked ? state.selected.add(p.row) : state.selected.delete(p.row));
  render();
});

$("search").addEventListener("input", (e) => { state.query = e.target.value; render(); });

$("filters").addEventListener("click", (e) => {
  const chip = e.target.closest(".chip");
  if (!chip) return;
  document.querySelectorAll(".chip").forEach((c) => c.classList.toggle("active", c === chip));
  state.filter = chip.dataset.filter;
  render();
});

/* ───────────────────────── ngăn chi tiết ───────────────────────── */

function openDrawer(row) {
  const p = state.people.find((x) => x.row === row);
  if (!p) return;
  state.drawerRow = row;
  $("drawerName").textContent = p.name;
  $("drawerRole").textContent = [p.type_label, p.role, p.employment].filter(Boolean).join(" · ");

  const alert = p.ready
    ? `<div class="alert ok">Đủ thông tin, có thể xuất hợp đồng.</div>`
    : p.type
      ? `<div class="alert warn">Còn thiếu ${p.missing.length} mục trong Excel:<ul>${
          p.missing.map((m) => `<li>${m}</li>`).join("")}</ul></div>`
      : `<div class="alert warn">Chưa chọn loại hợp đồng (HDLD hoặc HDKS) ở cột “Loại hợp đồng”.</div>`;

  const dl = p.detail.map(([k, v]) =>
    `<dt>${k}</dt><dd class="${v ? "" : "miss"}">${v || "chưa có"}</dd>`).join("");

  $("drawerBody").innerHTML = alert + `<dl class="dl">${dl}</dl>`;
  $("drawerPreview").disabled = !p.type;
  $("drawerPdf").disabled = !p.type;
  $("drawer").hidden = false;
  $("drawerBackdrop").hidden = false;
}

function closeDrawer() {
  $("drawer").hidden = true;
  $("drawerBackdrop").hidden = true;
  state.drawerRow = null;
}
$("drawerClose").addEventListener("click", closeDrawer);
$("drawerBackdrop").addEventListener("click", closeDrawer);
$("drawerPreview").addEventListener("click", () => preview([state.drawerRow]));
$("drawerPdf").addEventListener("click", () => exportFiles("pdf", [state.drawerRow], false));

/* ───────────────────────── xem trước & xuất ───────────────────────── */

async function preview(rows) {
  busy(true, "Đang dựng bản xem trước…");
  try {
    const url = URL.createObjectURL(await BACKEND.xemTruoc(rows));
    $("previewFrame").src = url;
    $("previewTitle").textContent = rows.length > 1
      ? `Xem trước ${rows.length} hợp đồng` : "Xem trước hợp đồng";
    $("previewModal").hidden = false;
  } catch (e) {
    toast(e.message, "err");
  } finally {
    busy(false);
  }
}

function closePreview() {
  $("previewModal").hidden = true;
  const src = $("previewFrame").src;
  $("previewFrame").src = "";
  if (src.startsWith("blob:")) URL.revokeObjectURL(src);
}
$("previewClose").addEventListener("click", closePreview);
$("previewModal").addEventListener("click", (e) => { if (e.target.id === "previewModal") closePreview(); });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function exportFiles(format, rows, merge) {
  if (!rows.length) return;
  const nhan = format === "pdf" ? "PDF" : "Word";
  try {
    if (merge && format === "pdf" && rows.length > 1) {
      busy(true, `Đang gộp ${rows.length} hợp đồng…`);
      const gop = await BACKEND.gopPdf(rows);
      download(gop.blob, gop.ten);
      toast(`Đã gộp ${rows.length} hợp đồng vào một file PDF.`, "ok");
      return;
    }
    // Tải lần lượt từng người để mỗi file mang đúng tên nhân sự đó.
    for (let i = 0; i < rows.length; i++) {
      busy(true, `Đang tạo hợp đồng ${i + 1}/${rows.length}…`);
      const mot = await BACKEND.xuatMot(rows[i], format);
      download(mot.blob, mot.ten);
      if (i < rows.length - 1) await sleep(350);
    }
    toast(rows.length === 1
      ? `Đã tải hợp đồng ${nhan}.`
      : `Đã tải ${rows.length} file ${nhan} riêng lẻ.`, "ok");
  } catch (e) {
    toast(e.message, "err");
  } finally {
    busy(false);
  }
}

const picked = () => [...state.selected];
$("btnPreview").addEventListener("click", () => preview(picked()));
$("btnPdf").addEventListener("click", () => exportFiles("pdf", picked(), $("mergeChk").checked));
$("btnDocx").addEventListener("click", () => exportFiles("docx", picked(), false));

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!$("previewModal").hidden) closePreview();
  else if (!$("drawer").hidden) closeDrawer();
});

/* ───────────────────────── khởi động ───────────────────────── */

const mauExcel = $("mauExcel");
if (mauExcel) {
  if (BACKEND.mauExcel) mauExcel.href = BACKEND.mauExcel;
  else mauExcel.closest(".hint-mau")?.remove();
}

/* Chuẩn bị sẵn ngay khi mở trang. Bản chạy trên máy xong tức thì; bản chạy
   trong trình duyệt phải tải Python nên báo tiến trình cho người dùng biết. */
const bootStatus = $("bootStatus");
const bao = (chu, kieu = "") => {
  if (!bootStatus) return;
  bootStatus.textContent = chu;
  bootStatus.className = `hint boot ${kieu}`;
  bootStatus.hidden = !chu;
};

BACKEND.khoiDong((buoc) => bao(buoc))
  .then(() => bao(""))
  .catch((e) => bao(`Không khởi động được: ${e.message}`, "err"));
