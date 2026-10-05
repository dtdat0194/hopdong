/* Lớp nối cho bản chạy thẳng trong trình duyệt.
 *
 * Pyodide nạp đúng những module Python trong CONTRACTS/app/, nên bản web này
 * và bản chạy trên máy dùng chung một lõi - sửa hợp đồng ở một nơi là cả hai
 * cùng đổi. File Excel không hề được gửi đi đâu: nó được đọc ngay trong tab.
 */

const BACKEND = (() => {
  const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/";
  const GOC = new URL(".", document.baseURI).href;

  // Thứ tự không quan trọng, Python tự giải quyết; liệt kê để biết cần tải gì.
  const MODULE = ["__init__", "util", "contract_spec", "measure", "docx_read",
                  "excel", "render_docx", "render_pdf", "browser"];
  const FONT = ["Tinos-Regular", "Tinos-Bold", "Tinos-Italic", "Tinos-BoldItalic"];
  const TEMPLATE = ["TEMPLATE_Hop_dong_lao_dong_Fulltime",
                    "TEMPLATE_Hop_dong_dich_vu_ky_su"];

  let py = null;
  let loi = null;

  function nap(src) {
    return new Promise((ok, fail) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = ok;
      s.onerror = () => fail(new Error("Không tải được " + src));
      document.head.appendChild(s);
    });
  }

  async function tai(duong) {
    const res = await fetch(GOC + duong);
    if (!res.ok) throw new Error(`Không tải được ${duong} (lỗi ${res.status})`);
    return res;
  }

  /** Python ném lỗi kèm nguyên traceback; chỉ lấy dòng cuối cho người dùng đọc. */
  function gon(e) {
    const dong = String(e && e.message ? e.message : e).trim().split("\n");
    return dong[dong.length - 1].replace(/^[\w.]*(Error|Exception):\s*/, "");
  }

  function goi(ten, ...args) {
    try {
      return py.globals.get("qcd").browser[ten](...args);
    } catch (e) {
      throw new Error(gon(e));
    }
  }

  /** Giá trị bytes của Python sang Uint8Array của JS. */
  function byte(proxy) {
    const u8 = proxy.toJs ? proxy.toJs() : proxy;
    if (proxy.destroy) proxy.destroy();
    return u8;
  }

  return {
    mauExcel: null,          // bản web không kèm file Excel mẫu

    async khoiDong(buoc) {
      if (py) return;
      if (loi) throw loi;
      try {
        buoc("Đang tải Python cho trình duyệt…");
        await nap(PYODIDE + "pyodide.js");
        py = await loadPyodide({ indexURL: PYODIDE });

        buoc("Đang nạp thư viện xử lý Word và PDF…");
        await py.loadPackage(["micropip", "lxml", "pillow", "jinja2", "typing-extensions"]);
        const micropip = py.pyimport("micropip");
        await micropip.install(["reportlab", "python-docx", "docxtpl", "openpyxl"]);

        buoc("Đang nạp mã nguồn và mẫu hợp đồng…");
        py.FS.mkdirTree("/app/qcd");
        py.FS.mkdirTree("/app/templates");
        py.FS.mkdirTree("/app/fonts");

        await Promise.all([
          ...MODULE.map(async (m) => py.FS.writeFile(
            `/app/qcd/${m}.py`, await (await tai(`CONTRACTS/app/${m}.py`)).text())),
          ...TEMPLATE.map(async (t) => py.FS.writeFile(
            `/app/templates/${t}.docx`,
            new Uint8Array(await (await tai(`CONTRACTS/templates/${t}.docx`)).arrayBuffer()))),
          ...FONT.map(async (f) => py.FS.writeFile(
            `/app/fonts/${f}.ttf`,
            new Uint8Array(await (await tai(`CONTRACTS/fonts/${f}.ttf`)).arrayBuffer()))),
        ]);

        buoc("Đang khởi động bộ dựng hợp đồng…");
        py.runPython(`
import sys
sys.path.insert(0, "/app")
import qcd.browser
import qcd.render_pdf as _rp
_rp.register_family("Times New Roman")   # nạp font sẵn cho lần xuất đầu nhanh hơn
`);
      } catch (e) {
        loi = new Error(gon(e));
        py = null;
        throw loi;
      }
    },

    async napExcel(file) {
      const u8 = new Uint8Array(await file.arrayBuffer());
      return JSON.parse(goi("nap_excel", u8));
    },

    async xemTruoc(rows) {
      return new Blob([byte(goi("xuat_pdf", py.toPy(rows)))], { type: "application/pdf" });
    },

    async xuatMot(row, format) {
      const ten = goi("ten_hop_dong", row);
      if (format === "docx") {
        return {
          blob: new Blob([byte(goi("xuat_docx", row))], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }),
          ten: `${ten}.docx`,
        };
      }
      return {
        blob: new Blob([byte(goi("xuat_pdf", py.toPy([row])))], { type: "application/pdf" }),
        ten: `${ten}.pdf`,
      };
    },

    async gopPdf(rows) {
      return {
        blob: new Blob([byte(goi("xuat_pdf", py.toPy(rows)))], { type: "application/pdf" }),
        ten: `Hop_dong_${rows.length}_nhan_su.pdf`,
      };
    },
  };
})();
