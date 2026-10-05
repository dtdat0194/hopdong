/* Lớp nối cho bản chạy trên máy: mọi việc nặng do máy chủ FastAPI làm. */

const BACKEND = (() => {
  let token = null;

  async function api(path, options) {
    const res = await fetch(path, options);
    if (!res.ok) {
      let msg = `Lỗi ${res.status}`;
      try {
        const j = await res.json();
        if (j.detail) msg = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
      } catch (_) { /* phản hồi không phải JSON */ }
      throw new Error(msg);
    }
    return res;
  }

  function tenFile(res, duPhong) {
    const cd = res.headers.get("Content-Disposition") || "";
    const sao = /filename\*=UTF-8''([^;]+)/i.exec(cd);
    if (sao) return decodeURIComponent(sao[1]);
    const thuong = /filename="([^"]+)"/i.exec(cd);
    return thuong ? thuong[1] : duPhong;
  }

  const xuat = (rows, format, merge) => api("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token, rows, format, merge }),
  });

  return {
    mauExcel: "/api/mau-excel",

    async khoiDong() { /* máy chủ đã sẵn sàng sẵn */ },

    async napExcel(file) {
      const fd = new FormData();
      fd.append("file", file);
      const data = await (await api("/api/upload", { method: "POST", body: fd })).json();
      token = data.token;
      return data;
    },

    async xemTruoc(rows) {
      const res = await api("/api/preview", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token, rows }),
      });
      return res.blob();
    },

    async xuatMot(row, format) {
      const res = await xuat([row], format, false);
      return { blob: await res.blob(), ten: tenFile(res, `hop-dong.${format}`) };
    },

    async gopPdf(rows) {
      const res = await xuat(rows, "pdf", true);
      return { blob: await res.blob(), ten: tenFile(res, "hop-dong.pdf") };
    },
  };
})();
