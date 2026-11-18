/* Bản web (GitHub Pages): thay lời gọi /api/... tới máy chủ bằng Python chạy ngay trong trình duyệt. */
(() => {
  const goc = document.currentScript.src;
  const tho = new Worker(new URL('tho.js?v=' + (window.SKT_PHIEN || ''), goc));
  const cho = new Map();
  let so = 0;

  const man = document.createElement('div');
  man.id = 'webTai';
  man.innerHTML = '<div><h2>Sổ kiểm thực</h2><p id="webTaiNd">Đang khởi động...</p><p class="mo" id="webTaiPhu">Bản web chạy ngay trên máy này. Thực đơn và sổ không gửi đi đâu.</p></div>';
  const ve = () => document.body.appendChild(man);
  document.body ? ve() : addEventListener('DOMContentLoaded', ve);

  let xong, hong;
  const sanSang = new Promise((a, b) => { xong = a; hong = b; });
  tho.onmessage = e => {
    const d = e.data;
    if (d.loai === 'tien_do') { const p = document.getElementById('webTaiNd'); if (p) p.textContent = d.nd; return; }
    if (d.loai === 'xong') { const p = document.getElementById('webTaiNd'); if (p) p.textContent = 'Đang mở thực đơn và sổ đã lưu...'; xong(); return; }
    if (d.loai === 'hong') {
      document.getElementById('webTaiNd').textContent = 'Không khởi động được: ' + d.nd;
      document.getElementById('webTaiPhu').textContent = 'Hãy kiểm tra mạng rồi tải lại trang. Lần đầu cần mạng để tải Python.';
      hong(new Error(d.nd)); return;
    }
    if (d.loai === 'loi_luu') { console.warn('Không lưu được vào trình duyệt:', d.nd); return; }
    const c = cho.get(d.id);
    if (c) { cho.delete(d.id); c(d); }
  };

  async function goiPy(url, opt = {}) {
    await sanSang;
    const u = new URL(url, location.href);
    let than = null;
    if (opt.body instanceof Blob) than = await opt.body.arrayBuffer();
    else if (typeof opt.body === 'string') than = new TextEncoder().encode(opt.body).buffer;
    const id = ++so;
    const kq = await new Promise(ok => {
      cho.set(id, ok);
      tho.postMessage({id, phuong: (opt.method || 'GET').toUpperCase(), duong: u.pathname.replace(/^.*\/api\//, '/api/') + u.search, than},
        than ? [than] : []);
    });
    return kq;
  }

  // goi() trong giao diện dùng như fetch
  async function webFetch(url, opt) {
    const kq = await goiPy(url, opt);
    return new Response(kq.than, {status: kq.ma, headers: {'Content-Type': kq.kieu || 'application/json'}});
  }

  // Bấm tên file để tải về: lấy file từ bộ nhớ trình duyệt
  document.addEventListener('click', async e => {
    const a = e.target.closest && e.target.closest('a[href*="/api/tai_ve"]');
    if (!a) return;
    e.preventDefault();
    const kq = await goiPy(a.getAttribute('href'));
    if (kq.ma !== 200) { alert('Không tìm thấy file'); return; }
    const url = URL.createObjectURL(new Blob([kq.than], {type: kq.kieu}));
    const t = document.createElement('a');
    t.href = url; t.download = kq.ten || 'so.xlsx';
    document.body.appendChild(t); t.click(); t.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }, true);

  window.WEB = {sanSang, fetch: webFetch, anTai: () => man.remove()};
})();
