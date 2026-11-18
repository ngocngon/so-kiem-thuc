/* Luồng chạy nền của bản web: tải Python (Pyodide), chạy code làm sổ, lưu dữ liệu vào trình duyệt (IndexedDB). */
const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/';
const FILE_PY = ['app.py', 'tao_so.py', 'so_buoc1.py', 'dinh_luong.json', 'mon_biet.json', 'so_tay_mon.json', 'web/cau_noi.py'];
const DATA = '/du_lieu_web';
importScripts(PYODIDE + 'pyodide.js');

let py, goi, dangLuu = Promise.resolve();
const bao = (loai, nd) => postMessage({loai, nd});

function luu() {            // ghi bộ nhớ xuống IndexedDB, lần lượt từng lần một
  dangLuu = dangLuu.then(() => new Promise(ok => py.FS.syncfs(false, e => { if (e) bao('loi_luu', String(e)); ok(); })));
  return dangLuu;
}

const sanSang = (async () => {
  bao('tien_do', 'Đang tải Python (lần đầu mất khoảng 1 phút)...');
  py = await loadPyodide({indexURL: PYODIDE});
  bao('tien_do', 'Đang tải thư viện Excel...');
  await py.loadPackage(['micropip']);
  await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl')");
  bao('tien_do', 'Đang mở dữ liệu đã lưu trên máy này...');
  py.FS.mkdirTree(DATA);
  py.FS.mount(py.FS.filesystems.IDBFS, {}, DATA);
  await new Promise((ok, loi) => py.FS.syncfs(true, e => e ? loi(e) : ok()));
  bao('tien_do', 'Đang nạp code làm sổ...');
  const phien = new URL(self.location.href).searchParams.get('v') || Date.now();
  py.FS.mkdirTree('/app/web');
  for (const f of FILE_PY) {
    const r = await fetch(new URL('../' + f + '?v=' + phien, self.location.href));
    if (!r.ok) throw new Error('Không tải được ' + f);
    py.FS.writeFile('/app/' + f, new Uint8Array(await r.arrayBuffer()));
  }
  py.runPython(`
import os, sys
os.environ['SKT_WEB'] = '1'
os.environ['SKT_DU_LIEU'] = '${DATA}'
os.chdir('/app')
sys.path[:0] = ['/app', '/app/web']
import cau_noi
`);
  goi = py.globals.get('cau_noi').goi;
})();

onmessage = async e => {
  const {id, phuong, duong, than} = e.data;
  try {
    await sanSang;
    const kq = goi(phuong, duong, new Uint8Array(than || 0));
    const [ma, kieu, ten, noi] = kq.toJs({create_pyproxies: false});
    kq.destroy();
    const b = noi instanceof Uint8Array ? noi.slice() : new Uint8Array(noi);
    if (phuong === 'POST') await luu();
    postMessage({id, ma, kieu, ten, than: b.buffer}, [b.buffer]);
  } catch (err) {
    postMessage({id, ma: 500, kieu: 'application/json', ten: '',
      than: new TextEncoder().encode(JSON.stringify({loi: 'Lỗi bản web: ' + (err.message || err)})).buffer});
  }
};
sanSang.then(() => bao('xong'), err => bao('hong', String(err.message || err)));
