// OPPO Shop KPI & Sales Analytics Dashboard Logic
// Supports GitHub Pages static hosting, LocalStorage persistence of latest Excel upload, and dynamic Excel parsing.

const STORAGE_KEY = 'oppo_shop_dashboard_latest_data_v1';
let currentData = null;
let charts = {};

// Formatters
const fmtNumber = (n) => new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 1 }).format(n || 0);
const fmtCurrency = (n) => new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND', maximumFractionDigits: 0 }).format(n || 0);
const fmtPercent = (n) => `${((n || 0) * 100).toFixed(1)}%`;

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  loadInitialData();
  setupEventListeners();
  renderAll();
});

function loadInitialData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed && parsed.categories && parsed.meta) {
        currentData = parsed;
        updateStorageBadge(true);
        return;
      }
    }
  } catch (e) {
    console.warn('Không thể đọc LocalStorage, dùng dữ liệu mặc định:', e);
  }
  currentData = JSON.parse(JSON.stringify(window.DEFAULT_SHOP_DATA));
  updateStorageBadge(false);
}

function updateStorageBadge(isLocalStorage) {
  const badge = document.getElementById('storageBadge');
  if (!badge) return;
  if (isLocalStorage) {
    badge.className = 'inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200';
    badge.innerHTML = '<i data-lucide="hard-drive" class="w-3.5 h-3.5"></i> Đã lưu bản upload gần nhất';
  } else {
    badge.className = 'inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md font-semibold bg-blue-50 text-blue-700 border border-blue-200';
    badge.innerHTML = '<i data-lucide="database" class="w-3.5 h-3.5"></i> Dữ liệu mặc định (GitHub)';
  }
  lucide.createIcons();
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  const bg = type === 'success' ? 'bg-emerald-700 text-white' : type === 'error' ? 'bg-rose-600 text-white' : 'bg-slate-800 text-white';
  toast.className = `${bg} px-4 py-3 rounded-xl shadow-lg text-xs sm:text-sm font-medium flex items-center gap-2 animate-fade-in max-w-sm`;
  toast.innerHTML = `<span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function setupEventListeners() {
  // Tab switching
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => {
        b.classList.remove('active', 'border-emerald-600', 'text-emerald-700', 'font-semibold');
        b.classList.add('border-transparent', 'text-slate-500', 'font-medium');
      });
      btn.classList.add('active', 'border-emerald-600', 'text-emerald-700', 'font-semibold');
      btn.classList.remove('border-transparent', 'text-slate-500', 'font-medium');

      const tabId = btn.getAttribute('data-tab');
      document.querySelectorAll('.tab-content').forEach(sec => sec.classList.add('hidden'));
      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
    });
  });

  // File Input Upload
  const fileInput = document.getElementById('excelUploadInput');
  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleExcelFile(e.target.files[0]);
      e.target.value = '';
    }
  });

  // Drag and drop across window
  const overlay = document.getElementById('dropZoneOverlay');
  let dragCounter = 0;
  window.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    overlay.classList.remove('hidden');
    overlay.classList.add('flex');
  });
  window.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      overlay.classList.add('hidden');
      overlay.classList.remove('flex');
    }
  });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    overlay.classList.add('hidden');
    overlay.classList.remove('flex');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleExcelFile(e.dataTransfer.files[0]);
    }
  });

  // Reset to default data
  document.getElementById('resetStorageBtn').addEventListener('click', () => {
    localStorage.removeItem(STORAGE_KEY);
    currentData = JSON.parse(JSON.stringify(window.DEFAULT_SHOP_DATA));
    updateStorageBadge(false);
    renderAll();
    showToast('Đã khôi phục về dữ liệu mặc định ban đầu.');
  });

  // Export current data as JS/JSON for updating GitHub repo
  document.getElementById('exportJsonBtn').addEventListener('click', () => {
    const content = 'window.DEFAULT_SHOP_DATA = ' + JSON.stringify(currentData, null, 2) + ';\n';
    const blob = new Blob([content], { type: 'application/javascript;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'default-data.js';
    a.click();
    URL.revokeObjectURL(url);
    showToast('Đã tải file default-data.js! Bạn có thể thay thế file này vào thư mục data/ trên GitHub.');
  });

  // Transaction filters
  document.getElementById('txSourceFilter').addEventListener('change', renderTransactionsTable);
  document.getElementById('txTypeFilter').addEventListener('change', renderTransactionsTable);
  document.getElementById('txSearchInput').addEventListener('input', renderTransactionsTable);
}

// Parse Uploaded Excel File using SheetJS
function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const parsedData = parseWorkbookToDashboardData(workbook, file.name);

      currentData = parsedData;
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(currentData));
        updateStorageBadge(true);
      } catch (storageErr) {
        showToast('Dữ liệu quá lớn nên chỉ hiển thị tạm thời (không lưu được vào LocalStorage).', 'info');
      }

      renderAll();
      showToast(`Đã cập nhật & lưu trữ dữ liệu từ file: ${file.name}`);
    } catch (err) {
      console.error(err);
      showToast('Lỗi đọc file Excel! Vui lòng kiểm tra đúng định dạng báo cáo Tổng Hợp Sale.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbookToDashboardData(wb, fileName) {
  const num = (v) => {
    if (v === null || v === undefined || v === '' || typeof v === 'boolean') return 0;
    if (typeof v === 'number') return isNaN(v) ? 0 : v;
    const cleaned = String(v).replace(/,/g, '').trim();
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : parsed;
  };
  const str = (v) => (v === null || v === undefined) ? '' : String(v).trim();
  const fmtDate = (v) => {
    if (!v) return '';
    if (v instanceof Date && !isNaN(v)) {
      const yyyy = v.getFullYear();
      const mm = String(v.getMonth() + 1).padStart(2, '0');
      const dd = String(v.getDate()).padStart(2, '0');
      const hh = String(v.getHours()).padStart(2, '0');
      const mi = String(v.getMinutes()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
    }
    return String(v).slice(0, 19);
  };

  // 1. Parse 'Tổng hợp' sheet
  const thSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('tổng hợp') || n.toLowerCase().includes('tong hop')) || wb.SheetNames[0];
  const wsTh = wb.Sheets[thSheetName];
  const thRows = XLSX.utils.sheet_to_json(wsTh, { header: 1, defval: null });

  const row3 = thRows[2] || [];
  const shopName = str(row3[0]) || 'OPPO Experience & Service Store';

  const categories = [
    { key: 'TOTAL', name: 'Tổng Điểm', target: num(row3[1]), actual: num(row3[2]) },
    { key: 'PHONE', name: 'Phone', target: num(row3[4]), actual: num(row3[5]) },
    { key: 'TAI NGHE', name: 'Tai Nghe', target: num(row3[7]), actual: num(row3[8]) },
    { key: 'PAD', name: 'Pad', target: num(row3[10]), actual: num(row3[11]) },
    { key: 'ỐP LƯNG', name: 'Ốp Lưng', target: num(row3[13]), actual: num(row3[14]) },
    { key: 'OPPO CARE', name: 'OPPO Care', target: num(row3[16]), actual: num(row3[17]) },
    { key: 'MDCL', name: 'MDCL (Miếng dán)', target: num(row3[19]), actual: num(row3[20]) },
    { key: 'SẠC DỰ PHÒNG', name: 'Sạc Dự Phòng', target: num(row3[22]), actual: num(row3[23]) },
    { key: 'WATCH', name: 'Watch', target: num(row3[26]), actual: num(row3[27]) },
  ];
  const a7ProCount = num(row3[25]);

  // Personal productivity (rows 11-16 in 0-indexed)
  const productivity = [];
  for (let r = 11; r <= 16; r++) {
    const rData = thRows[r] || [];
    if (rData[3]) {
      productivity.push({
        region: str(rData[1]),
        code: str(rData[2]),
        name: str(rData[3]),
        find_gap: num(rData[4]),
        may_thang: num(rData[5]),
        sua_chua: num(rData[6]),
        dich_vu_cong_them: num(rData[7]),
        so_find_n: num(rData[8]),
        so_thuong: num(rData[9]),
        volume: num(rData[10])
      });
    }
  }

  // KL-DL (row 18 in 0-indexed)
  const r19 = thRows[18] || [];
  const klDl = {
    khach_le: num(r19[1]),
    pct_khach_le: num(r19[2]),
    dl_ka: num(r19[3]),
    pct_dl_ka: num(r19[4]),
    md_cr: num(r19[5]),
    md_so: num(r19[6]),
    mdcl_so: num(r19[7]),
    ty_le_md: num(r19[8])
  };

  // TAT KPI (row 23 in 0-indexed)
  const r24 = thRows[23] || [];
  const tatKpi = {
    tat_ks: num(r24[1]),
    bonus_ks: num(r24[2]),
    tat_gcsm: num(r24[3]),
    bonus_gcsm: num(r24[4]),
    squid_game_tat: num(r24[5])
  };

  // KPI Report table (rows 31-36 in 0-indexed)
  const kpiReport = [];
  for (let r = 31; r <= 36; r++) {
    const rData = thRows[r] || [];
    if (rData[1]) {
      kpiReport.push({
        name: str(rData[1]).split('/')[0].trim(),
        full_name: str(rData[1]),
        today: num(rData[2]),
        accumulated: num(rData[3]),
        target: num(rData[4]),
        rate: num(rData[5]),
        status: str(rData[6]).split('/')[0].trim()
      });
    }
  }

  // 2. Parse 'GCSM SO'
  const gcsmRecords = [];
  const gcsmSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('gcsm'));
  if (gcsmSheetName) {
    const gcsmRows = XLSX.utils.sheet_to_json(wb.Sheets[gcsmSheetName], { defval: null });
    gcsmRows.forEach(row => {
      if (!row['Sales order No.']) return;
      const tStr = fmtDate(row['Sales time'] || row['Created on']);
      const creatorRaw = str(row['Created by']);
      const creator = creatorRaw.split('（')[0].split('(')[0].trim();
      gcsmRecords.push({
        order_no: str(row['Sales order No.']),
        shop: str(row['Name of service center']),
        status: str(row['Order status']),
        creator: creator,
        time: tStr,
        date: tStr.slice(0, 10),
        payment: str(row['Payment method']) || 'Khác',
        paid_amount: num(row['Paid amount (receivable after deduction)']),
        retail_price: num(row['Retail price']),
        product: str(row['Sales Product Name']),
        qty: num(row['Quantity']),
        model: str(row['Marketing model']),
        type: str(row['Type']) || 'Phụ kiện / Khác',
        points: num(row['Điểm'])
      });
    });
  }

  // 3. Parse 'Center SO' (header is usually at row 3 -> range: 2)
  const centerRecords = [];
  const centerSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('center'));
  if (centerSheetName) {
    const centerRows = XLSX.utils.sheet_to_json(wb.Sheets[centerSheetName], { range: 2, defval: null });
    centerRows.forEach(row => {
      if (!row['Đơn hàng chi tiết']) return;
      const tStr = fmtDate(row['Ngày xuất bán']);
      centerRecords.push({
        order_no: str(row['Đơn hàng tổng']),
        shop: str(row['Tên Brandshop']),
        date: tStr.slice(0, 10),
        time: tStr,
        category: str(row['Danh mục']),
        product: str(row['Tên sản phẩm']),
        qty: num(row['Số lượng']),
        retail_revenue: num(row['Doanh thu bán lẻ']),
        receivable: num(row['Phải thu khách hàng']),
        creator: str(row['Người lên đơn']),
        kpi_staff: str(row['Name KPI 1']) || str(row['Người lên đơn']),
        export_type: str(row['Loại xuất']),
        status: str(row['Trạng thái']),
        model: str(row['Model']),
        type: str(row['Type']),
        points: num(row['Điểm'])
      });
    });
  }

  const now = new Date();
  const uploadedAt = `Cập nhật lúc ${now.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })} ngày ${now.toLocaleDateString('vi-VN')}`;

  return {
    meta: {
      fileName: fileName,
      uploadedAt: uploadedAt,
      shopName: shopName
    },
    categories,
    a7ProCount,
    productivity,
    klDl,
    tatKpi,
    kpiReport,
    gcsmRecords,
    centerRecords
  };
}

// Render Everything
function renderAll() {
  if (!currentData) return;

  // Header metadata
  document.getElementById('shopTitle').textContent = currentData.meta.shopName || 'OPPO Experience & Service Store';
  document.getElementById('currentFileName').textContent = currentData.meta.fileName || 'Báo cáo Sale.xlsx';
  document.getElementById('lastUpdatedText').textContent = currentData.meta.uploadedAt || '';
  document.getElementById('recordCountSummary').textContent = `${currentData.gcsmRecords.length} dòng GCSM SO • ${currentData.centerRecords.length} dòng Center SO`;

  renderHeroCards();
  renderCategoryGrid();
  renderKpiReportTable();
  renderServiceMiniStats();
  renderCharts();
  renderTopProducts();
  renderKlDlBox();
  renderStaffTable();
  populateTransactionFilters();
  renderTransactionsTable();
  lucide.createIcons();
}

function renderHeroCards() {
  const totalCat = currentData.categories.find(c => c.key === 'TOTAL') || { target: 0, actual: 0 };
  const phoneCat = currentData.categories.find(c => c.key === 'PHONE') || { target: 0, actual: 0 };
  const totalPct = totalCat.target > 0 ? (totalCat.actual / totalCat.target) : 0;
  const phonePct = phoneCat.target > 0 ? (phoneCat.actual / phoneCat.target) : 0;

  // Calculate total retail revenue from Center SO (Xuất bán + Hiệu lực) & GCSM paid
  const centerRev = currentData.centerRecords
    .filter(r => r.export_type.toLowerCase().includes('bán') && r.status.toLowerCase().includes('hiệu lực'))
    .reduce((sum, r) => sum + (r.receivable || r.retail_revenue || 0), 0);
  const gcsmRev = currentData.gcsmRecords.reduce((sum, r) => sum + (r.retail_price || 0), 0);
  const totalRevenue = centerRev + gcsmRev;

  const totalVolume = currentData.productivity.reduce((sum, p) => sum + (p.volume || 0), 0);

  const cards = [
    {
      title: 'TỔNG ĐIỂM HOÀN THÀNH',
      value: `${fmtNumber(totalCat.actual)} / ${fmtNumber(totalCat.target)}`,
      sub: `Đạt ${fmtPercent(totalPct)} chỉ tiêu tháng`,
      badge: totalPct >= 1 ? 'Vượt Target' : 'Đang chạy',
      badgeColor: totalPct >= 1 ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800',
      icon: 'award',
      iconBg: 'bg-emerald-50 text-emerald-600'
    },
    {
      title: 'ĐIỂM BÁN PHONE',
      value: `${fmtNumber(phoneCat.actual)} / ${fmtNumber(phoneCat.target)}`,
      sub: `Hoàn thành ${fmtPercent(phonePct)} • A7 Pro: ${currentData.a7ProCount} máy`,
      badge: phonePct >= 1 ? 'Hoàn thành' : 'Cần tăng tốc',
      badgeColor: phonePct >= 1 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800',
      icon: 'smartphone',
      iconBg: 'bg-blue-50 text-blue-600'
    },
    {
      title: 'DOANH THU GHI NHẬN',
      value: fmtCurrency(totalRevenue),
      sub: `Center SO: ${fmtCurrency(centerRev)}`,
      badge: 'Thực thu & Bán lẻ',
      badgeColor: 'bg-indigo-100 text-indigo-800',
      icon: 'banknote',
      iconBg: 'bg-indigo-50 text-indigo-600'
    },
    {
      title: 'NĂNG SUẤT SERVICE (VOLUME)',
      value: fmtNumber(totalVolume),
      sub: `TAT-GCSM: ${fmtPercent(currentData.tatKpi.tat_gcsm)} • Khách lẻ: ${fmtPercent(currentData.klDl.pct_khach_le)}`,
      badge: `${currentData.productivity.length} Kỹ thuật viên`,
      badgeColor: 'bg-teal-100 text-teal-800',
      icon: 'wrench',
      iconBg: 'bg-teal-50 text-teal-600'
    }
  ];

  document.getElementById('heroKpiCards').innerHTML = cards.map(c => `
    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm hover:shadow-md transition">
      <div class="flex items-start justify-between">
        <div class="p-2.5 rounded-xl ${c.iconBg}">
          <i data-lucide="${c.icon}" class="w-5 h-5"></i>
        </div>
        <span class="px-2.5 py-0.5 rounded-full text-[11px] font-semibold ${c.badgeColor}">${c.badge}</span>
      </div>
      <div class="mt-4">
        <p class="text-[11px] font-semibold tracking-wider text-slate-500 uppercase">${c.title}</p>
        <p class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">${c.value}</p>
        <p class="text-xs text-slate-500 mt-1">${c.sub}</p>
      </div>
    </div>
  `).join('');
}

function renderCategoryGrid() {
  const subCats = currentData.categories.filter(c => c.key !== 'TOTAL');
  document.getElementById('categoryProgressGrid').innerHTML = subCats.map(cat => {
    const pct = cat.target > 0 ? (cat.actual / cat.target) : (cat.actual > 0 ? 1 : 0);
    const pctDisplay = cat.target > 0 ? `${(pct * 100).toFixed(1)}%` : (cat.actual > 0 ? '100%' : '0%');
    const barWidth = Math.min(pct * 100, 100);

    let barColor = 'bg-emerald-500';
    let textColor = 'text-emerald-700';
    let bgBadge = 'bg-emerald-50';
    if (cat.target > 0 && pct < 0.8) {
      barColor = 'bg-rose-500';
      textColor = 'text-rose-700';
      bgBadge = 'bg-rose-50';
    } else if (cat.target > 0 && pct < 1) {
      barColor = 'bg-amber-500';
      textColor = 'text-amber-700';
      bgBadge = 'bg-amber-50';
    }

    return `
      <div class="p-4 rounded-xl border border-slate-200/80 bg-slate-50/50 hover:bg-white transition">
        <div class="flex items-center justify-between mb-2">
          <span class="text-xs font-bold text-slate-800">${cat.name}</span>
          <span class="text-xs font-bold px-2 py-0.5 rounded-md ${bgBadge} ${textColor}">${pctDisplay}</span>
        </div>
        <div class="flex items-baseline justify-between text-xs mb-2">
          <span class="text-slate-500">Thực đạt: <strong class="text-slate-900 text-sm">${fmtNumber(cat.actual)}</strong></span>
          <span class="text-slate-500">Target: <strong class="text-slate-700">${fmtNumber(cat.target)}</strong></span>
        </div>
        <div class="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
          <div class="h-full ${barColor} progress-bar-fill rounded-full" style="width: ${barWidth}%"></div>
        </div>
      </div>
    `;
  }).join('');
}

function renderKpiReportTable() {
  const tbody = document.getElementById('kpiReportTableBody');
  tbody.innerHTML = currentData.kpiReport.map(item => {
    const isPass = item.status.toUpperCase().includes('ĐẠT') && !item.status.toUpperCase().includes('KHÔNG');
    const badge = isPass
      ? '<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">ĐẠT</span>'
      : '<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-100 text-rose-800">CHƯA ĐẠT</span>';

    const accDisplay = item.accumulated < 1 && item.accumulated > 0 ? fmtPercent(item.accumulated) : fmtNumber(item.accumulated);
    const targetDisplay = item.target === 0 ? '-' : fmtNumber(item.target);
    const rateDisplay = item.target === 0 ? '-' : fmtPercent(item.rate);

    return `
      <tr class="hover:bg-slate-50">
        <td class="py-2.5 pr-2 font-medium text-slate-800">${item.name}</td>
        <td class="py-2.5 px-2 text-right font-bold text-slate-900">${accDisplay}</td>
        <td class="py-2.5 px-2 text-right text-slate-500">${targetDisplay}</td>
        <td class="py-2.5 px-2 text-right font-semibold ${isPass ? 'text-emerald-700' : 'text-rose-600'}">${rateDisplay}</td>
        <td class="py-2.5 pl-2 text-center">${badge}</td>
      </tr>
    `;
  }).join('');
}

function renderServiceMiniStats() {
  const tat = currentData.tatKpi;
  const kl = currentData.klDl;
  document.getElementById('serviceMiniStats').innerHTML = `
    <div class="bg-slate-50 rounded-lg p-2">
      <div class="text-[10px] text-slate-500 font-medium">TAT-KS Tổng</div>
      <div class="text-sm font-bold text-emerald-700">${fmtPercent(tat.tat_ks)}</div>
    </div>
    <div class="bg-slate-50 rounded-lg p-2">
      <div class="text-[10px] text-slate-500 font-medium">TAT-GCSM</div>
      <div class="text-sm font-bold text-emerald-700">${fmtPercent(tat.tat_gcsm)}</div>
    </div>
    <div class="bg-slate-50 rounded-lg p-2">
      <div class="text-[10px] text-slate-500 font-medium">Tỷ lệ Miếng dán</div>
      <div class="text-sm font-bold text-blue-700">${fmtPercent(kl.ty_le_md)}</div>
    </div>
  `;
}

function destroyChart(id) {
  if (charts[id]) {
    charts[id].destroy();
    delete charts[id];
  }
}

function renderCharts() {
  // 1. Category Target vs Actual Bar Chart
  destroyChart('categoryBarChart');
  const subCats = currentData.categories.filter(c => c.key !== 'TOTAL');
  charts['categoryBarChart'] = new Chart(document.getElementById('categoryBarChart'), {
    type: 'bar',
    data: {
      labels: subCats.map(c => c.name),
      datasets: [
        {
          label: 'Target Tháng',
          data: subCats.map(c => c.target),
          backgroundColor: '#cbd5e1',
          borderRadius: 6,
          barPercentage: 0.7
        },
        {
          label: 'Thực Đạt',
          data: subCats.map(c => c.actual),
          backgroundColor: '#059669',
          borderRadius: 6,
          barPercentage: 0.7
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { usePointStyle: true, font: { size: 12 } } }
      },
      scales: {
        y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false } }
      }
    }
  });

  // 2. Category Doughnut Chart
  destroyChart('categoryDoughnutChart');
  const nonZeroCats = subCats.filter(c => c.actual > 0);
  charts['categoryDoughnutChart'] = new Chart(document.getElementById('categoryDoughnutChart'), {
    type: 'doughnut',
    data: {
      labels: nonZeroCats.map(c => c.name),
      datasets: [{
        data: nonZeroCats.map(c => c.actual),
        backgroundColor: ['#059669', '#0284c7', '#7c3aed', '#d97706', '#e11d48', '#0d9488', '#4f46e5'],
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '64%',
      plugins: {
        legend: { position: 'bottom', labels: { usePointStyle: true, padding: 14, font: { size: 11 } } }
      }
    }
  });

  // 3. Daily Trend Chart (Center SO + GCSM SO by date)
  destroyChart('dailyTrendChart');
  const dailyMap = {};
  currentData.centerRecords.forEach(r => {
    if (!r.date) return;
    if (!dailyMap[r.date]) dailyMap[r.date] = { points: 0, rev: 0 };
    if (r.export_type.toLowerCase().includes('bán') && r.status.toLowerCase().includes('hiệu lực')) {
      dailyMap[r.date].points += (r.points || 0);
      dailyMap[r.date].rev += (r.receivable || r.retail_revenue || 0) / 1000000;
    }
  });
  currentData.gcsmRecords.forEach(r => {
    if (!r.date) return;
    if (!dailyMap[r.date]) dailyMap[r.date] = { points: 0, rev: 0 };
    dailyMap[r.date].points += (r.points || 0);
    dailyMap[r.date].rev += (r.retail_price || 0) / 1000000;
  });

  const sortedDates = Object.keys(dailyMap).sort();
  charts['dailyTrendChart'] = new Chart(document.getElementById('dailyTrendChart'), {
    type: 'line',
    data: {
      labels: sortedDates.map(d => d.slice(5)),
      datasets: [
        {
          label: 'Điểm đạt được',
          data: sortedDates.map(d => dailyMap[d].points),
          borderColor: '#059669',
          backgroundColor: 'rgba(5, 150, 105, 0.12)',
          fill: true,
          tension: 0.35,
          yAxisID: 'y'
        },
        {
          label: 'Doanh thu (Triệu VNĐ)',
          data: sortedDates.map(d => Number(dailyMap[d].rev.toFixed(2))),
          borderColor: '#2563eb',
          backgroundColor: 'transparent',
          borderDash: [5, 5],
          tension: 0.35,
          yAxisID: 'y1'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        y: { type: 'linear', position: 'left', beginAtZero: true, title: { display: true, text: 'Điểm' } },
        y1: { type: 'linear', position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'Triệu VNĐ' } }
      }
    }
  });

  // 4. KL-DL Chart
  destroyChart('klDlChart');
  charts['klDlChart'] = new Chart(document.getElementById('klDlChart'), {
    type: 'pie',
    data: {
      labels: ['Khách lẻ (KL)', 'Đại lý (DL_KA)'],
      datasets: [{
        data: [currentData.klDl.khach_le, currentData.klDl.dl_ka],
        backgroundColor: ['#059669', '#3b82f6'],
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { usePointStyle: true } } }
    }
  });

  // 5. Staff Productivity Chart
  destroyChart('staffProductivityChart');
  charts['staffProductivityChart'] = new Chart(document.getElementById('staffProductivityChart'), {
    type: 'bar',
    data: {
      labels: currentData.productivity.map(p => p.name),
      datasets: [
        { label: 'Sửa chữa (*1)', data: currentData.productivity.map(p => p.sua_chua), backgroundColor: '#0284c7', borderRadius: 4 },
        { label: 'DV Cộng thêm (*0.5)', data: currentData.productivity.map(p => p.dich_vu_cong_them), backgroundColor: '#10b981', borderRadius: 4 },
        { label: 'Phiếu SO Thường', data: currentData.productivity.map(p => p.so_thuong), backgroundColor: '#f59e0b', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
    }
  });

  // 6. Staff Sales Points Chart
  destroyChart('staffSalesChart');
  const staffPoints = {};
  currentData.centerRecords.forEach(r => {
    if (r.export_type.toLowerCase().includes('bán') && r.status.toLowerCase().includes('hiệu lực')) {
      const name = r.kpi_staff || r.creator || 'Khác';
      staffPoints[name] = (staffPoints[name] || 0) + (r.points || 0);
    }
  });
  currentData.gcsmRecords.forEach(r => {
    const name = r.creator || 'Khác';
    staffPoints[name] = (staffPoints[name] || 0) + (r.points || 0);
  });
  const staffNames = Object.keys(staffPoints);
  charts['staffSalesChart'] = new Chart(document.getElementById('staffSalesChart'), {
    type: 'bar',
    data: {
      labels: staffNames,
      datasets: [{
        label: 'Tổng Điểm Bán Hàng (Center SO + GCSM SO)',
        data: staffNames.map(n => staffPoints[n]),
        backgroundColor: ['#059669', '#0d9488', '#2563eb', '#7c3aed'],
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { beginAtZero: true } }
    }
  });
}

function renderTopProducts() {
  const prodMap = {};
  currentData.centerRecords.forEach(r => {
    if (!r.export_type.toLowerCase().includes('bán') || !r.status.toLowerCase().includes('hiệu lực')) return;
    const key = r.model || r.product;
    if (!key) return;
    if (!prodMap[key]) prodMap[key] = { name: key, type: r.type || r.category, qty: 0, rev: 0, points: 0 };
    prodMap[key].qty += r.qty;
    prodMap[key].rev += (r.receivable || r.retail_revenue || 0);
    prodMap[key].points += r.points;
  });
  currentData.gcsmRecords.forEach(r => {
    if (r.points <= 0 && r.retail_price <= 0) return;
    const key = r.product || r.model;
    if (!key) return;
    if (!prodMap[key]) prodMap[key] = { name: key, type: r.type, qty: 0, rev: 0, points: 0 };
    prodMap[key].qty += r.qty;
    prodMap[key].rev += r.retail_price;
    prodMap[key].points += r.points;
  });

  const sorted = Object.values(prodMap).sort((a, b) => b.points - a.points || b.qty - a.qty).slice(0, 10);
  document.getElementById('topProductsTableBody').innerHTML = sorted.map((p, idx) => `
    <tr class="hover:bg-slate-50">
      <td class="py-2.5 pr-2 font-bold text-slate-400">${idx + 1}</td>
      <td class="py-2.5 px-2 font-semibold text-slate-800">${p.name}</td>
      <td class="py-2.5 px-2"><span class="px-2 py-0.5 rounded bg-slate-100 text-slate-600 text-[11px]">${p.type || '-'}</span></td>
      <td class="py-2.5 px-2 text-right font-medium">${fmtNumber(p.qty)}</td>
      <td class="py-2.5 px-2 text-right text-slate-600">${fmtCurrency(p.rev)}</td>
      <td class="py-2.5 pl-2 text-right font-bold text-emerald-700">${fmtNumber(p.points)}</td>
    </tr>
  `).join('');
}

function renderKlDlBox() {
  const kl = currentData.klDl;
  document.getElementById('klDlDetailsBox').innerHTML = `
    <div class="flex justify-between"><span>Khách lẻ (KL):</span><strong>${fmtNumber(kl.khach_le)} lượt (${fmtPercent(kl.pct_khach_le)})</strong></div>
    <div class="flex justify-between"><span>Đại lý (DL_KA):</span><strong>${fmtNumber(kl.dl_ka)} lượt (${fmtPercent(kl.pct_dl_ka)})</strong></div>
    <div class="border-t border-slate-200 my-1.5"></div>
    <div class="flex justify-between"><span>Miếng dán phiếu CR / SO:</span><strong>${fmtNumber(kl.md_cr)} / ${fmtNumber(kl.md_so)}</strong></div>
    <div class="flex justify-between"><span>MDCL phiếu SO:</span><strong>${fmtNumber(kl.mdcl_so)} (Tỷ lệ: ${fmtPercent(kl.ty_le_md)})</strong></div>
  `;
}

function renderStaffTable() {
  document.getElementById('staffTableBody').innerHTML = currentData.productivity.map(p => `
    <tr class="hover:bg-slate-50">
      <td class="py-3 px-3 font-mono text-slate-500">${p.code}</td>
      <td class="py-3 px-3 font-bold text-slate-900">${p.name}</td>
      <td class="py-3 px-3 text-slate-600">${p.region}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.find_gap)}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.may_thang)}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.sua_chua)}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.dich_vu_cong_them)}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.so_find_n)}</td>
      <td class="py-3 px-3 text-right">${fmtNumber(p.so_thuong)}</td>
      <td class="py-3 px-3 text-right font-extrabold text-emerald-700 bg-emerald-50/40">${fmtNumber(p.volume)}</td>
    </tr>
  `).join('');
}

function populateTransactionFilters() {
  const types = new Set();
  currentData.centerRecords.forEach(r => { if (r.type) types.add(r.type.toUpperCase()); });
  currentData.gcsmRecords.forEach(r => { if (r.type) types.add(r.type.toUpperCase()); });

  const select = document.getElementById('txTypeFilter');
  const currentVal = select.value;
  select.innerHTML = '<option value="ALL">Tất cả nhóm hàng</option>' +
    Array.from(types).sort().map(t => `<option value="${t}">${t}</option>`).join('');
  if (Array.from(types).includes(currentVal)) select.value = currentVal;
}

function renderTransactionsTable() {
  const sourceFilter = document.getElementById('txSourceFilter').value;
  const typeFilter = document.getElementById('txTypeFilter').value;
  const search = document.getElementById('txSearchInput').value.toLowerCase().trim();

  let unified = [];
  if (sourceFilter === 'ALL' || sourceFilter === 'CENTER') {
    currentData.centerRecords.forEach(r => {
      unified.push({
        source: 'Center SO',
        time: r.time,
        order_no: r.order_no,
        product: r.product || r.model,
        type: (r.type || r.category || '').toUpperCase(),
        staff: r.kpi_staff || r.creator,
        qty: r.qty,
        revenue: r.receivable || r.retail_revenue || 0,
        points: r.points,
        note: r.export_type
      });
    });
  }
  if (sourceFilter === 'ALL' || sourceFilter === 'GCSM') {
    currentData.gcsmRecords.forEach(r => {
      unified.push({
        source: 'GCSM SO',
        time: r.time,
        order_no: r.order_no,
        product: r.product,
        type: (r.type || '').toUpperCase(),
        staff: r.creator,
        qty: r.qty,
        revenue: r.retail_price || 0,
        points: r.points,
        note: r.status
      });
    });
  }

  const filtered = unified.filter(item => {
    if (typeFilter !== 'ALL' && item.type !== typeFilter) return false;
    if (search) {
      const hay = `${item.order_no} ${item.product} ${item.staff} ${item.type}`.toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  }).sort((a, b) => (b.time || '').localeCompare(a.time || ''));

  const displayList = filtered.slice(0, 150);
  document.getElementById('txTableBody').innerHTML = displayList.map(item => `
    <tr class="hover:bg-slate-50">
      <td class="py-2 px-3">
        <span class="px-2 py-0.5 rounded text-[10px] font-semibold ${item.source === 'Center SO' ? 'bg-blue-50 text-blue-700' : 'bg-purple-50 text-purple-700'}">
          ${item.source}
        </span>
      </td>
      <td class="py-2 px-3 text-slate-500 whitespace-nowrap">${item.time}</td>
      <td class="py-2 px-3 font-mono text-[11px] text-slate-600">${item.order_no}</td>
      <td class="py-2 px-3 font-medium text-slate-800">${item.product}</td>
      <td class="py-2 px-3 text-slate-500">${item.type || '-'}</td>
      <td class="py-2 px-3 text-slate-700">${item.staff}</td>
      <td class="py-2 px-3 text-right">${fmtNumber(item.qty)}</td>
      <td class="py-2 px-3 text-right text-slate-600">${fmtCurrency(item.revenue)}</td>
      <td class="py-2 px-3 text-right font-bold ${item.points > 0 ? 'text-emerald-700' : 'text-slate-400'}">${fmtNumber(item.points)}</td>
    </tr>
  `).join('');

  document.getElementById('txPaginationInfo').textContent = `Hiển thị ${displayList.length} / ${filtered.length} giao dịch phù hợp`;
}
