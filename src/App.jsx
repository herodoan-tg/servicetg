import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  FileSpreadsheet,
  BarChart3,
  Award,
  Smartphone,
  Headphones,
  Tablet,
  ShieldCheck,
  BatteryCharging,
  Watch,
  Layers,
  Download,
  Settings,
  CheckCircle2,
  Search
} from 'lucide-react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  Legend
} from 'recharts';

const DEFAULT_RULES = [
  { model: 'Find N6', points: 9, type: 'PHONE' },
  { model: 'Find X9 Ultra', points: 9, type: 'PHONE' },
  { model: 'Find X9s 12GB+256GB', points: 7, type: 'PHONE' },
  { model: 'Find X9s 12GB+512GB', points: 7, type: 'PHONE' },
  { model: 'Find X9 16GB+512GB', points: 7, type: 'PHONE' },
  { model: 'Find X9 12+256', points: 8, type: 'PHONE' },
  { model: 'Reno16 Pro 5G 12GB+256GB', points: 7, type: 'PHONE' },
  { model: 'Reno16 5G 8GB+256GB', points: 5, type: 'PHONE' },
  { model: 'Reno16 F 5G 8GB+256GB', points: 5, type: 'PHONE' },
  { model: 'Reno16 F 5G 8GB+128GB', points: 6, type: 'PHONE' },
  { model: 'Reno15 F 5G 12GB+256GB', points: 5, type: 'PHONE' },
  { model: 'Reno15 Pro 5G 12GB+256GB', points: 8, type: 'PHONE' },
  { model: 'Reno15 5G 12GB+256GB', points: 8, type: 'PHONE' },
  { model: 'Reno15 5G 8GB+256GB', points: 8, type: 'PHONE' },
  { model: 'A7 Pro 5G 8GB+256GB', points: 6, type: 'PHONE' },
  { model: 'A7 Pro 5G 8GB+128GB', points: 6, type: 'PHONE' },
  { model: 'A7 Pro 5G 6GB+256GB', points: 6, type: 'PHONE' },
  { model: 'A6 Pro 5G 12GB+256GB', points: 3, type: 'PHONE' },
  { model: 'A6 Pro 5G 8GB+256GB', points: 3, type: 'PHONE' },
  { model: 'A6 8GB+128GB', points: 3, type: 'PHONE' },
  { model: 'A6 Pro 8GB+256GB', points: 3, type: 'PHONE' },
  { model: 'A6t Pro 8GB+128GB', points: 3, type: 'PHONE' },
  { model: 'A6 Pro 8GB+128GB', points: 2, type: 'PHONE' },
  { model: 'A6x 6GB+256GB', points: 2, type: 'PHONE' },
  { model: 'A6x 4GB+128GB', points: 2, type: 'PHONE' },
  { model: 'A6x 4GB+64GB', points: 2, type: 'PHONE' },
  { model: 'A6c 4GB+128GB', points: 2, type: 'PHONE' },
  { model: 'A6c 4GB+64GB', points: 2, type: 'PHONE' },
  { model: 'A5i 4GB+64GB', points: 2, type: 'PHONE' },
  { model: 'A6t 6GB+256GB', points: 2, type: 'PHONE' },
  { model: 'A6t 4GB+128GB', points: 2, type: 'PHONE' },
  { model: 'A6t 4GB+64GB', points: 2, type: 'PHONE' },
  { model: 'Enco Buds 2', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Buds 2 Pro', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Buds3 Pro', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Buds3', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Air 3', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Air 2 Pro', points: 2, type: 'TAI NGHE' },
  { model: 'Enco Air 4', points: 3, type: 'TAI NGHE' },
  { model: 'Enco Air 4 Pro', points: 3, type: 'TAI NGHE' },
  { model: 'Enco Air5', points: 3, type: 'TAI NGHE' },
  { model: 'Enco Air5s', points: 3, type: 'TAI NGHE' },
  { model: 'Enco Air 5 Pro', points: 3, type: 'TAI NGHE' },
  { model: 'Enco X3S', points: 5, type: 'TAI NGHE' },
  { model: 'Enco Clip2', points: 5, type: 'TAI NGHE' },
  { model: 'Pad Neo 6GB', points: 2, type: 'PAD' },
  { model: 'Pad Neo 8GB', points: 2, type: 'PAD' },
  { model: 'Pad SE Wifi (4-128)', points: 4, type: 'PAD' },
  { model: 'Pad SE Wifi (4-128) Mate', points: 4, type: 'PAD' },
  { model: 'Pad SE LTE (4-128)', points: 4, type: 'PAD' },
  { model: 'Pad SE LTE (8-256) Mate', points: 4, type: 'PAD' },
  { model: 'Pad 3', points: 7, type: 'PAD' },
  { model: 'Pad 5', points: 8, type: 'PAD' },
  { model: 'OPPO Watch X2', points: 8, type: 'WATCH' },
  { model: 'OPPO Watch X2 Mini', points: 8, type: 'WATCH' },
  { model: 'OPPO Watch S (OWWE262)', points: 8, type: 'WATCH' },
  { model: 'OPPO Watch X3', points: 9, type: 'WATCH' },
  { model: 'OPPO Bubble', points: 8, type: '' },
  { model: 'Miếng dán có phí', points: 1, type: 'Miếng dán có phí' },
  { model: 'Sạc dự phòng', points: 2, type: 'Sạc dự phòng' },
  { model: 'Ốp Lưng', points: 3, type: 'ỐP LƯNG' },
  { model: 'Oppo Care', points: 4, type: 'Oppo Care' },
  { model: 'Ống kính Tele OPPO Hasselblad Earth Explorer Kit', points: 5, type: '' },
  { model: 'Find X10 Pro Max 16GB+512GB', points: 9, type: 'PHONE' },
  { model: 'Find X10 Pro Max in Titanium Grey Hasselblad Teleconverter Kit', points: 5, type: '' }
];

const DEFAULT_TARGETS = {
  total: 176,
  phone: 68,
  taiNghe: 6,
  pad: 4,
  opLung: 9,
  oppoCare: 20,
  mdcl: 64,
  sacDuPhong: 6,
  watch: 0
};

export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [shopName, setShopName] = useState('OPPO Experience & Service Store Tiền Giang');
  const [availableShops, setAvailableShops] = useState(['OPPO Experience & Service Store Tiền Giang']);
  const [targets, setTargets] = useState(DEFAULT_TARGETS);
  const [rules, setRules] = useState(DEFAULT_RULES);
  const [gcsmRows, setGcsmRows] = useState([]);
  const [centerRows, setCenterRows] = useState([]);
  const [fileName, setFileName] = useState('');
  const [searchTerm, setSearchTerm] = useState('');

  const lookupRule = (modelName, currentRules) => {
    if (!modelName) return { points: 0, type: '' };
    const clean = String(modelName).trim().toLowerCase();
    const found = currentRules.find((r) => r.model.trim().toLowerCase() === clean);
    return found ? { points: Number(found.points) || 0, type: (found.type || '').trim() } : { points: 0, type: '' };
  };

  const parseGcsmSheet = (rawRows, currentRules) => {
    return rawRows
      .filter((r) => r['Sales order No.'] || r['Name of service center'])
      .map((r, idx) => {
        const status = String(r['Order status'] || '').trim();
        const rawAmount = String(r['Accounts receivable after individual exemptions and reductions'] || '0').replace(/,/g, '');
        const amount = parseFloat(rawAmount) || 0;
        const salesType = String(r['Sales type'] || '').trim();
        const level3 = String(r['Level 3 material classification'] || '').trim();

        let computedType = '';
        if (
          status === 'Completed' &&
          (amount > 0 || (amount === 0 && salesType === 'Sales package fulfillment'))
        ) {
          if (amount > 0 && salesType === 'Warranty service') {
            computedType = 'Oppo Care';
          } else if (amount === 0 && salesType === 'Sales package fulfillment') {
            computedType = 'Miếng dán có phí';
          } else if (level3 === 'Paid Film' || level3 === 'Protective Film') {
            computedType = 'Miếng dán có phí';
          } else if (level3 === 'Power bank') {
            computedType = 'Sạc dự phòng';
          }
        }

        const ruleMatch = computedType ? lookupRule(computedType, currentRules) : { points: 0 };

        return {
          id: idx + 1,
          orderNo: r['Sales order No.'] || '',
          shop: String(r['Name of service center'] || '').trim(),
          createdBy: r['Created by'] || '',
          salesTime: r['Sales time'] || '',
          status,
          salesType,
          level3,
          productName: r['Sales Product Name'] || '',
          amount,
          computedType,
          points: ruleMatch.points
        };
      });
  };

  const parseCenterSheet = (rawMatrix, currentRules) => {
    let headerIdx = rawMatrix.findIndex(
      (row) => Array.isArray(row) && row.some((cell) => String(cell || '').trim() === 'Tên Brandshop')
    );
    if (headerIdx === -1) headerIdx = 2;

    const headers = rawMatrix[headerIdx].map((h) => String(h || '').trim());
    const dataRows = rawMatrix.slice(headerIdx + 1);

    const getCol = (row, name) => {
      const idx = headers.indexOf(name);
      return idx !== -1 ? row[idx] : '';
    };

    return dataRows
      .filter((row) => row && row.length > 5 && getCol(row, 'Tên Brandshop'))
      .map((row, idx) => {
        const shop = String(getCol(row, 'Tên Brandshop') || '').trim();
        const productName = String(getCol(row, 'Tên sản phẩm') || '').trim();
        const exportType = String(getCol(row, 'Loại xuất') || '').trim();
        const status = String(getCol(row, 'Trạng thái') || '').trim();
        const salesDate = String(getCol(row, 'Ngày xuất bán') || '').trim();
        const staff = String(getCol(row, 'Người lên đơn') || '').trim();
        const customer = String(getCol(row, 'Tên khách hàng') || '').trim();
        const revenue = Number(getCol(row, 'Doanh thu bán lẻ')) || 0;

        let normalizedModel = '';
        const lowerProd = productName.toLowerCase();
        if (productName !== '' && !lowerProd.includes('túi giấy')) {
          if (lowerProd.includes('ốp lưng')) {
            normalizedModel = 'Ốp Lưng';
          } else if (lowerProd.includes('sạc dự phòng')) {
            normalizedModel = 'Sạc dự phòng';
          } else if (lowerProd.includes('cường lực')) {
            normalizedModel = 'Miếng dán có phí';
          } else {
            normalizedModel = productName;
          }
        }

        const ruleMatch = normalizedModel ? lookupRule(normalizedModel, currentRules) : { points: 0, type: '' };

        return {
          id: idx + 1,
          orderNo: getCol(row, 'Đơn hàng tổng') || '',
          shop,
          salesDate,
          productName,
          normalizedModel,
          exportType,
          status,
          staff,
          customer,
          revenue,
          computedType: ruleMatch.type,
          points: ruleMatch.points
        };
      });
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });

      let activeRules = rules;
      if (wb.SheetNames.includes('Rule')) {
        const ruleJson = XLSX.utils.sheet_to_json(wb.Sheets['Rule']);
        const parsedRules = ruleJson
          .filter((r) => r['Model'])
          .map((r) => ({
            model: String(r['Model']).trim(),
            points: Number(r['Điểm tổng kết']) || 0,
            type: String(r['Type'] || '').trim()
          }));
        if (parsedRules.length > 0) {
          setRules(parsedRules);
          activeRules = parsedRules;
        }
      }

      if (wb.SheetNames.includes('Tổng hợp')) {
        const summaryMatrix = XLSX.utils.sheet_to_json(wb.Sheets['Tổng hợp'], { header: 1 });
        if (summaryMatrix[2]) {
          const r3 = summaryMatrix[2];
          if (r3[0]) setShopName(String(r3[0]).trim());
          setTargets({
            total: Number(r3[1]) || 176,
            phone: Number(r3[4]) || 68,
            taiNghe: Number(r3[7]) || 6,
            pad: Number(r3[10]) || 4,
            opLung: Number(r3[13]) || 9,
            oppoCare: Number(r3[16]) || 20,
            mdcl: Number(r3[19]) || 64,
            sacDuPhong: Number(r3[22]) || 6,
            watch: Number(r3[26]) || 0
          });
        }
      }

      const shopsSet = new Set();

      if (wb.SheetNames.includes('GCSM SO')) {
        const gcsmJson = XLSX.utils.sheet_to_json(wb.Sheets['GCSM SO']);
        const parsedGcsm = parseGcsmSheet(gcsmJson, activeRules);
        setGcsmRows(parsedGcsm);
        parsedGcsm.forEach((r) => r.shop && shopsSet.add(r.shop));
      }

      if (wb.SheetNames.includes('Center SO')) {
        const centerMatrix = XLSX.utils.sheet_to_json(wb.Sheets['Center SO'], { header: 1 });
        const parsedCenter = parseCenterSheet(centerMatrix, activeRules);
        setCenterRows(parsedCenter);
        parsedCenter.forEach((r) => r.shop && shopsSet.add(r.shop));
      }

      if (shopsSet.size > 0) {
        const shopList = Array.from(shopsSet);
        setAvailableShops(shopList);
        if (!shopsSet.has(shopName)) {
          setShopName(shopList[0]);
        }
      }
    };
    reader.readAsBinaryString(file);
  };

  const kpiSummary = useMemo(() => {
    const validCenter = centerRows.filter(
      (r) =>
        r.shop === shopName &&
        r.exportType.toLowerCase() === 'xuất bán' &&
        r.status.toLowerCase() === 'hiệu lực'
    );

    const validGcsm = gcsmRows.filter((r) => r.shop === shopName);

    const sumCenterByType = (typeStr) =>
      validCenter
        .filter((r) => r.computedType.toLowerCase() === typeStr.toLowerCase())
        .reduce((acc, cur) => acc + cur.points, 0);

    const sumCenterByModel = (modelStr) =>
      validCenter
        .filter((r) => r.normalizedModel.toLowerCase() === modelStr.toLowerCase())
        .reduce((acc, cur) => acc + cur.points, 0);

    const sumGcsmByType = (typeStr) =>
      validGcsm
        .filter((r) => r.computedType.toLowerCase() === typeStr.toLowerCase())
        .reduce((acc, cur) => acc + cur.points, 0);

    const phonePts = sumCenterByType('PHONE');
    const taiNghePts = sumCenterByType('TAI NGHE');
    const padPts = sumCenterByType('PAD');
    const opLungPts = sumCenterByType('ỐP LƯNG');
    const watchPts = sumCenterByType('WATCH');
    const oppoCarePts = sumGcsmByType('Oppo Care');
    const mdclPts = sumGcsmByType('Miếng dán có phí') + sumCenterByModel('Miếng dán có phí');
    const sacDuPhongPts = sumGcsmByType('Sạc dự phòng') + sumCenterByModel('Sạc dự phòng');

    const a7ProCount = validCenter.filter((r) =>
      r.normalizedModel.toLowerCase().includes('a7 pro')
    ).length;

    const totalPts =
      phonePts +
      taiNghePts +
      padPts +
      watchPts +
      opLungPts +
      oppoCarePts +
      mdclPts +
      sacDuPhongPts;

    const calcPct = (pts, target) => (target > 0 ? (pts / target) * 100 : 0);

    return {
      total: { points: totalPts, target: targets.total, pct: calcPct(totalPts, targets.total) },
      phone: { points: phonePts, target: targets.phone, pct: calcPct(phonePts, targets.phone) },
      taiNghe: { points: taiNghePts, target: targets.taiNghe, pct: calcPct(taiNghePts, targets.taiNghe) },
      pad: { points: padPts, target: targets.pad, pct: calcPct(padPts, targets.pad) },
      opLung: { points: opLungPts, target: targets.opLung, pct: calcPct(opLungPts, targets.opLung) },
      oppoCare: { points: oppoCarePts, target: targets.oppoCare, pct: calcPct(oppoCarePts, targets.oppoCare) },
      mdcl: { points: mdclPts, target: targets.mdcl, pct: calcPct(mdclPts, targets.mdcl) },
      sacDuPhong: {
        points: sacDuPhongPts,
        target: targets.sacDuPhong,
        pct: calcPct(sacDuPhongPts, targets.sacDuPhong)
      },
      watch: { points: watchPts, target: targets.watch, pct: calcPct(watchPts, targets.watch) },
      a7ProCount
    };
  }, [centerRows, gcsmRows, shopName, targets]);

  const chartData = [
    { name: 'PHONE', 'Chỉ tiêu': kpiSummary.phone.target, 'Đạt được': kpiSummary.phone.points },
    { name: 'TAI NGHE', 'Chỉ tiêu': kpiSummary.taiNghe.target, 'Đạt được': kpiSummary.taiNghe.points },
    { name: 'PAD', 'Chỉ tiêu': kpiSummary.pad.target, 'Đạt được': kpiSummary.pad.points },
    { name: 'ỐP LƯNG', 'Chỉ tiêu': kpiSummary.opLung.target, 'Đạt được': kpiSummary.opLung.points },
    { name: 'Oppo Care', 'Chỉ tiêu': kpiSummary.oppoCare.target, 'Đạt được': kpiSummary.oppoCare.points },
    { name: 'MDCL', 'Chỉ tiêu': kpiSummary.mdcl.target, 'Đạt được': kpiSummary.mdcl.points },
    { name: 'Sạc DP', 'Chỉ tiêu': kpiSummary.sacDuPhong.target, 'Đạt được': kpiSummary.sacDuPhong.points },
    { name: 'WATCH', 'Chỉ tiêu': kpiSummary.watch.target, 'Đạt được': kpiSummary.watch.points }
  ];

  const exportSummaryExcel = () => {
    const wb = XLSX.utils.book_new();
    const summaryData = [
      ['BÁO CÁO TỔNG HỢP KẾT QUẢ SALE & KPI'],
      ['Shop name', shopName],
      [],
      ['Hạng mục', 'Chỉ tiêu (Target)', 'Điểm đạt được', 'Tỷ lệ hoàn thành (%)'],
      ['TỔNG ĐIỂM', kpiSummary.total.target, kpiSummary.total.points, `${kpiSummary.total.pct.toFixed(1)}%`],
      ['PHONE', kpiSummary.phone.target, kpiSummary.phone.points, `${kpiSummary.phone.pct.toFixed(1)}%`],
      ['TAI NGHE', kpiSummary.taiNghe.target, kpiSummary.taiNghe.points, `${kpiSummary.taiNghe.pct.toFixed(1)}%`],
      ['PAD', kpiSummary.pad.target, kpiSummary.pad.points, `${kpiSummary.pad.pct.toFixed(1)}%`],
      ['ỐP LƯNG', kpiSummary.opLung.target, kpiSummary.opLung.points, `${kpiSummary.opLung.pct.toFixed(1)}%`],
      ['Oppo Care', kpiSummary.oppoCare.target, kpiSummary.oppoCare.points, `${kpiSummary.oppoCare.pct.toFixed(1)}%`],
      ['MDCL (Miếng dán có phí)', kpiSummary.mdcl.target, kpiSummary.mdcl.points, `${kpiSummary.mdcl.pct.toFixed(1)}%`],
      ['Sạc dự phòng', kpiSummary.sacDuPhong.target, kpiSummary.sacDuPhong.points, `${kpiSummary.sacDuPhong.pct.toFixed(1)}%`],
      ['WATCH', kpiSummary.watch.target, kpiSummary.watch.points, `${kpiSummary.watch.pct.toFixed(1)}%`],
      ['Số lượng A7 Pro bán ra', '-', kpiSummary.a7ProCount, '-']
    ];
    const ws = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws, 'Tổng hợp KPI');
    XLSX.writeFile(wb, `Bao_Cao_KPI_${shopName.replace(/\s+/g, '_')}.xlsx`);
  };

  const categoriesList = [
    { key: 'phone', label: 'PHONE', icon: Smartphone, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { key: 'taiNghe', label: 'TAI NGHE', icon: Headphones, color: 'text-blue-600', bg: 'bg-blue-50' },
    { key: 'pad', label: 'PAD', icon: Tablet, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { key: 'opLung', label: 'ỐP LƯNG', icon: Layers, color: 'text-amber-600', bg: 'bg-amber-50' },
    { key: 'oppoCare', label: 'Oppo Care', icon: ShieldCheck, color: 'text-teal-600', bg: 'bg-teal-50' },
    { key: 'mdcl', label: 'MDCL (Dán có phí)', icon: Smartphone, color: 'text-purple-600', bg: 'bg-purple-50' },
    { key: 'sacDuPhong', label: 'Sạc dự phòng', icon: BatteryCharging, color: 'text-rose-600', bg: 'bg-rose-50' },
    { key: 'watch', label: 'WATCH', icon: Watch, color: 'text-cyan-600', bg: 'bg-cyan-50' }
  ];

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      <header className="bg-emerald-700 text-white shadow-md sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-white/15 p-2.5 rounded-xl">
              <BarChart3 className="w-7 h-7 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">BÁO CÁO KẾT QUẢ SALE & KPI</h1>
              <p className="text-emerald-100 text-xs">{shopName}</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            {availableShops.length > 1 && (
              <select
                value={shopName}
                onChange={(e) => setShopName(e.target.value)}
                className="bg-emerald-800 text-white text-sm rounded-lg px-3 py-2 border border-emerald-600 focus:outline-none"
              >
                {availableShops.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            )}

            <label className="cursor-pointer bg-white text-emerald-800 hover:bg-emerald-50 font-semibold text-sm px-4 py-2 rounded-lg shadow-sm flex items-center gap-2 transition">
              <Upload className="w-4 h-4" />
              <span>{fileName ? 'Đổi file Excel' : 'Tải lên file Excel'}</span>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>

            <button
              onClick={exportSummaryExcel}
              className="bg-emerald-600 hover:bg-emerald-500 text-white border border-emerald-400/40 font-medium text-sm px-3.5 py-2 rounded-lg flex items-center gap-2 transition"
            >
              <Download className="w-4 h-4" />
              <span>Xuất Excel</span>
            </button>
          </div>
        </div>

        <div className="bg-emerald-800/60 border-t border-emerald-600/50">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex space-x-6 overflow-x-auto text-sm font-medium">
            {[
              { id: 'dashboard', label: 'Tổng hợp KPI' },
              { id: 'center', label: `Center SO (${centerRows.length})` },
              { id: 'gcsm', label: `GCSM SO (${gcsmRows.length})` },
              { id: 'settings', label: 'Cấu hình Target & Rule' }
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`py-3 border-b-2 whitespace-nowrap transition ${
                  activeTab === tab.id
                    ? 'border-white text-white font-semibold'
                    : 'border-transparent text-emerald-200 hover:text-white'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {!fileName && (
          <div className="mb-6 bg-amber-50 border border-amber-200 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <FileSpreadsheet className="w-6 h-6 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <h3 className="font-semibold text-amber-900 text-sm">
                  Chưa có dữ liệu đơn hàng được nạp
                </h3>
                <p className="text-amber-700 text-xs mt-0.5">
                  Hãy bấm <strong>"Tải lên file Excel"</strong> (file <code>Tổng Hợp Sale V1.4.xlsx</code>) để hệ thống tự động tính điểm từ sheet <strong>Center SO</strong>, <strong>GCSM SO</strong> và bảng <strong>Rule</strong>.
                </p>
              </div>
            </div>
          </div>
        )}

        {fileName && (
          <div className="mb-5 flex items-center gap-2 text-xs text-slate-500">
            <CheckCircle2 className="w-4 h-4 text-emerald-600" />
            <span>
              Đang hiển thị dữ liệu từ file: <strong className="text-slate-700">{fileName}</strong>
            </span>
          </div>
        )}

        {activeTab === 'dashboard' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="lg:col-span-2 bg-gradient-to-br from-emerald-600 to-teal-700 rounded-2xl p-6 text-white shadow-sm flex flex-col justify-between">
                <div className="flex items-start justify-between">
                  <div>
                    <span className="text-emerald-100 text-xs font-semibold uppercase tracking-wider">
                      Tổng Điểm Hoàn Thành KPI
                    </span>
                    <div className="mt-2 flex items-baseline gap-3">
                      <span className="text-5xl font-extrabold">{kpiSummary.total.points}</span>
                      <span className="text-emerald-200 text-lg font-medium">
                        / {kpiSummary.total.target} điểm
                      </span>
                    </div>
                  </div>
                  <div className="bg-white/20 backdrop-blur-sm px-4 py-2 rounded-xl text-right">
                    <span className="text-xs text-emerald-100 block">Tỷ lệ đạt</span>
                    <span className="text-2xl font-bold">{kpiSummary.total.pct.toFixed(1)}%</span>
                  </div>
                </div>

                <div className="mt-6">
                  <div className="w-full bg-black/20 rounded-full h-3 overflow-hidden">
                    <div
                      className="bg-white h-full rounded-full transition-all duration-500"
                      style={{ width: `${Math.min(kpiSummary.total.pct, 100)}%` }}
                    />
                  </div>
                  <div className="mt-2 flex justify-between text-xs text-emerald-100">
                    <span>Còn thiếu: {Math.max(0, kpiSummary.total.target - kpiSummary.total.points)} điểm</span>
                    <span>A7 Pro bán ra: {kpiSummary.a7ProCount} máy</span>
                  </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
                <div>
                  <h3 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                    <Award className="w-4 h-4 text-emerald-600" />
                    Trạng thái các nhóm KPI
                  </h3>
                  <div className="mt-4 space-y-3">
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-slate-600">Nhóm đạt ≥ 100% chỉ tiêu</span>
                      <span className="font-bold text-emerald-600">
                        {categoriesList.filter((c) => kpiSummary[c.key].pct >= 100).length} / 8 nhóm
                      </span>
                    </div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-slate-600">Đơn Center SO hợp lệ</span>
                      <span className="font-bold text-slate-800">
                        {
                          centerRows.filter(
                            (r) =>
                              r.shop === shopName &&
                              r.exportType.toLowerCase() === 'xuất bán' &&
                              r.status.toLowerCase() === 'hiệu lực'
                          ).length
                        }{' '}
                        dòng
                      </span>
                    </div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="text-slate-600">Đơn GCSM SO tính điểm</span>
                      <span className="font-bold text-slate-800">
                        {gcsmRows.filter((r) => r.shop === shopName && r.points > 0).length} dòng
                      </span>
                    </div>
                  </div>
                </div>

                <button
                  onClick={() => setActiveTab('settings')}
                  className="mt-4 w-full py-2 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition"
                >
                  <Settings className="w-3.5 h-3.5" />
                  Chỉnh sửa Target / Bảng điểm Rule
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {categoriesList.map((item) => {
                const stat = kpiSummary[item.key];
                const Icon = item.icon;
                const isDone = stat.target > 0 && stat.points >= stat.target;

                return (
                  <div
                    key={item.key}
                    className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm hover:shadow-md transition"
                  >
                    <div className="flex items-center justify-between">
                      <div className={`p-2.5 rounded-lg ${item.bg}`}>
                        <Icon className={`w-5 h-5 ${item.color}`} />
                      </div>
                      <span
                        className={`text-xs font-bold px-2.5 py-1 rounded-full ${
                          isDone
                            ? 'bg-emerald-100 text-emerald-700'
                            : stat.pct >= 70
                            ? 'bg-amber-100 text-amber-700'
                            : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {stat.target > 0 ? `${stat.pct.toFixed(1)}%` : '0%'}
                      </span>
                    </div>

                    <div className="mt-4">
                      <p className="text-xs font-semibold text-slate-500 uppercase">{item.label}</p>
                      <div className="mt-1 flex items-baseline gap-2">
                        <span className="text-2xl font-bold text-slate-800">{stat.points}</span>
                        <span className="text-xs text-slate-400">/ {stat.target} điểm</span>
                      </div>
                    </div>

                    <div className="mt-3 w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                      <div
                        className={`h-full rounded-full transition-all duration-500 ${
                          isDone ? 'bg-emerald-500' : 'bg-amber-500'
                        }`}
                        style={{ width: `${Math.min(stat.pct, 100)}%` }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4">
                Biểu đồ so sánh Điểm Đạt Được và Chỉ Tiêu theo nhóm
              </h3>
              <div className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                    <XAxis dataKey="name" tick={{ fontSize: 12, fill: '#475569' }} />
                    <YAxis tick={{ fontSize: 12, fill: '#475569' }} />
                    <Tooltip
                      contentStyle={{
                        borderRadius: '10px',
                        border: '1px solid #e2e8f0',
                        boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'
                      }}
                    />
                    <Legend />
                    <Bar dataKey="Chỉ tiêu" fill="#cbd5e1" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Đạt được" radius={[4, 4, 0, 0]}>
                      {chartData.map((entry, index) => (
                        <Cell
                          key={`cell-${index}`}
                          fill={entry['Đạt được'] >= entry['Chỉ tiêu'] ? '#059669' : '#f59e0b'}
                        />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'center' && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h3 className="font-bold text-slate-800">Chi tiết đơn hàng Center SO</h3>
                <p className="text-xs text-slate-500">
                  Chỉ tính điểm các dòng có Loại xuất = "Xuất bán" và Trạng thái = "Hiệu lực"
                </p>
              </div>
              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                <input
                  type="text"
                  placeholder="Tìm sản phẩm, khách hàng..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-9 pr-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 text-slate-600 border-b border-slate-200">
                    <th className="p-3 font-semibold">Ngày bán</th>
                    <th className="p-3 font-semibold">Tên sản phẩm</th>
                    <th className="p-3 font-semibold">Model quy đổi</th>
                    <th className="p-3 font-semibold">Loại xuất</th>
                    <th className="p-3 font-semibold">Trạng thái</th>
                    <th className="p-3 font-semibold">Nhóm (Type)</th>
                    <th className="p-3 font-semibold text-right">Điểm</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {centerRows
                    .filter(
                      (r) =>
                        r.shop === shopName &&
                        (r.productName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                          r.customer.toLowerCase().includes(searchTerm.toLowerCase()))
                    )
                    .map((r) => {
                      const isCounted =
                        r.exportType.toLowerCase() === 'xuất bán' &&
                        r.status.toLowerCase() === 'hiệu lực' &&
                        r.points > 0;

                      return (
                        <tr key={r.id} className={isCounted ? 'bg-emerald-50/30' : 'hover:bg-slate-50'}>
                          <td className="p-3 whitespace-nowrap text-slate-500">{r.salesDate}</td>
                          <td className="p-3 font-medium text-slate-800">{r.productName || '-'}</td>
                          <td className="p-3 text-slate-600">{r.normalizedModel || '-'}</td>
                          <td className="p-3">
                            <span
                              className={`px-2 py-0.5 rounded text-[11px] font-medium ${
                                r.exportType === 'Xuất bán'
                                  ? 'bg-blue-50 text-blue-700'
                                  : 'bg-slate-100 text-slate-600'
                              }`}
                            >
                              {r.exportType || '-'}
                            </span>
                          </td>
                          <td className="p-3">{r.status || '-'}</td>
                          <td className="p-3 font-semibold text-slate-700">{r.computedType || '-'}</td>
                          <td className="p-3 text-right font-bold">
                            {isCounted ? (
                              <span className="text-emerald-600">+{r.points}</span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'gcsm' && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h3 className="font-bold text-slate-800">Chi tiết đơn hàng dịch vụ GCSM SO</h3>
                <p className="text-xs text-slate-500">
                  Tự động phân loại Oppo Care, Miếng dán có phí, Sạc dự phòng theo điều kiện đơn hàng
                </p>
              </div>
              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                <input
                  type="text"
                  placeholder="Tìm mã đơn, sản phẩm..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-9 pr-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 text-slate-600 border-b border-slate-200">
                    <th className="p-3 font-semibold">Mã đơn (Order No.)</th>
                    <th className="p-3 font-semibold">Thời gian</th>
                    <th className="p-3 font-semibold">Sales Type</th>
                    <th className="p-3 font-semibold">Sản phẩm / Phân loại</th>
                    <th className="p-3 font-semibold text-right">Thực thu (VND)</th>
                    <th className="p-3 font-semibold">Nhóm KPI</th>
                    <th className="p-3 font-semibold text-right">Điểm</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {gcsmRows
                    .filter(
                      (r) =>
                        r.shop === shopName &&
                        (r.orderNo.toLowerCase().includes(searchTerm.toLowerCase()) ||
                          r.productName.toLowerCase().includes(searchTerm.toLowerCase()))
                    )
                    .map((r) => (
                      <tr key={r.id} className={r.points > 0 ? 'bg-emerald-50/30' : 'hover:bg-slate-50'}>
                        <td className="p-3 font-mono text-[11px] text-slate-600">{r.orderNo}</td>
                        <td className="p-3 whitespace-nowrap text-slate-500">{r.salesTime}</td>
                        <td className="p-3">{r.salesType}</td>
                        <td className="p-3 font-medium text-slate-800">{r.productName}</td>
                        <td className="p-3 text-right font-medium">
                          {r.amount.toLocaleString('vi-VN')}
                        </td>
                        <td className="p-3">
                          {r.computedType ? (
                            <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-semibold">
                              {r.computedType}
                            </span>
                          ) : (
                            '-'
                          )}
                        </td>
                        <td className="p-3 text-right font-bold">
                          {r.points > 0 ? (
                            <span className="text-emerald-600">+{r.points}</span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'settings' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
              <h3 className="font-bold text-slate-800 mb-4">Điều chỉnh Chỉ Tiêu (Target)</h3>
              <div className="space-y-3 text-sm">
                {[
                  { key: 'total', label: 'TỔNG ĐIỂM MỤC TIÊU' },
                  { key: 'phone', label: 'PHONE' },
                  { key: 'taiNghe', label: 'TAI NGHE' },
                  { key: 'pad', label: 'PAD' },
                  { key: 'opLung', label: 'ỐP LƯNG' },
                  { key: 'oppoCare', label: 'Oppo Care' },
                  { key: 'mdcl', label: 'MDCL (Miếng dán có phí)' },
                  { key: 'sacDuPhong', label: 'Sạc dự phòng' },
                  { key: 'watch', label: 'WATCH' }
                ].map((field) => (
                  <div key={field.key} className="flex items-center justify-between">
                    <label className="text-slate-600 font-medium">{field.label}</label>
                    <input
                      type="number"
                      value={targets[field.key]}
                      onChange={(e) =>
                        setTargets({ ...targets, [field.key]: Number(e.target.value) || 0 })
                      }
                      className="w-24 px-2.5 py-1 text-right border border-slate-300 rounded-lg font-bold text-slate-800"
                    />
                  </div>
                ))}
              </div>
            </div>

            <div className="lg:col-span-2 bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
              <h3 className="font-bold text-slate-800 mb-1">Bảng quy đổi điểm (Sheet Rule)</h3>
              <p className="text-xs text-slate-500 mb-4">
                Tự động cập nhật khi bạn tải file Excel có chứa sheet Rule
              </p>
              <div className="max-h-96 overflow-y-auto border border-slate-200 rounded-xl">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-100 sticky top-0">
                    <tr>
                      <th className="p-2.5 font-semibold">Tên Model / Hạng mục</th>
                      <th className="p-2.5 font-semibold">Nhóm (Type)</th>
                      <th className="p-2.5 font-semibold text-right">Điểm tổng kết</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {rules.map((r, i) => (
                      <tr key={i} className="hover:bg-slate-50">
                        <td className="p-2.5 font-medium text-slate-800">{r.model}</td>
                        <td className="p-2.5 text-slate-500">{r.type || '-'}</td>
                        <td className="p-2.5 text-right font-bold text-emerald-700">{r.points}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
