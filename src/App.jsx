import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip, Legend } from 'recharts';
import { 
  Plus, Trash2, TrendingUp, DollarSign, 
  ArrowUpRight, ArrowDownRight, Wallet, Landmark, 
  Settings, RefreshCw, CloudLightning, Loader2, CheckCircle2,
  Coins, X, Check, Edit3, Calculator, Calendar, FileText, Database, ShieldCheck, Search, Zap, Save,
  FileDown, FileUp
} from 'lucide-react';

// --- 資料庫區 (2025 最新財報數據) ---
const OFFICIAL_EPS_DB = {
  '2317': { cumEPS: 10.38, month: 9, note: 'Q3 財報 (賺破股本)' },
  '2892': { cumEPS: 1.51, month: 9, note: 'Q3 財報' },
  '2880': { cumEPS: 1.43, month: 9, note: 'Q3 財報' },
  '5880': { cumEPS: 1.14, month: 10, note: '前10月自結' },
  '1326': { cumEPS: -0.94, month: 9, note: 'Q3 財報' },
  '2330': { cumEPS: 32.50, month: 9, note: 'Q3 財報' },
  '2881': { cumEPS: 9.60, month: 10, note: '自結數' },
  '2882': { cumEPS: 7.20, month: 10, note: '自結數' },
  '2891': { cumEPS: 2.80, month: 10, note: '自結數' },
};

const TW_STOCKS = [
  { code: '2330', name: '台積電' }, { code: '2317', name: '鴻海' }, { code: '2454', name: '聯發科' },
  { code: '2308', name: '台達電' }, { code: '2303', name: '聯電' }, { code: '3711', name: '日月光投控' },
  { code: '2881', name: '富邦金' }, { code: '2882', name: '國泰金' }, { code: '2891', name: '中信金' },
  { code: '2886', name: '兆豐金' }, { code: '2884', name: '玉山金' }, { code: '2892', name: '第一金' },
  { code: '2880', name: '華南金' }, { code: '5880', name: '合庫金' }, { code: '2885', name: '元大金' },
  { code: '2890', name: '永豐金' }, { code: '2883', name: '凱基金' }, { code: '2887', name: '台新金' },
  { code: '1101', name: '台泥' }, { code: '1102', name: '亞泥' }, { code: '1301', name: '台塑' },
  { code: '1303', name: '南亞' }, { code: '1326', name: '台化' }, { code: '6505', name: '台塑化' },
  { code: '2002', name: '中鋼' }, { code: '2603', name: '長榮' }, { code: '2609', name: '陽明' },
  { code: '2615', name: '萬海' }, { code: '9904', name: '寶成' }, { code: '2912', name: '統一超' },
  { code: '0050', name: '元大台灣50' }, { code: '0056', name: '元大高股息' }, { code: '00878', name: '國泰永續高股息' },
  { code: '00929', name: '復華台灣科技優息' }, { code: '00919', name: '群益台灣精選高息' }, { code: '006208', name: '富邦台50' }
];

const DEFAULT_DATA = [
  { id: 1, code: '2892', name: '第一金', type: '金融', price: 27.80, quantity: 15000, cost: 24.50, cumEPS: 1.51, epsMonth: 9, cashPayoutRate: 60, stockPayoutRate: 20, isVerified: true },
  { id: 2, code: '2880', name: '華南金', type: '金融', price: 30.75, quantity: 12000, cost: 20.00, cumEPS: 1.43, epsMonth: 9, cashPayoutRate: 50, stockPayoutRate: 20, isVerified: true },
];

const COLORS = ['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899'];

// --- UI 元件 ---
const Card = ({ children, className = "" }) => (
  <div className={`bg-white rounded-xl shadow-sm border border-slate-100 p-5 ${className}`}>
    {children}
  </div>
);

const StatCard = ({ title, value, subValue, isPositive, icon: Icon, colorClass, alertLevel }) => {
  let valueColor = "text-slate-800";
  if (alertLevel === 'danger') valueColor = "text-rose-600";
  if (alertLevel === 'warning') valueColor = "text-amber-500";
  if (alertLevel === 'safe') valueColor = "text-emerald-600";

  return (
    <Card>
      <div className="flex items-start justify-between">
        <div>
          <p className="text-slate-500 text-xs font-medium mb-1">{title}</p>
          <h3 className={`text-2xl font-bold tracking-tight ${valueColor}`}>{value}</h3>
          {subValue && (
            <div className={`flex items-center mt-2 text-xs font-medium ${isPositive === true ? 'text-emerald-600' : isPositive === false ? 'text-rose-600' : 'text-slate-500'}`}>
              {isPositive === true && <ArrowUpRight size={14} className="mr-0.5" />}
              {isPositive === false && <ArrowDownRight size={14} className="mr-0.5" />}
              <span>{subValue}</span>
            </div>
          )}
        </div>
        <div className={`p-2.5 rounded-lg ${colorClass}`}>
          <Icon size={20} className="text-white" />
        </div>
      </div>
    </Card>
  );
};

export default function App() {
  // --- 0. 自動載入 Excel 函式庫 (CDN) ---
  const [xlsxReady, setXlsxReady] = useState(false);

  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      script.async = true;
      script.onload = () => setXlsxReady(true);
      document.body.appendChild(script);
    } else {
      setXlsxReady(true);
    }
  }, []);

  // --- 1. 資料持久層 (LocalStorage) ---
  const [assets, setAssets] = useState(() => {
    try {
      const saved = localStorage.getItem('stock_manager_excel_assets_v2');
      return saved ? JSON.parse(saved) : DEFAULT_DATA;
    } catch (e) { return DEFAULT_DATA; }
  });
  
  const [loanAmount, setLoanAmount] = useState(() => {
    try {
      const saved = localStorage.getItem('stock_manager_excel_loan_v2');
      return saved ? Number(saved) : 0;
    } catch (e) { return 0; }
  });

  const [lastUpdated, setLastUpdated] = useState(() => {
    const saved = localStorage.getItem('stock_manager_excel_updated_v2');
    return saved ? new Date(saved) : null;
  });

  useEffect(() => {
    localStorage.setItem('stock_manager_excel_assets_v2', JSON.stringify(assets));
  }, [assets]);

  useEffect(() => {
    localStorage.setItem('stock_manager_excel_loan_v2', loanAmount.toString());
  }, [loanAmount]);

  useEffect(() => {
    if (lastUpdated) localStorage.setItem('stock_manager_excel_updated_v2', lastUpdated.toISOString());
  }, [lastUpdated]);

  // --- UI State ---
  const [showAddForm, setShowAddForm] = useState(false);
  const [showPledgeSettings, setShowPledgeSettings] = useState(false);
  const [newAsset, setNewAsset] = useState({ 
    code: '', name: '', type: '電子', price: '', quantity: '', cost: '',
    cumEPS: '', epsMonth: 9, cashPayoutRate: '', stockPayoutRate: '' 
  });
  const [editingCell, setEditingCell] = useState({ id: null, field: null }); 
  const [editValue, setEditValue] = useState('');
  
  const [isUpdating, setIsUpdating] = useState(false);
  const [updateStatus, setUpdateStatus] = useState('');
  const [saveStatus, setSaveStatus] = useState('');
  const fileInputRef = useRef(null);

  const currentYear = new Date().getFullYear();

  // 搜尋相關
  const [searchTerm, setSearchTerm] = useState('');
  const [searchResults, setSearchResults] = useState([]);
  const [showResults, setShowResults] = useState(false);
  const searchRef = useRef(null);

  // 顯示儲存提示
  useEffect(() => {
    if (assets.length > 0) {
        setSaveStatus('saved');
        const timer = setTimeout(() => setSaveStatus(''), 2000);
        return () => clearTimeout(timer);
    }
  }, [assets, loanAmount]);

  // 搜尋邏輯
  useEffect(() => {
    if (searchTerm.length > 0) {
      const results = TW_STOCKS.filter(stock => 
        stock.code.includes(searchTerm) || stock.name.includes(searchTerm)
      );
      setSearchResults(results);
      setShowResults(true);
    } else {
      setSearchResults([]);
      setShowResults(false);
    }
  }, [searchTerm]);

  const selectStock = (stock) => {
    const dbData = OFFICIAL_EPS_DB[stock.code];
    setNewAsset(prev => ({
      ...prev,
      code: stock.code,
      name: stock.name,
      cumEPS: dbData ? dbData.cumEPS : prev.cumEPS,
      epsMonth: dbData ? dbData.month : prev.epsMonth,
    }));
    setSearchTerm(`${stock.code} ${stock.name}`);
    setShowResults(false);
    fetchRealTimePrice(stock.code).then(price => {
        if (price) setNewAsset(prev => ({ ...prev, price: price }));
    });
  };

  // 抓取即時股價
  const fetchRealTimePrice = async (code) => {
    try {
      const symbol = `${code}.TW`;
      const targetUrl = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d`;
      const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(targetUrl)}`;
      const response = await fetch(proxyUrl);
      if (!response.ok) throw new Error('Network response');
      const data = await response.json();
      const meta = data.chart?.result?.[0]?.meta;
      return meta && meta.regularMarketPrice ? meta.regularMarketPrice : null;
    } catch (error) {
      return null;
    }
  };

  // 一鍵更新 (整合版)
  const handleOneClickUpdate = useCallback(async () => {
    setIsUpdating(true);
    setUpdateStatus('updating');
    
    try {
      const pricePromises = assets.map(async (asset) => {
        if (!/^\d+$/.test(asset.code)) return asset;
        const newPrice = await fetchRealTimePrice(asset.code);
        return newPrice ? { ...asset, price: newPrice } : asset;
      });
      
      let updatedAssets = await Promise.all(pricePromises);

      await new Promise(resolve => setTimeout(resolve, 600)); // 模擬連線
      updatedAssets = updatedAssets.map(asset => {
        const officialData = OFFICIAL_EPS_DB[asset.code];
        if (officialData) {
            return {
                ...asset,
                cumEPS: officialData.cumEPS,
                epsMonth: officialData.month,
                isVerified: true
            };
        }
        return asset;
      });

      setAssets(updatedAssets);
      setLastUpdated(new Date());
      setUpdateStatus('success');
      setTimeout(() => setUpdateStatus(''), 3000);

    } catch (error) {
      setUpdateStatus('error');
    } finally {
      setIsUpdating(false);
    }
  }, [assets]);

  // Excel 匯出功能 (使用 window.XLSX)
  const handleExportExcel = () => {
    if (!window.XLSX) {
        alert("Excel 功能載入中，請稍後再試...");
        return;
    }
    const wsData = assets.map(asset => ({
        '股票代號': asset.code,
        '股票名稱': asset.name,
        '持有股數': asset.quantity,
        '成本均價': asset.cost,
        '目前股價': asset.price,
        '累計EPS': asset.cumEPS,
        '資料月份': asset.epsMonth,
        '現金配息率(%)': asset.cashPayoutRate,
        '股票配股率(%)': asset.stockPayoutRate
    }));

    const ws = window.XLSX.utils.json_to_sheet(wsData);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "我的存股清單");
    window.XLSX.writeFile(wb, `存股管家備份_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // Excel 匯入功能
  const handleImportExcel = (e) => {
    if (!window.XLSX) {
        alert("Excel 功能載入中，請稍後再試...");
        return;
    }
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const bstr = evt.target.result;
            const wb = window.XLSX.read(bstr, { type: 'binary' });
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const data = window.XLSX.utils.sheet_to_json(ws);

            const importedAssets = data.map((row, index) => ({
                id: Date.now() + index, 
                code: String(row['股票代號'] || ''),
                name: row['股票名稱'] || '',
                quantity: Number(row['持有股數'] || 0),
                cost: Number(row['成本均價'] || 0),
                price: Number(row['目前股價'] || 0),
                cumEPS: Number(row['累計EPS'] || 0),
                epsMonth: Number(row['資料月份'] || 12),
                cashPayoutRate: Number(row['現金配息率(%)'] || 0),
                stockPayoutRate: Number(row['股票配股率(%)'] || 0),
                isVerified: false 
            }));

            if (importedAssets.length > 0) {
                if(window.confirm(`成功讀取 ${importedAssets.length} 筆資料，是否覆蓋現有清單？`)) {
                    setAssets(importedAssets);
                    alert('匯入成功！');
                }
            } else {
                alert('Excel 檔案內容格式不符或無資料。');
            }
        } catch (error) {
            console.error(error);
            alert('匯入失敗，請確認檔案格式是否正確。');
        }
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; 
  };

  // 計算邏輯
  const summary = useMemo(() => {
    let totalValue = 0;
    let totalCost = 0;
    let totalCashDiv = 0;
    let totalStockDivValue = 0;
    
    assets.forEach(asset => {
      totalValue += asset.price * asset.quantity;
      totalCost += asset.cost * asset.quantity;
      
      const months = asset.epsMonth || 12;
      let annualizedEPS = (asset.cumEPS / months) * 12;
      if (asset.cumEPS <= 0.1) annualizedEPS = 0;
      
      const estCashDiv = annualizedEPS * (asset.cashPayoutRate / 100);
      const estStockDiv = annualizedEPS * (asset.stockPayoutRate / 100);

      const cashIncome = asset.quantity * estCashDiv;
      const stockIncome = asset.quantity * (estStockDiv / 10) * asset.price;
      
      totalCashDiv += cashIncome;
      totalStockDivValue += stockIncome;
    });

    const profit = totalValue - totalCost;
    const returnRate = totalCost > 0 ? (profit / totalCost) * 100 : 0;
    const totalDivValue = totalCashDiv + totalStockDivValue;
    const yieldRate = totalValue > 0 ? (totalDivValue / totalValue) * 100 : 0;
    
    const maxLoanable = totalValue * 0.6; 
    const availableLoan = Math.max(0, maxLoanable - loanAmount);
    const maintenanceRate = loanAmount > 0 ? (totalValue / loanAmount) * 100 : 0;

    return { totalValue, totalCost, profit, returnRate, totalCashDiv, totalStockDivValue, totalDivValue, yieldRate, maxLoanable, availableLoan, maintenanceRate };
  }, [assets, loanAmount]);

  const getRateStatus = (rate) => {
    if (loanAmount === 0) return { status: 'none', text: '無借款', color: 'text-slate-400', bg: 'bg-slate-100' };
    if (rate < 130) return { status: 'danger', text: '追繳中', color: 'text-rose-600', bg: 'bg-rose-100' };
    if (rate < 166) return { status: 'warning', text: '注意', color: 'text-amber-500', bg: 'bg-amber-100' };
    return { status: 'safe', text: '安全', color: 'text-emerald-600', bg: 'bg-emerald-100' };
  };

  const rateStatus = getRateStatus(summary.maintenanceRate);

  const chartData = useMemo(() => {
    return assets.map(asset => ({
      name: asset.name,
      value: asset.price * asset.quantity
    })).sort((a, b) => b.value - a.value);
  }, [assets]);

  const handleAddAsset = () => {
    if (!newAsset.code || !newAsset.name) return;
    const dbData = OFFICIAL_EPS_DB[newAsset.code];
    const newItem = {
      id: Date.now(),
      code: newAsset.code,
      name: newAsset.name,
      type: newAsset.type,
      price: Number(newAsset.price) || 0,
      quantity: Number(newAsset.quantity) || 0,
      cost: Number(newAsset.cost) || 0,
      cumEPS: dbData ? dbData.cumEPS : (Number(newAsset.cumEPS) || 0),
      epsMonth: dbData ? dbData.month : (Number(newAsset.epsMonth) || 12),
      cashPayoutRate: Number(newAsset.cashPayoutRate) || 0,
      stockPayoutRate: Number(newAsset.stockPayoutRate) || 0,
      isVerified: !!dbData
    };
    setAssets(prev => [...prev, newItem]);
    setNewAsset({ code: '', name: '', type: '電子', price: '', quantity: '', cost: '', cumEPS: '', epsMonth: 9, cashPayoutRate: '', stockPayoutRate: '' });
    setSearchTerm('');
    setShowAddForm(false);
  };

  const startEditing = (id, field, value) => {
    setEditingCell({ id, field });
    setEditValue(value.toString());
  };

  const saveEdit = () => {
    if (editingCell.id) {
        setAssets(assets.map(a => {
            if (a.id === editingCell.id) {
                const updated = { ...a, [editingCell.field]: Number(editValue) };
                if (editingCell.field === 'cumEPS') updated.isVerified = false;
                return updated;
            }
            return a;
        }));
    }
    setEditingCell({ id: null, field: null });
  };

  const handleDelete = (id) => {
    if(window.confirm('確定要刪除這檔股票嗎？')) {
        setAssets(assets.filter(a => a.id !== id));
    }
  };

  const formatCurrency = (val) => {
    return new Intl.NumberFormat('zh-TW', { style: 'currency', currency: 'TWD', maximumFractionDigits: 0 }).format(val);
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 pb-24">
      {/* Hidden File Input for Import */}
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        ref={fileInputRef} 
        onChange={handleImportExcel} 
        style={{ display: 'none' }} 
      />

      {/* Header */}
      <div className="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div className="max-w-2xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center gap-2.5">
            <div className="bg-blue-600 p-1.5 rounded-lg shadow-sm">
              <Wallet className="text-white" size={20} />
            </div>
            <div>
                <h1 className="text-lg font-bold text-slate-800 leading-tight">存股管家 <span className="text-blue-600">{currentYear}</span></h1>
                <div className="flex items-center gap-2 text-[10px] text-slate-500 font-medium">
                    {saveStatus === 'saved' ? 
                        <span className="text-emerald-500 flex items-center bg-emerald-50 px-1.5 py-0.5 rounded"><Save size={10} className="mr-1"/>已儲存</span> : 
                        (lastUpdated ? <span>{lastUpdated.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})} 更新</span> : <span>準備就緒</span>)
                    }
                    {updateStatus === 'success' && <span className="text-emerald-500 flex items-center bg-emerald-50 px-1.5 py-0.5 rounded ml-2"><CheckCircle2 size={10} className="mr-1"/>完成</span>}
                </div>
            </div>
          </div>
          
          <div className="flex gap-2">
            <button 
                onClick={handleOneClickUpdate}
                disabled={isUpdating}
                className={`h-9 px-3 rounded-full flex items-center justify-center transition-all shadow-sm gap-1.5 text-xs font-bold ${
                    isUpdating 
                    ? 'bg-slate-100 text-slate-400' 
                    : 'bg-indigo-50 text-indigo-600 hover:bg-indigo-100 border border-indigo-100'
                }`}
            >
                {isUpdating ? <Loader2 size={14} className="animate-spin"/> : <Zap size={14} className="fill-current"/>}
                {isUpdating ? '更新中' : '一鍵更新'}
            </button>
            <button 
                onClick={() => setShowAddForm(true)}
                className="w-9 h-9 bg-blue-600 text-white rounded-full flex items-center justify-center shadow-md hover:bg-blue-700 active:scale-95 transition-all"
            >
                <Plus size={20} />
            </button>
          </div>
        </div>
      </div>

      <div className="max-w-2xl mx-auto px-4 py-4 space-y-4">
        
        {/* Main Stats Cards */}
        <div className="grid grid-cols-2 gap-3">
          <StatCard 
            title="總資產市值" 
            value={formatCurrency(summary.totalValue)}
            subValue={`成本 ${formatCurrency(summary.totalCost)}`}
            isPositive={true}
            icon={DollarSign}
            colorClass="bg-blue-500"
          />
           <StatCard 
            title="未實現損益" 
            value={(summary.profit > 0 ? "+" : "") + formatCurrency(summary.profit)}
            subValue={`${summary.returnRate.toFixed(1)}%`}
            isPositive={summary.profit >= 0}
            icon={TrendingUp}
            colorClass={summary.profit >= 0 ? "bg-emerald-500" : "bg-rose-500"}
          />
          <StatCard 
            title={`預估 ${currentYear+1} 股利`} 
            value={formatCurrency(summary.totalDivValue)}
            subValue={`殖利率 ${summary.yieldRate.toFixed(1)}%`}
            isPositive={true}
            icon={Coins}
            colorClass="bg-violet-500"
          />
          <StatCard 
            title="整戶維持率" 
            value={loanAmount > 0 ? `${summary.maintenanceRate.toFixed(0)}%` : "無"}
            subValue={rateStatus.text}
            isPositive={summary.maintenanceRate >= 166}
            alertLevel={rateStatus.status}
            icon={Landmark}
            colorClass={
                rateStatus.status === 'danger' ? 'bg-rose-500' : 
                rateStatus.status === 'warning' ? 'bg-amber-500' : 'bg-emerald-500'
            }
          />
        </div>

        {/* Pledge Dashboard */}
        <Card className="bg-gradient-to-br from-slate-800 to-slate-900 text-white border-none p-4 overflow-hidden relative shadow-lg">
          <div className="flex justify-between items-center relative z-10">
            <div className="flex items-center gap-2">
                <RefreshCw size={16} className="text-blue-300"/>
                <span className="font-bold text-sm">質押試算</span>
            </div>
            <button onClick={() => setShowPledgeSettings(!showPledgeSettings)} className="text-xs bg-white/10 px-2 py-1 rounded hover:bg-white/20 transition">
                {showPledgeSettings ? '收起' : '展開'}
            </button>
          </div>

          {showPledgeSettings && (
            <div className="mt-4 space-y-4 relative z-10 animate-in fade-in slide-in-from-top-2">
                <div>
                    <label className="text-xs text-slate-400 block mb-1">目前借款</label>
                    <div className="relative">
                        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">$</span>
                        <input 
                            type="number" value={loanAmount} onChange={(e) => setLoanAmount(Number(e.target.value))}
                            className="w-full bg-slate-700/50 border border-slate-600 rounded-lg py-2 pl-6 pr-3 text-white outline-none font-mono"
                        />
                    </div>
                </div>
                <div className="flex justify-between text-sm">
                    <div className="text-slate-400">可借: <span className="text-white font-mono">{formatCurrency(summary.maxLoanable)}</span></div>
                    <div className="text-slate-400">剩餘: <span className="text-emerald-400 font-mono">{formatCurrency(summary.availableLoan)}</span></div>
                </div>
            </div>
          )}
        </Card>

        {/* Main Table */}
        <Card className="p-0 overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
                <h2 className="font-bold text-slate-700 flex items-center gap-2">
                    <FileText size={18}/> 持股與股利明細
                </h2>
                <div className="flex items-center gap-2">
                    <span className="text-[10px] text-slate-400">點擊數字可修改</span>
                </div>
            </div>
            <div className="overflow-x-auto">
                <table className="w-full text-left text-sm whitespace-nowrap">
                  <thead className="bg-slate-50 text-slate-500 border-b border-slate-200 text-xs">
                    <tr>
                      <th className="px-3 py-3 font-semibold sticky left-0 bg-slate-50 z-20 shadow-[1px_0_4px_-2px_rgba(0,0,0,0.1)] w-24">股票</th>
                      <th className="px-2 py-3 font-semibold text-right">現價</th>
                      <th className="px-2 py-3 font-semibold text-right">損益</th>
                      
                      {/* EPS Columns */}
                      <th className="px-2 py-3 font-semibold text-center bg-blue-50/50 text-blue-800 border-l border-slate-100">EPS</th>
                      <th className="px-2 py-3 font-semibold text-center bg-blue-50/50 text-blue-800">月</th>
                      <th className="px-2 py-3 font-semibold text-center bg-blue-50/50 text-blue-800 font-bold border-r border-slate-100">年化</th>
                      
                      <th className="px-2 py-3 font-semibold text-right text-orange-700 bg-orange-50/50">配息%</th>
                      <th className="px-2 py-3 font-semibold text-right text-indigo-700 bg-indigo-50/50">配股%</th>
                      
                      {/* Dividend Columns */}
                      <th className="px-2 py-3 font-semibold text-right border-l border-slate-100 bg-orange-50/30 text-orange-900">預估現金</th>
                      <th className="px-2 py-3 font-semibold text-right bg-violet-50/30 text-violet-900">預估配股</th>
                      <th className="px-2 py-3 font-semibold text-right bg-violet-100/50 text-violet-900 font-bold">總股利</th>
                      
                      <th className="px-2 py-3 font-semibold text-right">股數</th>
                      <th className="px-2 py-3 font-semibold text-right">成本</th>
                      <th className="px-2 py-3 font-semibold text-center w-8"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {assets.map((asset) => {
                      const marketValue = asset.price * asset.quantity;
                      const costValue = asset.cost * asset.quantity;
                      const profit = marketValue - costValue;
                      
                      const months = asset.epsMonth || 12;
                      let annualizedEPS = (asset.cumEPS / months) * 12;
                      if (asset.cumEPS <= 0.1) annualizedEPS = 0;
                      
                      const estCashDiv = annualizedEPS * (asset.cashPayoutRate / 100);
                      const estStockDiv = annualizedEPS * (asset.stockPayoutRate / 100);
                      
                      const totalCash = asset.quantity * estCashDiv;
                      const totalStockValue = asset.quantity * (estStockDiv / 10) * asset.price;
                      const totalRowDiv = totalCash + totalStockValue;

                      return (
                        <tr key={asset.id} className="hover:bg-slate-50 transition-colors group">
                          <td className="px-3 py-3 sticky left-0 bg-white z-10 shadow-[1px_0_4px_-2px_rgba(0,0,0,0.1)]">
                            <div className="flex flex-col">
                              <span className="font-bold text-slate-700 text-sm flex items-center gap-1">
                                  {asset.name}
                                  {asset.isVerified && <ShieldCheck size={10} className="text-emerald-500" title="已自動校正"/>}
                              </span>
                              <span className="text-[10px] text-slate-400">{asset.code}</span>
                            </div>
                          </td>
                          <td className="px-2 py-3 text-right font-medium">{asset.price.toFixed(1)}</td>
                          <td className={`px-2 py-3 text-right font-medium ${profit >= 0 ? 'text-emerald-600' : 'text-rose-600'}`}>
                            {formatCurrency(profit).replace('NT$', '')}
                          </td>
                          
                          {/* EPS Details */}
                          <td className="px-2 py-3 text-center bg-blue-50/20 border-l border-slate-100">
                             {editingCell.id === asset.id && editingCell.field === 'cumEPS' ? (
                                <input type="number" className="w-10 px-1 py-0.5 text-center border border-blue-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div className="text-slate-600 text-xs border-b border-dashed border-blue-200 inline-block cursor-pointer"
                                    onClick={() => startEditing(asset.id, 'cumEPS', asset.cumEPS)}>
                                    {asset.cumEPS.toFixed(2)}
                                </div>
                             )}
                          </td>
                          <td className="px-2 py-3 text-center bg-blue-50/20">
                             {editingCell.id === asset.id && editingCell.field === 'epsMonth' ? (
                                <input type="number" className="w-8 px-1 py-0.5 text-center border border-blue-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div className="text-slate-400 text-xs cursor-pointer"
                                    onClick={() => startEditing(asset.id, 'epsMonth', asset.epsMonth)}>
                                    {asset.epsMonth}
                                </div>
                             )}
                          </td>
                          <td className="px-2 py-3 text-center font-bold text-blue-700 bg-blue-50/20 border-r border-slate-100 text-xs">
                              {annualizedEPS.toFixed(1)}
                          </td>

                          {/* Payout % */}
                          <td className="px-2 py-3 text-right text-orange-700 bg-orange-50/20 text-xs">
                             {editingCell.id === asset.id && editingCell.field === 'cashPayoutRate' ? (
                                <input type="number" className="w-8 px-1 py-0.5 text-center border border-orange-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div onClick={() => startEditing(asset.id, 'cashPayoutRate', asset.cashPayoutRate)} className="cursor-pointer">{asset.cashPayoutRate}%</div>
                             )}
                          </td>
                          <td className="px-2 py-3 text-right text-indigo-700 bg-indigo-50/20 text-xs">
                             {editingCell.id === asset.id && editingCell.field === 'stockPayoutRate' ? (
                                <input type="number" className="w-8 px-1 py-0.5 text-center border border-indigo-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div onClick={() => startEditing(asset.id, 'stockPayoutRate', asset.stockPayoutRate)} className="cursor-pointer">{asset.stockPayoutRate}%</div>
                             )}
                          </td>

                          {/* Dividend Values */}
                          <td className="px-2 py-3 text-right border-l border-slate-100 bg-orange-50/20 text-orange-900 font-medium text-xs">
                             {formatCurrency(totalCash).replace('NT$', '')}
                          </td>
                          <td className="px-2 py-3 text-right bg-violet-50/20 text-violet-900 font-medium text-xs">
                             {formatCurrency(totalStockValue).replace('NT$', '')}
                          </td>
                          <td className="px-2 py-3 text-right bg-violet-100/30 text-violet-900 font-bold text-xs">
                             {formatCurrency(totalRowDiv).replace('NT$', '')}
                          </td>

                          {/* Quantity */}
                          <td className="px-2 py-3 text-right text-slate-500 text-xs">
                             {editingCell.id === asset.id && editingCell.field === 'quantity' ? (
                                <input type="number" className="w-14 px-1 py-0.5 text-right border border-blue-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div className="border-b border-dashed border-slate-200 inline-block cursor-pointer"
                                    onClick={() => startEditing(asset.id, 'quantity', asset.quantity)}>
                                    {asset.quantity.toLocaleString()}
                                </div>
                             )}
                          </td>
                          
                          {/* Cost */}
                          <td className="px-2 py-3 text-right">
                             {editingCell.id === asset.id && editingCell.field === 'cost' ? (
                                <input type="number" className="w-14 px-1 py-0.5 text-right border border-blue-400 rounded text-xs"
                                    autoFocus value={editValue} onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEdit} onKeyDown={(e) => e.key === 'Enter' && saveEdit()} />
                             ) : (
                                <div className="text-slate-500 text-xs border-b border-dashed border-slate-300 inline-block cursor-pointer"
                                    onClick={() => startEditing(asset.id, 'cost', asset.cost)}>
                                    {asset.cost.toFixed(1)}
                                </div>
                             )}
                          </td>
                          
                          <td className="px-2 py-3 text-center">
                            <button onClick={() => handleDelete(asset.id)} className="text-slate-300 hover:text-rose-500 p-1">
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
            </div>
        </Card>

        {/* Chart */}
        <Card className="flex flex-col items-center">
            <h3 className="text-sm font-bold text-slate-700 mb-4 self-start">資產配置圖</h3>
            <div className="h-48 w-full">
            <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                <Pie
                    data={chartData}
                    cx="50%" cy="50%"
                    innerRadius={50} outerRadius={70}
                    paddingAngle={2}
                    dataKey="value"
                >
                    {chartData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                    ))}
                </Pie>
                <RechartsTooltip formatter={(value) => formatCurrency(value)} />
                <Legend iconSize={8} wrapperStyle={{fontSize: '10px'}}/>
                </PieChart>
            </ResponsiveContainer>
            </div>
        </Card>

        {/* Excel Export/Import Buttons */}
        <div className="flex gap-2 justify-center pt-2">
            <button onClick={handleExportExcel} className="flex items-center gap-1 px-3 py-2 bg-slate-100 text-slate-600 rounded-lg text-xs hover:bg-slate-200 transition">
                <FileDown size={14}/> 匯出 Excel
            </button>
            <button onClick={() => fileInputRef.current.click()} className="flex items-center gap-1 px-3 py-2 bg-slate-100 text-slate-600 rounded-lg text-xs hover:bg-slate-200 transition">
                <FileUp size={14}/> 匯入 Excel
            </button>
        </div>
      </div>

      {/* Add Asset Modal */}
      {showAddForm && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-end sm:items-center justify-center z-50">
          <div className="bg-white w-full sm:max-w-lg sm:rounded-xl rounded-t-2xl p-6 animate-in slide-in-from-bottom-10 duration-300 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-bold text-slate-800">新增資產</h3>
                <button onClick={() => setShowAddForm(false)} className="p-2 bg-slate-100 rounded-full"><X size={20}/></button>
            </div>
            
            <div className="space-y-5">
              {/* Search */}
              <div className="relative">
                  <label className="text-xs font-bold text-slate-500 mb-1.5 block">搜尋股票</label>
                  <div className="relative">
                      <Search className="absolute left-3 top-3 text-slate-400" size={18} />
                      <input 
                        type="text" placeholder="輸入 2330 或 台積電" 
                        className="w-full pl-10 pr-4 py-2.5 border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50 focus:bg-white transition-all"
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        ref={searchRef}
                      />
                  </div>
                  {showResults && searchResults.length > 0 && (
                      <div className="absolute z-50 w-full mt-2 bg-white border border-slate-200 rounded-xl shadow-xl max-h-48 overflow-y-auto">
                          {searchResults.map(stock => (
                              <div 
                                key={stock.code} onClick={() => selectStock(stock)}
                                className="px-4 py-3 hover:bg-blue-50 cursor-pointer flex justify-between items-center border-b border-slate-50 last:border-0"
                              >
                                  <span className="font-bold text-slate-700">{stock.name}</span>
                                  <span className="text-xs text-slate-400 bg-slate-100 px-2 py-1 rounded-full">{stock.code}</span>
                              </div>
                          ))}
                      </div>
                  )}
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-xs font-bold text-slate-500 mb-1 block">代號</label>
                  <input type="text" className="w-full px-3 py-2.5 border border-slate-300 rounded-xl bg-slate-50 text-slate-500" value={newAsset.code} readOnly />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 mb-1 block">名稱</label>
                  <input type="text" className="w-full px-3 py-2.5 border border-slate-300 rounded-xl bg-slate-50 text-slate-500" value={newAsset.name} readOnly />
                </div>
              </div>
              
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-xs font-bold text-slate-500 mb-1 block">持有股數</label>
                  <input type="number" placeholder="0" className="w-full px-3 py-2.5 border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-blue-500"
                    value={newAsset.quantity} onChange={(e) => setNewAsset({...newAsset, quantity: e.target.value})} />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 mb-1 block">成本均價</label>
                  <input type="number" placeholder="0.00" className="w-full px-3 py-2.5 border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-blue-500"
                    value={newAsset.cost} onChange={(e) => setNewAsset({...newAsset, cost: e.target.value})} />
                </div>
              </div>

              <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-100 space-y-4">
                  <h4 className="text-sm font-bold text-blue-800 flex items-center gap-2"><Calendar size={16}/> {currentYear} 估值設定</h4>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                        <label className="text-[10px] font-bold text-blue-600 mb-1 block">累計EPS</label>
                        <input type="number" className="w-full px-2 py-2 border border-blue-200 rounded-lg outline-none focus:border-blue-500 text-center"
                            value={newAsset.cumEPS} onChange={(e) => setNewAsset({...newAsset, cumEPS: e.target.value})} />
                    </div>
                    <div>
                        <label className="text-[10px] font-bold text-blue-600 mb-1 block">月份</label>
                        <input type="number" className="w-full px-2 py-2 border border-blue-200 rounded-lg outline-none focus:border-blue-500 text-center"
                            value={newAsset.epsMonth} onChange={(e) => setNewAsset({...newAsset, epsMonth: e.target.value})} />
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                        <label className="text-[10px] font-bold text-indigo-600 mb-1 block">配息率%</label>
                        <input type="number" placeholder="50" className="w-full px-2 py-2 border border-indigo-200 rounded-lg outline-none focus:border-indigo-500 text-center"
                            value={newAsset.cashPayoutRate} onChange={(e) => setNewAsset({...newAsset, cashPayoutRate: e.target.value})} />
                    </div>
                    <div>
                        <label className="text-[10px] font-bold text-indigo-600 mb-1 block">配股率%</label>
                        <input type="number" placeholder="0" className="w-full px-2 py-2 border border-indigo-200 rounded-lg outline-none focus:border-indigo-500 text-center"
                            value={newAsset.stockPayoutRate} onChange={(e) => setNewAsset({...newAsset, stockPayoutRate: e.target.value})} />
                    </div>
                  </div>
              </div>

              <button onClick={handleAddAsset} className="w-full py-3.5 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-700 active:scale-95 transition-all shadow-lg shadow-blue-200 mt-2">
                  確認新增資產
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
