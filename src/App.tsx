/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  TrendingUp, 
  Calendar, 
  AlertCircle, 
  CheckCircle2, 
  BarChart3,
  Settings2,
  ClipboardPaste,
  Search,
  Star,
  Trash2,
  Info,
  ArrowRight,
  Download,
  Image as ImageIcon,
  FileSpreadsheet,
  ExternalLink,
  Copy,
  Check
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as htmlToImage from 'html-to-image';
import * as XLSX from 'xlsx';

// --- Types ---

type Supermarket = 'TGDD' | 'TOPZONE';
type ViewMode = 'DASHBOARD' | 'STAR_PUSH' | 'STAR_DEDUCTION';

interface DeductionRule {
  id: string;
  name: string;
  basePoints: number;
  penaltyPoints: number;
  keywords: string[];
}

const DEDUCTION_RULES: DeductionRule[] = [
  { id: 'phu-kien', name: 'Phụ Kiện', basePoints: 2, penaltyPoints: 3, keywords: [] },
  { id: 'dich-vu', name: 'Dịch Vụ', basePoints: 2, penaltyPoints: 1, keywords: [] },
  { id: 'tra-cham', name: 'Trả Chậm', basePoints: 2, penaltyPoints: 4, keywords: [] },
  { id: 'ict', name: 'ICT', basePoints: 2, penaltyPoints: 2, keywords: [] },
  { id: 'doanh-thu', name: 'Doanh Thu', basePoints: 1, penaltyPoints: 0, keywords: [] },
];

const DEDUCTION_SUBJECTS: Record<Supermarket, { category: string; group: string }[]> = {
  TGDD: [
    { category: 'Camera', group: 'Phụ Kiện' },
    { category: 'Pin sạc dự phòng', group: 'Phụ Kiện' },
    { category: 'TAI NGHE BLUETOOTH', group: 'Phụ Kiện' },
    { category: 'SIM MOBIFONE&VINAPHONE', group: 'Dịch Vụ' },
    { category: 'Sim Tổng', group: 'Dịch Vụ' },
    { category: 'VAS', group: 'Dịch Vụ' },
    { category: 'RÚT TIỀN TÀI KHOẢN NGÂN HÀNG', group: 'Dịch Vụ' },
    { category: 'TPBANK EVO VÀ VPBANK MWG', group: 'Trả Chậm' },
    { category: 'SMARTPHONE FLAGSHIP & TABLET ANDROID', group: 'ICT' },
    { category: 'ĐIỆN THOẠI REALME', group: 'ICT' },
    { category: 'ĐIỆN THOẠI VIVO', group: 'ICT' },
    { category: 'ĐỒNG HỒ', group: 'Phụ Kiện' },
    { category: 'Laptop', group: 'ICT' },
    { category: 'Phụ kiện - Đồng Hồ', group: 'Phụ Kiện' },
    { category: 'Homecredit', group: 'Trả Chậm' },
    { category: 'FECREDIT, SHINHAN, SAMSUNG FINANCE+', group: 'Trả Chậm' },
    { category: 'VÍ TRẢ SAU', group: 'Trả Chậm' },
    { category: 'VAY TIỀN MẶT CAKE VÀ FECREDIT', group: 'Trả Chậm' },
    { category: 'Bảo hiểm', group: 'Dịch Vụ' },
    { category: 'Bảo hiểm BÁN KÈM TRẢ CHẬM', group: 'Dịch Vụ' },
    { category: 'Doanh Thu', group: 'Doanh Thu' },
  ],
  TOPZONE: [
    { category: 'Pin sạc dự phòng', group: 'Phụ Kiện' },
    { category: 'VAS', group: 'Dịch Vụ' },
    { category: 'TPBANK EVO VÀ VPBANK MWG', group: 'Trả Chậm' },
    { category: 'Phụ kiện - Đồng Hồ', group: 'Phụ Kiện' },
    { category: 'Homecredit', group: 'Trả Chậm' },
    { category: 'FECREDIT, SHINHAN, SAMSUNG FINANCE+', group: 'Trả Chậm' },
    { category: 'VÍ TRẢ SAU', group: 'Trả Chậm' },
    { category: 'Bảo hiểm', group: 'Dịch Vụ' },
    { category: 'Bảo hiểm BÁN KÈM TRẢ CHẬM', group: 'Dịch Vụ' },
    { category: 'Doanh Thu', group: 'Doanh Thu' },
    { category: 'TAI NGHE BLUETOOTH', group: 'Phụ Kiện' },
  ]
};

interface StarThresholds {
  s1: number;
  s2: number;
  s3: number;
  s5: number;
  s10: number;
  s20: number;
  deduction: number;
}

interface KPIData {
  id: string;
  category: string;
  current: number;
  target: number;
  type: 'DTQĐ' | 'SLLK' | 'DTLK';
  status?: string;
  realtime?: number;
}

// --- Constants ---

const STAR_CONFIGS: Record<Supermarket, Record<string, StarThresholds>> = {
  TGDD: {
    'SIM': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'SIM MOBIFONE&VINAPHONE': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'TPBANK EVO VÀ VPBANK MWG': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'VAS': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'NẠP RÚT TIỀN TÀI KHOẢN NGÂN HÀNG': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'VAY TIỀN MẶT CAKE VÀ FECREDIT': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'VÍ TRẢ SAU': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'Homecredit': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'FECREDIT, SHINHAN, SAMSUNG FINANCE+': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'Bảo hiểm Xe Máy, Ô tô, SS Care +, BHRV Homecredit': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'Bảo hiểm': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'Camera': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'PIN SẠC DỰ PHÒNG': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'TAI NGHE BLUETOOTH': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'SMARTPHONE & TABLET ANDROID TRÊN 8 TRIỆU': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 40 },
    'ANDROID SẢN PHẨM MỚI': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 40 },
    'MÁY LỌC KHÔNG KHÍ - HÚT ẨM - HÚT BỤI': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'Laptop': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 40 },
    'ĐỒNG HỒ': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 40 },
    'PHỤ KIỆN - ĐỒNG HỒ': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'Realme': { s1: 100, s2: 120, s3: 150, s5: 200, s10: 220, s20: 250, deduction: 20 },
    'Vivo': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
  },
  TOPZONE: {
    'TPBANK EVO VÀ VPBANK MWG': { s1: 100, s2: 200, s3: 300, s5: 400, s10: 500, s20: 1000, deduction: 50 },
    'PIN SẠC DỰ PHÒNG': { s1: 100, s2: 200, s3: 300, s5: 400, s10: 500, s20: 1000, deduction: 50 },
    'VAS': { s1: 100, s2: 200, s3: 300, s5: 400, s10: 500, s20: 1000, deduction: 20 },
    'VÍ TRẢ SAU': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 50 },
    'PHỤ KIỆN - ĐỒNG HỒ': { s1: 100, s2: 200, s3: 300, s5: 400, s10: 500, s20: 1000, deduction: 20 },
    'Homecredit': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 20 },
    'FECREDIT, SHINHAN, SAMSUNG FINANCE+': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 20 },
    'Bảo hiểm Xe Máy, Ô tô, SS Care +, BHRV Homecredit': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
    'Bảo hiểm': { s1: 100, s2: 150, s3: 200, s5: 250, s10: 300, s20: 400, deduction: 30 },
  }
};

// --- Constants & Helpers ---

const getDaysInMonth = (year: number, month: number) => new Date(year, month + 1, 0).getDate();

const now = new Date();
const currentYear = now.getFullYear();
const currentMonth = now.getMonth();
const currentDay = now.getDate();
const daysInMonth = getDaysInMonth(currentYear, currentMonth);

// Số ngày còn lại tính cả ngày hôm nay
const DAYS_REMAINING = Math.max(1, daysInMonth - currentDay + 1);
const CURRENT_MONTH_NAME = `Tháng ${currentMonth + 1}`;
const CURRENT_MONTH_DISPLAY = `${(currentMonth + 1).toString().padStart(2, '0')}/${currentYear}`;
const CURRENT_DATE_DISPLAY = now.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' });

const BI_URLS = {
  TGDD: 'https://bi.thegioididong.com/thi-dua-st?id=92210.0&tab=1&rt=2&dm=2&mt=2',
  TOPZONE: 'https://bi.thegioididong.com/thi-dua-st?id=92765&tab=1&rt=2&dm=2&mt=2'
};

// --- Helper Functions ---

const parsePastedData = (text: string): KPIData[] => {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const results: KPIData[] = [];
  let currentType: 'DTQĐ' | 'SLLK' | 'DTLK' = 'DTLK';

  lines.forEach((line, index) => {
    // Detect section headers
    if (line.includes('DTQĐ')) currentType = 'DTQĐ';
    else if (line.includes('SLLK')) currentType = 'SLLK';
    else if (line.includes('DTLK')) currentType = 'DTLK';

    // Skip headers and "Tổng" rows
    if (line.includes('Ngành hàng') || line.startsWith('Tổng')) return;

    // Split by tab or multiple spaces
    const parts = line.split(/\t|\s{2,}/);
    if (parts.length >= 3) {
      const category = parts[0].trim();
      const current = parseFloat(parts[1].replace(/,/g, '')) || 0;
      const target = parseFloat(parts[2].replace(/,/g, '')) || 0;
      const status = parts[parts.length - 1]?.trim() || '';

      // Only add if target > 0
      if (target > 0) {
        results.push({
          id: `${currentType}-${index}-${category}`,
          category,
          current,
          target,
          type: currentType,
          status: status === 'TOP' || status === 'BOTTOM' ? status : undefined
        });
      }
    }
  });

  return results;
};

const parseRealtimeData = (text: string): { category: string, value: number }[] => {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const results: { category: string, value: number }[] = [];
  
  lines.forEach(line => {
    const parts = line.split(/\t|\s{2,}/);
    if (parts.length >= 2) {
      const category = parts[0].trim();
      const value = parseFloat(parts[1].replace(/,/g, '')) || 0;
      results.push({ category, value });
    }
  });
  return results;
};

// --- Components ---

export default function App() {
  const [activeTab, setActiveTab] = useState<Supermarket>('TOPZONE');
  const [viewMode, setViewMode] = useState<ViewMode>('DASHBOARD');
  const [data, setData] = useState<KPIData[]>([]);
  const [pastedTextTgdd, setPastedTextTgdd] = useState('');
  const [pastedTextTopzone, setPastedTextTopzone] = useState('');
  const [realtimeTextTgdd, setRealtimeTextTgdd] = useState('');
  const [realtimeTextTopzone, setRealtimeTextTopzone] = useState('');
  const [desiredTargetPercent, setDesiredTargetPercent] = useState(100);
  const [autoWatchMin, setAutoWatchMin] = useState(60);
  const [autoWatchMax, setAutoWatchMax] = useState(100);
  const [daysUsed, setDaysUsed] = useState(Math.max(1, currentDay - 1));
  const [searchQuery, setSearchQuery] = useState('');
  const [deductionManualPoints, setDeductionManualPoints] = useState<Record<string, number>>({});
  const [watchlist, setWatchlist] = useState<string[]>([]);
  const [manualUnstarred, setManualUnstarred] = useState<string[]>([]);
  const [enableHighlight, setEnableHighlight] = useState(true);
  const [showImport, setShowImport] = useState(false);
  const [showRealtimeImport, setShowRealtimeImport] = useState(false);
  const [copySuccess, setCopySuccess] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingValue, setEditingValue] = useState<string>('');
  const [isExporting, setIsExporting] = useState(false);
  const [exportState, setExportState] = useState<'IDLE' | 'CAPTURING_TGDD' | 'CAPTURING_TOPZONE'>('IDLE');
  const [originalTab, setOriginalTab] = useState<Supermarket | null>(null);
  const tableRef = useRef<HTMLDivElement>(null);
  const watchlistRef = useRef<HTMLDivElement>(null);
  const topzoneRef = useRef<HTMLDivElement>(null);

  // Load initial data or from localStorage
  useEffect(() => {
    const savedTab = localStorage.getItem('kpi_active_tab') as Supermarket;
    const tab = savedTab || 'TOPZONE';
    setActiveTab(tab);

    const saved = localStorage.getItem(`kpi_data_${tab}`);
    if (saved) {
      setData(JSON.parse(saved));
    }
    const savedWatchlist = localStorage.getItem(`kpi_watchlist_${tab}`);
    if (savedWatchlist) {
      setWatchlist(JSON.parse(savedWatchlist));
    }
    const savedManualUnstarred = localStorage.getItem(`kpi_manual_unstarred_${tab}`);
    if (savedManualUnstarred) {
      setManualUnstarred(JSON.parse(savedManualUnstarred));
    }
    const savedMin = localStorage.getItem('kpi_auto_min');
    const savedMax = localStorage.getItem('kpi_auto_max');
    if (savedMin) setAutoWatchMin(Number(savedMin));
    if (savedMax) setAutoWatchMax(Number(savedMax));
    
    const savedHighlight = localStorage.getItem('kpi_enable_highlight');
    if (savedHighlight !== null) {
      setEnableHighlight(savedHighlight === 'true');
    }
  }, []);

  // Save to localStorage
  useEffect(() => {
    if (data.length > 0 || watchlist.length > 0 || manualUnstarred.length > 0) {
      localStorage.setItem(`kpi_data_${activeTab}`, JSON.stringify(data));
      localStorage.setItem(`kpi_watchlist_${activeTab}`, JSON.stringify(watchlist));
      localStorage.setItem(`kpi_manual_unstarred_${activeTab}`, JSON.stringify(manualUnstarred));
    }
    localStorage.setItem('kpi_active_tab', activeTab);
    localStorage.setItem('kpi_auto_min', autoWatchMin.toString());
    localStorage.setItem('kpi_auto_max', autoWatchMax.toString());
    localStorage.setItem('kpi_deduction_manual', JSON.stringify(deductionManualPoints));
    localStorage.setItem('kpi_enable_highlight', enableHighlight.toString());
  }, [data, watchlist, manualUnstarred, activeTab, autoWatchMin, autoWatchMax, deductionManualPoints, enableHighlight]);

  useEffect(() => {
    const savedManual = localStorage.getItem('kpi_deduction_manual');
    if (savedManual) setDeductionManualPoints(JSON.parse(savedManual));
  }, []);

  const handleTabChange = (newTab: Supermarket) => {
    if (newTab === activeTab) return;
    
    // Save current state before switching
    localStorage.setItem(`kpi_data_${activeTab}`, JSON.stringify(data));
    localStorage.setItem(`kpi_watchlist_${activeTab}`, JSON.stringify(watchlist));
    localStorage.setItem(`kpi_manual_unstarred_${activeTab}`, JSON.stringify(manualUnstarred));

    // Load new state
    const savedData = localStorage.getItem(`kpi_data_${newTab}`);
    const savedWatchlist = localStorage.getItem(`kpi_watchlist_${newTab}`);
    const savedManualUnstarred = localStorage.getItem(`kpi_manual_unstarred_${newTab}`);
    
    setActiveTab(newTab);
    setData(savedData ? JSON.parse(savedData) : []);
    setWatchlist(savedWatchlist ? JSON.parse(savedWatchlist) : []);
    setManualUnstarred(savedManualUnstarred ? JSON.parse(savedManualUnstarred) : []);
  };

  const handleImport = () => {
    const parsedTgdd = pastedTextTgdd.trim() ? parsePastedData(pastedTextTgdd) : [];
    const parsedTopzone = pastedTextTopzone.trim() ? parsePastedData(pastedTextTopzone) : [];

    if (parsedTgdd.length > 0) {
      localStorage.setItem('kpi_data_TGDD', JSON.stringify(parsedTgdd));
      if (activeTab === 'TGDD') {
        setData(parsedTgdd);
      }
    }

    if (parsedTopzone.length > 0) {
      localStorage.setItem('kpi_data_TOPZONE', JSON.stringify(parsedTopzone));
      if (activeTab === 'TOPZONE') {
        setData(parsedTopzone);
      }
    }

    setShowImport(false);
    setPastedTextTgdd('');
    setPastedTextTopzone('');
  };

  const handleRealtimeImport = () => {
    const parsedTgdd = realtimeTextTgdd.trim() ? parseRealtimeData(realtimeTextTgdd) : [];
    const parsedTopzone = realtimeTextTopzone.trim() ? parseRealtimeData(realtimeTextTopzone) : [];

    // 1. Process TGDD Realtime Data
    if (parsedTgdd.length > 0) {
      if (activeTab === 'TGDD') {
        setData(prevData => prevData.map(item => {
          const match = parsedTgdd.find(r => r.category.toLowerCase() === item.category.toLowerCase());
          return match ? { ...item, realtime: match.value } : item;
        }));
      } else {
        const savedTgdd = localStorage.getItem('kpi_data_TGDD');
        if (savedTgdd) {
          const tgddData: KPIData[] = JSON.parse(savedTgdd);
          const updated = tgddData.map(item => {
            const match = parsedTgdd.find(r => r.category.toLowerCase() === item.category.toLowerCase());
            return match ? { ...item, realtime: match.value } : item;
          });
          localStorage.setItem('kpi_data_TGDD', JSON.stringify(updated));
        }
      }
    }

    // 2. Process TOPZONE Realtime Data
    if (parsedTopzone.length > 0) {
      if (activeTab === 'TOPZONE') {
        setData(prevData => prevData.map(item => {
          const match = parsedTopzone.find(r => r.category.toLowerCase() === item.category.toLowerCase());
          return match ? { ...item, realtime: match.value } : item;
        }));
      } else {
        const savedTopzone = localStorage.getItem('kpi_data_TOPZONE');
        if (savedTopzone) {
          const topzoneData: KPIData[] = JSON.parse(savedTopzone);
          const updated = topzoneData.map(item => {
            const match = parsedTopzone.find(r => r.category.toLowerCase() === item.category.toLowerCase());
            return match ? { ...item, realtime: match.value } : item;
          });
          localStorage.setItem('kpi_data_TOPZONE', JSON.stringify(updated));
        }
      }
    }

    setShowRealtimeImport(false);
    setRealtimeTextTgdd('');
    setRealtimeTextTopzone('');
  };

  const updateRealtimeValue = (id: string, value: number | undefined) => {
    setData(prevData => prevData.map(item => {
      if (item.id === id) {
        if (value === undefined) {
          const { realtime, ...rest } = item;
          return rest;
        }
        return { ...item, realtime: value };
      }
      return item;
    }));
  };

  const startEditing = (item: KPIData) => {
    setEditingId(item.id);
    setEditingValue((item.realtime || 0).toString());
  };

  const saveEditing = () => {
    if (editingId) {
      const val = parseFloat(editingValue);
      if (!isNaN(val)) {
        updateRealtimeValue(editingId, val);
      }
      setEditingId(null);
    }
  };

  const clearRealtime = () => {
    // Clear active state
    setData(prevData => prevData.map(item => {
      const { realtime, ...rest } = item;
      return rest;
    }));

    // Clear other tab's realtime data in localStorage
    const otherTab = activeTab === 'TOPZONE' ? 'TGDD' : 'TOPZONE';
    const savedOther = localStorage.getItem(`kpi_data_${otherTab}`);
    if (savedOther) {
      const otherData: KPIData[] = JSON.parse(savedOther);
      const cleared = otherData.map(item => {
        const { realtime, ...rest } = item;
        return rest;
      });
      localStorage.setItem(`kpi_data_${otherTab}`, JSON.stringify(cleared));
    }
  };

  const toggleWatchlist = (id: string) => {
    const item = data.find(d => d.id === id) || filteredData.find(d => d.id === id);
    if (!item) return;

    const progress = getCompletionPercent(item);
    const isAutoWatched = progress >= autoWatchMin && progress < autoWatchMax;
    const isCurrentlyWatched = watchlist.includes(id) || (isAutoWatched && !manualUnstarred.includes(id));

    if (isCurrentlyWatched) {
      // We want to unwatch it
      setWatchlist(prev => prev.filter(i => i !== id));
      if (isAutoWatched) {
        setManualUnstarred(prev => prev.includes(id) ? prev : [...prev, id]);
      }
    } else {
      // We want to watch it
      setWatchlist(prev => prev.includes(id) ? prev : [...prev, id]);
      setManualUnstarred(prev => prev.filter(i => i !== id));
    }
  };

  const getEffectiveCurrent = (item: KPIData) => item.current + (item.realtime || 0);

  const getStarInfo = (item: KPIData) => {
    const config = STAR_CONFIGS[activeTab][item.category] || STAR_CONFIGS[activeTab][item.category.toUpperCase()];
    if (!config) return null;

    const currentPercent = (item.target === 0) ? 0 : (getEffectiveCurrent(item) / item.target) * 100;
    
    let stars = 0;
    if (currentPercent >= config.s20) stars = 20;
    else if (currentPercent >= config.s10) stars = 10;
    else if (currentPercent >= config.s5) stars = 5;
    else if (currentPercent >= config.s3) stars = 3;
    else if (currentPercent >= config.s2) stars = 2;
    else if (currentPercent >= config.s1) stars = 1;

    const projectedPercent = (item.target === 0) ? 0 : ((getEffectiveCurrent(item) / daysUsed) * daysInMonth / item.target) * 100;
    
    let deduction = 0;
    if (projectedPercent < 100) {
      deduction = config.deduction;
    }

    // Find next threshold
    let nextThreshold = null;
    let nextStars = 0;
    if (currentPercent < config.s1) { nextThreshold = config.s1; nextStars = 1; }
    else if (currentPercent < config.s2) { nextThreshold = config.s2; nextStars = 2; }
    else if (currentPercent < config.s3) { nextThreshold = config.s3; nextStars = 3; }
    else if (currentPercent < config.s5) { nextThreshold = config.s5; nextStars = 5; }
    else if (currentPercent < config.s10) { nextThreshold = config.s10; nextStars = 10; }
    else if (currentPercent < config.s20) { nextThreshold = config.s20; nextStars = 20; }

    return { stars, deduction, config, currentPercent, nextThreshold, nextStars };
  };

  const getCompletionPercent = (item: KPIData) => {
    const current = getEffectiveCurrent(item);
    if (item.target === 0) return current > 0 ? 100 : 0;
    return (current / item.target) * 100;
  };

  const getProjectedPercent = (item: KPIData) => {
    const current = getEffectiveCurrent(item);
    if (item.target === 0) return 0;
    return ((current / daysUsed) * daysInMonth / item.target) * 100;
  };

  const calculateDaily = (item: KPIData) => {
    const current = getEffectiveCurrent(item);
    const totalNeeded = (item.target * desiredTargetPercent) / 100;
    const remaining = totalNeeded - current;
    const remainingDays = Math.max(1, daysInMonth - daysUsed + 1);
    return remaining > 0 ? remaining / remainingDays : 0;
  };

  const filteredData = useMemo(() => {
    // Get all subjects for the active supermarket
    const subjects = DEDUCTION_SUBJECTS[activeTab];
    
    // Create a base list from subjects
    const baseList = subjects.map(s => {
      const existing = data.find(item => item.category.toLowerCase() === s.category.toLowerCase());
      if (existing) return existing;
      return {
        id: `manual_${activeTab}_${s.category}`,
        category: s.category,
        current: 0,
        target: 0,
        lastUpdated: new Date().toISOString()
      };
    });

    // Also include any items in data that are NOT in DEDUCTION_SUBJECTS (if any)
    const extraItems = data.filter(item => 
      !subjects.some(s => s.category.toLowerCase() === item.category.toLowerCase())
    );

    const fullList = [...baseList, ...extraItems];

    return fullList
      .filter(item => {
        const matchesSearch = item.category.toLowerCase().includes(searchQuery.toLowerCase());
        // Only show items with target > 0
        return item.target > 0 && matchesSearch;
      })
      .sort((a, b) => {
        const progressA = getCompletionPercent(a);
        const progressB = getCompletionPercent(b);
        const isDoneA = progressA >= desiredTargetPercent;
        const isDoneB = progressB >= desiredTargetPercent;

        if (isDoneA && !isDoneB) return 1;
        if (!isDoneA && isDoneB) return -1;

        const projA = getProjectedPercent(a);
        const projB = getProjectedPercent(b);
        return projB - projA; // Sort descending by projected completion
      });
  }, [data, activeTab, searchQuery, desiredTargetPercent, daysUsed]);

  const topMissingIds = useMemo(() => {
    if (!enableHighlight) return [];
    return filteredData
      .map(item => {
        const current = getEffectiveCurrent(item);
        const missing = Math.max(0, (item.target * desiredTargetPercent / 100) - current);
        return { id: item.id, missing };
      })
      .filter(x => x.missing > 0)
      .sort((a, b) => b.missing - a.missing)
      .slice(0, 3)
      .map(x => x.id);
  }, [filteredData, desiredTargetPercent, enableHighlight]);

  const watchedItems = useMemo(() => {
    return data
      .filter(item => {
        if (item.target <= 0) return false;
        const progress = getCompletionPercent(item);
        const isAutoWatched = progress >= autoWatchMin && progress < autoWatchMax;
        const isManuallyUnstarred = manualUnstarred.includes(item.id);
        const isManuallyStarred = watchlist.includes(item.id);
        return isManuallyStarred || (isAutoWatched && !isManuallyUnstarred);
      })
      .sort((a, b) => getCompletionPercent(b) - getCompletionPercent(a));
  }, [data, watchlist, manualUnstarred, autoWatchMin, autoWatchMax]);

  const stats = useMemo(() => {
    if (filteredData.length === 0) return null;

    const items = filteredData.map(item => ({
      progress: getCompletionPercent(item),
      projected: getProjectedPercent(item)
    }));

    const completed = items.filter(i => i.projected >= 100).length;
    const high = items.filter(i => i.progress >= 80).length;
    const low = items.filter(i => i.progress < 50).length;
    const successRate = (completed / items.length) * 100;

    return { completed, high, low, successRate };
  }, [filteredData]);



  const exportToExcel = () => {
    const exportData = filteredData.map(item => {
      const currentPercent = getCompletionPercent(item);
      const projectedPercent = getProjectedPercent(item);
      const daily = calculateDaily(item);
      const deduction = deductionStats.find(s => s.category.toLowerCase() === item.category.toLowerCase() && s.supermarket === activeTab);
      
      return {
        'Ngành hàng': item.category,
        'Target': item.target,
        'Current': item.current,
        'Realtime': item.realtime || 0,
        '% Hoàn Thành': `${currentPercent.toFixed(1)}%`,
        '% Dự Kiến': `${projectedPercent.toFixed(1)}%`,
        'Cần thêm/ngày': daily.toFixed(1),
        'Trừ Sao': deduction?.starDeduction || 0
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'KPI');
    XLSX.writeFile(wb, `KPI_${activeTab}_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.xlsx`);
  };

  const copyMainTableToClipboard = () => {
    const header = 'Ngành hàng\tTarget\tCurrent\tRealtime\t% Hoàn Thành\t% Dự Kiến\tCần thêm/ngày\tTrừ Sao\n';
    const rows = filteredData.map(item => {
      const currentPercent = getCompletionPercent(item);
      const projectedPercent = getProjectedPercent(item);
      const daily = calculateDaily(item);
      const deduction = deductionStats.find(s => s.category.toLowerCase() === item.category.toLowerCase() && s.supermarket === activeTab);
      
      return `${item.category}\t${item.target}\t${item.current}\t${item.realtime || 0}\t${currentPercent.toFixed(1)}%\t${projectedPercent.toFixed(1)}%\t${daily.toFixed(1)}\t${deduction?.starDeduction || 0}`;
    }).join('\n');
    
    navigator.clipboard.writeText(header + rows).then(() => {
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    });
  };

  const exportDeductionsToExcel = () => {
    const exportData = deductionStats.map(s => ({
      'Siêu Thị': s.supermarket,
      'Ngành hàng': s.category,
      'Nhóm': s.group,
      'Sao ghi nhận 5 sao': '100%',
      'Sao ghi nhận 10 sao': '120%',
      'Sao ghi nhận 15 sao': '150%',
      'Sao ghi nhận 20 sao': '250%',
      'Số Sao bị trừ khi dùng sao': s.starDeduction > 0 ? `-${s.starDeduction}` : '0'
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Deductions');
    XLSX.writeFile(wb, `Deductions_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.xlsx`);
  };

  const copyDeductionsToClipboard = () => {
    const header = 'Siêu Thị\tNgành hàng\tNhóm\tSao ghi nhận 5 sao\tSao ghi nhận 10 sao\tSao ghi nhận 15 sao\tSao ghi nhận 20 sao\tSố Sao bị trừ khi dùng sao\n';
    const rows = deductionStats.map(s => 
      `${s.supermarket}\t${s.category}\t${s.group}\t100%\t120%\t150%\t250%\t${s.starDeduction > 0 ? `-${s.starDeduction}` : '0'}`
    ).join('\n');
    
    navigator.clipboard.writeText(header + rows).then(() => {
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    });
  };

  const exportTopzoneToImage = async () => {
    if (topzoneRef.current) {
      setIsExporting(true);
      // Wait for React to re-render without overflow classes
      setTimeout(async () => {
        try {
          const dataUrl = await htmlToImage.toPng(topzoneRef.current!, {
            backgroundColor: '#F8F9FA',
            pixelRatio: 3,
            cacheBust: true,
            filter: (node: any) => {
              return node.id === 'main-export-btn' || node.id === 'excel-export-btn' ? false : true;
            },
            style: {
              borderRadius: '0',
              margin: '0',
              padding: '40px',
              width: '1600px',
              height: 'auto',
              overflow: 'visible',
              maxWidth: 'none',
            }
          });
          const link = document.createElement('a');
          link.download = `KPI_${activeTab}_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.png`;
          link.href = dataUrl;
          link.click();
        } catch (error) {
          console.error('Error exporting image:', error);
        } finally {
          setIsExporting(false);
        }
      }, 100);
    }
  };

  const exportBothSupermarketsToImage = () => {
    setIsExporting(true);
    setOriginalTab(activeTab);
    setExportState('CAPTURING_TGDD');
  };

  useEffect(() => {
    if (exportState === 'IDLE') return;

    const runExport = async () => {
      // Small timeout to ensure DOM has fully painted
      await new Promise(resolve => setTimeout(resolve, 500));

      if (exportState === 'CAPTURING_TGDD') {
        if (activeTab !== 'TGDD') {
          handleTabChange('TGDD');
          return;
        }

        try {
          if (topzoneRef.current) {
            const tgddDataUrl = await htmlToImage.toPng(topzoneRef.current, {
              backgroundColor: '#F8F9FA',
              pixelRatio: 3,
              cacheBust: true,
              filter: (node: any) => {
                return node.id === 'main-export-btn' || node.id === 'excel-export-btn' || node.id === 'watchlist-export-btn' || node.id === 'watchlist-copy-btn' ? false : true;
              },
              style: {
                borderRadius: '0',
                margin: '0',
                padding: '40px',
                width: '1600px',
                height: 'auto',
                overflow: 'visible',
                maxWidth: 'none',
              }
            });
            const linkTgdd = document.createElement('a');
            linkTgdd.download = `KPI_TGDD_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.png`;
            linkTgdd.href = tgddDataUrl;
            linkTgdd.click();
          }
        } catch (error) {
          console.error('Error exporting TGDĐ:', error);
        }

        setExportState('CAPTURING_TOPZONE');
      } 
      else if (exportState === 'CAPTURING_TOPZONE') {
        if (activeTab !== 'TOPZONE') {
          handleTabChange('TOPZONE');
          return;
        }

        try {
          if (topzoneRef.current) {
            const topzoneDataUrl = await htmlToImage.toPng(topzoneRef.current, {
              backgroundColor: '#F8F9FA',
              pixelRatio: 3,
              cacheBust: true,
              filter: (node: any) => {
                return node.id === 'main-export-btn' || node.id === 'excel-export-btn' || node.id === 'watchlist-export-btn' || node.id === 'watchlist-copy-btn' ? false : true;
              },
              style: {
                borderRadius: '0',
                margin: '0',
                padding: '40px',
                width: '1600px',
                height: 'auto',
                overflow: 'visible',
                maxWidth: 'none',
              }
            });
            const linkTopzone = document.createElement('a');
            linkTopzone.download = `KPI_TOPZONE_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.png`;
            linkTopzone.href = topzoneDataUrl;
            linkTopzone.click();
          }
        } catch (error) {
          console.error('Error exporting TOPZONE:', error);
        }

        // Restore original tab
        if (originalTab && originalTab !== 'TOPZONE') {
          handleTabChange(originalTab);
        }
        
        // Reset state
        setExportState('IDLE');
        setOriginalTab(null);
        setIsExporting(false);
      }
    };

    runExport();
  }, [exportState, activeTab, originalTab]);

  const exportWatchlistToImage = async () => {
    if (watchlistRef.current) {
      setIsExporting(true);
      setTimeout(async () => {
        try {
          const dataUrl = await htmlToImage.toPng(watchlistRef.current!, {
            backgroundColor: '#F8F9FA',
            pixelRatio: 3,
            cacheBust: true,
            filter: (node: any) => {
              return node.id !== 'watchlist-export-btn' && node.id !== 'watchlist-copy-btn';
            },
            style: {
              borderRadius: '0',
              margin: '0',
              padding: '40px',
              width: '1600px',
              height: 'auto',
              overflow: 'visible',
              maxWidth: 'none',
            }
          });
          const link = document.createElement('a');
          link.download = `KPI_Watchlist_${activeTab}_${CURRENT_MONTH_DISPLAY.replace('/', '_')}.png`;
          link.href = dataUrl;
          link.click();
        } catch (error) {
          console.error('Error exporting watchlist image:', error);
        } finally {
          setIsExporting(false);
        }
      }, 100);
    }
  };

  const copyWatchlistNames = () => {
    const names = watchedItems.map(item => item.category).join('\n');
    navigator.clipboard.writeText(names).then(() => {
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    });
  };

  const copyItemName = (name: string) => {
    navigator.clipboard.writeText(name).then(() => {
      // Optional: show a small toast or feedback
    });
  };

  const activeSubjectsCount = data.filter(item => item.target > 0 && getEffectiveCurrent(item) > 0).length;

  const deductionStats = useMemo(() => {
    const allSubjects = [
      ...DEDUCTION_SUBJECTS.TGDD.map(s => ({ ...s, supermarket: 'TGDD' as Supermarket })),
      ...DEDUCTION_SUBJECTS.TOPZONE.map(s => ({ ...s, supermarket: 'TOPZONE' as Supermarket }))
    ];

    const stats = allSubjects.map(subject => {
      // Find matching item in data
      // We need to check data for both supermarkets
      const savedData = localStorage.getItem(`kpi_data_${subject.supermarket}`);
      const supermarketData: KPIData[] = savedData ? JSON.parse(savedData) : [];
      
      const item = supermarketData.find(d => d.category.toLowerCase() === subject.category.toLowerCase());
      
      // Use projected completion percentage
      const projectedPercent = item ? ((getEffectiveCurrent(item) / daysUsed) * daysInMonth / item.target) * 100 : 0;
      
      const rule = DEDUCTION_RULES.find(r => r.name === subject.group);
      const isUnder100 = projectedPercent < 100;
      
      let points = 0;
      if (isUnder100 && rule) {
        points = rule.basePoints + rule.penaltyPoints;
      } else if (!isUnder100) {
        points = 0; // Trên 100% thì bằng 0
      }
      
      const manual = deductionManualPoints[`${subject.supermarket}_${subject.category}`] || 0;
      const totalPoints = points + manual;

      return {
        ...subject,
        target: item?.target || 0,
        projectedPercent,
        isUnder100,
        points,
        manual,
        totalPoints,
        id: `${subject.supermarket}_${subject.category}`
      };
    }).filter(s => s.target > 0);

    // Sort by totalPoints descending to assign star deductions
    const sorted = [...stats].sort((a, b) => b.totalPoints - a.totalPoints);
    
    return stats.map(s => {
      const rank = sorted.findIndex(item => item.id === s.id);
      // Mapping: Rank based star deduction
      // We'll keep the previous formula or adjust if needed.
      // The user says "Chạy Theo Công Thức", let's use a rank-based one.
      const starDeduction = (s.totalPoints > 0 && s.projectedPercent < 100) ? Math.max(20, 100 - (rank * 5)) : 0;
      return { ...s, starDeduction };
    });
  }, [data, deductionManualPoints, activeTab, daysUsed]);

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-[#1A1A1A] selection:text-white">
      {/* Sidebar / Navigation Rail */}
      <div className="fixed left-0 top-0 bottom-0 w-16 bg-white border-r border-[#E5E7EB] z-50 flex flex-col items-center py-8 gap-8">
        <div className="w-10 h-10 bg-[#1A1A1A] rounded-xl flex items-center justify-center text-white font-serif italic text-xl">
          K
        </div>
        <div className="flex flex-col gap-4">
          <div className="w-12 h-12 rounded-xl flex flex-col items-center justify-center bg-[#1A1A1A] text-white shadow-lg relative" title="Tất cả">
            <BarChart3 size={18} />
            <span className="text-[8px] font-black mt-0.5">Tất cả</span>
            {activeSubjectsCount > 0 && (
              <div className="absolute -top-1 -right-1 bg-rose-500 text-white text-[8px] font-black px-1.5 py-0.5 rounded-full border-2 border-white shadow-sm">
                {activeSubjectsCount}
              </div>
            )}
          </div>
          <button 
            onClick={() => setShowImport(true)}
            className="w-12 h-12 rounded-xl flex flex-col items-center justify-center text-[#9CA3AF] hover:bg-indigo-50 hover:text-indigo-600 transition-all border border-transparent hover:border-indigo-100"
            title="Nhập dữ liệu Lũy Kế"
          >
            <ClipboardPaste size={18} />
            <span className="text-[8px] font-black mt-0.5">Lũy Kế</span>
          </button>
          <button 
            onClick={() => setShowRealtimeImport(true)}
            className={`w-12 h-12 rounded-xl flex flex-col items-center justify-center transition-all border ${data.some(d => d.realtime !== undefined) ? 'bg-amber-100 text-amber-600 border-amber-200' : 'text-[#9CA3AF] hover:bg-amber-50 hover:text-amber-600 border-transparent hover:border-amber-100'}`}
            title="Nhập Realtime"
          >
            <TrendingUp size={18} />
            <span className="text-[8px] font-black mt-0.5 uppercase">Realtime</span>
          </button>
          <a 
            href={BI_URLS.TGDD} 
            target="_blank" 
            rel="noopener noreferrer"
            className="w-12 h-12 rounded-xl flex flex-col items-center justify-center bg-white border border-[#E5E7EB] text-indigo-600 hover:bg-indigo-50 transition-all shadow-sm"
            title="BI TGDĐ"
          >
            <ExternalLink size={16} />
            <span className="text-[8px] font-black mt-0.5">TGDĐ</span>
          </a>
          <a 
            href={BI_URLS.TOPZONE} 
            target="_blank" 
            rel="noopener noreferrer"
            className="w-12 h-12 rounded-xl flex flex-col items-center justify-center bg-white border border-[#E5E7EB] text-indigo-600 hover:bg-indigo-50 transition-all shadow-sm"
            title="BI Topzone"
          >
            <ExternalLink size={16} />
            <span className="text-[8px] font-black mt-0.5">TZ</span>
          </a>
        </div>
        <div className="mt-auto">
          {/* Bottom space reserved */}
        </div>
      </div>

      {/* Main Content Area */}
      <div className="pl-16 min-h-screen bg-[#F8F9FA]">
        <div className="max-w-6xl mx-auto p-8 space-y-8">
          {/* Header */}
          <header className="bg-white border border-[#E5E7EB] rounded-2xl px-8 py-6 shadow-sm">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div className="flex-1">
                <div className="flex items-center gap-3 mb-1">
                  <h1 className="text-2xl font-bold tracking-tight">Tính Toán Mục Tiêu KPI</h1>
                  <span className="px-2 py-0.5 bg-indigo-50 text-indigo-600 text-xs font-bold rounded-md border border-indigo-100">
                    {CURRENT_MONTH_NAME}
                  </span>
                </div>
                <div className="flex items-center gap-2 mt-2">
                  <div className="flex p-1 bg-gray-100 rounded-xl border border-gray-200">
                    <button 
                      onClick={() => handleTabChange('TGDD')}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${activeTab === 'TGDD' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                      Siêu thị TGDĐ
                    </button>
                    <button 
                      onClick={() => handleTabChange('TOPZONE')}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${activeTab === 'TOPZONE' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                      Siêu thị Topzone
                    </button>
                  </div>
                  <div className="flex p-1 bg-gray-100 rounded-xl border border-gray-200 ml-2">
                    <button 
                      onClick={() => setViewMode('DASHBOARD')}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${viewMode === 'DASHBOARD' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                      Dashboard
                    </button>
                    <button 
                      onClick={() => setViewMode('STAR_PUSH')}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${viewMode === 'STAR_PUSH' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                      Thi Đua Sao
                    </button>
                    <button 
                      onClick={() => setViewMode('STAR_DEDUCTION')}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${viewMode === 'STAR_DEDUCTION' ? 'bg-white text-indigo-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                      Trừ Sao
                    </button>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-4">
                <button 
                  onClick={() => setShowRealtimeImport(true)}
                  className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold transition-all border ${
                    data.some(d => d.realtime !== undefined) 
                    ? 'bg-amber-50 border-amber-200 text-amber-600 shadow-sm' 
                    : 'bg-white border-[#E5E7EB] text-[#6B7280] hover:bg-amber-50 hover:text-amber-600'
                  }`}
                >
                  <TrendingUp size={16} />
                  <span className="hidden sm:inline">Nhập Realtime</span>
                </button>

                <div className="flex flex-col items-end">
                  <span className="text-[10px] font-bold uppercase tracking-widest text-[#9CA3AF]">Đã qua</span>
                  <span className="text-lg font-mono font-bold text-indigo-600">{daysUsed} ngày</span>
                </div>
                <div className="h-8 w-px bg-[#E5E7EB]" />
                <div className="flex flex-col items-end">
                  <span className="text-[10px] font-bold uppercase tracking-widest text-[#9CA3AF]">Còn lại</span>
                  <span className="text-lg font-mono font-bold">{Math.max(1, daysInMonth - daysUsed + 1)} ngày</span>
                </div>
                <div className="h-8 w-px bg-[#E5E7EB]" />
                
                <div className="flex items-center gap-2">
                  <div className="bg-[#F3F4F6] p-1.5 rounded-lg flex items-center gap-2 border border-[#E5E7EB]">
                    <Calendar size={14} className="text-[#6B7280]" />
                    <span className="text-xs font-bold uppercase">Ngày SD:</span>
                    <input 
                      type="number"
                      value={daysUsed}
                      onChange={(e) => setDaysUsed(Number(e.target.value))}
                      className="w-10 bg-transparent text-center font-mono font-bold focus:outline-none"
                    />
                    <span className="text-xs font-bold">/ {daysInMonth}</span>
                  </div>

                  <div className="bg-[#F3F4F6] p-1.5 rounded-lg flex items-center gap-2 border border-[#E5E7EB]">
                    <Settings2 size={14} className="text-[#6B7280]" />
                    <span className="text-xs font-bold uppercase">Mục tiêu:</span>
                    <input 
                      type="number"
                      value={desiredTargetPercent}
                      onChange={(e) => setDesiredTargetPercent(Number(e.target.value))}
                      className="w-12 bg-transparent text-center font-mono font-bold focus:outline-none"
                    />
                    <span className="text-xs font-bold">%</span>
                  </div>

                  <div className="bg-amber-50 p-1.5 rounded-lg flex items-center gap-2 border border-amber-200">
                    <Star size={14} className="text-amber-600" />
                    <span className="text-[10px] font-bold uppercase text-amber-700">Watchlist:</span>
                    <div className="flex items-center gap-1">
                      <input 
                        type="number"
                        value={autoWatchMin}
                        onChange={(e) => setAutoWatchMin(Number(e.target.value))}
                        className="w-10 bg-transparent text-center font-mono font-bold text-amber-900 focus:outline-none"
                      />
                      <span className="text-[10px] text-amber-400">-</span>
                      <input 
                        type="number"
                        value={autoWatchMax}
                        onChange={(e) => setAutoWatchMax(Number(e.target.value))}
                        className="w-10 bg-transparent text-center font-mono font-bold text-amber-900 focus:outline-none"
                      />
                      <span className="text-[10px] font-bold text-amber-700">%</span>
                    </div>
                    <button 
                      onClick={() => {
                        // Trigger a re-render/logic check (though automatic)
                        const currentData = [...data];
                        setData(currentData);
                      }}
                      className="p-1 hover:bg-amber-100 rounded text-amber-600 transition-colors"
                      title="Cập nhật Watchlist"
                    >
                      <TrendingUp size={12} />
                    </button>
                  </div>

                  {/* Highlight Toggle Button */}
                  <button 
                    onClick={() => setEnableHighlight(!enableHighlight)}
                    className={`p-1.5 rounded-lg flex items-center gap-2 border transition-all text-xs font-bold uppercase cursor-pointer select-none ${
                      enableHighlight 
                        ? 'bg-rose-50 border-rose-200 text-rose-700 shadow-sm' 
                        : 'bg-[#F3F4F6] border-[#E5E7EB] text-[#6B7280] hover:bg-gray-100'
                    }`}
                    title="Bật/Tắt Highlight TOP 3 thiếu hụt lớn nhất"
                  >
                    <span className={`w-2 h-2 rounded-full ${enableHighlight ? 'bg-rose-500 animate-pulse' : 'bg-gray-400'}`} />
                    <span>Highlight Top 3 Thiếu</span>
                  </button>
                </div>
              </div>
            </div>
          </header>

          <main className="mt-8">
            {viewMode === 'DASHBOARD' ? (
              <div className="space-y-8">
                {/* Watchlist Section */}
            {watchedItems.length > 0 && (
              <section id="watchlist-section" ref={watchlistRef} className="space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-[#1A1A1A]">
                    <Star size={18} className="fill-amber-400 text-amber-400" />
                    <h2 className="text-sm font-black uppercase tracking-widest">Môn quan tâm</h2>
                  </div>
                  <div className="flex items-center gap-2">
                    <button 
                      id="watchlist-copy-btn"
                      onClick={copyWatchlistNames}
                      className={`flex items-center gap-2 px-3 py-1.5 text-[10px] font-bold transition-all border rounded-lg bg-white ${copySuccess ? 'text-emerald-600 border-emerald-200 bg-emerald-50' : 'text-[#6B7280] hover:text-indigo-600 hover:bg-indigo-50 border-[#E5E7EB]'}`}
                    >
                      {copySuccess ? <Check size={12} /> : <Copy size={12} />}
                      <span>{copySuccess ? 'Đã copy' : 'Copy tên'}</span>
                    </button>
                    <button 
                      id="watchlist-export-btn"
                      onClick={exportWatchlistToImage}
                      className="flex items-center gap-2 px-3 py-1.5 text-[10px] font-bold text-[#6B7280] hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all border border-[#E5E7EB] bg-white"
                    >
                      <ImageIcon size={12} />
                      <span>Xuất ảnh</span>
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-2">
                  {watchedItems.map(item => {
                    const daily = calculateDaily(item);
                    const progress = getCompletionPercent(item);
                    const projected = getProjectedPercent(item);
                    const current = getEffectiveCurrent(item);
                    const missing = Math.max(0, (item.target * desiredTargetPercent / 100) - current);
                    const isTopMissing = topMissingIds.includes(item.id);
                    const topMissingRank = isTopMissing ? topMissingIds.indexOf(item.id) + 1 : 0;
                    return (
                      <motion.div 
                        layoutId={item.id}
                        key={item.id}
                        className={`p-2 rounded-xl border shadow-sm hover:shadow-md transition-shadow relative overflow-hidden group ${
                          isTopMissing 
                            ? topMissingRank === 1
                              ? 'bg-rose-50/80 border-rose-300 ring-2 ring-rose-500/10'
                              : topMissingRank === 2
                              ? 'bg-orange-50/60 border-orange-300'
                              : 'bg-amber-50/60 border-amber-200'
                            : 'bg-white border-[#E5E7EB]'
                        }`}
                      >
                        <button 
                          onClick={() => toggleWatchlist(item.id)}
                          className="absolute top-1.5 right-1.5 text-amber-400 z-10"
                        >
                          <Star size={12} className="fill-current" />
                        </button>
                        <div className="flex items-center gap-1 mb-0.5 flex-wrap">
                          <div className="text-[8px] font-black text-[#9CA3AF] uppercase truncate max-w-[40px]">{item.type}</div>
                          {item.realtime !== undefined && (
                            <span className="text-[7px] font-black bg-amber-100 text-amber-600 px-0.5 rounded uppercase">RT</span>
                          )}
                          {isTopMissing && (
                            <span className={`text-[7px] font-extrabold px-1 rounded uppercase shrink-0 ${
                              topMissingRank === 1 ? 'bg-rose-600 text-white animate-pulse' : topMissingRank === 2 ? 'bg-orange-500 text-white' : 'bg-amber-500 text-amber-950'
                            }`}>TOP {topMissingRank} THIẾU</span>
                          )}
                        </div>
                        <div className={`text-[11px] font-black mb-0.5 pr-4 truncate flex items-center gap-1 ${item.realtime !== undefined ? 'text-blue-900' : ''}`} title={item.category}>
                          <button 
                            onClick={() => copyItemName(item.category)}
                            className="hover:text-indigo-600 transition-colors text-left truncate flex-1"
                            title="Copy tên"
                          >
                            {item.category}
                          </button>
                          <span className="text-[7px] font-black text-slate-400 shrink-0">({item.target.toLocaleString()})</span>
                        </div>
                        
                        {item.realtime !== undefined && (
                          <div className="text-[15px] font-black text-red-600 mb-1 flex items-center gap-1">
                            <span className="opacity-60 text-[10px]">RT:</span>
                            <span>+{item.realtime.toLocaleString()}</span>
                          </div>
                        )}
                        
                        <div className="flex flex-col gap-1 mb-1">
                          <div className="grid grid-cols-2 gap-1">
                            <div className="p-1 bg-indigo-50 rounded-lg border border-indigo-100 flex flex-col items-center">
                              <div className="text-[7px] font-black text-indigo-600 uppercase leading-none mb-0.5">Mục tiêu/Ngày</div>
                              <div className="text-sm font-mono font-black text-indigo-700">
                                {daily > 0 ? daily.toLocaleString(undefined, { maximumFractionDigits: 1 }) : '0'}
                              </div>
                            </div>
                            <div className={`p-1 rounded-lg flex flex-col items-center border ${
                              isTopMissing
                                ? topMissingRank === 1
                                  ? 'bg-rose-600 border-rose-700 text-white shadow-xs'
                                  : topMissingRank === 2
                                  ? 'bg-orange-500 border-orange-600 text-white'
                                  : 'bg-amber-400 border-amber-500 text-amber-950'
                                : 'bg-rose-50 border-rose-100 text-rose-700'
                            }`}>
                              <div className={`text-[7px] font-black uppercase leading-none mb-0.5 ${isTopMissing && topMissingRank < 3 ? 'text-white' : isTopMissing ? 'text-amber-900' : 'text-rose-600'}`}>Còn thiếu</div>
                              <div className={`text-sm font-mono font-black ${isTopMissing && topMissingRank < 3 ? 'text-white' : isTopMissing ? 'text-amber-950' : 'text-rose-700'}`}>
                                {missing.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                              </div>
                            </div>
                          </div>
                        </div>
                        
                        <div className="flex justify-between items-center text-[7px] font-black text-[#9CA3AF] uppercase mb-0.5">
                          <div className="flex gap-1">
                            <span>Tiến độ: {progress.toFixed(0)}%</span>
                            <span className="opacity-40">|</span>
                            <span>Dự kiến: {projected.toFixed(0)}%</span>
                          </div>
                        </div>
                        <div className="w-full h-1 bg-[#F3F4F6] rounded-full overflow-hidden">
                          <div 
                            className={`h-full transition-all duration-500 ${progress >= desiredTargetPercent ? 'bg-emerald-500' : 'bg-indigo-500'}`}
                            style={{ width: `${Math.min(progress, 100)}%` }}
                          />
                        </div>
                      </motion.div>
                    );
                  })}
                </div>
              </section>
            )}

            {/* Main Tracking Section */}
            <div ref={topzoneRef} className="space-y-6 bg-[#F8F9FA] rounded-[2rem] p-8 border border-[#E5E7EB] shadow-sm">
                <div className="flex items-center gap-3 px-2">
                  <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shadow-lg">
                    <TrendingUp size={20} className="text-white" />
                  </div>
                  <h2 className="text-2xl font-black uppercase tracking-tighter text-[#1a1a1a]">
                    Push Số Môn Thi Đua {activeTab === 'TOPZONE' ? 'Topzone' : 'TGDĐ'}
                  </h2>
                </div>

                {/* Analysis Section */}
                {stats && (
                  <section className="grid grid-cols-1 md:grid-cols-5 gap-4">
                <div className="bg-white p-4 rounded-2xl border border-[#E5E7EB] shadow-sm">
                  <div className="text-[10px] font-bold text-[#9CA3AF] uppercase mb-1">Ngành hàng có target</div>
                  <div className="text-2xl font-mono font-black text-slate-700">{filteredData.length}</div>
                </div>
                <div className="bg-white p-4 rounded-2xl border border-[#E5E7EB] shadow-sm">
                  <div className="text-[10px] font-bold text-[#9CA3AF] uppercase mb-1">Số môn về số</div>
                  <div className="text-2xl font-mono font-black text-emerald-600">{stats.completed}</div>
                </div>
                <div className="bg-white p-4 rounded-2xl border border-[#E5E7EB] shadow-sm">
                  <div className="text-[10px] font-bold text-[#9CA3AF] uppercase mb-1">Số môn trên 80%</div>
                  <div className="text-2xl font-mono font-black text-indigo-600">{stats.high}</div>
                </div>
                <div className="bg-white p-4 rounded-2xl border border-[#E5E7EB] shadow-sm">
                  <div className="text-[10px] font-bold text-[#9CA3AF] uppercase mb-1">Số môn dưới 50%</div>
                  <div className="text-2xl font-mono font-black text-rose-600">{stats.low}</div>
                </div>
                <div className="bg-white p-4 rounded-2xl border border-[#E5E7EB] shadow-sm">
                  <div className="text-[10px] font-bold text-[#9CA3AF] uppercase mb-1">Tỉ lệ số môn dự kiến HT</div>
                  <div className="text-2xl font-mono font-black text-amber-600">{stats.successRate.toFixed(1)}%</div>
                </div>
              </section>
            )}



            <section ref={tableRef} className={`bg-white rounded-2xl border border-[#E5E7EB] shadow-sm ${isExporting ? '' : 'overflow-hidden'}`}>
              <div className="p-6 border-b border-[#E5E7EB] flex flex-col md:flex-row justify-between items-center gap-4">
              <div className="flex items-center gap-4 w-full md:w-auto">
                <div className="relative flex-1 md:w-64">
                  <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#9CA3AF]" />
                  <input 
                    type="text"
                    placeholder="Tìm kiếm ngành hàng..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 bg-[#F3F4F6] border-none rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 transition-all"
                  />
                </div>
                <div className="text-xs font-bold text-[#9CA3AF] uppercase">
                  {filteredData.length} ngành hàng có target
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button 
                  onClick={copyMainTableToClipboard}
                  className={`flex items-center gap-2 px-3 py-2 text-xs font-bold transition-all rounded-lg ${copySuccess ? 'text-emerald-600 bg-emerald-50' : 'text-[#6B7280] hover:text-indigo-600 hover:bg-indigo-50'}`}
                  title="Copy Data"
                >
                  {copySuccess ? <Check size={16} /> : <Copy size={16} />}
                  <span className="hidden md:inline">{copySuccess ? 'Đã copy' : 'Copy Data'}</span>
                </button>
                <button 
                  id="main-export-btn"
                  onClick={exportBothSupermarketsToImage}
                  className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-[#6B7280] hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all"
                  title="Xuất ảnh 2 siêu thị"
                >
                  <ImageIcon size={16} />
                  <span className="hidden md:inline">Xuất ảnh 2 siêu thị</span>
                </button>
                <button 
                  id="excel-export-btn"
                  onClick={exportToExcel}
                  className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-[#6B7280] hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                  title="Xuất Excel"
                >
                  <FileSpreadsheet size={16} />
                  <span className="hidden md:inline">Xuất Excel</span>
                </button>
                <button 
                  onClick={() => setData([])}
                  className="p-2 text-[#9CA3AF] hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                  title="Xóa tất cả"
                >
                  <Trash2 size={18} />
                </button>
                {data.some(d => d.realtime !== undefined) && (
                  <button 
                    onClick={clearRealtime}
                    className="flex items-center gap-2 px-3 py-2 text-xs font-bold text-amber-600 hover:bg-amber-50 rounded-lg transition-all"
                    title="Xóa Realtime"
                  >
                    <Trash2 size={16} />
                    <span className="hidden md:inline">Xóa Realtime</span>
                  </button>
                )}
              </div>
            </div>

            <div className={isExporting ? "" : "overflow-x-auto"}>
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-[#F9FAFB] text-[10px] font-bold uppercase tracking-widest text-[#6B7280] border-b border-[#E5E7EB]">
                    <th className="px-3 py-1.5 w-10"></th>
                    <th className="px-3 py-1.5 w-12 text-center">STT</th>
                    <th className="px-3 py-1.5">Ngành hàng</th>
                    <th className="px-3 py-1.5 text-right">Hiện tại</th>
                    <th className="px-3 py-1.5 text-right">Mục tiêu</th>
                    <th className="px-3 py-1.5 text-right">Tiến độ HT</th>
                    <th className="px-3 py-1.5 text-right">% Dự kiến HT</th>
                    <th className="px-3 py-1.5 text-right text-rose-600">Số còn thiếu</th>
                    <th className="px-3 py-1.5 text-right text-indigo-600">Mục tiêu/Ngày</th>
                    <th className="px-3 py-1.5 text-right text-rose-500">Trừ Sao</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB]">
                  <AnimatePresence mode="popLayout">
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="px-6 py-20 text-center">
                          <div className="flex flex-col items-center gap-3 opacity-40">
                            <BarChart3 size={48} />
                            <p className="text-sm font-medium">Chưa có dữ liệu hoặc không có ngành hàng nào có target.</p>
                            <button 
                              onClick={() => setShowImport(true)}
                              className="mt-2 px-4 py-2 bg-[#1A1A1A] text-white rounded-xl text-xs font-bold hover:bg-black transition-all"
                            >
                              Nhập dữ liệu ngay
                            </button>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      filteredData.map((item, index) => {
                        const daily = calculateDaily(item);
                        const progress = getCompletionPercent(item);
                        const projected = getProjectedPercent(item);
                        const isAutoWatched = progress >= autoWatchMin && progress < autoWatchMax;
                        const isWatched = watchlist.includes(item.id) || (isAutoWatched && !manualUnstarred.includes(item.id));
                        const current = getEffectiveCurrent(item);
                        const missing = Math.max(0, (item.target * desiredTargetPercent / 100) - current);
                        const isTopMissing = topMissingIds.includes(item.id);
                        const topMissingRank = isTopMissing ? topMissingIds.indexOf(item.id) + 1 : 0;

                        return (
                          <motion.tr 
                            key={item.id}
                            initial={{ opacity: 0 }}
                            animate={{ opacity: 1 }}
                            exit={{ opacity: 0 }}
                            className={`transition-colors group ${
                              isTopMissing 
                                ? topMissingRank === 1
                                  ? 'bg-rose-50/70 hover:bg-rose-100/70'
                                  : topMissingRank === 2
                                  ? 'bg-orange-50/55 hover:bg-orange-100/55'
                                  : 'bg-amber-50/55 hover:bg-amber-100/55'
                                : 'odd:bg-white even:bg-slate-50/50 hover:bg-indigo-50/40'
                            }`}
                          >
                            <td className="px-3 py-1">
                              <button 
                                onClick={() => toggleWatchlist(item.id)}
                                className={`transition-all ${isWatched ? 'text-amber-400' : 'text-[#E5E7EB] group-hover:text-[#9CA3AF]'}`}
                              >
                                <Star size={16} className={isWatched ? 'fill-current' : ''} />
                              </button>
                            </td>
                            <td className="px-3 py-1 text-center font-mono text-xs text-[#9CA3AF]">
                              {index + 1}
                            </td>
                            <td className="px-3 py-1">
                              <div className="flex flex-col">
                                <div className="flex items-center gap-2 flex-wrap">
                                  <span className={`text-sm font-black ${item.realtime !== undefined ? 'text-blue-900' : ''}`}>{item.category}</span>
                                  {item.realtime !== undefined && (
                                    <span className="text-[8px] font-black bg-amber-100 text-amber-600 px-1 rounded uppercase">Realtime</span>
                                  )}
                                  {isTopMissing && (
                                    <span className={`text-[9px] font-extrabold px-1.5 py-0.5 rounded flex items-center gap-1 uppercase tracking-wider ${
                                      topMissingRank === 1 
                                        ? 'bg-rose-600 text-white animate-pulse' 
                                        : topMissingRank === 2
                                        ? 'bg-orange-500 text-white'
                                        : 'bg-amber-400 text-amber-955'
                                    }`}>
                                      🔥 Top {topMissingRank} Thiếu
                                    </span>
                                  )}
                                </div>
                              </div>
                            </td>
                            <td className="px-3 py-1 text-right font-mono text-base">
                              <div className="flex flex-col items-end">
                                {editingId === item.id ? (
                                  <div className="flex items-center gap-1">
                                    <input 
                                      autoFocus
                                      type="number"
                                      value={editingValue}
                                      onChange={(e) => setEditingValue(e.target.value)}
                                      onBlur={saveEditing}
                                      onKeyDown={(e) => e.key === 'Enter' && saveEditing()}
                                      className="w-20 px-2 py-1 bg-white border border-amber-300 rounded text-right focus:outline-none focus:ring-2 focus:ring-amber-500"
                                    />
                                  </div>
                                ) : (
                                  <button 
                                    onClick={() => startEditing(item)}
                                    className={`hover:bg-amber-50 px-2 py-1 rounded transition-colors text-right w-full flex flex-col items-end ${item.realtime !== undefined ? 'text-amber-600' : ''}`}
                                    title="Click để nhập số Realtime"
                                  >
                                    <span className="font-black text-lg">{current.toLocaleString()}</span>
                                    {item.realtime !== undefined && (
                                      <div className="text-[15px] font-black text-red-600 flex items-center gap-1">
                                        <span>+{item.realtime.toLocaleString()} RT</span>
                                      </div>
                                    )}
                                  </button>
                                )}
                              </div>
                            </td>
                            <td className="px-3 py-1 text-right font-mono text-base text-[#9CA3AF] font-black">
                              {item.target.toLocaleString()}
                            </td>
                            <td className="px-3 py-1 text-right">
                              <span className={`text-sm font-black font-mono ${
                                progress >= desiredTargetPercent ? 'text-emerald-600' : 'text-[#1A1A1A]'
                              }`}>
                                {progress.toFixed(1)}%
                              </span>
                            </td>
                            <td className="px-3 py-1 text-right">
                              <span className={`text-base font-black font-mono ${
                                projected >= 100 ? 'text-emerald-600' : projected >= 80 ? 'text-amber-600' : 'text-rose-600'
                              }`}>
                                {projected.toFixed(1)}%
                              </span>
                            </td>
                            <td className="px-3 py-1 text-right">
                              {isTopMissing ? (
                                <span className={`inline-block text-base font-black font-mono px-2 py-0.5 rounded-lg border shadow-xs ${
                                  topMissingRank === 1 
                                    ? 'text-rose-700 bg-rose-100 border-rose-300 animate-pulse' 
                                    : topMissingRank === 2
                                    ? 'text-orange-700 bg-orange-100 border-orange-200'
                                    : 'text-amber-800 bg-amber-100 border-amber-200'
                                }`}>
                                  {missing.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                                </span>
                              ) : (
                                <span className="text-base font-black font-mono text-rose-600">
                                  {missing.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                                </span>
                              )}
                            </td>
                            <td className="px-3 py-1 text-right">
                              {progress >= desiredTargetPercent ? (
                                <div className="inline-flex items-center gap-1 text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md">
                                  <CheckCircle2 size={12} />
                                  <span className="text-[10px] font-black uppercase">Xong</span>
                                </div>
                              ) : (
                                <div className="flex flex-col items-end">
                                  <span className="text-3xl font-mono font-black text-indigo-600">
                                    {daily.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                                  </span>
                                  <span className="text-[9px] font-black text-[#9CA3AF] uppercase">/ ngày</span>
                                </div>
                              )}
                            </td>
                            <td className="px-3 py-1 text-right">
                              {(() => {
                                const deduction = deductionStats.find(s => s.category.toLowerCase() === item.category.toLowerCase() && s.supermarket === activeTab);
                                return deduction && deduction.starDeduction > 0 ? (
                                  <div className="flex flex-col items-end">
                                    <span className="text-xs font-mono font-black text-rose-600">-{deduction.starDeduction}</span>
                                    <span className="text-[8px] font-bold text-rose-400 uppercase">Sao</span>
                                  </div>
                                ) : (
                                  <span className="text-xs font-mono font-black text-emerald-600">0</span>
                                );
                              })()}
                            </td>
                          </motion.tr>
                        );
                      })
                    )}
                  </AnimatePresence>
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </div>
    ) : viewMode === 'STAR_PUSH' ? (
                /* STAR PUSH VIEW */
                <div className="space-y-6">
          <div className="bg-white p-6 rounded-2xl border border-[#E5E7EB] shadow-sm">
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 bg-amber-400 rounded-xl flex items-center justify-center shadow-lg">
                <Star size={20} className="text-white fill-current" />
              </div>
              <div>
                <h2 className="text-xl font-black uppercase tracking-tighter text-[#1a1a1a]">Thi Đua Sao & Push Số</h2>
                <p className="text-xs font-bold text-[#9CA3AF] uppercase">Phân bổ sao hợp lý để về số cuối tháng</p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {/* Important Items (High Deductions) */}
              <div className="md:col-span-2 space-y-4">
                <h3 className="text-sm font-black uppercase tracking-wider text-rose-600 flex items-center gap-2">
                  <AlertCircle size={16} />
                  Môn quan trọng cần Push (Dễ bị trừ sao cao)
                </h3>
                <div className="grid grid-cols-1 gap-3">
                  {data.filter(item => {
                    if (item.target <= 0) return false;
                    const info = getStarInfo(item);
                    return info && info.config.deduction >= 30 && info.currentPercent < 100;
                  }).sort((a, b) => {
                    const infoA = getStarInfo(a);
                    const infoB = getStarInfo(b);
                    return (infoB?.config.deduction || 0) - (infoA?.config.deduction || 0);
                  }).map(item => {
                    const info = getStarInfo(item)!;
                    const daily = calculateDaily(item);
                    return (
                      <div key={item.id} className="bg-rose-50 border border-rose-100 p-4 rounded-2xl flex justify-between items-center">
                        <div>
                          <div className="text-sm font-black text-rose-900">{item.category}</div>
                          <div className="text-[10px] font-bold text-rose-600 uppercase">Trừ {info.config.deduction} sao nếu không về số</div>
                          <div className="mt-2 flex items-center gap-4">
                            <div className="text-xs font-bold text-slate-500">Hiện tại: <span className="text-slate-900">{info.currentPercent.toFixed(1)}%</span></div>
                            <div className="text-xs font-bold text-slate-500">Cần/Ngày: <span className="text-indigo-600">{daily.toLocaleString()}</span></div>
                          </div>
                        </div>
                        <div className="text-right">
                          <div className="text-[10px] font-bold text-slate-400 uppercase mb-1">Cần thêm để có 1 sao</div>
                          <div className="text-lg font-mono font-black text-rose-700">
                            {Math.max(0, item.target - getEffectiveCurrent(item)).toLocaleString()}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Star Opportunities */}
              <div className="space-y-4">
                <h3 className="text-sm font-black uppercase tracking-wider text-amber-600 flex items-center gap-2">
                  <TrendingUp size={16} />
                  Cơ hội nhận thêm sao
                </h3>
                <div className="space-y-3">
                  {data.filter(item => {
                    if (item.target <= 0) return false;
                    const info = getStarInfo(item);
                    return info && info.nextThreshold !== null && info.currentPercent >= 80;
                  }).sort((a, b) => {
                    const infoA = getStarInfo(a);
                    const infoB = getStarInfo(b);
                    return (infoB?.currentPercent || 0) - (infoA?.currentPercent || 0);
                  }).map(item => {
                    const info = getStarInfo(item)!;
                    return (
                      <div key={item.id} className="bg-amber-50 border border-amber-100 p-3 rounded-xl">
                        <div className="text-xs font-black text-amber-900 truncate">{item.category}</div>
                        <div className="flex justify-between items-end mt-1">
                          <div className="text-[10px] font-bold text-amber-600 uppercase">
                            {info.currentPercent.toFixed(1)}% → {info.nextThreshold}%
                          </div>
                          <div className="text-xs font-black text-amber-700">
                            +{info.nextStars - info.stars} Sao
                          </div>
                        </div>
                        <div className="mt-2 h-1.5 bg-amber-200/50 rounded-full overflow-hidden">
                          <div 
                            className="h-full bg-amber-400" 
                            style={{ width: `${(info.currentPercent / info.nextThreshold!) * 100}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Full Star Table */}
            <div className="mt-12">
              <h3 className="text-sm font-black uppercase tracking-wider text-slate-700 mb-4">Bảng tổng hợp sao dự kiến</h3>
              <div className={`${isExporting ? '' : 'overflow-x-auto'} border border-slate-100 rounded-2xl`}>
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 text-[10px] font-bold uppercase text-slate-500 border-b border-slate-100">
                      <th className="px-4 py-3">Ngành hàng</th>
                      <th className="px-4 py-3 text-right">Tiến độ</th>
                      <th className="px-4 py-3 text-center">Sao hiện tại</th>
                      <th className="px-4 py-3 text-center">Dự kiến trừ</th>
                      <th className="px-4 py-3 text-right">Mục tiêu kế</th>
                      <th className="px-4 py-3 text-right">Cần thêm</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {data.filter(item => item.target > 0).map(item => {
                      const info = getStarInfo(item);
                      if (!info) return null;
                      return (
                        <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                          <td className="px-4 py-3 text-xs font-black text-slate-700">{item.category}</td>
                          <td className="px-4 py-3 text-right font-mono text-xs font-bold">{info.currentPercent.toFixed(1)}%</td>
                          <td className="px-4 py-3 text-center">
                            <div className="inline-flex items-center gap-1 px-2 py-0.5 bg-amber-100 text-amber-700 rounded-full text-[10px] font-black">
                              <Star size={10} className="fill-current" />
                              {info.stars}
                            </div>
                          </td>
                          <td className="px-4 py-3 text-center">
                            {info.deduction > 0 && (
                              <div className="inline-flex items-center gap-1 px-2 py-0.5 bg-rose-100 text-rose-700 rounded-full text-[10px] font-black">
                                -{info.deduction}
                              </div>
                            )}
                          </td>
                          <td className="px-4 py-3 text-right">
                            {info.nextThreshold ? (
                              <span className="text-[10px] font-bold text-slate-500">{info.nextThreshold}% (+{info.nextStars - info.stars} sao)</span>
                            ) : (
                              <span className="text-[10px] font-bold text-emerald-600 uppercase">Max Sao</span>
                            )}
                          </td>
                          <td className="px-4 py-3 text-right font-mono text-xs font-bold text-indigo-600">
                            {info.nextThreshold ? Math.max(0, (item.target * info.nextThreshold / 100) - getEffectiveCurrent(item)).toLocaleString() : '-'}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
          </div>
        ) : (
            <div className="space-y-8">
              {/* Star Deduction View */}
              <div className="bg-white rounded-[2rem] border border-[#E5E7EB] shadow-sm overflow-hidden">
                <div className="p-8 border-b border-[#E5E7EB] bg-[#F9FAFB] flex flex-col md:flex-row justify-between items-center gap-4">
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-12 bg-rose-600 rounded-2xl flex items-center justify-center shadow-lg">
                      <AlertCircle size={24} className="text-white" />
                    </div>
                    <div>
                      <h2 className="text-2xl font-black uppercase tracking-tighter text-[#1A1A1A]">Bảng Tính Trừ Sao</h2>
                      <p className="text-sm font-bold text-[#9CA3AF] uppercase tracking-widest">Dựa trên mức độ hoàn thành và quy tắc thi đua</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    <button 
                      onClick={copyDeductionsToClipboard}
                      className={`flex items-center gap-2 px-4 py-2 text-xs font-bold transition-all border rounded-xl ${copySuccess ? 'text-emerald-600 border-emerald-200 bg-emerald-50' : 'text-[#6B7280] hover:text-indigo-600 hover:bg-indigo-50 border-[#E5E7EB] bg-white'}`}
                    >
                      {copySuccess ? <Check size={14} /> : <Copy size={14} />}
                      <span>{copySuccess ? 'Đã copy' : 'Copy Data'}</span>
                    </button>
                    <button 
                      onClick={exportDeductionsToExcel}
                      className="flex items-center gap-2 px-4 py-2 text-xs font-bold text-[#6B7280] hover:text-emerald-600 hover:bg-emerald-50 rounded-xl transition-all border border-[#E5E7EB] bg-white"
                    >
                      <FileSpreadsheet size={14} />
                      <span>Xuất Excel</span>
                    </button>
                  </div>
                </div>

                <div className="p-8">
                  <div className={isExporting ? "" : "overflow-x-auto"}>
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-[#1A1A1A] text-white">
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20">Siêu Thị</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20">Ngành hàng</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20">Nhóm</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-center bg-rose-500">5 sao</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-center bg-amber-500">10 sao</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-center bg-yellow-400">15 sao</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-center bg-emerald-500">20 sao</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-right">Số Sao bị trừ</th>
                          <th className="px-4 py-3 text-[10px] font-black uppercase tracking-widest border border-white/20 text-center">Cộng Thêm</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#E5E7EB]">
                        {deductionStats.map((stat) => (
                          <tr key={stat.id} className="hover:bg-rose-50/30 transition-colors">
                            <td className="px-4 py-3 text-xs font-bold border border-[#E5E7EB]">{stat.supermarket}</td>
                            <td className="px-4 py-3 text-xs font-black border border-[#E5E7EB]">{stat.category}</td>
                            <td className="px-4 py-3 text-xs font-bold text-[#6B7280] border border-[#E5E7EB]">{stat.group}</td>
                            <td className="px-4 py-3 text-xs font-mono font-bold text-center border border-[#E5E7EB] bg-rose-50">100%</td>
                            <td className="px-4 py-3 text-xs font-mono font-bold text-center border border-[#E5E7EB] bg-amber-50">120%</td>
                            <td className="px-4 py-3 text-xs font-mono font-bold text-center border border-[#E5E7EB] bg-yellow-50">150%</td>
                            <td className="px-4 py-3 text-xs font-mono font-bold text-center border border-[#E5E7EB] bg-emerald-50">250%</td>
                            <td className="px-4 py-3 text-right border border-[#E5E7EB]">
                              <div className="flex flex-col items-end">
                                <div className={`text-sm font-mono font-black ${stat.starDeduction > 0 ? 'text-rose-600' : 'text-emerald-600'}`}>
                                  {stat.starDeduction > 0 ? `-${stat.starDeduction}` : '0'}
                                </div>
                                <div className="text-[8px] font-bold text-[#9CA3AF] uppercase">
                                  {stat.projectedPercent.toFixed(1)}% dự kiến
                                </div>
                              </div>
                            </td>
                            <td className="px-4 py-3 text-center border border-[#E5E7EB]">
                              <input 
                                type="number"
                                value={stat.manual || ''}
                                onChange={(e) => {
                                  const val = parseInt(e.target.value) || 0;
                                  setDeductionManualPoints(prev => ({ ...prev, [stat.id]: val }));
                                }}
                                placeholder="0"
                                className="w-12 px-1 py-1 bg-[#F3F4F6] border border-transparent rounded text-center font-mono font-bold focus:border-rose-500 focus:bg-white transition-all outline-none text-xs"
                              />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="mt-12 p-8 bg-rose-50 rounded-[2rem] border border-rose-100">
                    <div className="flex flex-col md:flex-row items-center justify-between gap-8">
                      <div className="flex items-center gap-4">
                        <div className="w-16 h-16 bg-rose-600 rounded-3xl flex items-center justify-center shadow-xl">
                          <Star size={32} className="text-white fill-current" />
                        </div>
                        <div>
                          <h3 className="text-xl font-black text-rose-900 uppercase">Tổng Sao Bị Trừ</h3>
                          <p className="text-sm font-bold text-rose-400 uppercase">Dựa trên mức điểm cao nhất của các hạng mục</p>
                        </div>
                      </div>
                      <div className="text-center md:text-right">
                        <div className="text-6xl font-mono font-black text-rose-600 tracking-tighter">
                          -{Math.max(...deductionStats.map(s => s.starDeduction))}
                        </div>
                        <div className="text-xs font-black text-rose-400 uppercase tracking-[0.2em] mt-2">Tổng Sao Khấu Trừ</div>
                      </div>
                    </div>
                  </div>

                  <div className="mt-8 grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="p-4 bg-white rounded-2xl border border-[#E5E7EB] flex items-start gap-3">
                      <Info size={20} className="text-indigo-500 shrink-0 mt-0.5" />
                      <p className="text-xs font-bold text-[#6B7280] leading-relaxed">
                        Hệ thống tự động tính điểm dựa trên dữ liệu nhập vào. Nếu % hoàn thành trung bình của các môn trong nhóm &lt; 100%, điểm vi phạm sẽ được cộng vào.
                      </p>
                    </div>
                    <div className="p-4 bg-white rounded-2xl border border-[#E5E7EB] flex items-start gap-3">
                      <Settings2 size={20} className="text-amber-500 shrink-0 mt-0.5" />
                      <p className="text-xs font-bold text-[#6B7280] leading-relaxed">
                        Bạn có thể nhập tay điểm cộng thêm cho từng hạng mục. Dữ liệu này sẽ được lưu lại tự động trên trình duyệt của bạn.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
      </main>
    </div>
  </div>

      {/* Import Modal */}
      <AnimatePresence>
        {showImport && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowImport(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white w-full max-w-4xl rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-8 space-y-6">
                <div className="flex justify-between items-center">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
                      <ClipboardPaste size={20} />
                    </div>
                    <div>
                      <h3 className="text-xl font-bold">Dán dữ liệu mới</h3>
                      <p className="text-sm text-[#6B7280]">Copy bảng dữ liệu từ báo cáo của 2 siêu thị và dán vào đây</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setShowImport(false)}
                    className="text-[#9CA3AF] hover:text-[#1A1A1A]"
                  >
                    Đóng
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-black uppercase text-[#1a1a1a] tracking-wider block">
                      Dữ liệu Siêu thị TGDĐ
                    </label>
                    <textarea 
                      value={pastedTextTgdd}
                      onChange={(e) => setPastedTextTgdd(e.target.value)}
                      placeholder="Dán dữ liệu TGDĐ tại đây... (Ví dụ: Ngành hàng DTQĐ Target...)"
                      className="w-full h-64 p-4 bg-[#F9FAFB] border border-[#E5E7EB] rounded-2xl font-mono text-xs focus:ring-2 focus:ring-indigo-500 focus:border-transparent transition-all resize-none"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black uppercase text-[#1a1a1a] tracking-wider block">
                      Dữ liệu Siêu thị Topzone
                    </label>
                    <textarea 
                      value={pastedTextTopzone}
                      onChange={(e) => setPastedTextTopzone(e.target.value)}
                      placeholder="Dán dữ liệu Topzone tại đây... (Ví dụ: Ngành hàng DTQĐ Target...)"
                      className="w-full h-64 p-4 bg-[#F9FAFB] border border-[#E5E7EB] rounded-2xl font-mono text-xs focus:ring-2 focus:ring-indigo-500 focus:border-transparent transition-all resize-none"
                    />
                  </div>
                </div>

                <div className="flex items-center gap-3 p-4 bg-amber-50 rounded-2xl border border-amber-100">
                  <Info size={18} className="text-amber-600 shrink-0" />
                  <p className="text-[11px] text-amber-800 leading-relaxed">
                    Hệ thống sẽ tự động nhận diện các cột: <strong>Ngành hàng</strong>, <strong>Hiện tại (DT/SL)</strong>, <strong>Target</strong> độc lập cho từng siêu thị.
                  </p>
                </div>

                <button 
                  onClick={handleImport}
                  disabled={!pastedTextTgdd.trim() && !pastedTextTopzone.trim()}
                  className="w-full py-4 bg-[#1A1A1A] text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-black disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                >
                  Xử lý dữ liệu
                  <ArrowRight size={18} />
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Realtime Import Modal */}
      <AnimatePresence>
        {showRealtimeImport && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowRealtimeImport(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white w-full max-w-4xl rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-8 space-y-6">
                <div className="flex justify-between items-center">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
                      <TrendingUp size={20} />
                    </div>
                    <div>
                      <h3 className="text-xl font-bold">Nhập dữ liệu Realtime</h3>
                      <p className="text-sm text-[#6B7280]">Cập nhật số liệu mới nhất của 2 siêu thị để tính toán lại</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setShowRealtimeImport(false)}
                    className="text-[#9CA3AF] hover:text-[#1A1A1A]"
                  >
                    Đóng
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-black uppercase text-[#1a1a1a] tracking-wider block">
                      Realtime Siêu thị TGDĐ
                    </label>
                    <textarea 
                      value={realtimeTextTgdd}
                      onChange={(e) => setRealtimeTextTgdd(e.target.value)}
                      placeholder="Dán dữ liệu realtime TGDĐ tại đây... (Cột 1: Ngành hàng, Cột 2: Số liệu mới)"
                      className="w-full h-64 p-4 bg-[#F9FAFB] border border-[#E5E7EB] rounded-2xl font-mono text-xs focus:ring-2 focus:ring-amber-500 focus:border-transparent transition-all resize-none"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black uppercase text-[#1a1a1a] tracking-wider block">
                      Realtime Siêu thị Topzone
                    </label>
                    <textarea 
                      value={realtimeTextTopzone}
                      onChange={(e) => setRealtimeTextTopzone(e.target.value)}
                      placeholder="Dán dữ liệu realtime Topzone tại đây... (Cột 1: Ngành hàng, Cột 2: Số liệu mới)"
                      className="w-full h-64 p-4 bg-[#F9FAFB] border border-[#E5E7EB] rounded-2xl font-mono text-xs focus:ring-2 focus:ring-amber-500 focus:border-transparent transition-all resize-none"
                    />
                  </div>
                </div>

                <div className="flex items-center gap-3 p-4 bg-indigo-50 rounded-2xl border border-indigo-100">
                  <Info size={18} className="text-indigo-600 shrink-0" />
                  <p className="text-[11px] text-indigo-800 leading-relaxed">
                    Hệ thống sẽ khớp độc lập <strong>Ngành hàng</strong> và cập nhật <strong>Số liệu mới</strong> cho từng siêu thị tương ứng.
                  </p>
                </div>

                <div className="flex gap-4">
                  <button 
                    onClick={clearRealtime}
                    className="flex-1 py-4 bg-white border border-[#E5E7EB] text-[#6B7280] rounded-2xl font-bold hover:bg-gray-50 transition-all"
                  >
                    Xóa Realtime hiện tại
                  </button>
                  <button 
                    onClick={handleRealtimeImport}
                    disabled={!realtimeTextTgdd.trim() && !realtimeTextTopzone.trim()}
                    className="flex-[2] py-4 bg-amber-600 text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-amber-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                  >
                    Cập nhật Realtime
                    <ArrowRight size={18} />
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Floating Info */}
      <div className="fixed bottom-8 right-8 flex flex-col gap-2">
        <div className="bg-white px-4 py-2 rounded-full border border-[#E5E7EB] shadow-lg flex items-center gap-2 text-[10px] font-bold uppercase tracking-widest">
          <Calendar size={12} className="text-indigo-600" />
          <span>Tháng {CURRENT_MONTH_DISPLAY}</span>
        </div>
      </div>
    </div>
  );
}
