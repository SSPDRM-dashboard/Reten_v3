import React, { useState, useMemo, useEffect } from 'react';
import { Printer, FileSpreadsheet, CalendarDays, CalendarRange, Users, Database, RefreshCw, AlertCircle, CheckCircle2, Download, User, FileText, TrendingUp } from 'lucide-react';
import Papa from 'papaparse';
import html2pdf from 'html2pdf.js';
import * as XLSX from 'xlsx';

// --- MOCK DATA FALLBACK ---
const tasksList = [
  'BIT/RONDAAN', 'PEJABAT PERTANYAAN', 'RCJ', 'SENTRI / PROWLER', 'SJR',
  'TUGAS PENTADBIRAN', 'JAGA AMAN / OP SEPADU', 'MESYUARAT / PERJUMPAAN',
  'TUGAS TRAFIK', 'PERHIMPUNAN', 'LDP / KURSUS', 'BILIK KAWALAN (DCC)', 'RCJ (MPV)'
];

const ranksList = ['ASP', 'INSP', 'SI', 'SM', 'SJN', 'KPL', 'L/KPL', 'KONSTABEL'];

const mockDailyData = [
  { id: 1, days: [7, null, 15, null, null, 8, 7, 8, 8, 15, 15, 7, null, 15, 23, 16, 8, 12, 16, 24, 8, 8, 8, null, null, null, null, null, null, null, null] },
  { id: 2, days: [8, 8, null, 8, 4, 4, 4, 4, 8, 8, null, 4, 9, 5, null, 11, 5, null, 8, 24, 11, 16, 8, null, null, 5, null, null, null, null, null] },
  { id: 3, days: Array(31).fill(null) },
  { id: 4, days: [...Array(16).fill(null), 12, ...Array(14).fill(null)] },
  { id: 5, days: Array(31).fill(null) },
  { id: 6, days: [null, null, null, null, null, 6, 8, null, 8, 8, 6, 6, null, null, null, null, 6, null, null, 18, 6, null, 5, 8, 12, 8, 8, 8, null, null, null] },
  ...Array.from({ length: 7 }, (_, i) => ({ id: i + 7, days: Array(31).fill(null) }))
];

const mockWeeklyData = [
  { id: 1, weeks: [22, 105, 98, 68, 0] },
  { id: 2, weeks: [36, 42, 70, 37, 0] },
  { id: 3, weeks: [0, 0, 0, 0, 0] },
  { id: 4, weeks: [0, 0, 12, 0, 0] },
  { id: 5, weeks: [0, 0, 0, 0, 0] },
  { id: 6, weeks: [18, 36, 24, 77, 18] },
  ...Array.from({ length: 7 }, (_, i) => ({ id: i + 7, weeks: [0, 0, 0, 0, 0] }))
];

const mockRankData = [
  { id: 1, ranks: [null, null, null, null, null, 121, 148, 24] },
  { id: 2, ranks: [null, null, null, null, null, 117, 68, null] },
  { id: 3, ranks: [null, null, null, null, null, null, null, null] },
  { id: 4, ranks: [null, null, null, null, null, 8, 4, null] },
  { id: 5, ranks: [null, null, null, null, null, null, null, null] },
  { id: 6, ranks: [24, 36, null, null, null, 77, 36, null] },
  ...Array.from({ length: 7 }, (_, i) => ({ id: i + 7, ranks: Array(8).fill(null) }))
];

const months = [
  'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
  'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'
];

const districts = ['SEMUA DAERAH', 'ALOR GAJAH', 'MELAKA TENGAH', 'JASIN', 'IPK SSPDRM'];
const years = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030];

// ==========================================
// CONFIGURATION
// ==========================================
// PASTE YOUR GOOGLE SHEET ID HERE
// Example: "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
const GOOGLE_SHEET_ID: string = "1mD8nfxGetTY1Xi4o4d471eCFOCDmbEJ_ZclBguqsnMI";

type TabType = 'MONTHLY' | 'PERSONAL' | 'ALLOWANCE' | 'ALLOWANCE_LIVE' | 'PENYALUR_MAKLUMAT' | 'FORECAST';

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [userRole, setUserRole] = useState('');
  const [userTab, setUserTab] = useState('');
  const [userDistrict, setUserDistrict] = useState('');
  const [loggedInName, setLoggedInName] = useState('');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const [activeTab, setActiveTab] = useState<TabType>('MONTHLY');
  const [selectedMonth, setSelectedMonth] = useState('JANUARY');
  const [selectedYear, setSelectedYear] = useState(2026);
  const [selectedYearFrom, setSelectedYearFrom] = useState(2021);
  const [selectedDistrict, setSelectedDistrict] = useState('ALOR GAJAH');
  const [selectedPerson, setSelectedPerson] = useState('ALL');
  const [searchNoBadan, setSearchNoBadan] = useState('');
  const [selectedNoBadanList, setSelectedNoBadanList] = useState<string[]>(Array(10).fill(''));
  const [voucherData, setVoucherData] = useState<any[]>([]);
  const [voucherDataLive, setVoucherDataLive] = useState<any[]>([]);
  const [attendanceDataLive, setAttendanceDataLive] = useState<any[]>([]);
  const [liveDataStatus, setLiveDataStatus] = useState<string>('Initializing...');
  const [maklumatData, setMaklumatData] = useState<any[]>([]);
  const MAKLUMAT_SHEET_ID = "140GxAx-6bU_PQipsPY97Aj1lkETvy3OmUnvbSYbHe1k";

  const fetchMaklumatLive = () => {
    Papa.parse(`https://docs.google.com/spreadsheets/d/${MAKLUMAT_SHEET_ID}/export?format=csv`, {
      download: true,
      header: true,
      complete: (results) => {
        console.log("Maklumat Data Headers:", results.meta.fields);
        console.log("Maklumat Data First Row:", results.data[0]);
        setMaklumatData(results.data);
      },
      error: (error) => {
        console.error("Error fetching maklumat", error);
      }
    });
  };

  useEffect(() => {
    if (activeTab === 'PENYALUR_MAKLUMAT') {
      fetchMaklumatLive();
    }
  }, [activeTab]);

  const [printMode, setPrintMode] = useState<'CURRENT' | 'ALL'>('CURRENT');
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  
  // Google Sheets State
  const [rawData, setRawData] = useState<any[]>([]);
  const [csvFields, setCsvFields] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState('');
  const [lastUpdated, setLastUpdated] = useState<Date | null>(null);

  // Load saved auth on mount
  useEffect(() => {
    const savedAuth = localStorage.getItem('pdrm_auth');
    const savedRole = localStorage.getItem('pdrm_user_role');
    const savedName = localStorage.getItem('pdrm_user_name');
    const savedTab = localStorage.getItem('pdrm_user_tab');
    const savedDistrict = localStorage.getItem('pdrm_user_district');

    if (savedAuth === 'true') {
      setIsAuthenticated(true);
      if (savedRole) setUserRole(savedRole);
      if (savedName) setLoggedInName(savedName);
      if (savedTab) setUserTab(savedTab);
      if (savedDistrict) {
        setUserDistrict(savedDistrict);
        if (savedRole && savedRole.toLowerCase() !== 'admin') {
          setSelectedDistrict(savedDistrict);
        }
      }
      
      if (savedRole && savedRole.toLowerCase() !== 'admin') {
        const tabStr = (savedTab || '').toUpperCase();
        let initialTab: TabType = 'PERSONAL';
        
        if (tabStr.includes('T1') || tabStr.includes('T2') || tabStr.includes('T3') || 
            tabStr.includes('WEEKLY') || tabStr.includes('DAILY') || tabStr.includes('RANK')) {
          initialTab = 'MONTHLY';
        } else if (tabStr.includes('T4') || tabStr.includes('PERSONAL')) {
          initialTab = 'PERSONAL';
        } else if (!tabStr) {
          initialTab = 'PERSONAL';
        }
        
        setActiveTab(initialTab);
      }
    }
  }, []);

  useEffect(() => {
    const defaultVoucherId = '1C7eChL2vbKsk6Yni5rklpx4lBRYQc2kI81V2vGixEa8';
    fetchVoucherData(defaultVoucherId);
  }, []);

  const fetchVoucherData = async (input: string) => {
    if (!input) return;
    let id = input;
    let gid = '';
    
    // Extract ID and GID if it's a full URL
    if (input.includes('docs.google.com/spreadsheets/d/')) {
      const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) id = match[1];
      const gidMatch = input.match(/gid=([0-9]+)/);
      if (gidMatch) gid = `&gid=${gidMatch[1]}`;
    } else if (id === '1C7eChL2vbKsk6Yni5rklpx4lBRYQc2kI81V2vGixEa8') {
      gid = '&gid=761351772';
    }

    try {
      // Extract ID and GID if it's a full URL
      let id = input;
      let gidValue = '761351772';
      
      if (input.includes('docs.google.com/spreadsheets/d/')) {
        const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match) id = match[1];
        const gidMatch = input.match(/gid=([0-9]+)/);
        if (gidMatch) gidValue = gidMatch[1];
      }

      const fetchUrl = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gidValue}`;
      
      const response = await fetch(fetchUrl);
      if (!response.ok) {
        console.error("Voucher fetch failed", response.status);
        return;
      }
      
      const csvText = await response.text();
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          const rows = results.data as any[][];
          if (rows && rows.length > 0) {
            const mapped = rows.map((row, index) => {
              // Skip header rows (usually first 1-2 rows)
              if (index < 2) return null;
              if (!row || row.length < 4) return null;
              
              // Columns D, F, H, J are indices 3, 5, 7, 9
              const noBadan = [row[3], row[5], row[7], row[9]].filter(Boolean).join(' ');
              if (!noBadan || noBadan.toUpperCase().includes('NOMBOR BADAN') || noBadan.toUpperCase().includes('NAMA')) return null;

              return {
                'NO KOD PVR': row[40] || '',
                'NO AKAUN BANK': row[22] || '',
                'NAMA BANK': row[21] || '',
                'NO TELEFON': row[14] || '',
                'No Badan': noBadan
              };
            }).filter(Boolean);
            
            setVoucherData(mapped);
          }
        }
      });
    } catch (e) {
      console.error("Voucher fetch failed", e);
    }
  };

  const fetchVoucherDataLive = async () => {
    setLiveDataStatus('Fetching...');
    
    // 1. Fetch Voucher Metadata (Bank info, etc.)
    const voucherId = '1C7eChL2vbKsk6Yni5rklpx4lBRYQc2kI81V2vGixEa8';
    const voucherGid = '761351772';
    const voucherUrl = `https://docs.google.com/spreadsheets/d/${voucherId}/export?format=csv&gid=${voucherGid}`;
    
    // 2. Fetch Attendance/Hours Data (Page 1 Elaun)
    const attendanceId = '1-suQYCmqWY38qlcniuqrNBLJQxAtrbbB5MWIW61iTP4';
    const attendanceGid = '1963976228';
    const attendanceUrl = `https://docs.google.com/spreadsheets/d/${attendanceId}/export?format=csv&gid=${attendanceGid}`;
    
    try {
      // Fetch Voucher Metadata
      const vResponse = await fetch(voucherUrl, { mode: 'cors' });
      if (vResponse.ok) {
        const vCsv = await vResponse.text();
        if (!vCsv.includes('<!DOCTYPE html>')) {
          Papa.parse(vCsv, {
            header: false,
            skipEmptyLines: true,
            complete: (results) => {
              const rows = results.data as any[][];
              const mapped = rows.map((row, index) => {
                if (index < 2) return null;
                if (!row || row.length < 4) return null;
                const noBadan = [row[3], row[5], row[7], row[9]].filter(Boolean).join(' ');
                if (!noBadan || noBadan.toUpperCase().includes('NOMBOR BADAN') || noBadan.toUpperCase().includes('NAMA')) return null;
                return {
                  'NO KOD PVR': row[40] || '',
                  'NO AKAUN BANK': row[22] || '',
                  'NAMA BANK': row[21] || '',
                  'NO TELEFON': row[14] || '',
                  'No Badan': noBadan,
                  'District': row[1] || '' 
                };
              }).filter(Boolean);
              setVoucherDataLive(mapped);
            }
          });
        }
      }

      // Fetch Attendance Data
      const aResponse = await fetch(attendanceUrl, { mode: 'cors' });
      if (aResponse.ok) {
        const aCsv = await aResponse.text();
        if (!aCsv.includes('<!DOCTYPE html>')) {
          Papa.parse(aCsv, {
            header: false,
            skipEmptyLines: true,
            complete: (results) => {
              const rows = results.data as any[][];
              const mapped = rows.map((row, index) => {
                if (index < 2) return null;
                if (!row || row.length < 4) return null;
                const noBadan = [row[3], row[5], row[7], row[9]].filter(Boolean).join(' ');
                if (!noBadan || noBadan.toUpperCase().includes('NOMBOR BADAN') || noBadan.toUpperCase().includes('NAMA')) return null;
                return {
                  'No Badan': noBadan,
                  'Duty Date': row[15] || '',
                  'Hours': row[18] || '',
                  'District': row[1] || '',
                  'Pangkat': row[14] || ''
                };
              }).filter(Boolean);
              setAttendanceDataLive(mapped);
              setLiveDataStatus(`Loaded ${mapped.length} records`);
            }
          });
        }
      } else {
        setLiveDataStatus(`Fetch failed: ${aResponse.status}`);
      }
    } catch (e) {
      setLiveDataStatus(`Error: Failed to fetch.`);
      console.error("Live Data fetch failed", e);
    }
  };

  useEffect(() => {
    fetchVoucherDataLive();
  }, []);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE") {
      setLoginError('Please set your GOOGLE_SHEET_ID in the code.');
      return;
    }

    if (!username || !password) {
      setLoginError('Please enter both username and password.');
      return;
    }

    setIsLoggingIn(true);
    setLoginError('');

    try {
      const response = await fetch(`https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv&sheet=users`);
      
      if (!response.ok) {
        throw new Error('Failed to fetch users data. Ensure the sheet is public and has a tab named "users".');
      }
      
      const csvText = await response.text();
      
      if (csvText.trim().toLowerCase().startsWith('<!doctype html>') || csvText.trim().toLowerCase().startsWith('<html')) {
         throw new Error('Received an HTML login page instead of data. Please ensure the Google Sheet sharing setting is set to "Anyone with the link".');
      }

      Papa.parse(csvText, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const users = results.data as any[];
          const user = users.find(u => {
            const uName = (u.Username || u.username || u.USERNAME || '').toString().trim();
            const uPass = (u.password || u.Password || u.PASSWORD || '').toString().trim();
            return uName.toLowerCase() === username.trim().toLowerCase() && uPass === password.trim();
          });
          
          if (user) {
            setIsAuthenticated(true);
            
            const getVal = (obj: any, targetKeys: string[]) => {
              const normalizedTargets = targetKeys.map(k => k.toLowerCase().replace(/[^a-z0-9]/g, ''));
              for (const key in obj) {
                const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
                if (normalizedTargets.includes(normalizedKey)) {
                  return obj[key];
                }
              }
              return '';
            };

            const role = getVal(user, ['Role', 'Peranan']).toString().trim();
            const name = getVal(user, ['Name', 'Nama', 'Username']).toString().trim().toUpperCase();
            const defaultTab = getVal(user, ['Tab', 'TabAccess', 'Capaian', 'Tab Access']).toString().trim().toUpperCase();
            const district = getVal(user, ['District', 'Daerah', 'Distric']).toString().trim().toUpperCase();
            
            setUserRole(role);
            setLoggedInName(name);
            setUserTab(defaultTab);
            setUserDistrict(district);
            
            localStorage.setItem('pdrm_auth', 'true');
            localStorage.setItem('pdrm_user_role', role);
            localStorage.setItem('pdrm_user_name', name);
            localStorage.setItem('pdrm_user_tab', defaultTab);
            localStorage.setItem('pdrm_user_district', district);
            
            if (role.toLowerCase() !== 'admin') {
              if (district) setSelectedDistrict(district);
              
              const normalizedTab = defaultTab.toUpperCase();
              let initialTab: TabType = 'PERSONAL';
              
              if (normalizedTab.includes('T1') || normalizedTab.includes('T2') || normalizedTab.includes('T3') || 
                  normalizedTab.includes('WEEKLY') || normalizedTab.includes('DAILY') || normalizedTab.includes('RANK')) {
                initialTab = 'MONTHLY';
              } else if (normalizedTab.includes('T4') || normalizedTab.includes('PERSONAL')) {
                initialTab = 'PERSONAL';
              } else if (!normalizedTab) {
                initialTab = 'PERSONAL';
              }
              
              setActiveTab(initialTab);
            } else {
              setActiveTab('MONTHLY');
              setSelectedPerson('ALL');
            }
            
            setLoginError('');
          } else {
            setLoginError('Invalid username or password');
            console.log("Parsed users:", users);
          }
          setIsLoggingIn(false);
        },
        error: (err: any) => {
          setLoginError('Error parsing users data.');
          setIsLoggingIn(false);
        }
      });
    } catch (err: any) {
      setLoginError(err.message || 'Error connecting to Google Sheets');
      setIsLoggingIn(false);
    }
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    setUserRole('');
    setLoggedInName('');
    setUserTab('');
    setUserDistrict('');
    setSelectedPerson('ALL');
    setSelectedNoBadanList(Array(10).fill(''));
    localStorage.removeItem('pdrm_auth');
    localStorage.removeItem('pdrm_user_role');
    localStorage.removeItem('pdrm_user_name');
    localStorage.removeItem('pdrm_user_tab');
    localStorage.removeItem('pdrm_user_district');
    localStorage.removeItem('pdrm_selected_nobadan');
  };

  // Fetch data when year changes
  useEffect(() => {
    if (isAuthenticated && GOOGLE_SHEET_ID && GOOGLE_SHEET_ID !== "YOUR_GOOGLE_SHEET_ID_HERE") {
      const fromYear = activeTab === 'PERSONAL' ? selectedYearFrom : selectedYear;
      fetchSheetData(GOOGLE_SHEET_ID, fromYear, selectedYear);
    }
  }, [isAuthenticated, selectedYearFrom, selectedYear, activeTab]);

  const fetchSheetData = async (id: string, startYear: number, endYear: number) => {
    if (!id) return;
    setIsLoading(true);
    setError('');
    try {
      // Always fetch years from 2021 up to max of selectedYear/endYear to ensure no cross-tab or historical entries are missed
      const yearFrom = 2021;
      const yearTo = Math.max(startYear, endYear, 2026);
      
      const tabsToFetch: string[] = [];
      for (let y = yearFrom; y <= yearTo; y++) {
        tabsToFetch.push(String(y));
        if (y >= 2022 && y <= 2025) {
          tabsToFetch.push("2" + y);
        }
      }

      const fetchPromises = tabsToFetch.map(tabName => 
        fetch(`https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(tabName)}`)
          .then(res => {
            if (!res.ok) {
              if (tabName === String(selectedYear)) {
                if (res.status === 401 || res.status === 403) {
                  throw new Error('Access Denied. Please ensure the Google Sheet sharing setting is set to "Anyone with the link".');
                }
                throw new Error(`Failed to fetch data. Ensure the sheet is public and has a tab named "${selectedYear}".`);
              }
              return null;
            }
            return res.text();
          })
          .catch(err => {
            if (tabName === String(selectedYear)) throw err;
            console.warn(`Could not fetch data for tab ${tabName}:`, err);
            return null;
          })
      );

      const results = await Promise.all(fetchPromises);
      
      const currentYearIndex = tabsToFetch.indexOf(String(selectedYear));
      const currentYearCsv = currentYearIndex !== -1 ? results[currentYearIndex] : null;
      if (!currentYearCsv) {
        throw new Error(`Failed to fetch data. Ensure the sheet is public and has a tab named "${selectedYear}".`);
      }

      if (currentYearCsv.trim().toLowerCase().startsWith('<!doctype html>') || currentYearCsv.trim().toLowerCase().startsWith('<html')) {
         throw new Error('Received an HTML login page instead of data. Please ensure the Google Sheet sharing setting is set to "Anyone with the link".');
      }

      let masterHeaderRow: any[] | null = null;
      const uniqueDataRowsMap = new Map<string, any[]>();

      results.forEach((csvText) => {
        if (csvText && !csvText.trim().toLowerCase().startsWith('<!doctype html>')) {
          Papa.parse(csvText, {
            header: false,
            skipEmptyLines: true,
            complete: (parseResult) => {
              if (parseResult.data && parseResult.data.length > 0) {
                let currentHeaderRowIndex = -1;
                for (let i = 0; i < Math.min(5, parseResult.data.length); i++) {
                  const r = parseResult.data[i] as any[];
                  if (r && Array.isArray(r) && r.some((c: any) => String(c).toUpperCase().includes('TARIKH')) &&
                      r.some((c: any) => String(c).toUpperCase().includes('DAERAH'))) {
                    currentHeaderRowIndex = i;
                    if (!masterHeaderRow) {
                      masterHeaderRow = r;
                    }
                    break;
                  }
                }

                const dataStartIdx = currentHeaderRowIndex !== -1 ? currentHeaderRowIndex + 1 : 0;

                for (let i = dataStartIdx; i < parseResult.data.length; i++) {
                  const row = parseResult.data[i] as any[];
                  if (!row || !Array.isArray(row) || row.length === 0) continue;

                  if (row.some((c: any) => String(c).toUpperCase().includes('TARIKH')) &&
                      row.some((c: any) => String(c).toUpperCase().includes('DAERAH'))) {
                    continue;
                  }

                  const timestampStr = String(row[0] || '').trim();
                  const dateStr = String(row[15] || '').trim();
                  const hoursStr = String(row[18] || '').trim();
                  
                  let nameStr = '';
                  const nameIndices = [3, 5, 7, 9];
                  nameIndices.forEach(idx => {
                    if (idx < row.length && row[idx]) {
                      nameStr += String(row[idx]).trim().toUpperCase();
                    }
                  });

                  if (!timestampStr && !dateStr && !hoursStr && !nameStr) {
                    continue;
                  }

                  const dedupeKey = `${timestampStr}|${dateStr}|${hoursStr}|${nameStr}`;
                  
                  if (!uniqueDataRowsMap.has(dedupeKey)) {
                    uniqueDataRowsMap.set(dedupeKey, row);
                  }
                }
              }
            }
          });
        }
      });

      const combinedData: any[] = [];
      if (masterHeaderRow) {
        combinedData.push(masterHeaderRow);
      }
      combinedData.push(...uniqueDataRowsMap.values());

      if (combinedData.length === 0) {
        setError('Error parsing CSV data. Check your sheet format.');
      } else {
        setRawData(combinedData);
        setCsvFields([]);
        setLastUpdated(new Date());
      }
      setIsLoading(false);
    } catch (err: any) {
      setError(err.message || 'Failed to connect to Google Sheets');
      setIsLoading(false);
    }
  };

  const [showDebug, setShowDebug] = useState(false);

  // --- DATA PROCESSING ---
  const processedData = useMemo(() => {
    const startY = Math.min(selectedYearFrom, selectedYear);
    const endY = Math.max(selectedYearFrom, selectedYear);

    if (!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE") {
      // Fallback to mock data only if no sheet is connected
      return { daily: mockDailyData, weekly: mockWeeklyData, rank: mockRankData, personal: [], districtPersonnel: [], debugLogs: [] };
    }

    // Initialize empty structures
    const daily = tasksList.map((task, i) => ({ id: i + 1, name: task, days: Array(31).fill(null) }));
    const weekly = tasksList.map((task, i) => ({ id: i + 1, name: task, weeks: Array(5).fill(0) }));
    const rank = tasksList.map((task, i) => ({ id: i + 1, name: task, ranks: Array(8).fill(null) }));
    const personalMap = new Map<string, any>();
    const debugLogs: any[] = [];

    if (!rawData || rawData.length === 0) {
      return { daily, weekly, rank, personal: [], districtPersonnel: [], debugLogs };
    }

    const normalizeStr = (s: string) => (s || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

    // Find header row
    let headerRowIndex = -1;
    let colIndices = {
      date: -1,
      district: -1,
      task: -1,
      hours: -1,
      rank: -1,
      colT: -1
    };

    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;
      
      console.log("Row " + i + ":", row);
      
      const dateIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
      const districtIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
      
      if (dateIdx !== -1 && districtIdx !== -1) {
        headerRowIndex = i;
        colIndices.date = dateIdx;
        colIndices.district = districtIdx;
        
        colIndices.task = row.findIndex(c => String(c).toUpperCase().includes('JENIS TUGASAN'));
        colIndices.hours = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
        
        // Fallback if not found dynamically (assuming they are after the 4 pairs of Balai/Name)
        if (colIndices.task === -1) colIndices.task = 10;
        if (colIndices.hours === -1) colIndices.hours = 11;

        colIndices.rank = row.findIndex(c => String(c).toUpperCase().includes('PANGKAT'));
        if (colIndices.rank === -1) colIndices.rank = 14; // Fallback to Column O

        colIndices.colT = row.findIndex(c => String(c).toUpperCase().includes('NYATAKAN') || String(c).toUpperCase().includes('LAIN-LAIN TUGAS') && !String(c).toUpperCase().includes('JENIS'));
        break;
      }
    }

    if (headerRowIndex === -1) {
      debugLogs.push({ reason: 'Could not find header row with TARIKH and DAERAH' });
      return { daily, weekly, rank, personal: [], districtPersonnel: [], debugLogs };
    }

    const isDistrictMatch = (rowDistrict: any, targetDistrict: string) => {
      if (targetDistrict === 'SEMUA DAERAH') return true;
      if (!rowDistrict || !targetDistrict) return false;
      const s = String(rowDistrict).trim().toUpperCase();
      const t = targetDistrict.trim().toUpperCase();
      
      if (s.includes(t) || t.includes(s)) return true;
      
      // Handle common abbreviations
      if (t === 'ALOR GAJAH' && (s === 'AG' || s.includes('ALOR'))) return true;
      if (t === 'MELAKA TENGAH' && (s === 'MT' || s.includes('TENGAH'))) return true;
      if (t === 'JASIN' && (s === 'JS' || s.includes('JASIN'))) return true;
      if (t === 'IPK SSPDRM' && (s === 'IPK')) return true;
      
      return false;
    };

    // Process data rows
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;

      const dateStr = colIndices.date !== -1 ? row[colIndices.date] : null;
      const districtStr = colIndices.district !== -1 ? row[colIndices.district] : null;
      const rankStr = colIndices.rank !== -1 ? row[colIndices.rank] : null;

      let rowMonth = -1, rowYear = -1, rowDay = -1;
      
      if (dateStr) {
        // Extract date part (ignore time)
        const dateOnly = String(dateStr).split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            // YYYY-MM-DD
            rowYear = parseInt(parts[0], 10);
            rowMonth = parseInt(parts[1], 10) - 1;
            rowDay = parseInt(parts[2], 10);
          } else {
            // DD/MM/YYYY or MM/DD/YYYY
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            
            if (p1 > 12) {
              rowMonth = p0 - 1;
              rowDay = p1;
              rowYear = p2;
            } else {
              rowDay = p0;
              rowMonth = p1 - 1;
              rowYear = p2;
            }
            if (rowYear < 100) rowYear += 2000;
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            rowYear = d.getFullYear();
            rowMonth = d.getMonth();
            rowDay = d.getDate();
          }
        }
      }

      const isYearMatch = rowYear !== -1 && rowYear >= 2021 && rowYear <= 2030;
      
      if (!isYearMatch) continue;

      const rowIsDistrictMatch = isDistrictMatch(districtStr, selectedDistrict);

      let taskStr = colIndices.task !== -1 && colIndices.task < row.length ? row[colIndices.task] : null;
      let hoursStr = colIndices.hours !== -1 && colIndices.hours < row.length ? row[colIndices.hours] : null;

      const totalHours = parseFloat(String(hoursStr)) || 0;
      
      // Construct a comparable date value (YYYYMMDD) for determining latest submission
      const rowDateValue = rowYear * 10000 + (Math.max(0, rowMonth) + 1) * 100 + Math.max(0, rowDay);

      // Personnel data (Tahunan) - Process for everyone in this year to determine their latest district
      const nameIndices = [3, 5, 7, 9];
      nameIndices.forEach(idx => {
        if (idx < row.length && row[idx] && String(row[idx]).trim() !== '') {
          const personName = String(row[idx]).trim().toUpperCase();
          const currentRank = String(rankStr || '').toUpperCase().trim();
          
          // Ultra-aggressive NO. BADAN extraction: find any numeric sequence in name OR rank
          const numMatch = (personName + " " + currentRank).match(/\d+/);
          const finalNoBadan = numMatch ? numMatch[0] : '';
          
          // Ultra-aggressive name normalization for key: Remove EVERYTHING except letters
          const cleanNameLong = personName.replace(finalNoBadan, '').replace(/\s+/g, ' ').trim();
          const cleanNameForKey = cleanNameLong.replace(/[^A-Z]/g, '');
          const normalizedKey = `${finalNoBadan}|${cleanNameForKey}`;
          const balaiStr = idx > 0 && row[idx - 1] ? String(row[idx - 1]).trim().toUpperCase() : '';
          
          if (!personalMap.has(normalizedKey)) {
            const initialYears: any = {};
            years.forEach(y => {
              initialYears[y] = { months: Array(12).fill(0), total: 0, balai: '' };
            });
            if (initialYears[rowYear]) {
              initialYears[rowYear].balai = balaiStr;
            }
            personalMap.set(normalizedKey, {
              name: cleanNameLong,
              rank: currentRank,
              noBadan: finalNoBadan,
              balai: balaiStr,
              years: initialYears,
              months: Array(12).fill(0),
              total: 0,
              latestDistrict: String(districtStr || '').trim().toUpperCase(),
              latestDateValue: rowDateValue,
              districts: new Set([String(districtStr || '').trim().toUpperCase()]),
              districtsByYear: {
                [rowYear]: new Set([String(districtStr || '').trim().toUpperCase()])
              }
            });
          } else {
            const pData = personalMap.get(normalizedKey)!;
            // Prefer name version with spaces if current doesn't have them
            if (cleanNameLong.includes(' ') && !pData.name.includes(' ')) {
              pData.name = cleanNameLong;
            }
            
            if (districtStr) {
              const dStr = String(districtStr).trim().toUpperCase();
              pData.districts.add(dStr);
              if (!pData.districtsByYear) pData.districtsByYear = {};
              if (!pData.districtsByYear[rowYear]) pData.districtsByYear[rowYear] = new Set();
              pData.districtsByYear[rowYear].add(dStr);
            }
            // Update latest district and rank if this row is newer or same date (last row wins)
            if (rowDateValue >= pData.latestDateValue) {
              pData.latestDistrict = String(districtStr || '').trim().toUpperCase();
              pData.latestDateValue = rowDateValue;
              pData.rank = currentRank;
              pData.noBadan = finalNoBadan;
              if (balaiStr) pData.balai = balaiStr;
            }
            if (balaiStr && pData.years[rowYear]) {
              pData.years[rowYear].balai = balaiStr;
            }
          }
          
          const pData = personalMap.get(normalizedKey)!;
          if (rowMonth >= 0 && rowMonth < 12) {
            if (pData.years[rowYear]) {
              pData.years[rowYear].months[rowMonth] += totalHours;
              pData.years[rowYear].total += totalHours;
            }
            if (rowYear === selectedYear && rowIsDistrictMatch) {
              pData.months[rowMonth] += totalHours;
              pData.total += totalHours;
            }
          }
        }
      });

      if (rowYear !== selectedYear) continue;

      // Daily, Weekly, Rank - these MUST still match the district filter
      if (!rowIsDistrictMatch) {
        continue;
      }

      const isMonthMatch = rowMonth !== -1 && months[rowMonth] === selectedMonth;
      if (!isMonthMatch) {
        continue;
      }

      if (!taskStr || !hoursStr || String(taskStr).trim() === '' || String(hoursStr).trim() === '') {
        continue;
      }

      // If task is LAIN-LAIN TUGAS, check Column T for the specific task name
      if (normalizeStr(String(taskStr)) === normalizeStr('LAIN-LAIN TUGAS') && colIndices.colT !== -1 && row.length > colIndices.colT) {
        const colTValue = String(row[colIndices.colT]).trim();
        if (colTValue) {
          // Check if colTValue matches any of our known tasks
          const matchedTask = tasksList.find(t => normalizeStr(t) === normalizeStr(colTValue));
          if (matchedTask) {
            taskStr = matchedTask;
          } else {
            // Some specific mappings based on common entries
            const normalizedColT = normalizeStr(colTValue);
            if (normalizedColT.includes('trafik') || normalizedColT.includes('traffik')) taskStr = 'TUGAS TRAFIK';
            else if (normalizedColT.includes('lalulintas')) taskStr = 'KAWALAN LALULINTAS';
            else if (normalizedColT.includes('mahkamah')) taskStr = 'TUGASAN MAHKAMAH';
            else if (normalizedColT.includes('khas')) taskStr = 'TUGASAN KHAS';
            else if (normalizedColT.includes('mesyuarat')) taskStr = 'MESYUARAT';
            else if (normalizedColT.includes('jagaaman')) taskStr = 'JAGA AMAN';
            else if (normalizedColT.includes('operasi')) taskStr = 'OPERASI';
          }
        }
      }

      const taskIndex = tasksList.findIndex(t => normalizeStr(t) === normalizeStr(String(taskStr)));
      
      if (taskIndex === -1) {
        continue;
      }

      // Daily
      if (rowDay >= 1 && rowDay <= 31) {
        daily[taskIndex].days[rowDay - 1] = (daily[taskIndex].days[rowDay - 1] || 0) + totalHours;
      }

      // Weekly
      const weekIndex = Math.min(Math.floor((rowDay - 1) / 7), 4);
      weekly[taskIndex].weeks[weekIndex] += totalHours;

      // Rank
      if (rankStr) {
        const rawRank = String(rankStr).toUpperCase().replace(/\/SP$/, '').trim();
        let rankIndex = ranksList.findIndex(r => r === rawRank);
        
        if (rankIndex === -1) {
          if (rawRank === 'KONST') {
            rankIndex = ranksList.indexOf('KONSTABEL');
          } else {
            // Fallback fuzzy match
            rankIndex = ranksList.findIndex(r => normalizeStr(r) === normalizeStr(String(rankStr)));
          }
        }
        
        if (rankIndex !== -1) {
          rank[taskIndex].ranks[rankIndex] = (rank[taskIndex].ranks[rankIndex] || 0) + totalHours;
        }
      }
    }

    const getRankPriority = (rank: string) => {
      const normalized = rank.toUpperCase().trim();
      const hierarchy = [
        'SUPT', 'DSP', 'ASP', 'INSP', 'SI', 'SM', 'SJN', 'KPL', 'L/KPL', 'KONST'
      ];
      for (let i = 0; i < hierarchy.length; i++) {
        if (normalized.startsWith(hierarchy[i])) return i;
      }
      return 100;
    };

    const extractNoBadan = (name: string, rank: string) => {
      const nameMatch = name.match(/\d+/);
      if (nameMatch) return parseInt(nameMatch[0], 10);
      const rankMatch = rank.match(/\d+/);
      if (rankMatch) return parseInt(rankMatch[0], 10);
      return 999999;
    };

    const personal = Array.from(personalMap.values())
      .filter(p => {
        // Always include the selected person (or logged in user) regardless of district filter
        if (selectedPerson !== 'ALL' && p.name === selectedPerson) return true;
        if (userRole.toLowerCase() !== 'admin' && (p.name.includes(loggedInName) || loggedInName.includes(p.name))) return true;

        if (searchNoBadan.trim() !== '') {
          const searchTerm = searchNoBadan.trim().toUpperCase();
          return (p.noBadan || '').toUpperCase().includes(searchTerm) || (p.name || '').toUpperCase().includes(searchTerm);
        }

        if (selectedPerson === 'ALL') {
          // Broaden filter: Include if they worked in this district in any of the selected years
          let isMatch = false;
          for (let y = startY; y <= endY; y++) {
            const districtsYear = p.districtsByYear?.[y] || new Set();
            if (Array.from(districtsYear).some(d => isDistrictMatch(d, selectedDistrict))) {
              isMatch = true;
              break;
            }
          }
          if (!isMatch) return false;
        } else {
          // For single person selection, also check any year in the window
          let isMatch = false;
          for (let y = startY; y <= endY; y++) {
            const districtsYear = p.districtsByYear?.[y] || new Set();
            if (Array.from(districtsYear).some(d => isDistrictMatch(d, selectedDistrict))) {
              isMatch = true;
              break;
            }
          }
          if (!isMatch) return false;
        }

        return true;
      })
      .sort((a, b) => {
        const priorityA = getRankPriority(a.rank);
        const priorityB = getRankPriority(b.rank);
        
        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }
        
        const noBadanA = extractNoBadan(a.name, a.rank);
        const noBadanB = extractNoBadan(b.name, b.rank);
        
        if (noBadanA !== noBadanB) {
          return noBadanA - noBadanB;
        }
        
        return a.name.localeCompare(b.name);
      });

    const districtPersonnel = Array.from(personalMap.values())
      .filter(p => {
        // Only checking district for districtPersonnel
        let isMatch = false;
        for (let y = startY; y <= endY; y++) {
          const districtsYear = p.districtsByYear?.[y] || new Set();
          if (Array.from(districtsYear).some(d => isDistrictMatch(d, selectedDistrict))) {
            isMatch = true;
            break;
          }
        }
        return isMatch;
      })
      .sort((a, b) => {
        const priorityA = getRankPriority(a.rank);
        const priorityB = getRankPriority(b.rank);
        
        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }
        
        const noBadanA = extractNoBadan(a.name, a.rank);
        const noBadanB = extractNoBadan(b.name, b.rank);
        
        if (noBadanA !== noBadanB) {
          return noBadanA - noBadanB;
        }
        
        return a.name.localeCompare(b.name);
      });

    return { daily, weekly, rank, personal, districtPersonnel, debugLogs };
  }, [rawData, csvFields, selectedMonth, selectedYear, selectedYearFrom, selectedDistrict, searchNoBadan, selectedPerson, loggedInName, userRole]);

  useEffect(() => {
    if (selectedPerson !== 'ALL' && processedData.districtPersonnel.length > 0) {
      const exists = processedData.districtPersonnel.some(p => p.name === selectedPerson);
      if (!exists && userRole.toLowerCase() === 'admin') {
        setSelectedPerson('ALL');
      }
    }
  }, [selectedDistrict, processedData.districtPersonnel, selectedPerson, userRole]);

  // --- CALCULATIONS ---
  const dailyWithTotals = useMemo(() => {
    return processedData.daily.map((row, index) => {
      const total = row.days.reduce((sum, val) => (sum || 0) + (val || 0), 0);
      return { ...row, name: tasksList[index], total: total === 0 ? null : total };
    });
  }, [processedData.daily]);

  const dailyColumnTotals = useMemo(() => {
    const totals = Array(31).fill(0);
    let grandTotal = 0;
    dailyWithTotals.forEach(row => {
      row.days.forEach((val, idx) => {
        if (val) {
          totals[idx] += val;
          grandTotal += val;
        }
      });
    });
    return { days: totals.map(t => t === 0 ? null : t), grandTotal };
  }, [dailyWithTotals]);

  const weeklyWithTotals = useMemo(() => {
    return processedData.weekly.map((row, index) => {
      const total = row.weeks.reduce((sum, val) => sum + val, 0);
      return { ...row, name: tasksList[index], total };
    });
  }, [processedData.weekly]);

  const weeklyColumnTotals = useMemo(() => {
    const totals = Array(5).fill(0);
    let grandTotal = 0;
    weeklyWithTotals.forEach(row => {
      row.weeks.forEach((val, idx) => {
        totals[idx] += val;
        grandTotal += val;
      });
    });
    return { weeks: totals, grandTotal };
  }, [weeklyWithTotals]);

  const rankWithTotals = useMemo(() => {
    return processedData.rank.map((row, index) => {
      const total = row.ranks.reduce((sum, val) => (sum || 0) + (val || 0), 0);
      return { ...row, name: tasksList[index], total: total === 0 ? null : total };
    });
  }, [processedData.rank]);

  const rankColumnTotals = useMemo(() => {
    const totals = Array(8).fill(0);
    let grandTotal = 0;
    rankWithTotals.forEach(row => {
      row.ranks.forEach((val, idx) => {
        if (val) {
          totals[idx] += val;
          grandTotal += val;
        }
      });
    });
    return { ranks: totals.map(t => t === 0 ? null : t), grandTotal };
  }, [rankWithTotals]);

  // --- PAYMENT BY YEAR SUMMARY CALCULATION ---
  const paymentByYearData = useMemo(() => {
    // Return mock data if no sheet is connected or rawData is empty
    if (!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE" || !rawData || rawData.length === 0) {
      return getSamplePaymentByYear();
    }

    const structure: Record<string, any> = {};
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];
    
    districtsList.forEach(d => {
      structure[d] = {
        PEG: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
        APR: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } }
      };
    });

    // Find header row
    let headerRowIndex = -1;
    let colIndices = {
      date: -1,
      district: -1,
      hours: -1,
      rank: -1
    };

    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;
      
      const dateIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
      const districtIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
      
      if (dateIdx !== -1 && districtIdx !== -1) {
        headerRowIndex = i;
        colIndices.date = dateIdx;
        colIndices.district = districtIdx;
        colIndices.hours = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
        if (colIndices.hours === -1) colIndices.hours = 11;
        colIndices.rank = row.findIndex(c => String(c).toUpperCase().includes('PANGKAT'));
        if (colIndices.rank === -1) colIndices.rank = 14;
        break;
      }
    }

    if (headerRowIndex === -1) {
      return getSamplePaymentByYear();
    }

    const isRowDistrictMatch = (rowDistrict: any, targetDistrict: string) => {
      if (!rowDistrict || !targetDistrict) return false;
      const s = String(rowDistrict).trim().toUpperCase();
      const t = targetDistrict.trim().toUpperCase();
      if (s.includes(t) || t.includes(s)) return true;
      if (t === 'ALOR GAJAH' && (s === 'AG' || s.includes('ALOR'))) return true;
      if (t === 'MELAKA TENGAH' && (s === 'MT' || s.includes('TENGAH'))) return true;
      if (t === 'JASIN' && (s === 'JS' || s.includes('JASIN'))) return true;
      if (t === 'IPK SSPDRM' && (s === 'IPK')) return true;
      return false;
    };

    const personMonthMap = new Map<string, any>();

    // Process data rows
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;

      const dateStr = colIndices.date !== -1 ? row[colIndices.date] : null;
      const districtStr = colIndices.district !== -1 ? row[colIndices.district] : null;
      const rankStr = colIndices.rank !== -1 ? row[colIndices.rank] : null;

      let rowMonth = -1, rowYear = -1, rowDay = -1;
      
      if (dateStr) {
        const dateOnly = String(dateStr).split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            rowYear = parseInt(parts[0], 10);
            rowMonth = parseInt(parts[1], 10) - 1;
            rowDay = parseInt(parts[2], 10);
          } else {
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            if (p1 > 12) {
              rowMonth = p0 - 1;
              rowDay = p1;
              rowYear = p2;
            } else {
              rowDay = p0;
              rowMonth = p1 - 1;
              rowYear = p2;
            }
            if (rowYear < 100) rowYear += 2000;
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            rowYear = d.getFullYear();
            rowMonth = d.getMonth();
            rowDay = d.getDate();
          }
        }
      }

      if (rowYear !== selectedYear) continue;

      let rowDistrict = '';
      for (const d of districtsList) {
        if (isRowDistrictMatch(districtStr, d)) {
          rowDistrict = d;
          break;
        }
      }
      if (!rowDistrict) continue;

      let hoursStr = colIndices.hours !== -1 && colIndices.hours < row.length ? row[colIndices.hours] : null;
      const totalHours = parseFloat(String(hoursStr)) || 0;
      if (totalHours <= 0) continue;

      const nameIndices = [3, 5, 7, 9];
      nameIndices.forEach(idx => {
        if (idx < row.length && row[idx] && String(row[idx]).trim() !== '') {
          const personName = String(row[idx]).trim().toUpperCase();
          const currentRank = String(rankStr || '').toUpperCase().trim();
          
          const numMatch = (personName + " " + currentRank).match(/\d+/);
          const finalNoBadan = numMatch ? numMatch[0] : '';
          const cleanNameLong = personName.replace(finalNoBadan, '').replace(/\s+/g, ' ').trim();
          const cleanNameForKey = cleanNameLong.replace(/[^A-Z]/g, '');
          const normalizedKey = `${finalNoBadan}|${cleanNameForKey}`;
          
          const key = `${normalizedKey}|${rowMonth}`;
          if (!personMonthMap.has(key)) {
            personMonthMap.set(key, {
              name: cleanNameLong,
              rank: currentRank,
              noBadan: finalNoBadan,
              district: rowDistrict,
              month: rowMonth,
              hours: 0
            });
          }
          personMonthMap.get(key).hours += totalHours;
        }
      });
    }

    personMonthMap.forEach((entry) => {
      let h = entry.hours;
      h = Math.min(h, 48); // Cap total duty hours to 48
      
      const r = entry.rank.toUpperCase();
      const isPeg = r.includes('SUPT') || r.includes('DSP') || r.includes('ASP') || r.includes('INSP');
      const category = isPeg ? 'PEG' : 'APR';
      const rate = isPeg ? 9.80 : 8.00;

      const dist = entry.district;
      if (!structure[dist]) return;

      const paidHours = Math.min(h, 48);

      if (h <= 48) {
        structure[dist][category].bracket24_48.bil += 1;
        structure[dist][category].bracket24_48.rm += paidHours * rate;
      } else if (h <= 96) {
        structure[dist][category].bracket49_96.bil += 1;
        structure[dist][category].bracket49_96.rm += paidHours * rate;
      } else {
        structure[dist][category].bracket97_128.bil += 1;
        structure[dist][category].bracket97_128.rm += paidHours * rate;
      }
    });

    districtsList.forEach(d => {
      const peg = structure[d].PEG;
      const apr = structure[d].APR;
      
      const calcTotal = (k: 'bracket1_23' | 'bracket24_48' | 'bracket49_96' | 'bracket97_128') => {
        return {
          bil: peg[k].bil + apr[k].bil,
          rm: peg[k].rm + apr[k].rm
        };
      };

      structure[d].JUMLAH = {
        bracket1_23: calcTotal('bracket1_23'),
        bracket24_48: calcTotal('bracket24_48'),
        bracket49_96: calcTotal('bracket49_96'),
        bracket97_128: calcTotal('bracket97_128')
      };

      ['PEG', 'APR', 'JUMLAH'].forEach(role => {
        const item = structure[d][role];
        item.total = {
          bil: item.bracket24_48.bil,
          rm: item.bracket24_48.rm
        };
      });
    });

    const keseluruhan: Record<string, any> = {
      PEG: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:0, rm:0 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:0, rm:0 } },
      APR: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:0, rm:0 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:0, rm:0 } },
      JUMLAH: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:0, rm:0 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:0, rm:0 } }
    };

    districtsList.forEach(d => {
      ['PEG', 'APR', 'JUMLAH'].forEach(role => {
        const from = structure[d][role];
        if (!from) return;
        const to = keseluruhan[role];

        to.total.bil += from.total.bil;
        to.total.rm += from.total.rm;
      });
    });

    return {
      districts: structure,
      keseluruhan
    };
  }, [rawData, selectedYear]);

  // --- MONTHLY FORECAST SUMMARY CALCULATION ---
  const monthlyForecastData = useMemo(() => {
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];

    // Initialize clean structure
    const structure: Record<string, any> = {};
    districtsList.forEach(d => {
      structure[d] = {
        PEG: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        },
        APR: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        },
        JUMLAH: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        }
      };
    });

    const keseluruhan: Record<string, any> = {
      PEG: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 0, rm: 0 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 0, rm: 0 }
      },
      APR: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 0, rm: 0 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 0, rm: 0 }
      },
      JUMLAH: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 0, rm: 0 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 0, rm: 0 }
      }
    };

    // If no sheet connected or rawData is empty, return a simulated state or mock data
    if (!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE" || !rawData || rawData.length === 0) {
      return getSamplePaymentByYear();
    }

    // Process real data!
    // 1. Find Header
    let headerRowIndex = -1;
    let colIndices = {
      date: -1,
      district: -1,
      hours: -1,
      rank: -1
    };

    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;
      
      const dateIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
      const districtIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
      
      if (dateIdx !== -1 && districtIdx !== -1) {
        headerRowIndex = i;
        colIndices.date = dateIdx;
        colIndices.district = districtIdx;
        colIndices.hours = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
        if (colIndices.hours === -1) colIndices.hours = 11;
        colIndices.rank = row.findIndex(c => String(c).toUpperCase().includes('PANGKAT'));
        if (colIndices.rank === -1) colIndices.rank = 14;
        break;
      }
    }

    if (headerRowIndex === -1) {
      return getSamplePaymentByYear();
    }

    const isRowDistrictMatch = (rowDistrict: any, targetDistrict: string) => {
      if (!rowDistrict || !targetDistrict) return false;
      const s = String(rowDistrict).trim().toUpperCase();
      const t = targetDistrict.trim().toUpperCase();
      if (s.includes(t) || t.includes(s)) return true;
      if (t === 'ALOR GAJAH' && (s === 'AG' || s.includes('ALOR'))) return true;
      if (t === 'MELAKA TENGAH' && (s === 'MT' || s.includes('TENGAH'))) return true;
      if (t === 'JASIN' && (s === 'JS' || s.includes('JASIN'))) return true;
      if (t === 'IPK SSPDRM' && (s === 'IPK')) return true;
      return false;
    };

    const personMonthMap = new Map<string, any>();

    // Process data rows
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;

      const dateStr = colIndices.date !== -1 ? row[colIndices.date] : null;
      const districtStr = colIndices.district !== -1 ? row[colIndices.district] : null;
      const rankStr = colIndices.rank !== -1 ? row[colIndices.rank] : null;

      let rowMonth = -1, rowYear = -1;
      
      if (dateStr) {
        const dateOnly = String(dateStr).split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            rowYear = parseInt(parts[0], 10);
            rowMonth = parseInt(parts[1], 10) - 1;
          } else {
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            if (p1 > 12) {
              rowMonth = p0 - 1;
              rowYear = p2;
            } else {
              rowMonth = p1 - 1;
              rowYear = p2;
            }
            if (rowYear < 100) rowYear += 2000;
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            rowYear = d.getFullYear();
            rowMonth = d.getMonth();
          }
        }
      }

      // Check both Month and Year!
      if (rowYear !== selectedYear) continue;
      const currentMonthIndex = months.indexOf(selectedMonth);
      if (rowMonth !== currentMonthIndex) continue;

      let rowDistrict = '';
      for (const d of districtsList) {
        if (isRowDistrictMatch(districtStr, d)) {
          rowDistrict = d;
          break;
        }
      }
      if (!rowDistrict) continue;

      let hoursStr = colIndices.hours !== -1 && colIndices.hours < row.length ? row[colIndices.hours] : null;
      const totalHours = parseFloat(String(hoursStr)) || 0;
      if (totalHours <= 0) continue;

      const nameIndices = [3, 5, 7, 9];
      nameIndices.forEach(idx => {
        if (idx < row.length && row[idx] && String(row[idx]).trim() !== '') {
          const personName = String(row[idx]).trim().toUpperCase();
          const currentRank = String(rankStr || '').toUpperCase().trim();
          
          const numMatch = (personName + " " + currentRank).match(/\d+/);
          const finalNoBadan = numMatch ? numMatch[0] : '';
          const cleanNameLong = personName.replace(finalNoBadan, '').replace(/\s+/g, ' ').trim();
          const cleanNameForKey = cleanNameLong.replace(/[^A-Z]/g, '');
          const normalizedKey = `${finalNoBadan}|${cleanNameForKey}`;
          
          const key = `${normalizedKey}`;
          if (!personMonthMap.has(key)) {
            personMonthMap.set(key, {
              name: cleanNameLong,
              rank: currentRank,
              noBadan: finalNoBadan,
              district: rowDistrict,
              hours: 0
            });
          }
          personMonthMap.get(key).hours += totalHours;
        }
      });
    }

    personMonthMap.forEach((entry) => {
      let h = entry.hours;
      h = Math.min(h, 24); // Cap total duty hours to 24 so calculation correctly falls into the 24 bracket

      const r = entry.rank.toUpperCase();
      const isPeg = r.includes('SUPT') || r.includes('DSP') || r.includes('ASP') || r.includes('INSP');
      const category = isPeg ? 'PEG' : 'APR';
      const rate = isPeg ? 9.80 : 8.00;

      const dist = entry.district;
      if (!structure[dist]) return;

      const paidHours = Math.min(h, 24); // Keeping Math.min here is redundant but safe

      if (h <= 24) {
        structure[dist][category].bracket24_48.bil += 1;
        structure[dist][category].bracket24_48.rm += paidHours * rate;
      } else if (h <= 96) {
        structure[dist][category].bracket49_96.bil += 1;
        structure[dist][category].bracket49_96.rm += paidHours * rate;
      } else {
        structure[dist][category].bracket97_128.bil += 1;
        structure[dist][category].bracket97_128.rm += paidHours * rate;
      }
    });

    // Populate total and JUMLAH row for each district
    districtsList.forEach(d => {
      ['PEG', 'APR'].forEach(role => {
        const item = structure[d][role];
        
        // Keseluruhan is A + B + C
        item.total = {
          bil: item.bracket24_48.bil + item.bracket49_96.bil + item.bracket97_128.bil,
          rm: item.bracket24_48.rm + item.bracket49_96.rm + item.bracket97_128.rm
        };
      });

      // Calculate JUMLAH = PEG + APR
      const peg = structure[d].PEG;
      const apr = structure[d].APR;
      
      const calcTotal = (k: 'bracket1_23' | 'bracket24_48' | 'bracket49_96' | 'bracket97_128' | 'total') => {
        return {
          bil: peg[k].bil + apr[k].bil,
          rm: peg[k].rm + apr[k].rm
        };
      };

      structure[d].JUMLAH = {
        bracket1_23: calcTotal('bracket1_23'),
        bracket24_48: calcTotal('bracket24_48'),
        bracket49_96: calcTotal('bracket49_96'),
        bracket97_128: calcTotal('bracket97_128'),
        total: calcTotal('total')
      };
    });

    // Calculate Keseluruhan across all districts
    districtsList.forEach(d => {
      ['PEG', 'APR', 'JUMLAH'].forEach(role => {
        const from = structure[d][role];
        const to = keseluruhan[role];

        to.bracket1_23.bil += from.bracket1_23.bil;
        to.bracket1_23.rm += from.bracket1_23.rm;

        to.bracket24_48.bil += from.bracket24_48.bil;
        to.bracket24_48.rm += from.bracket24_48.rm;

        to.bracket49_96.bil += from.bracket49_96.bil;
        to.bracket49_96.rm += from.bracket49_96.rm;

        to.bracket97_128.bil += from.bracket97_128.bil;
        to.bracket97_128.rm += from.bracket97_128.rm;

        to.total.bil += from.total.bil;
        to.total.rm += from.total.rm;
      });
    });

    return {
      districts: structure,
      keseluruhan
    };
  }, [rawData, selectedMonth, selectedYear]);

  // --- YEARLY FORECAST SUMMARY CALCULATION ---
  const yearlyForecastData = useMemo(() => {
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];

    // Initialize clean structure
    const structure: Record<string, any> = {};
    districtsList.forEach(d => {
      structure[d] = {
        PEG: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        },
        APR: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        },
        JUMLAH: {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        }
      };
    });

    const keseluruhan: Record<string, any> = {
      PEG: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
      APR: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
      JUMLAH: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } }
    };

    if (!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE" || !rawData || rawData.length === 0) {
      return getSamplePaymentByYear();
    }

    let headerRowIndex = -1;
    let colIndices = { date: -1, district: -1, hours: -1, rank: -1 };

    for (let i = 0; i < Math.min(20, rawData.length); i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;
      
      const dateIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
      const districtIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
      
      if (dateIdx !== -1 && districtIdx !== -1) {
        headerRowIndex = i;
        colIndices.date = dateIdx;
        colIndices.district = districtIdx;
        colIndices.hours = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
        if (colIndices.hours === -1) colIndices.hours = 11;
        colIndices.rank = row.findIndex(c => String(c).toUpperCase().includes('PANGKAT'));
        if (colIndices.rank === -1) colIndices.rank = 14;
        break;
      }
    }

    if (headerRowIndex === -1) {
      return getSamplePaymentByYear();
    }

    const isRowDistrictMatch = (rowDistrict: any, targetDistrict: string) => {
      if (!rowDistrict || !targetDistrict) return false;
      const s = String(rowDistrict).trim().toUpperCase();
      const t = targetDistrict.trim().toUpperCase();
      if (s.includes(t) || t.includes(s)) return true;
      if (t === 'ALOR GAJAH' && (s === 'AG' || s.includes('ALOR'))) return true;
      if (t === 'MELAKA TENGAH' && (s === 'MT' || s.includes('TENGAH'))) return true;
      if (t === 'JASIN' && (s === 'JS' || s.includes('JASIN'))) return true;
      if (t === 'IPK SSPDRM' && (s === 'IPK')) return true;
      return false;
    };

    // key format: "nobadan|name|month"
    const personMonthMap = new Map<string, any>();

    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!Array.isArray(row)) continue;

      const dateStr = colIndices.date !== -1 ? row[colIndices.date] : null;
      const districtStr = colIndices.district !== -1 ? row[colIndices.district] : null;
      const rankStr = colIndices.rank !== -1 ? row[colIndices.rank] : null;

      let rowMonth = -1, rowYear = -1;
      
      if (dateStr) {
        const dateOnly = String(dateStr).split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            rowYear = parseInt(parts[0], 10);
            rowMonth = parseInt(parts[1], 10) - 1;
          } else {
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            if (p1 > 12) {
              rowMonth = p0 - 1;
              rowYear = p2;
            } else {
              rowMonth = p1 - 1;
              rowYear = p2;
            }
            if (rowYear < 100) rowYear += 2000;
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            rowYear = d.getFullYear();
            rowMonth = d.getMonth();
          }
        }
      }

      if (rowYear !== selectedYear) continue; // ONLY CHECK YEAR
      if (rowMonth === -1) continue; // safety

      let rowDistrict = '';
      for (const d of districtsList) {
        if (isRowDistrictMatch(districtStr, d)) {
          rowDistrict = d;
          break;
        }
      }
      if (!rowDistrict) continue;

      let hoursStr = colIndices.hours !== -1 && colIndices.hours < row.length ? row[colIndices.hours] : null;
      const totalHours = parseFloat(String(hoursStr)) || 0;
      if (totalHours <= 0) continue;

      const nameIndices = [3, 5, 7, 9];
      nameIndices.forEach(idx => {
        if (idx < row.length && row[idx] && String(row[idx]).trim() !== '') {
          const personName = String(row[idx]).trim().toUpperCase();
          const currentRank = String(rankStr || '').toUpperCase().trim();
          
          const numMatch = (personName + " " + currentRank).match(/\d+/);
          const finalNoBadan = numMatch ? numMatch[0] : '';
          const cleanNameLong = personName.replace(finalNoBadan, '').replace(/\s+/g, ' ').trim();
          const cleanNameForKey = cleanNameLong.replace(/[^A-Z]/g, '');
          const normalizedKey = `${finalNoBadan}|${cleanNameForKey}`;
          
          // CRITICAL DIFFERENCE: Key includes month!
          const key = `${normalizedKey}|${rowMonth}`;
          if (!personMonthMap.has(key)) {
            personMonthMap.set(key, {
              name: cleanNameLong,
              rank: currentRank,
              noBadan: finalNoBadan,
              district: rowDistrict,
              month: rowMonth,
              hours: 0
            });
          }
          personMonthMap.get(key).hours += totalHours;
        }
      });
    }

    personMonthMap.forEach((entry) => {
      let h = entry.hours;
      h = Math.min(h, 24); // Cap total duty hours to 24 PER MONTH

      const r = entry.rank.toUpperCase();
      const isPeg = r.includes('SUPT') || r.includes('DSP') || r.includes('ASP') || r.includes('INSP');
      const category = isPeg ? 'PEG' : 'APR';
      const rate = isPeg ? 9.80 : 8.00;

      const dist = entry.district;
      if (!structure[dist]) return;

      const paidHours = Math.min(h, 24);

      if (h <= 24) {
        structure[dist][category].bracket24_48.bil += 1;
        structure[dist][category].bracket24_48.rm += paidHours * rate;
      } else if (h <= 96) {
        structure[dist][category].bracket49_96.bil += 1;
        structure[dist][category].bracket49_96.rm += paidHours * rate;
      } else {
        structure[dist][category].bracket97_128.bil += 1;
        structure[dist][category].bracket97_128.rm += paidHours * rate;
      }
    });

    districtsList.forEach(d => {
      ['PEG', 'APR'].forEach(role => {
        const item = structure[d][role];
        item.total = {
          bil: item.bracket24_48.bil + item.bracket49_96.bil + item.bracket97_128.bil,
          rm: item.bracket24_48.rm + item.bracket49_96.rm + item.bracket97_128.rm
        };
      });

      const peg = structure[d].PEG;
      const apr = structure[d].APR;
      
      const calcTotal = (k: 'bracket1_23' | 'bracket24_48' | 'bracket49_96' | 'bracket97_128' | 'total') => {
        return {
          bil: peg[k].bil + apr[k].bil,
          rm: peg[k].rm + apr[k].rm
        };
      };

      structure[d].JUMLAH = {
        bracket1_23: calcTotal('bracket1_23'),
        bracket24_48: calcTotal('bracket24_48'),
        bracket49_96: calcTotal('bracket49_96'),
        bracket97_128: calcTotal('bracket97_128'),
        total: calcTotal('total')
      };
    });

    districtsList.forEach(d => {
      ['PEG', 'APR', 'JUMLAH'].forEach(role => {
        const from = structure[d][role];
        const to = keseluruhan[role];

        to.bracket1_23.bil += from.bracket1_23.bil;
        to.bracket1_23.rm += from.bracket1_23.rm;

        to.bracket24_48.bil += from.bracket24_48.bil;
        to.bracket24_48.rm += from.bracket24_48.rm;

        to.bracket49_96.bil += from.bracket49_96.bil;
        to.bracket49_96.rm += from.bracket49_96.rm;

        to.bracket97_128.bil += from.bracket97_128.bil;
        to.bracket97_128.rm += from.bracket97_128.rm;

        to.total.bil += from.total.bil;
        to.total.rm += from.total.rm;
      });
    });

    return {
      districts: structure,
      keseluruhan
    };
  }, [rawData, selectedYear]);

  const handlePrint = () => {
    try {
      const result = window.print();
      if (result === undefined) {
        if (window.self !== window.top) {
           setShowPrintModal(true);
        }
      }
    } catch (e) {
      console.error("Print failed:", e);
      setShowPrintModal(true);
    }
  };

  const handlePrintAll = () => {
    setPrintMode('ALL');
    setTimeout(() => {
      try {
        const result = window.print();
        if (result === undefined) {
          if (window.self !== window.top) {
             setShowPrintModal(true);
          }
        }
      } catch (e) {
        console.error("Print failed:", e);
        setShowPrintModal(true);
      }
    }, 500);
  };

  const handlePrintCurrent = () => {
    setPrintMode('CURRENT');
    setTimeout(() => {
      try {
        const result = window.print();
        if (result === undefined) {
          if (window.self !== window.top) {
             setShowPrintModal(true);
          }
        }
      } catch (e) {
        console.error("Print failed:", e);
        setShowPrintModal(true);
      }
    }, 100);
  };

  const handleSaveAllPDF = () => {
    if (isGeneratingPDF) return;
    
    setIsGeneratingPDF(true);
    setPrintMode('ALL');
    
    setTimeout(() => {
      window.scrollTo(0, 0);
      const element = document.getElementById('report-container');
      
      if (!element) {
        setIsGeneratingPDF(false);
        alert("Report container not found.");
        return;
      }
      
      const originalStyle = element.style.overflow;
      element.style.overflow = 'visible';
      
      const opt: any = {
        margin:       [10, 10, 10, 10],
        filename:     `SSPDRM_Full_Report_${selectedMonth}_${selectedYear}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { 
          scale: 2, 
          useCORS: true,
          logging: false,
          letterRendering: true,
          allowTaint: true,
          scrollX: 0,
          scrollY: 0,
          windowWidth: 1400
        },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape', compress: true },
        pagebreak:    { mode: ['css', 'legacy'], avoid: ['.break-inside-avoid', 'tr'] }
      };
 
      try {
        html2pdf().set(opt).from(element).save().then(() => {
          element.style.overflow = originalStyle;
          setIsGeneratingPDF(false);
          setPrintMode('CURRENT');
        }).catch((err: any) => {
          console.error("PDF generation error:", err);
          element.style.overflow = originalStyle;
          setIsGeneratingPDF(false);
          setPrintMode('CURRENT');
        });
      } catch (err) {
        console.error("html2pdf initialization error:", err);
        element.style.overflow = originalStyle;
        setIsGeneratingPDF(false);
        setPrintMode('CURRENT');
      }
    }, 1000);
  };

  const handleSaveCurrentPDF = () => {
    if (isGeneratingPDF) return;
    
    setIsGeneratingPDF(true);
    setPrintMode('CURRENT');
    
    // Give more time for the UI to settle
    setTimeout(() => {
      window.scrollTo(0, 0);
      const element = document.getElementById('report-container');
      
      if (!element) {
        setIsGeneratingPDF(false);
        alert("Report container not found. Please try again.");
        return;
      }
      
      // Temporarily adjust element for better capture
      const originalStyle = element.style.overflow;
      element.style.overflow = 'visible';
      
      const opt: any = {
        margin:       [10, 10, 10, 10],
        filename:     `SSPDRM_${activeTab}_Report_${selectedMonth}_${selectedYear}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { 
          scale: 2, 
          useCORS: true,
          logging: false,
          letterRendering: true,
          allowTaint: true,
          scrollX: 0,
          scrollY: 0,
          windowWidth: 1400 // Ensure a consistent width for scaling
        },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape', compress: true },
        pagebreak:    { mode: ['css', 'legacy'], avoid: ['.break-inside-avoid', 'tr'] }
      };

      try {
        // Correct order: .set(opt).from(element).save()
        html2pdf().set(opt).from(element).save().then(() => {
          element.style.overflow = originalStyle;
          setIsGeneratingPDF(false);
        }).catch((err: any) => {
          console.error("PDF generation error:", err);
          element.style.overflow = originalStyle;
          setIsGeneratingPDF(false);
          alert("PDF generation failed. This usually happens if the report is too large for the browser's memory. \n\nSolution: Click 'Print Current' and choose 'Save as PDF' in the print window.");
        });
      } catch (err) {
        console.error("html2pdf initialization error:", err);
        element.style.overflow = originalStyle;
        setIsGeneratingPDF(false);
        alert("Could not start PDF generation. Please use the 'Print Current' button instead.");
      }
    }, 1000);
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-gray-100 flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-lg shadow-md w-full max-w-md">
          <div className="flex justify-center mb-6">
            <div className="bg-blue-900 p-4 rounded-full">
              <Users className="w-12 h-12 text-white" />
            </div>
          </div>
          <h1 className="text-2xl font-bold text-center text-gray-800 mb-6">SSPDRM Report System</h1>
          
          {(!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE") ? (
            <div className="text-center">
              <p className="text-red-600 font-medium mb-4">
                Please configure your GOOGLE_SHEET_ID in the code to continue.
              </p>
            </div>
          ) : (
            <form onSubmit={handleLogin}>
              <div className="mb-4">
                <label className="block text-gray-700 text-sm font-bold mb-2" htmlFor="username">
                  Username
                </label>
                <input
                  id="username"
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  className="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                  placeholder="Enter username"
                  disabled={isLoggingIn}
                />
              </div>
              <div className="mb-6">
                <label className="block text-gray-700 text-sm font-bold mb-2" htmlFor="password">
                  Password
                </label>
                <input
                  id="password"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                  placeholder="Enter password"
                  disabled={isLoggingIn}
                />
                {loginError && <p className="text-red-500 text-xs italic mt-2">{loginError}</p>}
              </div>
              <div className="flex flex-col gap-3">
                <button
                  type="submit"
                  disabled={isLoggingIn}
                  className="bg-blue-900 hover:bg-blue-800 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline w-full flex justify-center items-center gap-2 disabled:opacity-50"
                >
                  {isLoggingIn ? <RefreshCw className="w-4 h-4 animate-spin" /> : null}
                  {isLoggingIn ? 'Signing In...' : 'Sign In'}
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    );
  }

  // --- RENDERERS ---
  const renderDailyTable = () => (
    <table className="w-full border-collapse border border-black text-xs sm:text-sm text-center font-medium">
      <thead>
        <tr>
          <th className="border border-black p-1 sm:p-2 w-8" rowSpan={2}>Bil</th>
          <th className="border border-black p-1 sm:p-2 text-left min-w-[200px]" rowSpan={2}>PENUGASAN</th>
          <th className="border border-black p-1 sm:p-2" colSpan={31}>DALAM BULAN TERSEBUT</th>
          <th className="border border-black p-1 sm:p-2 w-16 text-[10px] leading-tight" rowSpan={2}>
            JUMLAH<br/>JAM
          </th>
        </tr>
        <tr>
          {Array.from({ length: 31 }, (_, i) => (
            <th key={i + 1} className="border border-black p-1 w-6 sm:w-8 font-bold">
              {i + 1}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {dailyWithTotals.map((row) => (
          <tr key={row.id} className="even:bg-gray-50 print:even:bg-transparent break-inside-avoid">
            <td className="border border-black p-1 font-bold">{row.id}</td>
            <td className="border border-black p-1 text-left font-bold pl-2">{row.name}</td>
            {row.days.map((val, idx) => (
              <td key={idx} className="border border-black p-1 font-bold">
                {val || ''}
              </td>
            ))}
            <td className="border border-black p-1 font-bold bg-gray-50 print:bg-transparent">
              {row.total || ''}
            </td>
          </tr>
        ))}
        <tr className="bg-gray-50 print:bg-transparent">
          <td className="border border-black p-1" colSpan={2}></td>
          {dailyColumnTotals.days.map((val, idx) => (
            <td key={idx} className="border border-black p-1 font-bold text-gray-600 print:text-black">
              {val || ''}
            </td>
          ))}
          <td className="border border-black p-1 font-bold text-blue-600 print:text-black">
            {dailyColumnTotals.grandTotal}
          </td>
        </tr>
      </tbody>
    </table>
  );

  const renderWeeklyTable = () => (
    <table className="w-full border-collapse border border-black text-xs sm:text-sm print:text-[11px] text-center font-medium mt-4">
      <thead>
        <tr>
          <th className="border border-black p-1 print:p-0.5 w-12" rowSpan={2}>BIL</th>
          <th className="border border-black p-1 print:p-0.5 text-left w-48 min-w-[150px]" rowSpan={2}>PROGRAM / AKTIVITI</th>
          <th className="border border-black p-1 print:p-0.5" colSpan={5}>DALAM BULAN TERSEBUT</th>
          <th className="border border-black p-1 print:p-0.5 w-32" rowSpan={2}>JUMLAH JAM<br/>KESELURUHAN</th>
        </tr>
        <tr>
          <th className="border border-black p-1 print:p-0.5 w-24">MINGGU<br/>PERTAMA</th>
          <th className="border border-black p-1 print:p-0.5 w-24">MINGGU<br/>KEDUA</th>
          <th className="border border-black p-1 print:p-0.5 w-24">MINGGU<br/>KETIGA</th>
          <th className="border border-black p-1 print:p-0.5 w-24">MINGGU<br/>KEEMPAT</th>
          <th className="border border-black p-1 print:p-0.5 w-24">MINGGU<br/>KELIMA</th>
        </tr>
      </thead>
      <tbody>
        {weeklyWithTotals.map((row) => (
          <tr key={row.id} className="even:bg-gray-50 print:even:bg-transparent break-inside-avoid">
            <td className="border border-black p-1 print:p-0.5">{row.id}</td>
            <td className="border border-black p-1 print:p-0.5 text-left pl-2">{row.name}</td>
            {row.weeks.map((val, idx) => (
              <td key={idx} className="border border-black p-1 print:p-0.5">
                {val}
              </td>
            ))}
            <td className="border border-black p-1 print:p-0.5 bg-gray-50 print:bg-transparent">
              {row.total}
            </td>
          </tr>
        ))}
        <tr className="bg-gray-50 print:bg-transparent">
          <td className="border border-black p-1 print:p-0.5" colSpan={2}></td>
          {weeklyColumnTotals.weeks.map((val, idx) => (
            <td key={idx} className="border border-black p-1 print:p-0.5 font-bold text-gray-600 print:text-black">
              {val === 0 ? '' : val}
            </td>
          ))}
          <td className="border border-black p-1 print:p-0.5 font-bold text-blue-600 print:text-black">
            {weeklyColumnTotals.grandTotal}
          </td>
        </tr>
      </tbody>
    </table>
  );

  const renderRankTable = () => (
    <table className="w-full border-collapse border border-black text-xs sm:text-sm text-center font-medium">
      <thead>
        <tr>
          <th className="border border-black p-2 w-48 bg-white" colSpan={2}>DAERAH</th>
          <th className="border border-black p-2 bg-[#00b0f0] text-black font-bold" colSpan={9}>
            PENUGASAN PEGAWAI DAN ANGGOTA SSPDRM DAERAH {selectedDistrict} BAGI BULAN
          </th>
        </tr>
        <tr>
          <th className="border border-black p-2 bg-[#ffff00] text-black font-bold text-lg" colSpan={2}>
            {selectedDistrict}
          </th>
          <th className="border border-black p-2 bg-[#d9d9d9] text-black text-lg font-bold" colSpan={9}>
            {selectedMonth} {selectedYear}
          </th>
        </tr>
        <tr className="bg-white">
          <th className="border border-black p-2 w-12">BIL</th>
          <th className="border border-black p-2 text-left w-48 min-w-[150px]">JENIS TUGAS</th>
          <th className="border border-black p-1 w-16">ASP/SP</th>
          <th className="border border-black p-1 w-16">INSP/SP</th>
          <th className="border border-black p-1 w-16">SI/SP</th>
          <th className="border border-black p-1 w-16">SM/SP</th>
          <th className="border border-black p-1 w-16">SJN/SP</th>
          <th className="border border-black p-1 w-16">KPL/SP</th>
          <th className="border border-black p-1 w-16">L/KPL/SP</th>
          <th className="border border-black p-1 w-20">KONST/SP</th>
          <th className="border border-black p-2 w-24 font-bold">JUMLAH</th>
        </tr>
      </thead>
      <tbody>
        {rankWithTotals.map((row) => (
          <tr key={row.id} className="even:bg-gray-50 print:even:bg-transparent break-inside-avoid">
            <td className="border border-black p-1">{row.id}</td>
            <td className="border border-black p-1 text-left pl-2">{row.name}</td>
            {row.ranks.map((val, idx) => (
              <td key={idx} className="border border-black p-1">
                {val || ''}
              </td>
            ))}
            <td className="border border-black p-1 bg-gray-50 print:bg-transparent">
              {row.total || ''}
            </td>
          </tr>
        ))}
        <tr className="bg-gray-50 print:bg-transparent">
          <td className="border border-black p-1" colSpan={2}></td>
          {rankColumnTotals.ranks.map((val, idx) => (
            <td key={idx} className="border border-black p-1 font-bold text-gray-600 print:text-black">
              {val || ''}
            </td>
          ))}
          <td className="border border-black p-1 font-bold text-blue-600 print:text-black">
            {rankColumnTotals.grandTotal}
          </td>
        </tr>
      </tbody>
    </table>
  );

  const renderPersonalTable = () => {
    const monthNames = ['JAN', 'FEB', 'MAC', 'APR', 'MEI', 'JUN', 'JUL', 'OGOS', 'SEPT', 'OKT', 'NOV', 'DIS'];
    
    let displayedPersonnel = processedData.personal;
    
    if (selectedPerson !== 'ALL') {
      displayedPersonnel = processedData.personal.filter(p => p.name === selectedPerson);
      
      // If no exact match and it's the logged in user, try relaxed match (for initial login state)
      if (displayedPersonnel.length === 0 && selectedPerson === loggedInName) {
        displayedPersonnel = processedData.personal.filter(p => 
          p.name.includes(loggedInName) || loggedInName.includes(p.name)
        );
      }
    } else if (userRole.toLowerCase() !== 'admin' && !userTab.toUpperCase().includes('ADMIN')) {
      // If not admin and 'ALL' is selected, we might still want to default to themselves 
      // UNLESS they explicitly chose 'ALL' and we want to allow it.
      // The user said "the data for 'semua anggota' is not showing", so we allow it.
      // No extra filtering here means it shows everyone in processedData.personal (which is district-filtered).
    }

    const monthTotals = Array(12).fill(0);
    let grandTotal = 0;
    
    const startY = Math.min(selectedYearFrom, selectedYear);
    const endY = Math.max(selectedYearFrom, selectedYear);

    const yearsToRender: number[] = [];
    for (let y = startY; y <= endY; y++) {
      yearsToRender.push(y);
    }

    displayedPersonnel.forEach(p => {
      // In "ALL" view, we sum over the selected year range
      yearsToRender.forEach(year => {
        const yearData = p.years && p.years[year] ? p.years[year] : { months: Array(12).fill(0), total: 0 };
        yearData.months.forEach((m: number, i: number) => {
          monthTotals[i] += (m || 0);
          grandTotal += (m || 0);
        });
      });
    });

    if (selectedPerson !== 'ALL' && displayedPersonnel.length > 0) {
      const person = displayedPersonnel[0];
      
      // Extract PANGKAT from Rank field (cleaning it for display)
      let pangkat = person.rank;
      const rankParts = person.rank.split(' ');
      if (rankParts.length > 1) {
        pangkat = rankParts[0];
      } else if (rankParts.length === 1 && /^\d+$/.test(rankParts[0])) {
        pangkat = ''; 
      }

      return (
        <div className="w-full">
          <div className="w-full flex font-bold mb-2 text-[25px] mt-2">
            <div className="w-24 flex-shrink-0"></div>
            <div className="flex-1">
              <div className="grid grid-cols-12 gap-0 mb-1">
                <div className="col-span-8 text-left pl-2">NAMA : {person.name}</div>
                <div className="col-span-4 text-left pl-2">NO.BADAN : {person.noBadan}</div>
              </div>
              <div className="grid grid-cols-12 gap-0 mb-1">
                <div className="col-span-8 text-left pl-2 uppercase">BALAI BERTUGAS : {person.balai || ''}</div>
                <div className="col-span-4 text-left pl-2">PANGKAT : {pangkat}</div>
              </div>
            </div>
            <div className="w-24 flex-shrink-0"></div>
          </div>

          <div className="mt-8">
            <table className="w-full border-collapse border border-black text-center font-bold text-[17px]">
              <thead>
                <tr className="bg-[#135DD8] text-white">
                  <th className="border border-black py-[6.5px] px-2 w-24 bg-[#135DD8]" style={{ backgroundColor: '#135DD8', color: 'white' }}>TAHUN</th>
                  {monthNames.map(m => (
                    <th key={m} className="border border-black py-[6.5px] px-2 w-16 bg-[#135DD8]" style={{ backgroundColor: '#135DD8', color: 'white' }}>{m}</th>
                  ))}
                  <th className="border border-black py-[6.5px] px-2 w-24 bg-[#135DD8]" style={{ backgroundColor: '#135DD8', color: 'white' }}>JUMLAH<br/>JAM</th>
                </tr>
              </thead>
              <tbody>
                {yearsToRender.map((year) => {
                  const yearData = person.years && person.years[year] ? person.years[year] : { months: Array(12).fill(0), total: 0 };
                  return (
                    <tr key={year} className="bg-white">
                      <td className="border border-black py-[6.5px] px-2">{year}</td>
                      {yearData.months.map((hours: number, i: number) => (
                        <td key={i} className="border border-black py-[6.5px] px-2">{hours || 0}</td>
                      ))}
                      <td className="border border-black py-[6.5px] px-2">{yearData.total}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="mt-8 text-left text-[17px] font-bold flex justify-center">
            <div className="w-64">
              <div className="mb-12 text-center">Disahkan oleh</div>
              <div className="border-b border-black w-full mb-2"></div>
              <div className="text-center">(Nama & Jawatan)</div>
              <div className="text-center">Tarikh</div>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="w-full">
        {displayedPersonnel.length === 0 && (userRole.toLowerCase() !== 'admin' || selectedPerson !== 'ALL') && (
          <div className="p-8 text-center text-gray-500 bg-gray-50 rounded-lg border border-dashed border-gray-300 mb-6">
            Tiada data tugasan dijumpai untuk {userRole.toLowerCase() !== 'admin' ? 'anda' : 'anggota ini'} dari {selectedYearFrom} hingga {selectedYear} di daerah {selectedDistrict}.
          </div>
        )}
        
        <table className="w-full border-collapse border border-black text-xs sm:text-sm text-center font-medium">
          <thead>
            <tr className="bg-[#135DD8] text-white">
              <th className="border border-black p-2 w-10">BIL</th>
              <th className="border border-black p-2 w-32">NO. BADAN</th>
              <th className="border border-black p-2 w-24">PANGKAT</th>
              <th className="border border-black p-2 text-left min-w-[200px]">NAMA</th>
              <th className="border border-black p-2 w-32">DAERAH</th>
              <th className="border border-black p-2 w-32">BALAI PENDAFTARAN</th>
              <th className="border border-black p-2 w-20">TAHUN</th>
              {monthNames.map(m => (
                <th key={m} className="border border-black p-1 w-12">{m}</th>
              ))}
              <th className="border border-black p-2 w-20">JUMLAH<br/>JAM</th>
            </tr>
          </thead>
          <tbody>
            {displayedPersonnel.map((person, idx) => {
              // Extract PANGKAT from Rank field (cleaning it for display if it contains No Badan)
              let pangkat = person.rank;
              const rankParts = person.rank.split(' ');
              if (rankParts.length > 1) {
                pangkat = rankParts[0];
              } else if (rankParts.length === 1 && /^\d+$/.test(rankParts[0])) {
                pangkat = ''; 
              }

              return yearsToRender.map((y, yIdx) => {
                const yearData = person.years && person.years[y] ? person.years[y] : { months: Array(12).fill(0), total: 0 };
                
                return (
                  <tr key={`${idx}-${y}`} className="even:bg-gray-50 print:even:bg-transparent break-inside-avoid">
                    {yIdx === 0 && (
                      <>
                        <td className="border border-black p-1" rowSpan={yearsToRender.length}>{idx + 1}</td>
                        <td className="border border-black p-1" rowSpan={yearsToRender.length}>{person.noBadan}</td>
                        <td className="border border-black p-1" rowSpan={yearsToRender.length}>{pangkat}</td>
                        <td className="border border-black p-1 text-left pl-2" rowSpan={yearsToRender.length}>{person.name}</td>
                        <td className="border border-black p-1 uppercase" rowSpan={yearsToRender.length}>{person.latestDistrict || selectedDistrict}</td>
                      </>
                    )}
                    <td className="border border-black p-1">{yearData.balai || person.balai || ''}</td>
                    <td className="border border-black p-1">{y}</td>
                    {yearData.months.map((hours: number, i: number) => (
                      <td key={i} className="border border-black p-1">{hours || 0}</td>
                    ))}
                    <td className="border border-black p-1 font-bold bg-gray-50 print:bg-transparent">{yearData.total}</td>
                  </tr>
                );
              });
            })}
            {displayedPersonnel.length === 0 && (
              <tr>
                <td colSpan={20} className="border border-black p-4 text-gray-500">Tiada rekod anggota dijumpai untuk tempoh ini</td>
              </tr>
            )}
          </tbody>
        </table>
        
        {/* Signature Area */}
        <div className="mt-16 text-sm font-bold print:block">
          <div className="w-64">
            <div className="mb-16 text-center">Disahkan oleh</div>
            <div className="border-b border-black w-full mb-2"></div>
            <div className="text-center">(Nama & Jawatan)</div>
            <div className="text-center">Tarikh</div>
          </div>
        </div>
      </div>
    );
  };

  const exportToWord = () => {
    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML to Word</title><style>body { font-family: Arial, sans-serif; } h1 { font-family: 'Times New Roman', serif; } table { border-collapse: collapse; width: 100%; } .td-main { border: 1px solid black; padding: 6px 12px; font-size: 13px; } .td-opt { border: none; padding: 2px 4px; font-size: 13px; } .font-bold { font-weight: bold; } .text-center { text-align: center; } .uppercase { text-transform: uppercase; } .bg-gray { background-color: #f3f4f6; } </style></head><body>";
    const footer = "</body></html>";
    let rawHTML = document.getElementById("maklumat-container")?.innerHTML || "";
    
    // Replace the marker with MS Word specific page break
    rawHTML = rawHTML.replace(/<div class="page-break-mark".*?<\/div>/g, '<br clear="all" style="mso-special-character:line-break;page-break-before:always" />');
    
    const sourceHTML = header + rawHTML + footer;
    
    const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
    const fileDownload = document.createElement("a");
    document.body.appendChild(fileDownload);
    fileDownload.href = source;
    fileDownload.download = `Borang_Maklumat_${selectedMonth}_${selectedYear}.doc`;
    fileDownload.click();
    document.body.removeChild(fileDownload);
  };

  const exportPersonalToExcel = () => {
    const monthNames = ['JAN', 'FEB', 'MAC', 'APR', 'MEI', 'JUN', 'JUL', 'OGOS', 'SEPT', 'OKT', 'NOV', 'DIS'];
    
    let displayedPersonnel = processedData.personal;
    
    if (selectedPerson !== 'ALL') {
      displayedPersonnel = processedData.personal.filter(p => p.name === selectedPerson);
      if (displayedPersonnel.length === 0 && selectedPerson === loggedInName) {
        displayedPersonnel = processedData.personal.filter(p => 
          p.name.includes(loggedInName) || loggedInName.includes(p.name)
        );
      }
    }

    const startY = Math.min(selectedYearFrom, selectedYear);
    const endY = Math.max(selectedYearFrom, selectedYear);
    const yearsToRender: number[] = [];
    for (let y = startY; y <= endY; y++) {
      yearsToRender.push(y);
    }

    const wb = XLSX.utils.book_new();
    const dataRows: any[][] = [];

    // Header title rows
    dataRows.push(["SUKARELAWAN SIMPANAN POLIS DIRAJA MALAYSIA (SSPDRM)"]);
    dataRows.push([`KONTINJEN : MELAKA`]);
    dataRows.push([`DAERAH : ${selectedDistrict}`]);
    const yearRangeStr = selectedYearFrom !== selectedYear ? `${startY} - ${endY}` : `${selectedYear}`;
    dataRows.push([`JUMLAH JAM PENUGASAN BULANAN BAGI TAHUN ${yearRangeStr}`]);
    dataRows.push([]); // spacer row

    if (selectedPerson !== 'ALL' && displayedPersonnel.length > 0) {
      // Single person view
      const person = displayedPersonnel[0];
      let pangkat = person.rank;
      const rankParts = person.rank.split(' ');
      if (rankParts.length > 1) {
        pangkat = rankParts[0];
      } else if (rankParts.length === 1 && /^\d+$/.test(rankParts[0])) {
        pangkat = ''; 
      }

      dataRows.push([`NAMA: ${person.name}`, "", "", `NO. BADAN: ${person.noBadan}`]);
      dataRows.push([`BALAI BERTUGAS: ${person.balai || ''}`, "", "", `PANGKAT: ${pangkat}`]);
      dataRows.push([]); // spacer row

      // Table Header
      dataRows.push([
        "TAHUN",
        ...monthNames,
        "JUMLAH JAM"
      ]);

      // Table Body
      yearsToRender.forEach((year) => {
        const yearData = person.years && person.years[year] ? person.years[year] : { months: Array(12).fill(0), total: 0 };
        dataRows.push([
          year,
          ...yearData.months.map(m => m || 0),
          yearData.total
        ]);
      });

      const ws = XLSX.utils.aoa_to_sheet(dataRows);

      // Merge titles and personnel info
      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 13 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 13 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 13 } },
        { s: { r: 3, c: 0 }, e: { r: 3, c: 13 } },
        { s: { r: 5, c: 0 }, e: { r: 5, c: 2 } }, // NAMA merge
        { s: { r: 5, c: 3 }, e: { r: 5, c: 5 } }, // NO BADAN merge
        { s: { r: 6, c: 0 }, e: { r: 6, c: 2 } }, // BALAI merge
        { s: { r: 6, c: 3 }, e: { r: 6, c: 5 } }  // PANGKAT merge
      ];

      ws['!cols'] = [
        { wch: 12 }, // TAHUN
        ...Array(12).fill({ wch: 8 }), // months JAN-DIS
        { wch: 15 } // JUMLAH JAM
      ];

      XLSX.utils.book_append_sheet(wb, ws, "Jam Penugasan");
      XLSX.writeFile(wb, `Jam_Penugasan_Tahunan_${person.name.replace(/\s+/g, '_')}_${yearRangeStr}.xlsx`);
    } else {
      // Multiple personnel / all view
      // Table Header
      dataRows.push([
        "BIL",
        "NO. BADAN",
        "PANGKAT",
        "NAMA",
        "DAERAH",
        "BALAI PENDAFTARAN",
        "TAHUN",
        ...monthNames,
        "JUMLAH JAM"
      ]);

      let excelRowIndex = dataRows.length; // 1-based indexing for merges, after title blocks
      const merges: any[] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 19 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 19 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 19 } },
        { s: { r: 3, c: 0 }, e: { r: 3, c: 19 } }
      ];

      displayedPersonnel.forEach((person, idx) => {
        let pangkat = person.rank;
        const rankParts = person.rank.split(' ');
        if (rankParts.length > 1) {
          pangkat = rankParts[0];
        } else if (rankParts.length === 1 && /^\d+$/.test(rankParts[0])) {
          pangkat = ''; 
        }

        yearsToRender.forEach((y, yIdx) => {
          const yearData = person.years && person.years[y] ? person.years[y] : { months: Array(12).fill(0), total: 0 };
          
          dataRows.push([
            yIdx === 0 ? idx + 1 : "",
            yIdx === 0 ? person.noBadan : "",
            yIdx === 0 ? pangkat : "",
            yIdx === 0 ? person.name : "",
            yIdx === 0 ? (person.latestDistrict || selectedDistrict) : "",
            yearData.balai || person.balai || '',
            y,
            ...yearData.months.map(m => m || 0),
            yearData.total
          ]);
        });

        // Add merges for the current person's multi-year row cells
        if (yearsToRender.length > 1) {
          const startR = excelRowIndex;
          const endR = excelRowIndex + yearsToRender.length - 1;
          merges.push({ s: { r: startR, c: 0 }, e: { r: endR, c: 0 } }); // BIL
          merges.push({ s: { r: startR, c: 1 }, e: { r: endR, c: 1 } }); // NO. BADAN
          merges.push({ s: { r: startR, c: 2 }, e: { r: endR, c: 2 } }); // PANGKAT
          merges.push({ s: { r: startR, c: 3 }, e: { r: endR, c: 3 } }); // NAMA
          merges.push({ s: { r: startR, c: 4 }, e: { r: endR, c: 4 } }); // DAERAH
        }
        excelRowIndex += yearsToRender.length;
      });

      const ws = XLSX.utils.aoa_to_sheet(dataRows);
      ws['!merges'] = merges;

      ws['!cols'] = [
        { wch: 6 },  // BIL
        { wch: 15 }, // NO. BADAN
        { wch: 12 }, // PANGKAT
        { wch: 30 }, // NAMA
        { wch: 15 }, // DAERAH
        { wch: 25 }, // BALAI PENDAFTARAN
        { wch: 10 }, // TAHUN
        ...Array(12).fill({ wch: 7 }), // JAN-DIS
        { wch: 15 }  // JUMLAH JAM
      ];

      XLSX.utils.book_append_sheet(wb, ws, "Jam Penugasan Semua");
      XLSX.writeFile(wb, `Jam_Penugasan_Tahunan_Semua_${selectedDistrict}_${yearRangeStr}.xlsx`);
    }
  };

  const exportYearlyToExcel = () => {
    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const dataRows: any[][] = [];
    
    // Title Banner rows
    dataRows.push(["KONTINJEN : MELAKA"]);
    dataRows.push([`TUNTUTAN ELAUN PENUGASAN SSPDRM - SEPANJANG TAHUN ${selectedYear}`]);
    dataRows.push([]); // spacer
    
    // Header Row 1
    dataRows.push([
      "BIL",
      "FORMASI",
      "PANGKAT",
      "JUMLAH JAM BERTUGAS",
      "24 (A)", "",
      "49 - 96 (B)", "",
      "97 - 128 (C)", "",
      "KESELURUHAN (A+B+C)", ""
    ]);
    
    // Header Row 2
    dataRows.push([
      "",
      "",
      "",
      "1 - 23 JAM BIL",
      "BIL", "RM",
      "BIL", "RM",
      "BIL", "RM",
      "Jumlah BIL", "Jumlah RM"
    ]);
    
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];
    let bilIndex = 1;
    
    districtsList.forEach((district) => {
      const dData = yearlyForecastData.districts[district] || {
        PEG: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
        APR: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
        JUMLAH: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } }
      };
      
      const roles = ['PEG', 'APR', 'JUMLAH'];
      roles.forEach((role, rIdx) => {
        const rData = dData[role] || {
          bracket1_23: { bil: 0, rm: 0 },
          bracket24_48: { bil: 0, rm: 0 },
          bracket49_96: { bil: 0, rm: 0 },
          bracket97_128: { bil: 0, rm: 0 },
          total: { bil: 0, rm: 0 }
        };
        
        dataRows.push([
          rIdx === 0 ? bilIndex : "",
          rIdx === 0 ? district : "",
          role,
          rData.bracket1_23.bil,
          rData.bracket24_48.bil,
          rData.bracket24_48.rm,
          rData.bracket49_96.bil,
          rData.bracket49_96.rm,
          rData.bracket97_128.bil,
          rData.bracket97_128.rm,
          rData.total.bil,
          rData.total.rm
        ]);
      });
      bilIndex++;
    });
    
    // Overall total rows
    const rolesOverall = ['PEG', 'APR', 'JUMLAH'];
    rolesOverall.forEach((role, idx) => {
      const rData = yearlyForecastData.keseluruhan[role] || {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 0, rm: 0 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 0, rm: 0 }
      };
      
      dataRows.push([
        idx === 0 ? "KESELURUHAN IMPLIKASI KEWANGAN" : "",
        "",
        role,
        rData.bracket1_23.bil,
        rData.bracket24_48.bil,
        rData.bracket24_48.rm,
        rData.bracket49_96.bil,
        rData.bracket49_96.rm,
        rData.bracket97_128.bil,
        rData.bracket97_128.rm,
        rData.total.bil,
        rData.total.rm
      ]);
    });
    
    // Footer row
    dataRows.push([]);
    dataRows.push(["SILA ISIKAN BUTIRAN MENGGUNAKAN FORMAT YANG TELAH DISEDIAKAN MENGIKUT KONTINJEN"]);
    
    const ws = XLSX.utils.aoa_to_sheet(dataRows);
    
    // Set merges
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 11 } },
      
      // BIL, FORMASI, PANGKAT, JUMLAH JAM header merges
      { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } },
      { s: { r: 3, c: 1 }, e: { r: 4, c: 1 } },
      { s: { r: 3, c: 2 }, e: { r: 4, c: 2 } },
      { s: { r: 3, c: 3 }, e: { r: 4, c: 3 } },
      
      // Horizontal merges for brackets
      { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
      { s: { r: 3, c: 6 }, e: { r: 3, c: 7 } },
      { s: { r: 3, c: 8 }, e: { r: 3, c: 9 } },
      { s: { r: 3, c: 10 }, e: { r: 3, c: 11 } }
    ];
    
    // Add row spans for districts
    for (let d = 0; d < 4; d++) {
      const baseRow = 5 + d * 3;
      ws['!merges'].push({ s: { r: baseRow, c: 0 }, e: { r: baseRow + 2, c: 0 } });
      ws['!merges'].push({ s: { r: baseRow, c: 1 }, e: { r: baseRow + 2, c: 1 } });
    }
    
    // Merge first two columns of KESELURUHAN row
    ws['!merges'].push({ s: { r: 17, c: 0 }, e: { r: 19, c: 1 } });
    
    // Column widths
    ws['!cols'] = [
      { wch: 10 }, // BIL / IMPLIKASI
      { wch: 20 }, // FORMASI
      { wch: 12 }, // PANGKAT
      { wch: 20 }, // JUMLAH JAM BERTUGAS
      { wch: 8 },  // 24 (A) BIL
      { wch: 14 }, // 24 (A) RM
      { wch: 8 },  // 49-96 BIL
      { wch: 14 }, // 49-96 RM
      { wch: 8 },  // 97-128 BIL
      { wch: 14 }, // 97-128 RM
      { wch: 16 }, // KESELURUHAN BIL
      { wch: 16 }  // KESELURUHAN RM
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Yearly Forecast");
    XLSX.writeFile(wb, `Unjuran_Penugasan_Tahunan_${selectedYear}.xlsx`);
  };


  const renderYearlyForecastTable = () => {
    const data = yearlyForecastData;
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];

    const formatRM = (val: number) => {
      if (val === 0) return '0.00';
      return val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    const formatBil = (val: number) => {
      return val === 0 ? '0' : val.toString();
    };

    return (
      <div className="w-full max-w-7xl mx-auto bg-white p-2 sm:p-6 select-none font-sans print:p-0 print-page-container mt-8">
        {/* Yellow Header Banners */}
        <div className="border border-black mb-1 bg-[#ffff00] text-black text-center py-2 font-bold text-lg sm:text-xl uppercase tracking-wider">
          KONTINJEN : MELAKA
        </div>
        <div className="border border-black bg-[#ffff00] text-black text-center py-2 font-bold text-sm sm:text-base uppercase tracking-wide flex justify-between px-4 sm:px-8">
          <span className="w-1/3 text-left">TUNTUTAN ELAUN PENUGASAN SSPDRM - </span>
          <span className="w-1/3 text-center text-xl font-black">SEPANJANG TAHUN {selectedYear}</span>
          <span className="w-1/3"></span>
        </div>

        {/* The Main Table with exact colors from photo */}
        <table className="w-full border-collapse border border-black text-center font-bold text-[11px] sm:text-xs mt-4">
          <thead>
            {/* Header Row 1 */}
            <tr className="bg-[#ebd300] text-black text-[10px] sm:text-[11px] h-[34px]">
              <th className="border border-black p-1 w-10 text-center" rowSpan={2}>BIL</th>
              <th className="border border-black p-1 w-40 text-left pl-2" rowSpan={2}>FORMASI</th>
              <th className="border border-black p-1 w-24 text-center" rowSpan={2}>PANGKAT</th>
              <th className="border border-black p-1 bg-[#3a60a1] text-white w-20 text-center font-bold" colSpan={1}>JUMLAH JAM BERTUGAS</th>
              <th className="border border-black p-1 bg-[#1e7c53] text-white text-center font-bold" colSpan={2}>24 (A)</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white text-center font-bold" colSpan={2}>49 - 96 (B)</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white text-center font-bold" colSpan={2}>97 - 128 (C)</th>
              <th className="border border-black p-1 bg-[#cca300] text-black text-center font-bold" colSpan={2}>KESELURUHAN</th>
            </tr>
            {/* Header Row 2 */}
            <tr className="text-black text-[10px] sm:text-[11px]">
              <th className="border border-black p-1 bg-[#4f81bd] text-white w-20 text-center font-bold">1 - 23 JAM<br/>BIL</th>
              <th className="border border-black p-1 bg-[#375623] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#375623] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#ffc000] text-black w-14 text-center font-bold">Jumlah (A+B+C)</th>
              <th className="border border-black p-1 bg-[#ffc000] text-black w-24 text-center font-bold">RM</th>
            </tr>
          </thead>
          <tbody>
            {districtsList.flatMap((district, dIdx) => {
              const dData = data.districts[district] || {
                PEG: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
                APR: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
                JUMLAH: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } }
              };

              const roles = ['PEG', 'APR', 'JUMLAH'];

              return roles.map((role, rIdx) => {
                const rData = dData[role] || {
                  bracket1_23: { bil: 0, rm: 0 },
                  bracket24_48: { bil: 0, rm: 0 },
                  bracket49_96: { bil: 0, rm: 0 },
                  bracket97_128: { bil: 0, rm: 0 },
                  total: { bil: 0, rm: 0 }
                };

                const isJumlah = role === 'JUMLAH';
                const colFG_Bg = isJumlah ? 'bg-[#ffc000] text-black' : '';

                return (
                  <tr key={`${district}-${role}`} className="bg-white">
                    {rIdx === 0 && (
                      <>
                        <td className="border border-black p-1 text-center font-bold" rowSpan={3}>{dIdx + 1}</td>
                        <td className="border border-black p-1 text-left uppercase font-bold text-xs" rowSpan={3}>
                          {district}
                        </td>
                      </>
                    )}

                    {/* Role / Pangkat column */}
                    <td className={`border border-black p-1 uppercase text-center font-bold ${colFG_Bg}`}>{role}</td>

                    {/* 1-23 JAM Column */}
                    <td className={`border border-black p-1 text-center font-bold ${colFG_Bg}`}>
                      {formatBil(rData.bracket1_23.bil)}
                    </td>

                    {/* 24-48 (A) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colFG_Bg}`}>
                      {formatBil(rData.bracket24_48.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colFG_Bg}`}>
                      {formatRM(rData.bracket24_48.rm)}
                    </td>

                    {/* 49-96 (B) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colFG_Bg}`}>
                      {formatBil(rData.bracket49_96.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colFG_Bg}`}>
                      {formatRM(rData.bracket49_96.rm)}
                    </td>

                    {/* 97-128 (C) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colFG_Bg}`}>
                      {formatBil(rData.bracket97_128.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colFG_Bg}`}>
                      {formatRM(rData.bracket97_128.rm)}
                    </td>

                    {/* KESELURUHAN (A+B+C) */}
                    <td className="border border-black p-1 text-center font-black bg-[#ffc000] text-black">
                      {formatBil(rData.total.bil)}
                    </td>
                    <td className="border border-black p-1 text-right pr-2 font-black bg-[#ffc000] text-black">
                      {formatRM(rData.total.rm)}
                    </td>
                  </tr>
                );
              });
            })}

            {/* KESELURUHAN ROW (Bottom) */}
            {['PEG', 'APR', 'JUMLAH'].map((role, idx) => {
              const rData = data.keseluruhan[role] || {
                bracket1_23: { bil: 0, rm: 0 },
                bracket24_48: { bil: 0, rm: 0 },
                bracket49_96: { bil: 0, rm: 0 },
                bracket97_128: { bil: 0, rm: 0 },
                total: { bil: 0, rm: 0 }
              };

              return (
                <tr key={`keseluruhan-${role}`}>
                  {idx === 0 && (
                    <td className="border border-black p-1 text-center font-black uppercase bg-[#ebd300] text-black" colSpan={2} rowSpan={3}>
                      KESELURUHAN IMPLIKASI KEWANGAN
                    </td>
                  )}
                  {/* Role / Pangkat column */}
                  <td className="border border-black p-1 uppercase text-center font-black bg-[#ffff00] text-black">{role}</td>

                  {/* 1-23 JAM */}
                  <td className="border border-black p-1 font-black bg-[#4f81bd] text-white text-center">
                    {formatBil(rData.bracket1_23.bil)}
                  </td>

                  {/* 24-48 (A) */}
                  <td className="border border-black p-1 font-black bg-[#375623] text-white text-center">
                    {formatBil(rData.bracket24_48.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#375623] text-white text-right pr-2">
                    {formatRM(rData.bracket24_48.rm)}
                  </td>

                  {/* 49-96 (B) */}
                  <td className="border border-black p-1 font-black bg-[#4a154b] text-white text-center">
                    {formatBil(rData.bracket49_96.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#4a154b] text-white text-right pr-2">
                    {formatRM(rData.bracket49_96.rm)}
                  </td>

                  {/* 97-128 (C) */}
                  <td className="border border-black p-1 font-black bg-[#7b113a] text-white text-center">
                    {formatBil(rData.bracket97_128.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#7b113a] text-white text-right pr-2">
                    {formatRM(rData.bracket97_128.rm)}
                  </td>

                  {/* KESELURUHAN (A+B+C) */}
                  <td className="border border-black p-1 font-black bg-[#ffc000] text-black text-center">
                    {formatBil(rData.total.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#ffc000] text-black text-right pr-2">
                    {formatRM(rData.total.rm)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Instructions Banner Footer */}
        <div className="border border-black bg-[#ffffea] text-black text-left pl-4 py-2 mt-4 font-bold text-xs uppercase tracking-wider">
          SILA ISIKAN BUTIRAN MENGGUNAKAN FORMAT YANG TELAH DISEDIAKAN MENGIKUT KONTINJEN
        </div>
      </div>
    );
  };

  const renderPenyalurMaklumat = () => {
    // Filter data based on selected month and year
    const filteredData = maklumatData.filter(row => {
      const tarikh = String(row['TARIKH MAKLUMAT'] || '').trim();
      if (!tarikh) return false;
      const parts = tarikh.split('/');
      if (parts.length < 3) return false;
      const m = parseInt(parts[1], 10) - 1;
      const y = parseInt(parts[2], 10);
      return m === months.indexOf(selectedMonth) && y === selectedYear;
    });

    if (filteredData.length === 0) {
      return (
        <div className="w-full max-w-4xl mx-auto bg-white p-8 text-center print:hidden">
          <p className="text-gray-500">Tiada data untuk bulan dan tahun yang dipilih.</p>
        </div>
      );
    }

    const sumPegawai = { total: 0, jenayah: 0, narkotik: 0, trafic: 0, komersil: 0, lain: 0 };
    const sumAnggota = { total: 0, jenayah: 0, narkotik: 0, trafic: 0, komersil: 0, lain: 0 };
    const sumTotal = { total: 0, jenayah: 0, narkotik: 0, trafic: 0, komersil: 0, lain: 0 };

    filteredData.forEach(row => {
      const pejawatan = String(row['PEJAWATAN ANDA'] || '').trim().toUpperCase();
      const isPegawai = pejawatan === 'PEG';
      const jenisRaw = String(row['JENIS MAKLUMAT'] || '').trim().toUpperCase();
      
      const checkType = (type: string) => {
        if (type === 'JENAYAH' && jenisRaw.includes('JENAYAH')) return 1;
        if (type === 'NARKOTIK' && (jenisRaw.includes('DADAH') || jenisRaw.includes('NARKOTIK'))) return 1;
        if (type === 'TRAFIC' && (jenisRaw.includes('TRAFIK') || jenisRaw.includes('TRAFIC'))) return 1;
        if (type === 'KOMERSIL' && jenisRaw.includes('KOMERSIL')) return 1;
        if (type === 'LAIN-LAIN' && jenisRaw.includes('LAIN-LAIN')) return 1;
        return 0;
      };

      const j = checkType('JENAYAH');
      const n = checkType('NARKOTIK');
      const t = checkType('TRAFIC');
      const k = checkType('KOMERSIL');
      const l = checkType('LAIN-LAIN');

      if (isPegawai) {
        sumPegawai.total += 1;
        sumPegawai.jenayah += j;
        sumPegawai.narkotik += n;
        sumPegawai.trafic += t;
        sumPegawai.komersil += k;
        sumPegawai.lain += l;
      } else {
        sumAnggota.total += 1;
        sumAnggota.jenayah += j;
        sumAnggota.narkotik += n;
        sumAnggota.trafic += t;
        sumAnggota.komersil += k;
        sumAnggota.lain += l;
      }
      
      sumTotal.total += 1;
      sumTotal.jenayah += j;
      sumTotal.narkotik += n;
      sumTotal.trafic += t;
      sumTotal.komersil += k;
      sumTotal.lain += l;
    });

    const formatS = (val: number) => val === 0 ? 'TIADA' : String(val);

    const namesByDaerah: Record<string, Set<string>> = {};
    filteredData.forEach(row => {
      let daerahRaw = String(row['DAERAH ANDA'] || '').trim().toUpperCase();
      let daerah = daerahRaw.replace('SSPDRM', '').trim();
      if (daerah === '') daerah = 'LAIN-LAIN';
      
      if (!namesByDaerah[daerah]) {
        namesByDaerah[daerah] = new Set();
      }
      const rawName = String(row['NAMA'] || '').trim();
      
      if (rawName) {
        namesByDaerah[daerah].add(rawName);
      }
    });

    const sortedDaerahs = Object.keys(namesByDaerah).sort();

    return (
      <div id="maklumat-container">
        <style>{`
          .td-main { border: 1px solid black; padding: 6px 12px; }
          .td-opt { border: none; padding: 2px 4px; }
          .bg-gray { background-color: #f3f4f6 !important; }
          .page-break-after { page-break-after: always; break-after: page; } 
          @media print { 
            .page-break-after { page-break-after: always !important; break-after: page !important; } 
          }
        `}</style>
        {filteredData.map((row, index) => {
          const rawName = String(row[' NOMBOR BADAN DAN NAMA'] || row['NOMBOR BADAN DAN NAMA'] || '').trim();
          const noBadanMatch = rawName.match(/\d+/);
          const noBadan = noBadanMatch ? noBadanMatch[0] : '';
          const namaPenyalur = rawName.replace(noBadan, '').trim();
          const pangkat = String(row['PANGKAT'] || '').trim();
          const tarikhMasa = `${String(row['TARIKH MAKLUMAT'] || '')} / ${String(row['MASA MAKLUMAT'] || '')}`;
          const noSd = String(row['NO SD'] || '-').trim();

          const namaPemberi = String(row['NAMA'] || '').trim();
          const lokasi = String(row['LOKASI TERIMA MAKLUMAT'] || '').trim();
          const kategori = String(row['KATEGORI MAKLUMAT'] || '').toUpperCase().trim();
          const jenis = String(row['JENIS MAKLUMAT'] || '').toUpperCase().trim();
          const butiran = String(row['BUTIR-BUTIR MAKLUMAT'] || '').trim();

          const lastDayOfMonth = new Date(selectedYear, months.indexOf(selectedMonth) + 1, 0).getDate();
          const strLastDay = `${lastDayOfMonth}/${months.indexOf(selectedMonth) + 1}/${selectedYear}`;

          return (
            <div key={index}>
              <div className="w-full max-w-4xl mx-auto bg-white p-8 split-page page-break-after doc-page-break" style={{ marginBottom: '2rem' }}>
                <h1 className="text-center font-bold font-serif" style={{ fontSize: '16px', textAlign: 'center', marginTop: '-4px', marginBottom: '16px' }}>BORANG MAKLUMAT</h1>
                
                <table className="w-full border-collapse border border-black text-sm text-left">
                  <tbody>
                    {/* BUTIR-BUTIR PENYALUR MAKLUMAT */}
                  <tr className="bg-gray">
                    <td colSpan={2} className="td-main font-bold uppercase text-center" style={{ textAlign: 'center' }}>BUTIR-BUTIR PENYALUR MAKLUMAT</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white w-[35%]" style={{ width: '35%' }}>NAMA</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{namaPenyalur}</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">NO.BADAN</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{noBadan}</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">PANGKAT</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{pangkat}</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">TARIKH/MASA</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{tarikhMasa}</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">NO.SD</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{noSd}</td>
                  </tr>

                  {/* BUTIR-BUTIR PEMBERI MAKLUMAT */}
                  <tr className="bg-gray">
                    <td colSpan={2} className="td-main font-bold uppercase text-center" style={{ textAlign: 'center' }}>BUTIR-BUTIR PEMBERI MAKLUMAT</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">NAMA</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{namaPemberi}</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">KP</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>RAHSIA</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">PEKERJAAN</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>ANGGOTA SSPDRM</td>
                  </tr>
                  <tr>
                    <td className="td-main uppercase bg-white">NO. TEL</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>RAHSIA</td>
                  </tr>

                  {/* LOKASI TERIMA MAKLUMAT */}
                  <tr>
                    <td className="td-main uppercase bg-white align-top">LOKASI <br />TERIMA MAKLUMAT</td>
                    <td className="td-main font-bold uppercase bg-white text-center" style={{ textAlign: 'center' }}>{lokasi}</td>
                  </tr>

                  {/* KATEGORI MAKLUMAT */}
                  <tr>
                    <td className="td-main uppercase bg-white align-top">KATEGORI MAKLUMAT</td>
                    <td className="td-main bg-white align-top">
                      <table style={{width: '100%', border: 'none'}}>
                        <tbody>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>1) UMUM {kategori.includes('UMUM') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>2) RAHSIA {kategori.includes('RAHSIA') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                        </tbody>
                      </table>
                    </td>
                  </tr>

                  {/* JENIS MAKLUMAT */}
                  <tr>
                    <td className="td-main uppercase bg-white align-top">JENIS MAKLUMAT</td>
                    <td className="td-main bg-white align-top">
                      <table style={{width: '100%', border: 'none'}}>
                        <tbody>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>1) JENAYAH {jenis.includes('JENAYAH') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>2) DADAH {jenis.includes('DADAH') || jenis.includes('NARKOTIK') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>3) TRAFIK {jenis.includes('TRAFIK') || jenis.includes('TRAFIC') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>4) KOMERSIL {jenis.includes('KOMERSIL') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                          <tr>
                            <td className="td-opt" style={{width: '120px'}}>5) LAIN-LAIN {jenis.includes('LAIN-LAIN') ? <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;√</strong> : ''}</td>
                          </tr>
                        </tbody>
                      </table>
                    </td>
                  </tr>

                  {/* BUTIR-BUTIR MAKLUMAT */}
                  <tr className="bg-gray">
                    <td colSpan={2} className="td-main font-bold uppercase text-center" style={{ textAlign: 'center' }}>BUTIR-BUTIR MAKLUMAT</td>
                  </tr>
                  <tr>
                    <td colSpan={2} className="td-main bg-white align-middle text-center" style={{ minHeight: '120px', height: '120px' }}>
                      <div style={{ textAlign: 'center', minHeight: '120px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        {butiran}
                      </div>
                    </td>
                  </tr>

                  {/* ULASAN & TANDATANGAN */}
                  <tr>
                    <td className="td-main bg-white" style={{ height: '150px', verticalAlign: 'bottom', width: '35%' }} valign="bottom">
                      <div className="font-bold" style={{ marginBottom: '8px' }}>TANDATANGAN :</div>
                      <div className="font-bold">TARIKH : {strLastDay}</div>
                    </td>
                    <td className="td-main bg-white" style={{ height: '150px', verticalAlign: 'top' }} valign="top">
                      <div className="font-bold">ULASAN:</div>
                    </td>
                  </tr>
                </tbody>
              </table>
              </div>
              <div className="page-break-mark"></div>
            </div>
          );
        })}

        {/* Summary Table at the last page */}
        <div className="w-full max-w-4xl mx-auto bg-white p-8 doc-page-break split-page page-break-after mt-8" style={{ marginTop: '2rem' }}>
          <h1 className="text-center font-bold text-xl mb-4 font-serif" style={{ textAlign: 'center', fontWeight: 'bold', fontSize: '1.25rem', marginBottom: '1rem', fontFamily: 'serif' }}>PELAPORAN MAKLUMAT SSPDRM</h1>
          <table className="w-full border-collapse border border-black text-center" style={{ fontSize: '14px', fontFamily: 'Arial, sans-serif', width: '100%', borderCollapse: 'collapse', textAlign: 'center' }}>
            <tbody>
              <tr style={{ backgroundColor: '#00A2E8', fontWeight: 'bold', fontSize: '16px' }}>
                  <td colSpan={2} className="border border-black p-2 uppercase text-left w-1/2" style={{ border: '1px solid black', padding: '8px', textAlign: 'left' }}>KONTINJEN: MELAKA</td>
                  <td colSpan={5} className="border border-black p-2 uppercase text-left pl-12 w-1/2" style={{ border: '1px solid black', padding: '8px', textAlign: 'left', paddingLeft: '48px' }}>BULAN &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{selectedMonth.toUpperCase()}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{selectedYear}</td>
              </tr>
              <tr style={{ backgroundColor: '#BDD7EE' }}>
                  <td rowSpan={2} className="border border-black p-2 font-bold uppercase w-1/5" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>PERKARA</td>
                  <td rowSpan={2} className="border border-black p-2 font-bold uppercase w-1/6" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>JUMLAH MAKLUMAT<br/>DITERIMA</td>
                  <td colSpan={5} className="border border-black p-2 font-bold uppercase" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>KATEGORI</td>
              </tr>
              <tr style={{ backgroundColor: '#BDD7EE', fontSize: '12px' }} className="uppercase">
                  <td className="border border-black p-2" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>JENAYAH</td>
                  <td className="border border-black p-2" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>NARKOTIK</td>
                  <td className="border border-black p-2" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>TRAFIC</td>
                  <td className="border border-black p-2" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>KOMERSIL</td>
                  <td className="border border-black p-2" style={{ border: '1px solid black', padding: '8px', textAlign: 'center' }}>LAIN-LAIN</td>
              </tr>
              <tr style={{ height: '100px' }}>
                  <td className="border border-black p-4 font-bold text-left uppercase" style={{ border: '1px solid black', padding: '16px', textAlign: 'left' }}>JUMLAH PEGAWAI</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.total)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.jenayah)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.narkotik)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.trafic)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.komersil)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumPegawai.lain)}</td>
              </tr>
              <tr style={{ height: '100px' }}>
                  <td className="border border-black p-4 font-bold text-left uppercase" style={{ border: '1px solid black', padding: '16px', textAlign: 'left' }}>JUMLAH ANGGOTA</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.total)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.jenayah)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.narkotik)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.trafic)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.komersil)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumAnggota.lain)}</td>
              </tr>
              <tr style={{ height: '100px' }}>
                  <td className="border border-black p-4 font-bold text-left uppercase" style={{ border: '1px solid black', padding: '16px', textAlign: 'left' }}>JUMLAH KESELURUHAN</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.total)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.jenayah)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.narkotik)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.trafic)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.komersil)}</td>
                  <td className="border border-black p-4 font-bold text-lg" style={{ border: '1px solid black', padding: '16px', textAlign: 'center', fontSize: '18px' }}>{formatS(sumTotal.lain)}</td>
              </tr>
            </tbody>
          </table>
          <div className="page-break-mark"></div>
        </div>

        {/* Senarai Nama Penyalur by Daerah */}
        <div className="w-full max-w-4xl mx-auto bg-white p-8 doc-page-break split-page mt-8" style={{ marginTop: '2rem' }}>
          <h1 className="text-center font-bold font-serif uppercase" style={{ fontSize: '18px', marginBottom: '2rem', textAlign: 'center', fontFamily: 'serif', fontWeight: 'bold' }}>SENARAI NAMA ANGGOTA / PEGAWAI HADIR MEMBERI MAKLUMAT BAGI<br/>BULAN {selectedMonth.toUpperCase()} {selectedYear}</h1>
          <div className="flex flex-col gap-6" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {sortedDaerahs.map((daerah, idx) => (
              <div key={idx} className="bg-white break-inside-avoid" style={{ breakInside: 'avoid', pageBreakInside: 'avoid', border: '1px solid black', padding: '24px' }}>
                <h2 className="font-bold uppercase text-[16px]" style={{ fontWeight: 'bold', fontSize: '16px', marginBottom: '16px', fontFamily: 'Arial, sans-serif' }}>{daerah}</h2>
                <div style={{ borderBottom: '1px solid black', marginBottom: '16px' }}></div>
                <div className="flex flex-col gap-2" style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {Array.from(namesByDaerah[daerah]).map((nama, nIdx) => (
                    <div key={nIdx} className="text-[14px] font-medium uppercase" style={{ fontSize: '14px', fontFamily: 'Arial, sans-serif' }}>{nIdx + 1}. {nama}</div>
                  ))}
                  {namesByDaerah[daerah].size === 0 && <div className="text-[14px] italic text-gray-500" style={{ fontSize: '14px', fontStyle: 'italic', color: '#6b7280' }}>Tiada rekod</div>}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  };

  const renderAllowanceTable = (isLive: boolean = false) => {
    const currentVoucherData = isLive ? voucherDataLive : voucherData;
    const daysInMonth = new Date(selectedYear, months.indexOf(selectedMonth) + 1, 0).getDate();
    const daysArray = Array.from({ length: 31 }, (_, i) => i + 1);

    // Filter rawData to get daily hours for the 10 selected people
    const getPersonnelDailyHours = (noBadan: string) => {
      const dailyHours = Array(31).fill(0);
      let name = '';
      let rank = '';
      
      const normalize = (s: string) => String(s || '').replace(/[^0-9]/g, '');
      const targetNo = normalize(noBadan);

      const isDistrictMatch = (rowDistrict: any, targetDistrict: string) => {
        if (!rowDistrict || !targetDistrict) return false;
        const s = String(rowDistrict).trim().toUpperCase();
        const t = targetDistrict.trim().toUpperCase();
        
        if (s.includes(t) || t.includes(s)) return true;
        
        // Handle common abbreviations
        if (t === 'ALOR GAJAH' && (s === 'AG' || s.includes('ALOR'))) return true;
        if (t === 'MELAKA TENGAH' && (s === 'MT' || s.includes('TENGAH'))) return true;
        if (t === 'JASIN' && (s === 'JS' || s.includes('JASIN'))) return true;
        if (t === 'IPK SSPDRM' && (s === 'IPK')) return true;
        
        return false;
      };

      if (isLive) {
        // Use attendanceDataLive for hours and dates
        if (!attendanceDataLive || attendanceDataLive.length === 0) return { dailyHours, name, rank, totalHours: 0 };

        // Find the person in processedData.districtPersonnel first to get their name and rank (consistent with backup)
        const person = processedData.districtPersonnel.find(p => {
          const pNoStr = String(p.noBadan).replace(/[^0-9]/g, '');
          if (!pNoStr) return false;
          const pNo = parseInt(pNoStr, 10);
          const tNo = parseInt(targetNo, 10);
          return pNo === tNo;
        });

        if (person) {
          name = person.name.replace(/[0-9]/g, '').trim();
          rank = person.rank;
        } else {
          // Fallback to voucher data info if not in attendance sheet
          const personInfo = voucherDataLive.find(v => normalize(v['No Badan']) === targetNo);
          if (personInfo) {
            name = personInfo['No Badan'].replace(/[0-9]/g, '').trim();
          }
        }

        attendanceDataLive.forEach(row => {
          if (normalize(row['No Badan']) !== targetNo) return;
          
          if (!rank && row['Pangkat']) {
            rank = String(row['Pangkat']).trim();
          }
          if (!name && row['No Badan']) {
            name = String(row['No Badan']).replace(/[0-9]/g, '').trim();
          }
          
          // District filtering removed to ensure searched members from other districts show data correctly
          
          
          const dateStr = String(row['Duty Date'] || '');
          const hours = parseFloat(String(row['Hours'] || '0')) || 0;
          
          let rowYear = -1, rowMonth = -1, rowDay = -1;
          const dateOnly = dateStr.split(' ')[0];
          const parts = dateOnly.split(/[\/\-]/);
          if (parts.length === 3) {
            if (parts[0].length === 4) {
              rowYear = parseInt(parts[0], 10);
              rowMonth = parseInt(parts[1], 10) - 1;
              rowDay = parseInt(parts[2], 10);
            } else {
              const p0 = parseInt(parts[0], 10);
              const p1 = parseInt(parts[1], 10);
              const p2 = parseInt(parts[2], 10);
              if (p1 > 12) { rowMonth = p0 - 1; rowDay = p1; rowYear = p2; }
              else { rowDay = p0; rowMonth = p1 - 1; rowYear = p2; }
              if (rowYear < 100) rowYear += 2000;
            }
          } else {
            const d = new Date(dateStr);
            if (!isNaN(d.getTime())) {
              rowYear = d.getFullYear();
              rowMonth = d.getMonth();
              rowDay = d.getDate();
            }
          }

          if (rowYear === selectedYear && rowMonth === months.indexOf(selectedMonth)) {
            if (rowDay >= 1 && rowDay <= 31) {
              dailyHours[rowDay - 1] += hours;
            }
          }
        });

        const totalHours = dailyHours.reduce((a, b) => a + b, 0);
        return { dailyHours, name, rank, totalHours };
      }

      // BACKUP DATA LOGIC (Original)
      if (!noBadan || !rawData || rawData.length === 0) return { dailyHours, name, rank, totalHours: 0 };

      // Find the person in processedData.districtPersonnel first to get their name
      const person = processedData.districtPersonnel.find(p => {
        const pNoStr = String(p.noBadan).replace(/[^0-9]/g, '');
        if (!pNoStr) return false;
        const pNo = parseInt(pNoStr, 10);
        const tNo = parseInt(targetNo, 10);
        return pNo === tNo;
      });

      if (!person) return { dailyHours, name, rank, totalHours: 0 };
      
      name = person.name.replace(/[0-9]/g, '').trim();
      rank = person.rank;
      
      // Find header row and indices (same logic as useMemo for consistency)
      let dateIdx = -1, distIdx = -1, hoursIdx = -1;
      let headerRowIndex = -1;

      for (let i = 0; i < Math.min(20, rawData.length); i++) {
        const row = rawData[i];
        if (!Array.isArray(row)) continue;
        const dIdx = row.findIndex(c => String(c).toUpperCase().includes('TARIKH'));
        const dsIdx = row.findIndex(c => String(c).toUpperCase().includes('DAERAH'));
        if (dIdx !== -1 && dsIdx !== -1) {
          headerRowIndex = i;
          dateIdx = dIdx;
          distIdx = dsIdx;
          hoursIdx = row.findIndex(c => String(c).toUpperCase().includes('JUMLAH JAM'));
          if (hoursIdx === -1) hoursIdx = 11; // Fallback
          break;
        }
      }

      if (headerRowIndex === -1) return { dailyHours, name, rank, totalHours: 0 };

      rawData.slice(headerRowIndex + 1).forEach((row) => {
        if (!Array.isArray(row)) return;
        
        // District filtering removed to ensure searched members from other districts show data correctly


        const dateStr = String(row[dateIdx] || '');
        let rowYear = -1, rowMonth = -1, rowDay = -1;

        // Robust date parsing (same as useMemo)
        const dateOnly = dateStr.split(' ')[0];
        const parts = dateOnly.split(/[\/\-]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            rowYear = parseInt(parts[0], 10);
            rowMonth = parseInt(parts[1], 10) - 1;
            rowDay = parseInt(parts[2], 10);
          } else {
            const p0 = parseInt(parts[0], 10);
            const p1 = parseInt(parts[1], 10);
            const p2 = parseInt(parts[2], 10);
            if (p1 > 12) { rowMonth = p0 - 1; rowDay = p1; rowYear = p2; }
            else { rowDay = p0; rowMonth = p1 - 1; rowYear = p2; }
            if (rowYear < 100) rowYear += 2000;
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            rowYear = d.getFullYear();
            rowMonth = d.getMonth();
            rowDay = d.getDate();
          }
        }

        const hours = parseFloat(String(row[hoursIdx] || '0')) || 0;
        
        if (rowYear === selectedYear && months[rowMonth] === selectedMonth) {
          // Check if this person is in this row
          const nameIndices = [3, 5, 7, 9];
          const isPresent = nameIndices.some((idx) => {
            if (idx >= row.length) return false;
            const rowContent = String(row[idx] || '').toUpperCase();
            return rowContent.includes(targetNo) || (person.name && rowContent.includes(person.name.toUpperCase()));
          });

          if (isPresent && rowDay >= 1 && rowDay <= 31) {
            dailyHours[rowDay - 1] += hours;
          }
        }
      });

      const totalHours = dailyHours.reduce((a, b) => a + b, 0);
      return { dailyHours, name, rank, totalHours };
    };

    const getRate = (rank: string) => {
      const r = rank.toUpperCase();
      if (r.includes('SUPT') || r.includes('DSP') || r.includes('ASP') || r.includes('INSP')) return 9.80;
      return 8.00;
    };

    const numberToMalayWords = (n: number) => {
      const units = ['', 'SATU', 'DUA', 'TIGA', 'EMPAT', 'LIMA', 'ENAM', 'TUJUH', 'LAPAN', 'SEMBILAN'];
      const teens = ['SEPULUH', 'SEBELAS', 'DUA BELAS', 'TIGA BELAS', 'EMPAT BELAS', 'LIMA BELAS', 'ENAM BELAS', 'TUJUH BELAS', 'LAPAN BELAS', 'SEMBILAN BELAS'];
      const tens = ['', 'SEPULUH', 'DUA PULUH', 'TIGA PULUH', 'EMPAT PULUH', 'LIMA PULUH', 'ENAM PULUH', 'TUJUH PULUH', 'LAPAN PULUH', 'SEMBILAN PULUH'];
      
      const convert = (num: number): string => {
        if (num === 0) return '';
        if (num < 10) return units[num];
        if (num < 20) return teens[num - 10];
        if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 !== 0 ? ' ' + units[num % 10] : '');
        if (num < 1000) return (num < 200 ? 'SERATUS' : units[Math.floor(num / 100)] + ' RATUS') + (num % 100 !== 0 ? ' ' + convert(num % 100) : '');
        if (num < 1000000) return (num < 2000 ? 'SERIBU' : convert(Math.floor(num / 1000)) + ' RIBU') + (num % 1000 !== 0 ? ' ' + convert(num % 1000) : '');
        return '';
      };

      const ringgit = Math.floor(n);
      const sen = Math.round((n - ringgit) * 100);
      
      let result = convert(ringgit);
      if (sen > 0) {
        result += ' DAN SEN ' + convert(sen);
      }
      return result + ' SAHAJA';
    };

    let grandTotalElaun = 0;
    let grandTotalHoursWorked = 0;
    let grandTotalCappedHours = 0;
    let grandTotalPenugasan = 0;

    return (
      <React.Fragment>
        {/* Report 1: Attendance & Allowance */}
        <div className="print-page-container relative pb-8 print:pb-0">
          <div className="relative z-10 pt-4">
            <div className="absolute top-0 right-0 text-right text-[10px] font-bold">
              SPDRM MELAKA BR.NO............<br/>
              LAMPIRAN 'A1'<br/>
              PDRM (H) 49<br/>
              <span className="text-[8px]">PNMB.,K.L</span>
            </div>

            <div className="text-center mb-2">
              <h2 className="text-[30px] font-bold uppercase print:text-[26px]">PASUKAN SUKARELAWAN SIMPANAN POLIS</h2>
              <div className="text-[14px] font-bold mt-1 print:text-[14px]">
                Daftar Kedatangan dan Jadual Elaun bagi Bulan: <span className="border-b border-black px-4">{selectedMonth} {selectedYear}</span>
              </div>
              <div className="text-[14px] font-bold mt-2 flex justify-center items-center gap-8 print:mt-1 print:gap-4">
                <span className="text-[14px] print:text-[14px]">Nama Pasukan :</span>
                <span className="text-[24px] font-black tracking-widest print:text-[14px]">{selectedDistrict}</span>
              </div>
            </div>

            <div className="overflow-x-auto relative">
              {isLoading && (
                <div className="absolute inset-0 flex items-center justify-center z-20 print:hidden" style={{ backgroundColor: 'rgba(255, 255, 255, 0.5)' }}>
                  <div className="flex flex-col items-center gap-2">
                    <RefreshCw className="animate-spin text-blue-600" size={32} />
                    <span className="text-xs font-bold text-blue-600 uppercase tracking-widest">Fetching Data...</span>
                  </div>
                </div>
              )}
              <table className="w-full border-collapse border border-black text-[9px] text-center font-bold table-fixed">
                <colgroup>
                  <col style={{ width: '2%' }} />
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '16.8%' }} />
                  {Array(31).fill(0).map((_, i) => (
                    <col key={i} style={{ width: '1.2%' }} />
                  ))}
                  <col style={{ width: '2.5%' }} />
                  <col style={{ width: '2.5%' }} />
                  <col style={{ width: '3.5%' }} />
                  <col style={{ width: '3%' }} />
                  <col style={{ width: '3%' }} />
                  <col style={{ width: '3%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '4%' }} />
                </colgroup>
                <thead>
                  <tr className="bg-gray-50 h-[22px]">
                    <th className="border border-black py-0 px-1" rowSpan={2}>Bil</th>
                    <th className="border border-black py-0 px-1" rowSpan={2}>NO. BADAN</th>
                    <th className="border border-black py-0 px-1" rowSpan={2}>Pangkat</th>
                    <th className="border border-black py-0 px-1" rowSpan={2}>Nama</th>
                    <th className="border border-black py-0 px-1" colSpan={31}>Jumlah jam bertugas / berlatih pada tarikh berikut</th>
                    <th className="border border-black py-0 px-0.5 text-[8px] leading-tight font-bold" rowSpan={2}>Jumlah<br/>kedatangan</th>
                    <th className="border border-black py-0 px-0.5 text-[8px] leading-tight font-bold" colSpan={2} rowSpan={2}>Jumlah<br/>Jam<br/>bertugas/<br/>berlatih</th>
                    <th className="border border-black py-0 px-0.5 text-[10px] font-bold" colSpan={3}>ELAUN KENDERAAN</th>
                    <th className="border border-black py-0 px-0.5 text-[8px] leading-tight font-bold" rowSpan={2}>Belanja Elaun Latihan<br/>(Peringatan A)</th>
                    <th className="border border-black py-0 px-0.5 text-[8px] leading-tight font-bold" rowSpan={2}>Jumlah Elaun yang akan<br/>di bayar</th>
                    <th className="border border-black py-0 px-0.5 text-[8px] leading-tight font-bold" rowSpan={2}>Tanda tangan<br/>penerima</th>
                  </tr>
                  <tr className="bg-gray-50 h-[22px]">
                    {daysArray.map(d => <th key={d} className="border border-black p-0 text-[8px]">{d}</th>)}
                    <th className="border border-black py-0 px-0.5 text-[7px] leading-tight font-bold">Jenis<br/>(Peringa<br/>tan B)</th>
                    <th className="border border-black py-0 px-0.5 text-[7px] leading-tight font-bold">Jumlah tiap-<br/>tiap<br/>kedatangan</th>
                    <th className="border border-black py-0 px-0.5 text-[7px] leading-tight font-bold">Elaun<br/>gantian<br/>tetap<br/>basikal</th>
                  </tr>
                </thead>
                <tbody className="text-[12px] print:text-[9px]">
                  {Array.from({ length: 15 }).map((_, idx) => {
                    if (idx < 10) {
                      const noBadan = selectedNoBadanList[idx] || '';
                      const { dailyHours, name, rank, totalHours } = getPersonnelDailyHours(noBadan);
                      const rate = getRate(rank);
                      const cappedHours = Math.min(totalHours, 24);
                      const allowance = cappedHours * rate;
                      const kedatangan = dailyHours.filter(h => h > 0).length;
                      
                      grandTotalElaun += allowance;
                      grandTotalHoursWorked += totalHours;
                      grandTotalCappedHours += cappedHours;
                      grandTotalPenugasan += kedatangan;

                      return (
                        <tr key={idx} className="h-[28px] print:h-[18px]">
                          <td className="border border-black p-1">{idx + 1}</td>
                          <td className="border border-black p-0 relative">
                            <input 
                              type="text" 
                              list="personnel-list"
                              value={noBadan}
                              onChange={(e) => {
                                const newList = [...selectedNoBadanList];
                                newList[idx] = e.target.value;
                                setSelectedNoBadanList(newList);
                              }}
                              className="w-full h-full text-center bg-transparent border-none outline-none print:hidden font-bold"
                            />
                            <span className="hidden print:block w-full text-center font-bold">{noBadan}</span>
                          </td>
                          <td className="border border-black p-1">{rank}</td>
                          <td className="border border-black p-1 text-left truncate print:text-[8px] leading-tight max-w-[150px]">{name || (noBadan ? 'NOT FOUND' : '')}</td>
                          {dailyHours.map((h, i) => (
                            <td key={i} className="border border-black p-0 text-[10px]">{h || ''}</td>
                          ))}
                          <td className="border border-black p-1">{kedatangan || ''}</td>
                          <td className="border border-black p-1">{totalHours || ''}</td>
                          <td className="border border-black p-1">{cappedHours || ''}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1">{noBadan ? rate.toFixed(2) : ''}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1">{noBadan ? allowance.toFixed(2) : ''}</td>
                          <td className="border border-black p-1 font-bold">{noBadan ? allowance.toFixed(2) : ''}</td>
                          <td className="border border-black p-1"></td>
                        </tr>
                      );
                    } else if (idx < 13) {
                      // Empty rows 11, 12, 13
                      return (
                        <tr key={idx} className="h-[28px] print:h-[18px]">
                          <td className="border border-black p-1">{idx + 1}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          {Array(31).fill(0).map((_, i) => <td key={i} className="border border-black p-0 text-[10px]"></td>)}
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                        </tr>
                      );
                    } else if (idx === 13) {
                      // RINGGIT Row (Bil 14)
                      return (
                        <tr key={idx} className="h-[28px] print:h-[18px] font-bold">
                          <td className="border border-black p-1">{idx + 1}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1 text-left pl-4 uppercase" colSpan={38}>
                            RINGGIT : {numberToMalayWords(grandTotalElaun)}
                          </td>
                          <td className="border border-black p-1">{grandTotalElaun.toFixed(2)}</td>
                          <td className="border border-black p-1 font-bold">{grandTotalElaun.toFixed(2)}</td>
                          <td className="border border-black p-1"></td>
                        </tr>
                      );
                    } else {
                      // Final Total Row (Bil 15)
                      return (
                        <tr key={idx} className="h-[28px] print:h-[18px] font-bold">
                          <td className="border border-black p-1">{idx + 1}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          {Array(31).fill(0).map((_, i) => {
                            let dayTotal = 0;
                            selectedNoBadanList.forEach(nb => {
                              const { dailyHours } = getPersonnelDailyHours(nb);
                              dayTotal += dailyHours[i] || 0;
                            });
                            return <td key={i} className="border border-black p-0 text-[10px]">{dayTotal || ''}</td>;
                          })}
                          <td className="border border-black p-1 text-[10px] font-bold">{grandTotalPenugasan || ''}</td>
                          <td className="border border-black p-1 text-[10px] font-bold">{grandTotalHoursWorked || ''}</td>
                          <td className="border border-black p-1 text-[10px] font-bold">{grandTotalCappedHours || ''}</td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1"></td>
                          <td className="border border-black p-1">{grandTotalElaun.toFixed(2)}</td>
                          <td className="border border-black p-1 font-bold">{grandTotalElaun.toFixed(2)}</td>
                          <td className="border border-black p-1"></td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
              <datalist id="personnel-list">
                {processedData.districtPersonnel
                  .filter(p => {
                    const num = String(p.noBadan).replace(/[^0-9]/g, '');
                    return !!num && !selectedNoBadanList.includes(num);
                  })
                  .map((p, i) => {
                    const num = String(p.noBadan).replace(/[^0-9]/g, '');
                    // Show name and district in the label to help user identify the person
                    return num ? <option key={i} value={num}>{p.name} ({p.latestDistrict})</option> : null;
                  })}
              </datalist>
            </div>

            <div className="mt-2 print:mt-1 text-[9px] font-bold uppercase">
              JUMLAH JAM KEDATANGAN BAGI TIAP-TIAP KAWAD
            </div>

            <div className="mt-2 print:mt-1 grid grid-cols-2 gap-4 text-[10px] print:text-[9px]">
              <div className="space-y-4 print:space-y-2">
                <div className="font-bold">(Pegawai Simpanan yang diwartakan atau Inspektor)</div>
                <div className="pt-4 print:pt-2 w-64"></div>
              </div>
              <div className="space-y-1">
                <div className="font-bold">PERINGATAN A : Elaun belania latihan yang terbanyak dalam tiap-tiap bulan ialah mengenai latihan/tugas 24 jam</div>
                <div className="font-bold">PERINGATAN B : Elaun kenderaan ialah satu daripada berikut:</div>
                <div className="grid grid-cols-3 gap-2">
                  <div>(a) Elaun kenderaan (E.K) tetap basikal</div>
                  <div>(b) Elaun Hitungan Batu yang dibenarkan atau</div>
                  <div>(c) Elaun ganti</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Report 2: Voucher */}
        <div className="html2pdf__page-break"></div>
        <div className="print-page-container pt-[188px] print:pt-0 border-t-2 border-dashed border-gray-300 print:border-none">
          <div className="flex justify-between items-start mb-6">
            <div className="text-xs font-bold underline">SSPDRM MELAKA - BAUCER NO :</div>
            <div className="text-xs font-bold">SPDRM MELAKA BR.NO: ...........................</div>
          </div>

          <table className="w-full border-collapse border border-black text-xs text-center font-bold">
            <thead>
              <tr className="bg-gray-50 h-[27px]">
                <th className="border border-black p-1 w-12">BIL</th>
                <th className="border border-black p-1 w-32">NO KOD PVR</th>
                <th className="border border-black p-1">NAMA</th>
                <th className="border border-black p-1 w-40">NO AKAUN BANK</th>
                <th className="border border-black p-1 w-48">NAMA BANK</th>
                <th className="border border-black p-1 w-32">NO TELEFON</th>
                <th className="border border-black p-1 w-32">JUMLAH ( RM )</th>
              </tr>
            </thead>
            <tbody>
              {selectedNoBadanList.map((noBadan, idx) => {
                const { name, rank, totalHours } = getPersonnelDailyHours(noBadan);
                const rate = getRate(rank);
                const cappedHours = Math.min(totalHours, 24);
                const allowance = cappedHours * rate;
                const targetNo = noBadan.replace(/[^0-9]/g, '');
                
                // Find extra info from currentVoucherData (the other sheet)
                const extra = targetNo ? (currentVoucherData.find(v => {
                  const vNo = String(v['No Badan'] || v['NO BADAN'] || v['NO KOD PVR'] || '').replace(/[^0-9]/g, '');
                  return vNo === targetNo;
                }) || {}) : {};

                return (
                  <tr key={idx} className="h-[30px] break-inside-avoid">
                    <td className="border border-black p-1">{idx + 1}</td>
                    <td className="border border-black p-2">{extra['NO KOD PVR'] || ''}</td>
                    <td className="border border-black p-2 text-center">{name}</td>
                    <td className="border border-black p-2">{extra['NO AKAUN BANK'] || ''}</td>
                    <td className="border border-black p-2">{extra['NAMA BANK'] || ''}</td>
                    <td className="border border-black p-2">{extra['NO TELEFON'] || ''}</td>
                    <td className="border border-black p-2">
                      {noBadan ? `RM ${allowance.toFixed(2)}` : ''}
                    </td>
                  </tr>
                );
              })}
              <tr className="h-[30px] font-bold bg-gray-50">
                <td className="border border-black p-1 text-right pr-4" colSpan={6}>JUMLAH</td>
                <td className="border border-black p-1">RM {grandTotalElaun.toFixed(2)}</td>
              </tr>
            </tbody>
          </table>

          <div className="mt-8 grid grid-cols-2 gap-12 text-xs font-bold">
            <div>
              <div>Tandatangan Pegawai Bahagian & Cop</div>
              <div className="mt-[120px] border-t border-black border-dotted w-full"></div>
            </div>
            <div>
              <div>Tandatangan Komandan/Ejutan & Cop</div>
              <div className="mt-[120px] border-t border-black border-dotted w-full"></div>
            </div>
          </div>
          
        </div>
      </React.Fragment>
    );
  };

  const renderPaymentByYearTable = () => {
    const data = paymentByYearData;
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];

    const formatRM = (val: number) => {
      if (val === 0) return '0.00';
      return val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    const formatBil = (val: number) => {
      return val === 0 ? '0' : val.toString();
    };

    return (
      <div className="w-full max-w-7xl mx-auto bg-white p-4 sm:p-8">
        {/* Yellow Header Banners */}
        <div className="border-2 border-black mb-1 bg-[#ffff00] text-black text-center py-2 font-bold text-xl uppercase tracking-wider">
          KONTINJEN : MELAKA
        </div>
        <div className="border-2 border-black bg-[#ffff00] text-black text-center py-2 font-bold text-lg uppercase tracking-wide flex justify-between px-8">
          <span className="w-1/3 text-left">TUNTUTAN ELAUN PENUGASAN SSPDRM - </span>
          <span className="w-1/3 text-center text-2xl font-black">{selectedYear}</span>
          <span className="w-1/3"></span>
        </div>

        {/* The Main Table */}
        <table className="w-full border-collapse border-2 border-black text-center font-bold text-xs sm:text-sm mt-4">
          <thead>
            <tr className="bg-[#404040] text-white text-xs">
              <th className="border border-black p-3 w-12 text-center">BIL</th>
              <th className="border border-black p-3 w-48 text-left uppercase">FORMASI</th>
              <th className="border border-black p-3 w-28 text-center">PANGKAT</th>
              <th className="border border-black p-3 w-40 bg-[#166534] text-center">BILANGAN ANGGOTA (&gt;= 24 JAM)</th>
              <th className="border border-black p-3 bg-[#eab308] text-black text-center">JUMLAH TUNTUTAN ELAUN (RM)<br/>(MAKSIMUM 24 JAM)</th>
            </tr>
          </thead>
          <tbody>
            {districtsList.flatMap((district, dIdx) => {
              const dData = data.districts[district] || {
                PEG: { total: { bil:0, rm:0 } },
                APR: { total: { bil:0, rm:0 } },
                JUMLAH: { total: { bil:0, rm:0 } }
              };

              const roles = ['PEG', 'APR', 'JUMLAH'];

              return roles.map((role, rIdx) => {
                const rData = dData[role] || { total: { bil: 0, rm: 0 } };
                const isJumlahRow = role === 'JUMLAH';
                const rowBg = isJumlahRow ? 'bg-[#ffff99]' : 'bg-white';

                return (
                  <tr key={`${district}-${role}`} className={`${rowBg} even:bg-opacity-80 text-xs`}>
                    {/* Render BIL & FORMASI on the first sub-row (PEG) of each district */}
                    {rIdx === 0 && (
                      <>
                        <td className="border border-black p-2 text-center" rowSpan={3}>{dIdx + 1}</td>
                        <td className="border border-black p-2 text-left uppercase font-black" rowSpan={3}>{district}</td>
                      </>
                    )}

                    {/* Role / Pangkat column */}
                    <td className="border border-black p-2 uppercase text-center">{role}</td>

                    {/* Keseluruhan BIL & RM */}
                    <td className="border border-black p-2 bg-[#ffffcc] text-black font-black text-center">{formatBil(rData.total.bil)}</td>
                    <td className="border border-black p-2 bg-[#ffffcc] text-right font-black text-black">{formatRM(rData.total.rm)}</td>
                  </tr>
                );
              });
            })}

            {/* JUMLAH KESELURUHAN (Bottom section) */}
            {['PEG', 'APR', 'JUMLAH'].map((role, rIdx) => {
              const rData = data.keseluruhan[role] || { total: { bil: 0, rm: 0 } };
              const isJumlahRow = role === 'JUMLAH';
              const rowBg = isJumlahRow ? 'bg-[#ffff00] text-black font-extrabold text-sm' : 'bg-[#ffff99] text-black font-bold text-xs';

              return (
                <tr key={`keseluruhan-${role}`} className={`${rowBg}`}>
                  {rIdx === 0 && (
                    <td className="border border-black p-2 uppercase text-center font-black" colSpan={2} rowSpan={3}>
                      JUMLAH KESELURUHAN
                    </td>
                  )}
                  {/* Role / Pangkat column */}
                  <td className="border border-black p-2 uppercase text-center">{role}</td>

                  {/* Keseluruhan BIL & RM */}
                  <td className="border border-black p-2 bg-[#ffff00] font-black text-center">{formatBil(rData.total.bil)}</td>
                  <td className="border border-black p-2 bg-[#ffff00] text-right font-black">{formatRM(rData.total.rm)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Instructions Banner Footer */}
        <div className="border border-black bg-[#ffffea] text-black text-left pl-4 py-2 mt-4 font-bold text-xs uppercase tracking-wider">
          SILA ISIKAN BUTIRAN MENGGUNAKAN FORMAT YANG TELAH DISEDIAKAN MENGIKUT KONTINJEN
        </div>
      </div>
    );
  };

  const renderForecastTable = () => {
    const data = monthlyForecastData;
    const districtsList = ['IPK SSPDRM', 'MELAKA TENGAH', 'ALOR GAJAH', 'JASIN'];

    const formatRM = (val: number) => {
      if (val === 0) return '0.00';
      return val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    const formatBil = (val: number) => {
      return val === 0 ? '0' : val.toString();
    };

    return (
      <div className="w-full max-w-7xl mx-auto bg-white p-2 sm:p-6 select-none font-sans print:p-0">
        {/* Yellow Header Banners */}
        <div className="border border-black mb-1 bg-[#ffff00] text-black text-center py-2 font-bold text-lg sm:text-xl uppercase tracking-wider">
          KONTINJEN : MELAKA
        </div>
        <div className="border border-black bg-[#ffff00] text-black text-center py-2 font-bold text-sm sm:text-base uppercase tracking-wide flex justify-between px-4 sm:px-8">
          <span className="w-1/3 text-left">TUNTUTAN ELAUN PENUGASAN SSPDRM - </span>
          <span className="w-1/3 text-center text-xl font-black">{selectedMonth} {selectedYear}</span>
          <span className="w-1/3"></span>
        </div>

        {/* The Main Table with exact colors from photo */}
        <table className="w-full border-collapse border border-black text-center font-bold text-[11px] sm:text-xs mt-4">
          <thead>
            {/* Header Row 1 */}
            <tr className="bg-[#ebd300] text-black text-[10px] sm:text-[11px] h-[34px]">
              <th className="border border-black p-1 w-10 text-center" rowSpan={2}>BIL</th>
              <th className="border border-black p-1 w-40 text-left pl-2" rowSpan={2}>FORMASI</th>
              <th className="border border-black p-1 w-24 text-center" rowSpan={2}>PANGKAT</th>
              <th className="border border-black p-1 bg-[#3a60a1] text-white w-20 text-center font-bold" colSpan={1}>JUMLAH JAM BERTUGAS</th>
              <th className="border border-black p-1 bg-[#1e7c53] text-white text-center font-bold" colSpan={2}>24 (A)</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white text-center font-bold" colSpan={2}>49 - 96 (B)</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white text-center font-bold" colSpan={2}>97 - 128 (C)</th>
              <th className="border border-black p-1 bg-[#cca300] text-black text-center font-bold" colSpan={2}>KESELURUHAN</th>
            </tr>
            {/* Header Row 2 */}
            <tr className="text-black text-[10px] sm:text-[11px]">
              <th className="border border-black p-1 bg-[#4f81bd] text-white w-20 text-center font-bold">1 - 23 JAM<br/>BIL</th>
              <th className="border border-black p-1 bg-[#375623] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#375623] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#4a154b] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white w-14 text-center font-bold">BIL</th>
              <th className="border border-black p-1 bg-[#7b113a] text-white w-24 text-center font-bold">RM</th>
              <th className="border border-black p-1 bg-[#ffc000] text-black w-14 text-center font-bold">Jumlah (A+B+C)</th>
              <th className="border border-black p-1 bg-[#ffc000] text-black w-24 text-center font-bold">RM</th>
            </tr>
          </thead>
          <tbody>
            {districtsList.flatMap((district, dIdx) => {
              const dData = data.districts[district] || {
                PEG: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
                APR: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } },
                JUMLAH: { bracket1_23: { bil: 0, rm: 0 }, bracket24_48: { bil: 0, rm: 0 }, bracket49_96: { bil: 0, rm: 0 }, bracket97_128: { bil: 0, rm: 0 }, total: { bil: 0, rm: 0 } }
              };

              const roles = ['PEG', 'APR', 'JUMLAH'];

              return roles.map((role, rIdx) => {
                const rData = dData[role] || {
                  bracket1_23: { bil: 0, rm: 0 },
                  bracket24_48: { bil: 0, rm: 0 },
                  bracket49_96: { bil: 0, rm: 0 },
                  bracket97_128: { bil: 0, rm: 0 },
                  total: { bil: 0, rm: 0 }
                };
                const isJumlahRow = role === 'JUMLAH';
                
                // Color formatting matching the photo
                const rowBg = isJumlahRow ? 'bg-[#ffc000] text-black font-extrabold' : 'bg-[#d9e1f2] text-black font-medium';

                // Bracket columns backgrounds
                const colE_Bg = isJumlahRow ? 'bg-[#ffc000]' : 'bg-[#b4c6e7]';
                const colFG_Bg = isJumlahRow ? 'bg-[#ffc000]' : 'bg-[#e2efda]';
                const colHI_Bg = isJumlahRow ? 'bg-[#ffc000]' : 'bg-[#f2dbdb]';
                const colJK_Bg = isJumlahRow ? 'bg-[#ffc000]' : 'bg-[#eec5d3] text-black';
                const colLM_Bg = 'bg-[#fff2cc]';

                return (
                  <tr key={`${district}-${role}`} className={`${rowBg} h-[28px] text-[11px]`}>
                    {/* Render BIL & FORMASI on the first sub-row (PEG) of each district */}
                    {rIdx === 0 && (
                      <>
                        <td className="border border-black p-1 text-center font-bold bg-white text-black" rowSpan={3}>{dIdx + 1}</td>
                        <td className="border border-black p-1 text-left uppercase font-bold bg-white text-black pl-2" rowSpan={3}>{district}</td>
                      </>
                    )}

                    {/* Role / Pangkat column */}
                    <td className="border border-black p-1 uppercase text-center font-bold">{role}</td>

                    {/* 1-23 JAM Column */}
                    <td className={`border border-black p-1 text-center font-bold ${colE_Bg}`}>
                      {formatBil(rData.bracket1_23.bil)}
                    </td>

                    {/* 24-48 (A) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colFG_Bg}`}>
                      {formatBil(rData.bracket24_48.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colFG_Bg}`}>
                      {formatRM(rData.bracket24_48.rm)}
                    </td>

                    {/* 49-96 (B) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colHI_Bg}`}>
                      {formatBil(rData.bracket49_96.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colHI_Bg}`}>
                      {formatRM(rData.bracket49_96.rm)}
                    </td>

                    {/* 97-128 (C) Columns */}
                    <td className={`border border-black p-1 text-center font-bold ${colJK_Bg}`}>
                      {formatBil(rData.bracket97_128.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-bold ${colJK_Bg}`}>
                      {formatRM(rData.bracket97_128.rm)}
                    </td>

                    {/* KESELURUHAN (A+B+C) Columns */}
                    <td className={`border border-black p-1 text-center font-black ${colLM_Bg} text-black`}>
                      {formatBil(rData.total.bil)}
                    </td>
                    <td className={`border border-black p-1 text-right pr-2 font-black ${colLM_Bg} text-black`}>
                      {formatRM(rData.total.rm)}
                    </td>
                  </tr>
                );
              });
            })}

            {/* JUMLAH KESELURUHAN (Bottom yellow section) */}
            {['PEG', 'APR', 'JUMLAH'].map((role, rIdx) => {
              const rData = data.keseluruhan[role] || {
                bracket1_23: { bil: 0, rm: 0 },
                bracket24_48: { bil: 0, rm: 0 },
                bracket49_96: { bil: 0, rm: 0 },
                bracket97_128: { bil: 0, rm: 0 },
                total: { bil: 0, rm: 0 }
              };
              const isJumlahRow = role === 'JUMLAH';
              
              const rowBg = isJumlahRow ? 'bg-[#c6e0b4] text-black font-extrabold text-xs' : 'bg-[#e2efda] text-black font-bold text-[11px]';

              return (
                <tr key={`keseluruhan-${role}`} className={`${rowBg} h-[30px]`}>
                  {rIdx === 0 && (
                    <td className="border border-black p-1 uppercase text-left font-black bg-[#ffff00] text-black pl-4" colSpan={2} rowSpan={3}>
                      JUMLAH KESELURUHAN
                    </td>
                  )}
                  {/* Role / Pangkat column */}
                  <td className="border border-black p-1 uppercase text-center font-black bg-[#ffff00] text-black">{role}</td>

                  {/* 1-23 JAM */}
                  <td className="border border-black p-1 font-black bg-[#4f81bd] text-white text-center">
                    {formatBil(rData.bracket1_23.bil)}
                  </td>

                  {/* 24-48 (A) */}
                  <td className="border border-black p-1 font-black bg-[#375623] text-white text-center">
                    {formatBil(rData.bracket24_48.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#375623] text-white text-right pr-2">
                    {formatRM(rData.bracket24_48.rm)}
                  </td>

                  {/* 49-96 (B) */}
                  <td className="border border-black p-1 font-black bg-[#4a154b] text-white text-center">
                    {formatBil(rData.bracket49_96.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#4a154b] text-white text-right pr-2">
                    {formatRM(rData.bracket49_96.rm)}
                  </td>

                  {/* 97-128 (C) */}
                  <td className="border border-black p-1 font-black bg-[#7b113a] text-white text-center">
                    {formatBil(rData.bracket97_128.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#7b113a] text-white text-right pr-2">
                    {formatRM(rData.bracket97_128.rm)}
                  </td>

                  {/* KESELURUHAN (A+B+C) */}
                  <td className="border border-black p-1 font-black bg-[#ffc000] text-black text-center">
                    {formatBil(rData.total.bil)}
                  </td>
                  <td className="border border-black p-1 font-black bg-[#ffc000] text-black text-right pr-2">
                    {formatRM(rData.total.rm)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Instructions Banner Footer */}
        <div className="border border-black bg-[#ffffea] text-black text-left pl-4 py-2 mt-4 font-bold text-xs uppercase tracking-wider">
          SILA ISIKAN BUTIRAN MENGGUNAKAN FORMAT YANG TELAH DISEDIAKAN MENGIKUT KONTINJEN
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gray-100 p-4 sm:p-8 print:p-0 print:bg-white">
      <style>{`
        @media print {
          @page {
            size: A4 landscape;
            margin: 10mm;
          }
          body {
            -webkit-print-color-adjust: exact;
            background-color: white !important;
          }
          .print-page-container {
            page-break-after: always !important;
            page-break-inside: avoid !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            box-sizing: border-box !important;
            display: block !important;
            height: auto !important;
          }
          .print-page-container:last-child {
            page-break-after: auto !important;
          }
          #report-container {
            border: none !important;
            box-shadow: none !important;
            padding: 0 !important;
            width: 100% !important;
            max-width: none !important;
            margin: 0 !important;
          }
          table {
            width: 100% !important;
          }
          th, td {
            word-wrap: break-word !important;
          }
        }
        
        /* PDF Generation Compact Mode */
        .pdf-compact-mode {
          width: 1400px !important; /* Expanded width to prevent cutting off table */
          padding: 15px !important;
          background: white !important;
          border: none !important;
          box-shadow: none !important;
          margin: 0 auto !important;
        }
        .pdf-compact-mode * {
          box-shadow: none !important;
        }
        .pdf-compact-mode .print\\:hidden,
        .pdf-compact-mode [class*="print:hidden"] {
          display: none !important;
        }
        .pdf-compact-mode [class*="print:pt-0"] {
          padding-top: 0 !important;
        }
        .pdf-compact-mode [class*="print:pt-2"] {
          padding-top: 0.5rem !important;
        }
        .pdf-compact-mode [class*="print:pb-0"] {
          padding-bottom: 0 !important;
        }
        .pdf-compact-mode [class*="print:p-0"] {
          padding: 0 !important;
        }
        .pdf-compact-mode [class*="print:-mt-[45px]"] {
          margin-top: -45px !important;
        }
        .pdf-compact-mode [class*="print:mt-1"] {
          margin-top: 0.25rem !important;
        }
        .pdf-compact-mode [class*="print:border-none"] {
          border: none !important;
        }
        .pdf-compact-mode [class*="print:bg-transparent"] {
          background-color: transparent !important;
        }
        .pdf-compact-mode [class*="print:bg-white"] {
          background-color: white !important;
        }
        .pdf-compact-mode [class*="print:shadow-none"] {
          box-shadow: none !important;
        }
        .pdf-compact-mode [class*="print:max-w-none"] {
          max-width: none !important;
        }
        .pdf-compact-mode [class*="print:h-[18px]"] {
          height: 18px !important;
        }
        .pdf-compact-mode [class*="print:text-black"] {
          color: black !important;
        }
        .pdf-compact-mode [class*="print:text-[8px]"] { font-size: 8px !important; line-height: 8px !important; }
        .pdf-compact-mode [class*="print:text-[9px]"] { font-size: 9px !important; line-height: 9px !important; }
        .pdf-compact-mode [class*="print:text-[11px]"] { font-size: 11px !important; line-height: 11px !important; }
        .pdf-compact-mode [class*="print:text-[14px]"] { font-size: 14px !important; line-height: 14px !important; }
        .pdf-compact-mode [class*="print:text-[16px]"] { font-size: 16px !important; line-height: 16px !important; }
        .pdf-compact-mode [class*="print:text-[26px]"] { font-size: 26px !important; line-height: 26px !important; }
        .pdf-compact-mode [class*="print:text-[30px]"] { font-size: 30px !important; line-height: 30px !important; }
        .pdf-compact-mode [class*="print:gap-4"] { gap: 1rem !important; }
        .pdf-compact-mode [class*="print:space-y-2"] > :not([hidden]) ~ :not([hidden]) {
          --tw-space-y-reverse: 0;
          margin-top: calc(0.5rem * calc(1 - var(--tw-space-y-reverse))) !important;
          margin-bottom: calc(0.5rem * var(--tw-space-y-reverse)) !important;
        }
        .pdf-compact-mode .overflow-x-auto {
          overflow: visible !important;
          overflow-x: visible !important;
        }
        .pdf-compact-mode .page-break-before {
          page-break-before: always !important;
          break-before: page !important;
        }
        .pdf-compact-mode .print-page-container {
          margin-bottom: 0 !important;
          padding: 0 !important;
        }
        .pdf-compact-mode .print-page-container:last-child {
          page-break-after: auto !important;
        }
        .pdf-compact-mode table {
          width: 100% !important;
        }
        .pdf-compact-mode th, .pdf-compact-mode td {
          word-wrap: break-word !important;
        }
      `}</style>
      {/* Controls - Hidden when printing */}
      <div className="max-w-7xl mx-auto mb-6 bg-white p-6 rounded-xl shadow-sm border border-gray-200 print:hidden">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h1 className="text-2xl font-bold text-gray-900 flex items-center gap-2">
              <FileSpreadsheet className="w-6 h-6 text-blue-600" />
              SSPDRM Dashboard
            </h1>
            <div className="flex items-center gap-3 mt-2">
              <p className="text-sm text-gray-500">
                Select parameters and view type to generate the report.
              </p>
              {(GOOGLE_SHEET_ID && GOOGLE_SHEET_ID !== "YOUR_GOOGLE_SHEET_ID_HERE") ? (
                <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full bg-green-100 text-green-700 text-xs font-medium">
                  <CheckCircle2 className="w-3 h-3" /> Connected to Sheets
                </span>
              ) : (
                <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full bg-amber-100 text-amber-700 text-xs font-medium">
                  <AlertCircle className="w-3 h-3" /> Using Mock Data
                </span>
              )}
            </div>
          </div>
          
          <div className="flex flex-wrap items-center gap-3">
            <select 
              value={selectedDistrict}
              onChange={(e) => setSelectedDistrict(e.target.value)}
              disabled={userRole.toLowerCase() !== 'admin'}
              className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none disabled:opacity-75 disabled:bg-gray-100"
            >
              {userRole.toLowerCase() === 'admin' ? (
                districts.map(d => <option key={d} value={d}>{d}</option>)
              ) : (
                <option value={userDistrict}>{userDistrict}</option>
              )}
            </select>
            
            <select 
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
            >
              {months.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
            
            {activeTab === 'PERSONAL' ? (
              <div className="flex items-center gap-1">
                <span className="text-sm font-medium text-gray-700">Tahun Dari:</span>
                <select 
                  value={selectedYearFrom}
                  onChange={(e) => setSelectedYearFrom(Number(e.target.value))}
                  className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
                >
                  {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
                <span className="text-sm font-medium text-gray-700 ml-1">Hingga:</span>
                <select 
                  value={selectedYear}
                  onChange={(e) => setSelectedYear(Number(e.target.value))}
                  className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
                >
                  {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
            ) : (
              <select 
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
                className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              >
                {years.map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            )}
            
            {activeTab === 'PERSONAL' && (
              <>
                <input
                  type="text"
                  placeholder="Carian No Badan / Nama..."
                  value={searchNoBadan}
                  onChange={(e) => {
                    setSearchNoBadan(e.target.value);
                    if (e.target.value !== '' && selectedPerson !== 'ALL') {
                      setSelectedPerson('ALL');
                    }
                  }}
                  className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none w-48"
                />
                <select 
                  value={selectedPerson}
                  onChange={(e) => setSelectedPerson(e.target.value)}
                  className="px-3 py-2 bg-gray-50 border border-gray-300 rounded-lg text-sm font-medium focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none max-w-[200px] truncate"
                >
                  <option value="ALL">SEMUA ANGGOTA</option>
                  {processedData.districtPersonnel.map(p => (
                    <option key={`${p.noBadan}-${p.name}`} value={p.name}>{p.name} ({p.latestDistrict})</option>
                  ))}
                </select>
              </>
            )}

            {(GOOGLE_SHEET_ID && GOOGLE_SHEET_ID !== "YOUR_GOOGLE_SHEET_ID_HERE") && (
              <button 
                onClick={() => {
                  const fromYear = activeTab === 'PERSONAL' ? selectedYearFrom : selectedYear;
                  fetchSheetData(GOOGLE_SHEET_ID, fromYear, selectedYear);
                  fetchVoucherDataLive();
                }}
                disabled={isLoading}
                className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 rounded-lg text-sm font-medium transition-colors shadow-sm disabled:opacity-50"
              >
                <RefreshCw className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} />
                Refresh
              </button>
            )}
            
            <button 
              onClick={handleLogout}
              className="flex items-center gap-2 px-4 py-2 bg-red-50 border border-red-200 hover:bg-red-100 text-red-600 rounded-lg text-sm font-medium transition-colors shadow-sm"
            >
              Logout
            </button>

            <div className="flex flex-col items-end gap-1">
              <div className="flex gap-2">
                <button 
                  onClick={handlePrintCurrent}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm"
                >
                  <Printer className="w-4 h-4" />
                  Print Current
                </button>
                <button 
                  onClick={handlePrintAll}
                  className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm"
                >
                  <Printer className="w-4 h-4" />
                  Print All Reports
                </button>
                <button 
                  onClick={handleSaveCurrentPDF}
                  disabled={isGeneratingPDF}
                  className={`flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm ${isGeneratingPDF ? 'opacity-50 cursor-not-allowed' : ''}`}
                >
                  {isGeneratingPDF && printMode === 'CURRENT' ? (
                    <RefreshCw className="w-4 h-4 animate-spin" />
                  ) : (
                    <Download className="w-4 h-4" />
                  )}
                  {isGeneratingPDF && printMode === 'CURRENT' ? 'Generating...' : 'Save Current PDF'}
                </button>
                <button 
                  onClick={handleSaveAllPDF}
                  disabled={isGeneratingPDF}
                  className={`flex items-center gap-2 px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm ${isGeneratingPDF ? 'opacity-50 cursor-not-allowed' : ''}`}
                >
                  {isGeneratingPDF && printMode === 'ALL' ? (
                    <RefreshCw className="w-4 h-4 animate-spin" />
                  ) : (
                    <Download className="w-4 h-4" />
                  )}
                  {isGeneratingPDF && printMode === 'ALL' ? 'Generating...' : 'Save All to PDF'}
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Tabs */}
        <div className="flex gap-2 mt-6 border-b border-gray-200 pb-px">
          {(userRole.toLowerCase() === 'admin' || 
            userTab.toUpperCase().includes('T1') || userTab.toUpperCase().includes('WEEKLY') ||
            userTab.toUpperCase().includes('T2') || userTab.toUpperCase().includes('DAILY') ||
            userTab.toUpperCase().includes('T3') || userTab.toUpperCase().includes('RANK') ||
            !userTab || userTab.trim() === '') && (
            <button
              onClick={() => setActiveTab('MONTHLY')}
              className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
                activeTab === 'MONTHLY' 
                  ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                  : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
              }`}
            >
              <CalendarDays className="w-4 h-4" />
              Laporan Bulanan (Monthly Report)
            </button>
          )}
          {(userRole.toLowerCase() === 'admin' || userTab.toUpperCase().includes('T4') || userTab.toUpperCase().includes('PERSONAL') || !userTab || userTab.trim() === '') && (
            <button
              onClick={() => setActiveTab('PERSONAL')}
              className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
                activeTab === 'PERSONAL' 
                  ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                  : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
              }`}
            >
              <User className="w-4 h-4" />
              Jam Penugasan (Tahunan)
            </button>
          )}
          <button
            onClick={() => setActiveTab('ALLOWANCE')}
            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
              activeTab === 'ALLOWANCE' 
                ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
            }`}
          >
            <Download className="w-4 h-4" />
            Paysheet (backup data)
          </button>
          <button
            onClick={() => setActiveTab('ALLOWANCE_LIVE')}
            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
              activeTab === 'ALLOWANCE_LIVE' 
                ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
            }`}
          >
            <Download className="w-4 h-4" />
            Paysheet(live data)
          </button>
          <button
            onClick={() => setActiveTab('PENYALUR_MAKLUMAT')}
            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
              activeTab === 'PENYALUR_MAKLUMAT' 
                ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
            }`}
          >
            <FileText className="w-4 h-4" />
            Penyalur Maklumat
          </button>
          <button
            onClick={() => setActiveTab('FORECAST')}
            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-t-lg transition-colors ${
              activeTab === 'FORECAST' 
                ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
                : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
            }`}
          >
            <TrendingUp className="w-4 h-4" />
            Payment Forecast
          </button>
        </div>
      </div>

      {/* Printable Report Area */}
      {(!GOOGLE_SHEET_ID || GOOGLE_SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE") && (
        <div className="max-w-7xl mx-auto mb-6 bg-amber-50 border border-amber-200 rounded-xl p-6 text-center print:hidden shadow-sm">
          <Database className="w-10 h-10 text-amber-500 mx-auto mb-3" />
          <h2 className="text-lg font-bold text-amber-900 mb-2">Currently showing sample data</h2>
          <p className="text-amber-700 mb-4 max-w-2xl mx-auto">
            Please configure your GOOGLE_SHEET_ID in the code to view and manage your actual SSPDRM volunteer data.
          </p>
        </div>
      )}

      <div id="report-container" className={`max-w-[1400px] mx-auto bg-white print:max-w-none print:shadow-none shadow-lg border-4 border-blue-600 print:border-none p-4 sm:p-8 print:p-0 overflow-x-auto ${isGeneratingPDF ? 'pdf-compact-mode' : ''}`}>
        
        {/* Data Tables */}
        <div className="w-full overflow-x-auto">
          {(printMode === 'ALL' || activeTab === 'MONTHLY') && (
            <>
              {/* Bulanan (Monthly) */}
              <div className="print-page-container">
                <div className="text-center mb-6">
                  <h2 className="text-xl sm:text-2xl font-bold uppercase tracking-wide text-gray-900">
                    SUKARELAWAN POLIS DIRAJA MALAYSIA KONTINJEN MELAKA
                  </h2>
                  <div className="text-sm sm:text-base font-semibold mt-1 uppercase">
                    BAHAGIAN PENTADBIRAN <span className="ml-2">{selectedDistrict}</span>
                  </div>
                  <div className="text-sm sm:text-base font-semibold mt-1">
                    PROGRAM / AKTIVITI PASUKAN (BULANAN)
                  </div>
                  <div className="text-sm sm:text-base font-semibold mt-1 uppercase">
                    BULAN : <span className="ml-2">{selectedMonth}</span> <span className="ml-4">{selectedYear}</span>
                  </div>
                </div>
                {renderDailyTable()}
              </div>

              {/* Pangkat (Rank) - Moved to 2nd page */}
              <div className="print-page-container">
                {renderRankTable()}
              </div>

              {/* Mingguan (Weekly) - Moved to 3rd page */}
              <div className="print-page-container">
                <div className="text-center mb-6">
                  <h2 className="text-xl sm:text-2xl font-bold uppercase tracking-wide text-gray-900">
                    SUKARELAWAN POLIS DIRAJA MALAYSIA KONTINJEN MELAKA
                  </h2>
                  <div className="text-sm sm:text-base font-semibold mt-1 uppercase">
                    BAHAGIAN PENTADBIRAN <span className="ml-2">{selectedDistrict}</span>
                  </div>
                  <div className="text-sm sm:text-base font-semibold mt-1">
                    PROGRAM / AKTIVITI PASUKAN (MINGGUAN)
                  </div>
                  <div className="text-sm sm:text-base font-semibold mt-1 uppercase">
                    BULAN : <span className="ml-2">{selectedMonth}</span> <span className="ml-4">{selectedYear}</span>
                  </div>
                </div>
                {renderWeeklyTable()}
              </div>
            </>
          )}

          {(printMode === 'ALL' || activeTab === 'PERSONAL') && (
            <div className="print-page-container">
              {activeTab === 'PERSONAL' && (
                <div className="mb-4 flex gap-4 print:hidden justify-end">
                  <button onClick={exportPersonalToExcel} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm cursor-pointer">
                    <FileSpreadsheet className="w-4 h-4" />
                    Muat Turun Excel (Tahunan)
                  </button>
                  <button onClick={() => window.print()} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm cursor-pointer">
                    <Printer className="w-4 h-4" />
                    Cetak Borang
                  </button>
                </div>
              )}
              <div className="text-center mb-6">
                <h2 className="text-xl sm:text-2xl font-bold uppercase tracking-wide text-gray-900">
                  SUKARELAWAN SIMPANAN POLIS DIRAJA MALAYSIA (SSPDRM)
                </h2>
                <div className="text-xl sm:text-2xl font-bold mt-1 uppercase">
                  KONTINJEN : <span className="ml-2">MELAKA</span>
                </div>
                <div className="text-xl sm:text-2xl font-bold mt-1 uppercase">
                  DAERAH : <span className="ml-2">{selectedDistrict}</span>
                </div>
                <div className="text-xl sm:text-2xl font-bold mt-4 uppercase">
                  JUMLAH JAM PENUGSAN BULANAN BAGI TAHUN <span className="ml-2 border-b border-black pb-1 px-4">
                    {selectedYearFrom !== selectedYear ? `${Math.min(selectedYearFrom, selectedYear)} - ${Math.max(selectedYearFrom, selectedYear)}` : selectedYear}
                  </span>
                </div>
              </div>
              {renderPersonalTable()}
            </div>
          )}

          {(printMode === 'ALL' || activeTab === 'ALLOWANCE') && (
            <React.Fragment>
              {renderAllowanceTable(false)}
            </React.Fragment>
          )}

          {(activeTab === 'ALLOWANCE_LIVE') && (
            <React.Fragment>
              {activeTab === 'ALLOWANCE_LIVE' && (
                <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg flex items-center justify-between print:hidden">
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full ${voucherDataLive.length > 0 ? 'bg-green-500' : 'bg-amber-500'}`}></div>
                    <div className="flex flex-col">
                      <span className="text-xs font-medium text-blue-800">Live Data Status: {liveDataStatus}</span>
                    </div>
                  </div>
                  <button 
                    onClick={fetchVoucherDataLive}
                    className="text-[10px] bg-blue-600 text-white px-2 py-1 rounded hover:bg-blue-700 transition-colors"
                  >
                    Refresh Live Data
                  </button>
                </div>
              )}
              {renderAllowanceTable(true)}
            </React.Fragment>
          )}

          {(printMode === 'ALL' || activeTab === 'PENYALUR_MAKLUMAT') && (
            <div className="print-page-container mt-4 pt-4 border-t border-gray-200">
              {activeTab === 'PENYALUR_MAKLUMAT' && (
                <div className="mb-4 flex gap-4 print:hidden justify-end">
                  <button onClick={exportToWord} className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm">
                    <Download className="w-4 h-4" />
                    Muat Turun (Word)
                  </button>
                  <button onClick={() => window.print()} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm">
                    <Printer className="w-4 h-4" />
                    Cetak Borang
                  </button>
                </div>
              )}
              {renderPenyalurMaklumat()}
            </div>
          )}

          {(printMode === 'ALL' || activeTab === 'FORECAST') && (
            <div className="print-page-container">
              {activeTab === 'FORECAST' && (
                <div className="mb-4 flex gap-4 print:hidden justify-end">
                  <button onClick={exportYearlyToExcel} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm cursor-pointer">
                    <FileSpreadsheet className="w-4 h-4" />
                    Muat Turun Excel (Tahunan)
                  </button>
                  <button onClick={() => window.print()} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm cursor-pointer">
                    <Printer className="w-4 h-4" />
                    Cetak Borang
                  </button>
                </div>
              )}
              {renderForecastTable()}
              {renderYearlyForecastTable()}
            </div>
          )}
        </div>
        
      </div>

      {/* Print Modal */}
      {showPrintModal && (
        <div className="fixed inset-0 flex items-center justify-center z-50 p-4 print:hidden" style={{ backgroundColor: 'rgba(0, 0, 0, 0.5)' }}>
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6">
            <h3 className="text-xl font-bold text-gray-900 mb-2">Printing Unavailable Here</h3>
            <p className="text-gray-600 mb-6">
              Because this app is running inside a preview window, the browser's print function is blocked. 
              <br/><br/>
              To print or save as PDF, please click the <strong>"Open in new tab"</strong> button at the top right of your screen (the square icon with an arrow pointing out), and try printing from there.
            </p>
            <div className="flex justify-end">
              <button 
                onClick={() => setShowPrintModal(false)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 rounded-lg font-medium transition-colors"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function getSamplePaymentByYear() {
  const structure: Record<string, any> = {
    'IPK SSPDRM': {
      PEG: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 11, rm: 4459 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 11, rm: 4459 }
      },
      APR: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 38, rm: 13248 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 38, rm: 13248 }
      },
      JUMLAH: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 49, rm: 17707 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 49, rm: 17707 }
      }
    },
    'MELAKA TENGAH': {
      PEG: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 22, rm: 9378.60 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 22, rm: 9378.60 }
      },
      APR: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 651, rm: 209000 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 651, rm: 209000 }
      },
      JUMLAH: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 673, rm: 218378.60 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 673, rm: 218378.60 }
      }
    },
    'ALOR GAJAH': {
      PEG: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 22, rm: 9996 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 22, rm: 9996 }
      },
      APR: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 219, rm: 74048 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 219, rm: 74048 }
      },
      JUMLAH: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 241, rm: 84044 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 241, rm: 84044 }
      }
    },
    'JASIN': {
      PEG: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 11, rm: 9996 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 11, rm: 9996 }
      },
      APR: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 167, rm: 74048 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 167, rm: 74048 }
      },
      JUMLAH: {
        bracket1_23: { bil: 0, rm: 0 },
        bracket24_48: { bil: 178, rm: 84044 },
        bracket49_96: { bil: 0, rm: 0 },
        bracket97_128: { bil: 0, rm: 0 },
        total: { bil: 178, rm: 84044 }
      }
    }
  };

  const keseluruhan = {
    PEG: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:66, rm:33829.60 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:66, rm:33829.60 } },
    APR: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:1075, rm:370344 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:1075, rm:370344 } },
    JUMLAH: { bracket1_23: { bil:0, rm:0 }, bracket24_48: { bil:1141, rm:404173.60 }, bracket49_96: { bil:0, rm:0 }, bracket97_128: { bil:0, rm:0 }, total: { bil:1141, rm:404173.60 } }
  };

  return { districts: structure, keseluruhan };
}