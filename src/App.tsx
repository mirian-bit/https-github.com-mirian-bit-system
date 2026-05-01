/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from "react";
import { 
  LayoutGrid, 
  Users, 
  Settings, 
  Calendar, 
  ChevronRight, 
  ChevronLeft,
  ChevronDown,
  MoreHorizontal,
  Play,
  Download,
  Upload,
  Plus,
  Trash2,
  CheckCircle2,
  AlertCircle,
  X,
  FileSpreadsheet,
  Clock,
  CalendarDays,
  RotateCcw,
  ShieldCheck,
  Award,
  UserPlus,
  Coffee
} from "lucide-react";
import { format, addDays, startOfDay, parseISO, isWeekend, isWithinInterval } from "date-fns";
import { zhTW } from "date-fns/locale";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "motion/react";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { 
  db, auth, googleProvider, 
  setDoc, doc, getDoc, getDocs, collection, query, where, onSnapshot, writeBatch, deleteDoc,
  signInWithPopup, onAuthStateChanged, type User
} from "./lib/firebase";
import { LeaveType, type Department, type Employee, type ShiftRecord, type ScheduleState, type DailyRequirement } from "./types";

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const STORAGE_KEY = "LOGISTICS_SCHEDULE_DATA";

// --- Constants ---
const ATTENDANCE_TIMES = [
  "無",
  "08:00-17:00",
  "09:00-18:00",
  "10:00-19:00",
  "13:00-22:00",
  "22:00-07:00",
  "自休",
  "系休",
  "國定"
];

const INITIAL_DEPTS: Department[] = [
  {
    id: "dept-1",
    name: "大溪理貨一課",
    skills: ["理貨", "驗收", "廠退", "事務"],
    groups: [
      { id: "g1-1", name: "日班理貨一組", maxLeavePerDay: 3, skillRequirements: [{ skill: "理貨", count: 5 }] },
      { id: "g1-2", name: "日班理貨二組", maxLeavePerDay: 3, skillRequirements: [{ skill: "理貨", count: 5 }] },
      { id: "g1-3", name: "事務組", maxLeavePerDay: 1, skillRequirements: [{ skill: "事務", count: 2 }] },
      { id: "g1-3b", name: "日班驗收組", maxLeavePerDay: 2, skillRequirements: [{ skill: "驗收", count: 3 }] },
      { id: "g1-4", name: "夜班驗收組", maxLeavePerDay: 2, skillRequirements: [{ skill: "驗收", count: 3 }] },
      { id: "g1-5", name: "EC廠退組", maxLeavePerDay: 2, skillRequirements: [{ skill: "廠退", count: 2 }] },
    ]
  },
  {
    id: "dept-2",
    name: "大溪理貨二課",
    skills: ["店訂", "退貨", "分揀", "POP"],
    groups: [
      { id: "g2-1", name: "日班店訂組", maxLeavePerDay: 3, skillRequirements: [{ skill: "店訂", count: 4 }] },
      { id: "g2-2", name: "日班退貨組", maxLeavePerDay: 2, skillRequirements: [{ skill: "退貨", count: 3 }] },
      { id: "g2-3", name: "中班分揀組", maxLeavePerDay: 3, skillRequirements: [{ skill: "分揀", count: 5 }] },
      { id: "g2-4", name: "POP組", maxLeavePerDay: 1, skillRequirements: [{ skill: "POP", count: 2 }] },
    ]
  },
  {
    id: "dept-3",
    name: "倉儲管理課",
    skills: ["庫存", "出貨", "廠退", "收發"],
    groups: [
      { id: "g3-1", name: "日班庫存組", maxLeavePerDay: 2, skillRequirements: [{ skill: "庫存", count: 3 }] },
      { id: "g3-2", name: "日班出貨組", maxLeavePerDay: 3, skillRequirements: [{ skill: "出貨", count: 5 }] },
      { id: "g3-3", name: "日班廠退組", maxLeavePerDay: 1, skillRequirements: [{ skill: "廠退", count: 2 }] },
      { id: "g3-4", name: "中班庫存組", maxLeavePerDay: 2, skillRequirements: [{ skill: "庫存", count: 3 }] },
      { id: "g3-5", name: "收發組", maxLeavePerDay: 1, skillRequirements: [{ skill: "收發", count: 2 }] },
    ]
  },
  {
    id: "dept-4",
    name: "運務課",
    skills: ["運務"],
    groups: [
      { id: "g4-1", name: "運務組", maxLeavePerDay: 5, skillRequirements: [{ skill: "運務", count: 10 }] },
    ]
  }
];

const INITIAL_EMPLOYEES: Employee[] = [
  // 範例資料，分配至新課組
  { id: "emp-1", name: "王理貨", deptId: "dept-1", groupId: "g1-1", skills: ["理貨"], nationalHolidayQuota: 2, systemLeaveQuota: 6 },
  { id: "emp-2", name: "李驗收", deptId: "dept-1", groupId: "g1-3b", skills: ["驗收"], nationalHolidayQuota: 1, systemLeaveQuota: 7 },
  { id: "emp-3", name: "張店訂", deptId: "dept-2", groupId: "g2-1", skills: ["店訂"], nationalHolidayQuota: 2, systemLeaveQuota: 6 },
  { id: "emp-4", name: "趙庫存", deptId: "dept-3", groupId: "g3-1", skills: ["庫存"], nationalHolidayQuota: 2, systemLeaveQuota: 6 },
  { id: "emp-5", name: "關運務", deptId: "dept-4", groupId: "g4-1", skills: ["運務"], nationalHolidayQuota: 1, systemLeaveQuota: 7 },
];

const DEFAULT_STATE: ScheduleState = {
  departments: INITIAL_DEPTS,
  employees: INITIAL_EMPLOYEES,
  records: {},
  dailyRequirements: {},
  currentPeriodStart: format(new Date(), "yyyy-MM-dd"),
  nationalHolidays: ["2026-05-01", "2026-06-10"],
  attendanceTimes: ["無", "08:00-17:00", "09:00-18:00", "10:00-19:00", "13:00-22:00", "22:00-07:00", "自休", "系休", "國定"],
  leaveInterval: {
    start: format(new Date(), "yyyy-MM-dd"),
    end: format(addDays(new Date(), 27), "yyyy-MM-dd"),
  }
};

export default function App() {
  const [viewMode, setViewMode] = useState<"front" | "back">("front");
  const [activeDept, setActiveDept] = useState<string>(INITIAL_DEPTS[0].id);
  const [startDate, setStartDate] = useState<string>(format(new Date(), "yyyy-MM-dd"));
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [selectedGroupForDaily, setSelectedGroupForDaily] = useState<string>("");
  const [isBalancing, setIsBalancing] = useState(false);
  
  const [state, setState] = useState<ScheduleState>(DEFAULT_STATE);
  const [isClearingAll, setIsClearingAll] = useState(false);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [importErrors, setImportErrors] = useState<{ row: number; name: string; reason: string }[]>([]);
  const [shiftDeptMode, setShiftDeptMode] = useState<string>("global"); // "global" or deptId
  const [quickMode, setQuickMode] = useState<"shift" | "leave">("shift"); // 快速模式：填班或標假
  const [activeMenu, setActiveMenu] = useState<{ empId: string; dateStr: string; x: number; y: number } | null>(null);
  const [lastClickTime, setLastClickTime] = useState<number>(0);

  const [notification, setNotification] = useState<{ message: string; type: "success" | "error" } | null>(null);

  // Auth Listener
  useEffect(() => {
    return onAuthStateChanged(auth, (u) => {
      setUser(u);
      if (!u) setLoading(false);
    });
  }, []);

  // Sync with Firestore when logged in
  useEffect(() => {
    if (!user) return;

    setLoading(true);
    const unsubDepts = onSnapshot(collection(db, "departments"), (snap) => {
      if (!snap.empty) {
        const depts = snap.docs.map(d => d.data() as Department);
        setState(prev => ({ ...prev, departments: depts }));
      } else {
        // Seed if empty
        INITIAL_DEPTS.forEach(d => setDoc(doc(db, "departments", d.id), d));
      }
    });

    const unsubEmps = onSnapshot(collection(db, "employees"), (snap) => {
      if (!snap.empty) {
        const emps = snap.docs.map(d => d.data() as Employee);
        setState(prev => ({ ...prev, employees: emps }));
      } else {
        INITIAL_EMPLOYEES.forEach(e => setDoc(doc(db, "employees", e.id), e));
      }
    });

    const unsubConfig = onSnapshot(doc(db, "config", "global"), (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        setState(prev => ({
          ...prev,
          nationalHolidays: data.nationalHolidays || [],
          attendanceTimes: data.attendanceTimes || DEFAULT_STATE.attendanceTimes,
          currentPeriodStart: data.currentPeriodStart || format(new Date(), "yyyy-MM-dd"),
          leaveInterval: data.leaveInterval || DEFAULT_STATE.leaveInterval
        }));
      } else {
        setDoc(doc(db, "config", "global"), {
          nationalHolidays: DEFAULT_STATE.nationalHolidays,
          attendanceTimes: DEFAULT_STATE.attendanceTimes,
          currentPeriodStart: DEFAULT_STATE.currentPeriodStart,
          leaveInterval: DEFAULT_STATE.leaveInterval
        });
      }
    });

    const unsubRecords = onSnapshot(collection(db, "shiftRecords"), (snap) => {
      const records: Record<string, ShiftRecord> = {};
      snap.docs.forEach(d => {
        records[d.id] = d.data() as ShiftRecord;
      });
      setState(prev => ({ ...prev, records }));
    });

    const unsubDailyReqs = onSnapshot(collection(db, "dailyRequirements"), (snap) => {
      const dailyRequirements: Record<string, DailyRequirement> = {};
      snap.docs.forEach(d => {
        dailyRequirements[d.id] = d.data() as DailyRequirement;
      });
      setState(prev => ({ ...prev, dailyRequirements }));
      setLoading(false);
    });

    return () => {
      unsubDepts();
      unsubEmps();
      unsubConfig();
      unsubRecords();
      unsubDailyReqs();
    };
  }, [user]);

  // Sync internal date state with storage (start date)
  useEffect(() => {
    if (state.currentPeriodStart !== startDate) {
      setStartDate(state.currentPeriodStart);
    }
  }, [state.currentPeriodStart]);

  // 當起始日期變更時，自動強制同步可休假區間 (28天週期)
  useEffect(() => {
    const startDateObj = parseISO(startDate);
    const endDateObj = addDays(startDateObj, 27);
    const start = format(startDateObj, "yyyy-MM-dd");
    const end = format(endDateObj, "yyyy-MM-dd");

    if (state.leaveInterval.start !== start || state.leaveInterval.end !== end) {
      setState(prev => ({
        ...prev,
        leaveInterval: { start, end }
      }));
      
      if (user) {
        setDoc(doc(db, "config", "global"), { 
          leaveInterval: { start, end } 
        }, { merge: true });
      }
    }
  }, [startDate, user]);

  // Ensure activeDept is valid when departments load
  useEffect(() => {
    if (state.departments.length > 0) {
      const exists = state.departments.some(d => d.id === activeDept);
      if (!exists) {
        setActiveDept(state.departments[0].id);
      }
    }
  }, [state.departments, activeDept]);

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
      showNotice("登入成功");
    } catch (e) {
      console.error(e);
      showNotice("登入失敗", "error");
    }
  };

  const handleLogout = () => {
    auth.signOut();
    showNotice("已登出");
  };

  // 當手動調整起始日期時更新 state
  const handleStartDateChange = async (val: string) => {
    setStartDate(val);
    setState(prev => ({ ...prev, currentPeriodStart: val }));
    if (user) {
      await setDoc(doc(db, "config", "global"), { currentPeriodStart: val }, { merge: true });
    }
  };

  const sortTimes = (times: string[]) => {
    const systemStart = ["無"];
    const leaves = ["自休", "系休", "國定"];
    const shifts = times.filter(t => !systemStart.includes(t) && !leaves.includes(t));
    const sortedShifts = shifts.sort((a, b) => {
      const timeA = a.match(/\d{2}:\d{2}/)?.[0] || "";
      const timeB = b.match(/\d{2}:\d{2}/)?.[0] || "";
      return timeA.localeCompare(timeB);
    });
    return [...systemStart, ...sortedShifts, ...leaves];
  };

  const toggleAttendanceTime = async (time: string, mode: "add" | "remove", deptId?: string) => {
    if (["無", "自休", "系休", "國定"].includes(time) && mode === "remove") {
      showNotice("基本系統選項不可刪除", "error");
      return;
    }
    
    if (deptId && deptId !== "global") {
      const dept = state.departments.find(d => d.id === deptId);
      if (!dept) return;
      
      const currentTimes = dept.attendanceTimes || [];
      // 確保不包含系統保留字
      const updatedList = mode === "add"
        ? Array.from(new Set([...currentTimes, time])).filter(t => !["自休", "系休", "國定"].includes(t))
        : currentTimes.filter(t => t !== time);
      
      const sortedTimes = sortTimes(updatedList);
      const updatedDept = { ...dept, attendanceTimes: sortedTimes };
      
      setState(prev => ({
        ...prev,
        departments: prev.departments.map(d => d.id === deptId ? updatedDept : d)
      }));
      
      if (user) {
        await setDoc(doc(db, "departments", deptId), updatedDept);
      }
      showNotice(`${dept.name} ${mode === "add" ? "已新增班別" : "已刪除班別"}`);
    } else {
      const updatedList = mode === "add" 
        ? Array.from(new Set([...state.attendanceTimes, time])).filter(t => !["自休", "系休", "國定"].includes(t))
        : state.attendanceTimes.filter(t => t !== time);
      const sortedTimes = sortTimes(updatedList);
      setState(prev => ({ ...prev, attendanceTimes: sortedTimes }));
      if (user) {
        await setDoc(doc(db, "config", "global"), { attendanceTimes: sortedTimes }, { merge: true });
      }
      showNotice(`全域設定 ${mode === "add" ? "已新增班別" : "已刪除班別"}`);
    }
  };

  const handleImportEmployees = async (e: React.ChangeEvent<HTMLInputElement>, targetGroupId?: string) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportErrors([]); // Clear previous errors
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];
        const newEmployees: Employee[] = [];
        const batch = writeBatch(db);
        const errors: { row: number; name: string; reason: string }[] = [];

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const rowNum = i + 2; // Excel row number (1-indexed + header)
          const name = row["姓名"]?.toString().trim();
          let deptName = row["課別"]?.toString().trim();
          let groupName = row["組別"]?.toString().trim();
          const skillsStr = row["具備技能"]?.toString().trim() || "";
          const systemQuota = parseInt(row["系休額度"]) || 0;
          const nationalQuota = parseInt(row["國定額度"]) || 0;
          
          if (!name) {
            errors.push({ row: rowNum, name: "未知", reason: "姓名欄位空白" });
            continue;
          }

          let dept, group;
          if (targetGroupId) {
            dept = state.departments.find(d => d.groups.some(g => g.id === targetGroupId));
            group = dept?.groups.find(g => g.id === targetGroupId);
          } else {
            dept = state.departments.find(d => d.name === deptName);
            group = dept?.groups.find(g => g.name === groupName);
          }

          if (!dept || !group) {
            errors.push({ 
              row: rowNum, 
              name, 
              reason: targetGroupId 
                ? "系統內部組別匹配失敗" 
                : `找不到對應課別(${deptName || "未填"})或組別(${groupName || "未填"})` 
            });
            continue;
          }

          const empId = "emp_" + Date.now() + "_" + Math.random().toString(36).substr(2, 5);
          const emp: Employee = {
            id: empId, name, deptId: dept.id, groupId: group.id,
            skills: skillsStr.split(/[基,，\s\/]+/).filter(Boolean),
            systemLeaveQuota: systemQuota, nationalHolidayQuota: nationalQuota
          };
          newEmployees.push(emp);
          if (user) batch.set(doc(db, "employees", empId), emp);
        }

        setImportErrors(errors);
        
        if (errors.length > 0) {
          setTimeout(() => {
            document.getElementById('import-error-anchor')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }, 300);
        }

        if (newEmployees.length > 0) {
          setState(prev => ({ ...prev, employees: [...prev.employees, ...newEmployees] }));
          if (user) await batch.commit();
          showNotice("成功匯入 " + newEmployees.length + " 名員工" + (errors.length > 0 ? " (另有 " + errors.length + " 筆異常)" : ""));
        } else if (errors.length > 0) {
          showNotice("匯入未完成，共發現 " + errors.length + " 筆格式異常", "error");
        } else {
          showNotice("匯入失敗：表單內無有效資料", "error");
        }
      } catch (err) {
        showNotice("檔案讀取出錯", "error");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const downloadEmployeeTemplate = () => {
    const mainData = [
      { "姓名": "張三", "課別": "大溪理貨一課", "組別": "日班理貨一組", "具備技能": "理貨, 驗收", "系休額度": 8, "國定額度": 1 },
      { "姓名": "李四", "課別": "大溪理貨一課", "組別": "日班理貨二組", "具備技能": "理貨", "系休額度": 7, "國定額度": 2 }
    ];
    const wsMain = XLSX.utils.json_to_sheet(mainData);
    const refData: any[] = [];
    state.departments.forEach(d => {
      d.groups.forEach(g => { refData.push({ "課別名稱": d.name, "組別名稱": g.name, "可用技能": d.skills.join(", ") }); });
    });
    const wsRef = XLSX.utils.json_to_sheet(refData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsMain, "人事匯入填寫區");
    XLSX.utils.book_append_sheet(wb, wsRef, "系統課別組別參考");
    XLSX.writeFile(wb, "物流排班_人事公版匯入檔(支援多組別).xlsx");
  };

  const showNotice = (message: string, type: "success" | "error" = "success") => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  };
  
const days = useMemo(() => {
    const start = parseISO(startDate);
    return Array.from({ length: 28 }, (_, i) => addDays(start, i));
  }, [startDate]);

  // 處理儲存格下拉選單變更
  const handleTimeChange = async (empId: string, date: string, value: string) => {
    // 檢查是否在可休區間內
    if (["自休", "系休", "國定"].includes(value)) {
      if (date < state.leaveInterval.start || date > state.leaveInterval.end) {
        showNotice(`選定日期 ${date} 不在可休假區間內 (${state.leaveInterval.start} ~ ${state.leaveInterval.end})`, "error");
        return;
      }
    }

    const key = `${empId}_${date}`;
    let type: "shift" | "leave" = "shift";
    let code = "出勤";
    let time = "08:00-17:00";
    let isLeaveValue = false;

    if (value === "自休") {
      const selfLeavesCount = (Object.values(state.records) as ShiftRecord[]).filter(r => r.employeeId === empId && r.code === LeaveType.Self && r.date !== date).length;
      if (selfLeavesCount >= 4) {
        showNotice("自選休(自休)限制每週期最多四天", "error");
        return;
      }
      type = "leave"; code = LeaveType.Self; isLeaveValue = true; 
    }
    else if (value === "系休") { type = "leave"; code = LeaveType.System; isLeaveValue = true; }
    else if (value === "國定") {
      const nationalLeavesCount = (Object.values(state.records) as ShiftRecord[]).filter(r => r.employeeId === empId && r.code === LeaveType.National && r.date !== date).length;
      const emp = state.employees.find(e => e.id === empId);
      if (emp && nationalLeavesCount >= emp.nationalHolidayQuota) {
        showNotice(`已達該員國定假日休假額度上限 (${emp.nationalHolidayQuota} 天)`, "error");
        return;
      }
      type = "leave"; code = LeaveType.National; isLeaveValue = true; 
    }
    else if (value === "無") { code = ""; time = ""; }
    else {
      // 處理 "出勤" 或 "出勤(技能)"
      code = value;
      type = "shift";
      isLeaveValue = false;
    }

    const newRecord: ShiftRecord = {
      employeeId: empId,
      date,
      type,
      code,
      time: isLeaveValue ? "" : time,
      isManual: value !== "無"
    };

    // 規則 4: 不可連續出勤六天
    if (!isLeaveValue && value !== "無") {
      const recs = { ...state.records, [key]: newRecord };
      let consecutive = 0;
      for (const d of days) {
        const dStr = format(d, "yyyy-MM-dd");
        if (recs[`${empId}_${dStr}`]?.type === "shift") {
          consecutive++;
          if (consecutive > 6) {
            showNotice("違反排班規定：不可連續出勤超過六天", "error");
            return;
          }
        } else {
          consecutive = 0;
        }
      }
    }

    if (user) {
      if (value === "無") {
        await deleteDoc(doc(db, "shiftRecords", key));
      } else {
        await setDoc(doc(db, "shiftRecords", key), newRecord);
      }
    } else {
      setState(prev => ({
        ...prev,
        records: {
          ...prev.records,
          [key]: newRecord
        }
      }));
    }
  };

  const toggleHoliday = async (dateStr: string) => {
    const isRemove = state.nationalHolidays.includes(dateStr);
    const newHolidays = isRemove 
      ? state.nationalHolidays.filter(d => d !== dateStr) 
      : [...state.nationalHolidays, dateStr];

    setState(prev => ({ ...prev, nationalHolidays: newHolidays }));
    
    if (user) {
      await setDoc(doc(db, "config", "global"), { nationalHolidays: newHolidays }, { merge: true });
    }
  };

  const updateLeaveInterval = async (key: "start" | "end", val: string) => {
    const newInterval = { ...state.leaveInterval, [key]: val };
    setState(prev => ({ ...prev, leaveInterval: newInterval }));

    if (user) {
      await setDoc(doc(db, "config", "global"), { leaveInterval: newInterval }, { merge: true });
    }
  };

  // 取得當前排班週期內的國定假日
  const periodHolidays = useMemo(() => {
    if (days.length === 0) return [];
    const interval = { start: days[0], end: days[days.length - 1] };
    return state.nationalHolidays.filter(h => {
      try { return isWithinInterval(parseISO(h), interval); } 
      catch(e) { return false; }
    }).sort();
  }, [days, state.nationalHolidays]);

  const applyHolidayQuotaToAll = async () => {
    const count = periodHolidays.length;
    const newEmployees = state.employees.map(emp => ({
      ...emp,
      nationalHolidayQuota: count
    }));
    setState(prev => ({ ...prev, employees: newEmployees }));
    if (user) {
      const batch = writeBatch(db);
      newEmployees.forEach(emp => {
        batch.update(doc(db, "employees", emp.id), { nationalHolidayQuota: count });
      });
      await batch.commit();
    }
    showNotice(`已將本週期國定假日天數 (${count} 天) 套用至所有員工`);
  };

  const fetchGovHolidays = async () => {
    try {
      showNotice("正在搜集政府國定假日資料...");
      const year = parseISO(startDate).getFullYear();
      const commonHolidays: Record<number, string[]> = {
        2024: ["2024-01-01", "2024-02-08", "2024-02-09", "2024-02-10", "2024-02-11", "2024-02-12", "2024-02-13", "2024-02-14", "2024-02-28", "2024-04-04", "2024-04-05", "2024-05-01", "2024-06-10", "2024-09-17", "2024-10-10"],
        2025: ["2025-01-01", "2025-01-27", "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31", "2025-02-01", "2025-02-02", "2025-02-28", "2025-04-04", "2025-04-05", "2025-05-01", "2025-05-31", "2025-10-06", "2025-10-10"],
        2026: ["2026-01-01", "2026-01-19", "2026-02-17", "2026-02-28", "2026-04-03", "2026-04-04", "2026-04-05", "2026-05-01", "2026-06-19", "2026-09-25", "2026-10-10"],
        2027: ["2027-01-01", "2027-02-06", "2027-02-07", "2027-02-08", "2027-02-09", "2027-02-10", "2027-02-11", "2027-02-28", "2027-04-04", "2027-04-05", "2027-05-01", "2027-06-08", "2027-09-15", "2027-10-10"]
      };
      const selectedHolidays = commonHolidays[year] || [];
      if (selectedHolidays.length === 0) {
        showNotice(`目前暫無 ${year} 年度的預設國定假日資料庫。`, "error");
        return;
      }
      const unitedHolidays = Array.from(new Set([...state.nationalHolidays, ...selectedHolidays])).sort();
      setState(prev => ({ ...prev, nationalHolidays: unitedHolidays }));
      if (user) {
        await setDoc(doc(db, "config", "global"), { nationalHolidays: unitedHolidays }, { merge: true });
      }
      showNotice(`已自動同步 ${year} 年度國定假日`);
    } catch (e) {
      showNotice("搜集失敗", "error");
    }
  };

  const updateDailyReq = async (groupId: string, date: string, type: 'maxLeave' | 'skillCount', skill?: string, val?: number) => {
    const key = `${groupId}_${date}`;
    const group = state.departments.flatMap(d => d.groups).find(g => g.id === groupId);
    if (!group) return;

    const existing = state.dailyRequirements[key] || {
      id: key,
      groupId,
      date,
      maxLeave: group.maxLeavePerDay,
      skillRequirements: [...group.skillRequirements]
    };

    let newReq: DailyRequirement;
    if (type === 'maxLeave') {
      newReq = { ...existing, maxLeave: val || 0 };
    } else {
      const skills = [...existing.skillRequirements];
      const idx = skills.findIndex(s => s.skill === skill);
      if (idx > -1) {
        skills[idx] = { ...skills[idx], count: val || 0 };
      } else if (skill) {
        skills.push({ skill, count: val || 0 });
      }
      newReq = { ...existing, skillRequirements: skills };
    }

    setState(prev => ({
      ...prev,
      dailyRequirements: { ...prev.dailyRequirements, [key]: newReq }
    }));

    if (user) {
      await setDoc(doc(db, "dailyRequirements", key), newReq);
    }
  };

  const rebalanceGroupLeaveLimits = async (groupId: string, changedDate: string, newValue: number) => {
    const group = state.departments.flatMap(d => d.groups).find(g => g.id === groupId);
    if (!group) return;
    
    setIsBalancing(true);
    const totalEmps = state.employees.filter(e => e.groupId === groupId).length;
    if (totalEmps === 0) {
      setIsBalancing(false);
      return;
    }

    // 計算本週期該組總應休天數 (8天週休 + 國定配額)
    const totalTargetOffs = state.employees
      .filter(e => e.groupId === groupId)
      .reduce((acc, emp) => acc + 8 + emp.nationalHolidayQuota, 0);

    const periodDaysCount = days.length; // 通常是 28
    const otherDays = days
      .map(d => format(d, "yyyy-MM-dd"))
      .filter(d => d !== changedDate);

    // 剩餘要填補的休假總額
    const remainingToFill = totalTargetOffs - newValue;
    const avgPerDay = Math.max(0, Math.floor(remainingToFill / otherDays.length));
    const remainder = Math.max(0, remainingToFill % otherDays.length);

    const batch = writeBatch(db);
    const updates: Record<string, DailyRequirement> = {};

    otherDays.forEach((dateStr, idx) => {
      const key = `${groupId}_${dateStr}`;
      const existing = state.dailyRequirements[key] || {
        id: key,
        groupId,
        date: dateStr,
        maxLeave: group.maxLeavePerDay,
        skillRequirements: [...group.skillRequirements]
      };
      
      const adjustedMax = avgPerDay + (idx < remainder ? 1 : 0);
      const newReq = { ...existing, maxLeave: adjustedMax };
      
      updates[key] = newReq;
      if (user) batch.set(doc(db, "dailyRequirements", key), newReq);
    });

    // 更新當前變動的那一天
    const changedKey = `${groupId}_${changedDate}`;
    const currentExisting = state.dailyRequirements[changedKey] || {
      id: key,
      groupId,
      date: changedDate,
      maxLeave: newValue,
      skillRequirements: [...group.skillRequirements]
    };
    const currentNewReq = { ...currentExisting, maxLeave: newValue };
    updates[changedKey] = currentNewReq;
    if (user) batch.set(doc(db, "dailyRequirements", changedKey), currentNewReq);

    setState(prev => ({
      ...prev,
      dailyRequirements: { ...prev.dailyRequirements, ...updates }
    }));

    if (user) await batch.commit();
    setIsBalancing(false);
    showNotice(`已自動平均分配組別「${group.name}」剩餘休假額度。`);
  };

  const autoPreFillGroupLeaveLimits = async (groupId: string) => {
    const group = state.departments.flatMap(d => d.groups).find(g => g.id === groupId);
    if (!group) return;

    const totalEmps = state.employees.filter(e => e.groupId === groupId).length;
    if (totalEmps === 0) {
      showNotice("該組別目前無員工，無法計算平均休假值。", "error");
      return;
    }

    // 平均計算：(總人數 * (8天週休 + 平均國休配額)) / 週期天數
    const avgQuota = state.employees.filter(e => e.groupId === groupId).reduce((acc, emp) => acc + emp.nationalHolidayQuota, 0) / totalEmps;
    const totalOffsPerEmp = 8 + avgQuota;
    const totalGroupOffsNeeded = totalEmps * totalOffsPerEmp;
    const recommendedPerDay = Math.ceil(totalGroupOffsNeeded / days.length);

    const batch = writeBatch(db);
    const updates: Record<string, DailyRequirement> = {};

    days.forEach(day => {
      const dateStr = format(day, "yyyy-MM-dd");
      const key = `${groupId}_${dateStr}`;
      const newReq: DailyRequirement = {
        id: key,
        groupId,
        date: dateStr,
        maxLeave: recommendedPerDay,
        skillRequirements: [...group.skillRequirements]
      };
      updates[key] = newReq;
      if (user) batch.set(doc(db, "dailyRequirements", key), newReq);
    });

    setState(prev => ({
      ...prev,
      dailyRequirements: { ...prev.dailyRequirements, ...updates }
    }));

    if (user) await batch.commit();
    showNotice(`已根據組別人數 (${totalEmps}人) 自動預設平均每日休假上限為 ${recommendedPerDay} 人。`);
  };

  const applyGroupDefaultsToAllDays = async (groupId: string) => {
    const group = state.departments.flatMap(d => d.groups).find(g => g.id === groupId);
    if (!group) return;

    const batch = writeBatch(db);
    const updates: Record<string, DailyRequirement> = {};

    days.forEach(day => {
      const dateStr = format(day, "yyyy-MM-dd");
      const key = `${groupId}_${dateStr}`;
      const newReq: DailyRequirement = {
        id: key,
        groupId,
        date: dateStr,
        maxLeave: group.maxLeavePerDay,
        skillRequirements: [...group.skillRequirements]
      };
      updates[key] = newReq;
      if (user) {
        batch.set(doc(db, "dailyRequirements", key), newReq);
      }
    });

    setState(prev => ({
      ...prev,
      dailyRequirements: { ...prev.dailyRequirements, ...updates }
    }));

    if (user) {
      await batch.commit();
    }
    showNotice(`已將「${group.name}」基本需求套用至全週期`);
  };

  const updateGroupConfig = async (deptId: string, groupId: string, field: 'maxLeavePerDay' | 'skillRequirements', value: any) => {
    setState(prev => ({
      ...prev,
      departments: prev.departments.map(d => {
        if (d.id !== deptId) return d;
        return {
          ...d,
          groups: d.groups.map(g => {
            if (g.id !== groupId) return g;
            return { ...g, [field]: value };
          })
        };
      })
    }));

    if (user) {
      const dept = state.departments.find(d => d.id === deptId);
      if (dept) {
        const updatedGroups = dept.groups.map(g => g.id === groupId ? { ...g, [field]: value } : g);
        await setDoc(doc(db, "departments", deptId), { groups: updatedGroups }, { merge: true });
      }
    }
  };

  const updateGroupSkills = async (deptId: string, groupId: string, mode: 'add' | 'remove', skillName: string) => {
    const dept = state.departments.find(d => d.id === deptId);
    if (!dept) return;
    const group = dept.groups.find(g => g.id === groupId);
    if (!group) return;

    let newSkills = [...group.skillRequirements];
    if (mode === 'add') {
      if (newSkills.some(s => s.skill === skillName)) return;
      newSkills.push({ skill: skillName, count: 0 });
    } else {
      newSkills = newSkills.filter(s => s.skill !== skillName);
    }

    await updateGroupConfig(deptId, groupId, 'skillRequirements', newSkills);
    showNotice(mode === 'add' ? `已為「${group.name}」新增技能項目：${skillName}` : `已移除技能項目：${skillName}`);
  };

  const updateDepartmentSkills = async (deptId: string, mode: 'add' | 'remove', skillName: string) => {
    // This is now less critical but we keep it for global department skill tracking if needed
    const dept = state.departments.find(d => d.id === deptId);
    if (!dept) return;

    let newSkills = [...dept.skills];
    if (mode === 'add') {
      if (newSkills.includes(skillName)) return;
      newSkills.push(skillName);
    } else {
      newSkills = newSkills.filter(s => s !== skillName);
    }

    setState(prev => ({
      ...prev,
      departments: prev.departments.map(d => d.id === deptId ? { ...d, skills: newSkills } : d)
    }));

    if (user) {
      await setDoc(doc(db, "departments", deptId), { skills: newSkills }, { merge: true });
    }
  };

  // 執行自動排班演算法 (物流專業版 v2：勞基法一例一休 + 國休彈性排入)
  const runAutoSchedule = async () => {
    const newRecords = { ...state.records };
    const periodDays = days.map(d => format(d, "yyyy-MM-dd"));
    
    // 預處理：清除非手動紀錄
    Object.keys(newRecords).forEach(key => {
      if (!newRecords[key]?.isManual) delete newRecords[key];
    });

    // 取得組別與日期的需求快取
    const getRequirements = (groupId: string, dayStr: string) => {
      const group = state.departments.flatMap(d => d.groups).find(g => g.id === groupId);
      if (!group) return null;
      const daily = state.dailyRequirements[`${groupId}_${dayStr}`];
      return {
        maxLeave: daily ? daily.maxLeave : group.maxLeavePerDay,
        skillReqs: daily ? daily.skillRequirements : group.skillRequirements,
        group
      };
    };

    // 檢查某員工在某日是否可以休假 (考慮總人數與技能數)
    const canGrantLeave = (empId: string, dayStr: string, records: Record<string, ShiftRecord>) => {
      const emp = state.employees.find(e => e.id === empId);
      if (!emp) return false;
      const req = getRequirements(emp.groupId, dayStr);
      if (!req) return false;

      // 1. 檢查組別總休假上限
      const groupEmps = state.employees.filter(e => e.groupId === emp.groupId);
      const currentLeaves = groupEmps.filter(e => records[`${e.id}_${dayStr}`]?.type === "leave").length;
      if (currentLeaves >= req.maxLeave) return false;

      // 2. 檢查各項技能剩餘人數是否充足
      for (const skillReq of req.skillReqs) {
        if (emp.skills.includes(skillReq.skill)) {
          // 計算具備此技能且「目前尚未排休」的人數
          const totalWithSkill = groupEmps.filter(e => e.skills.includes(skillReq.skill)).length;
          const leavesWithSkill = groupEmps.filter(e => e.skills.includes(skillReq.skill) && records[`${e.id}_${dayStr}`]?.type === "leave").length;
          
          if (totalWithSkill - leavesWithSkill <= skillReq.count) {
            return false; // 若休了這個人，該技能人數將低於需求
          }
        }
      }

      return true;
    };

    // 分配週次
    const weeks: string[][] = [];
    let currentWeek: string[] = [];
    periodDays.forEach(day => {
      currentWeek.push(day);
      const dayOfWeek = parseISO(day).getDay();
      if (dayOfWeek === 0 || day === periodDays[periodDays.length - 1]) {
        weeks.push(currentWeek);
        currentWeek = [];
      }
    });

    // 員工排班
    state.employees.forEach(emp => {
      const empRecords = (Object.values(newRecords) as ShiftRecord[]).filter(r => r?.employeeId === emp.id);
      const manualStandardOffs = empRecords.filter(r => r.code === LeaveType.Self || r.code === LeaveType.System).length;
      let remainingStandardOffs = Math.max(0, 8 - manualStandardOffs);
      
      const manualNationalOffs = empRecords.filter(r => r.code === LeaveType.National).length;
      let remainingNationalOffs = Math.max(0, emp.nationalHolidayQuota - manualNationalOffs);

      let continuousWork = 0;

      // 階段 1：滿足強制休假 (連勤 6 天或週休不足)
      for (let w = 0; w < weeks.length; w++) {
        const weekDays = weeks[w];
        let weekStandardOffs = weekDays.filter(d => {
          const r = newRecords[`${emp.id}_${d}`];
          return r?.type === "leave" && (r.code === LeaveType.Self || r.code === LeaveType.System);
        }).length;

        for (const dayStr of weekDays) {
          const currentRec = newRecords[`${emp.id}_${dayStr}`];
          if (currentRec) {
            if (currentRec.type === "leave") continuousWork = 0;
            else continuousWork++;
            continue;
          }

          // 強制條件
          const mustOff = continuousWork >= 5 || (weekStandardOffs < 2 && weekDays.indexOf(dayStr) >= weekDays.length - (2 - weekStandardOffs));

          if (mustOff && canGrantLeave(emp.id, dayStr, newRecords)) {
            if (remainingStandardOffs > 0) {
              newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "leave", code: LeaveType.System, isManual: false };
              remainingStandardOffs--; weekStandardOffs++; continuousWork = 0;
            } else {
              newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "leave", code: "強休", isManual: false };
              continuousWork = 0;
            }
          } else {
            // 若不能休或不需強休，暫定上班 (後面會再細排)
            // 不在此處寫入，留給階段 2
          }
        }
      }

      // 階段 2：分配國定假與剩餘排休
      for (const dayStr of periodDays) {
        if (newRecords[`${emp.id}_${dayStr}`]) continue;

        if (canGrantLeave(emp.id, dayStr, newRecords)) {
          if (remainingNationalOffs > 0) {
            newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "leave", code: LeaveType.National, isManual: false };
            remainingNationalOffs--;
          } else if (remainingStandardOffs > 0) {
            newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "leave", code: LeaveType.System, isManual: false };
            remainingStandardOffs--;
          } else {
            newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "shift", code: "出勤", time: "08:00-17:00", isManual: false };
          }
        } else {
          newRecords[`${emp.id}_${dayStr}`] = { employeeId: emp.id, date: dayStr, type: "shift", code: "出勤", time: "08:00-17:00", isManual: false };
        }
      }
    });

    // 階段 3：技能標記 (Optimization for display)
    periodDays.forEach(dayStr => {
      state.departments.forEach(dept => {
        dept.groups.forEach(group => {
          const req = getRequirements(group.id, dayStr);
          if (!req) return;
          const assignedIds = new Set<string>();

          req.skillReqs.forEach(sReq => {
            let matches = 0;
            const available = state.employees.filter(e => e.groupId === group.id && newRecords[`${e.id}_${dayStr}`]?.type === "shift");
            // 優先排序：具備較少技能的人優先佔坑 (避免技能重疊浪費)
            const sortedAvailable = [...available].sort((a, b) => a.skills.length - b.skills.length);
            
            for (const emp of sortedAvailable) {
              if (matches >= sReq.count) break;
              if (assignedIds.has(emp.id)) continue;
              if (emp.skills.includes(sReq.skill)) {
                newRecords[`${emp.id}_${dayStr}`].code = `出勤(${sReq.skill})`;
                assignedIds.add(emp.id); matches++;
              }
            }
          });
        });
      });
    });


    if (user) {
      try {
        const entries = Object.entries(newRecords);
        // Firestore batch limit is 500
        for (let i = 0; i < entries.length; i += 500) {
          const batch = writeBatch(db);
          const chunk = entries.slice(i, i + 500);
          chunk.forEach(([key, rec]) => batch.set(doc(db, "shiftRecords", key), rec));
          await batch.commit();
        }
        showNotice("物流專用引擎 v2：已成功執行自動排班。");
      } catch (err) {
        console.error("Auto-schedule failed:", err);
        showNotice("自動排班存檔失敗，請檢查網路或分批執行。", "error");
      }
    } else {
      setState(prev => ({ ...prev, records: newRecords }));
      showNotice("物流專用引擎 v2：已將國休獨立於週休二日外，並嚴守一例一休原則。");
    }
  };

  // 匯出功能 (.xlsx)
  const exportToExcel = () => {
    const dept = state.departments.find(d => d.id === activeDept);
    if (!dept) return;

    const data: any[] = [];
    const dateHeaders = days.map(d => format(d, "MM/dd(EEEEEE)", { locale: zhTW }));
    
    // Header Row
    data.push(["組別", "姓名", "具備技能", "系休額度", "國定額度", ...dateHeaders]);

    // Employee Rows
    const deptGroups = dept.groups;
    deptGroups.forEach(group => {
      state.employees.filter(e => e.groupId === group.id).forEach(emp => {
        const row = [
          group.name,
          emp.name,
          emp.skills.join("/"),
          emp.systemLeaveQuota,
          emp.nationalHolidayQuota
        ];
        days.forEach(day => {
          const dateStr = format(day, "yyyy-MM-dd");
          const rec = state.records[`${emp.id}_${dateStr}`];
          row.push(rec ? (rec.type === "leave" ? rec.code : `${rec.code} ${rec.time || ""}`) : "");
        });
        data.push(row);
      });
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "班表");
    XLSX.writeFile(wb, `物流排班表_${dept.name}_${startDate}.xlsx`);
    showNotice("成功匯出 .xlsx 檔案");
  };

  // 匯入功能 (.xlsx)
  const importFromExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target?.result as string;
      const wb = XLSX.read(bstr, { type: "binary" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[];

      if (rows.length < 2) return;
      const newRecords = { ...state.records };
      const dateCols = rows[0].slice(5); 

      const empMap = new Map<string, Employee>();
      state.employees.forEach(emp => empMap.set(emp.name, emp));

      const batch = user ? writeBatch(db) : null;

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const empName = row[1];
        const emp = empMap.get(empName);
        if (!emp) continue;

        dateCols.forEach((_, colIndex) => {
          const day = days[colIndex];
          if (!day) return;
          const dateStr = format(day, "yyyy-MM-dd");
          const val = row[colIndex + 5];

          const key = `${emp.id}_${dateStr}`;
          if (!val) {
            delete newRecords[key];
            if (batch) batch.delete(doc(db, "shiftRecords", key));
            return;
          }

          let type: "shift" | "leave" = "shift";
          let code = "900";
          let time = "";

          if (["自休", "系休", "國定"].includes(val)) {
            type = "leave"; code = val;
          } else {
            const parts = String(val).split(" ");
            code = parts[0];
            time = parts[1] || "";
          }

          const newRec: ShiftRecord = { employeeId: emp.id, date: dateStr, type, code, time, isManual: true };
          newRecords[key] = newRec;
        });
      }

      // 檢查匯入後是否違反連勤 6 天規定
      for (const emp of state.employees) {
        let consecutive = 0;
        for (const day of days) {
          const dStr = format(day, "yyyy-MM-dd");
          if (newRecords[`${emp.id}_${dStr}`]?.type === "shift") {
            consecutive++;
            if (consecutive > 6) {
              showNotice(`匯入失敗：員工 ${emp.name} 於 ${dStr} 違反連勤 6 天規定`, "error");
              return;
            }
          } else {
            consecutive = 0;
          }
        }
      }

      if (batch) {
        Object.entries(newRecords).forEach(([key, rec]) => {
          batch.set(doc(db, "shiftRecords", key), rec);
        });
        await batch.commit();
      } else {
        setState(prev => ({ ...prev, records: newRecords }));
      }
      showNotice("班表資料已從 Excel 匯入並同步至雲端");
    };
    reader.readAsBinaryString(file);
  };

  const deleteEmployee = async (id: string, name: string) => {
    try {
      if (user) await deleteDoc(doc(db, "employees", id));
      setState(prev => ({ ...prev, employees: prev.employees.filter(e => e.id !== id) }));
      showNotice(`已移除員工：${name}`);
    } catch (err) {
      console.error(err);
      showNotice("移除失敗", "error");
    }
  };

  const deleteDepartment = async (id: string, name: string) => {
    setState(prev => ({
      ...prev,
      departments: prev.departments.filter(d => d.id !== id)
    }));

    if (user) {
      try {
        await deleteDoc(doc(db, "departments", id));
        showNotice(`已移除課別：${name}`);
      } catch (err) {
        console.error(err);
        showNotice("資料庫同步失敗", "error");
      }
    }
  };

  // --- UI Components ---
  const ScheduleTable = () => {
    let dept = state.departments.find(d => d.id === activeDept);
    if (!dept && state.departments.length > 0) {
      dept = state.departments[0];
    }
    
    if (!dept) {
      return (
        <div className="h-full flex flex-col items-center justify-center text-gray-400 bg-white/50 rounded-3xl border border-dashed border-gray-200">
          <LayoutGrid className="w-12 h-12 mb-4 opacity-20" />
          <p className="text-sm font-bold italic">請從左側選單選擇一個課別以開始排班</p>
        </div>
      );
    }

    // 輔助函式：取得可用班別 (不含休假)
    const getAvailableShiftsForEmp = (empId: string) => {
      const emp = state.employees.find(e => e.id === empId);
      if (!emp) return ["出勤"];
      
      const deptOfEmp = state.departments.find(d => d.id === emp.deptId);
      const skills = deptOfEmp?.skills || [];
      
      const options = ["出勤"];
      skills.forEach(s => options.push(`出勤(${s})`));
      return options;
    };

    return (
      <div className="flex flex-col h-full relative" onClick={() => activeMenu && setActiveMenu(null)}>
        {/* Dashboard Header: Simplified Tooltip or Summary */}
        <div className="flex items-center justify-between mb-4 bg-white/60 p-3 px-5 rounded-2xl border border-gray-100 backdrop-blur-sm">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-green-500 rounded-full"></div>
              <span className="text-[11px] font-black text-gray-600">點擊儲存格：標記「自休」</span>
            </div>
            <div className="flex items-center gap-2">
              <Clock className="w-3.5 h-3.5 text-morandi-sidebar" />
              <span className="text-[11px] font-black text-gray-600">滑鼠移至右側選單：設定「出勤時間」</span>
            </div>
          </div>
          <div className="text-[10px] text-gray-400 font-bold bg-gray-50 px-3 py-1 rounded-full border border-gray-100">
             * 系休與國定假由系統自動判定，必要時可從選單修正
          </div>
        </div>

        <div className="flex-1 overflow-auto border border-gray-200 rounded-3xl bg-white shadow-xl custom-scrollbar relative">
          <div className="grid min-w-max" style={{ gridTemplateColumns: `160px repeat(${days.length}, 64px)` }}>
            {/* Header Column... */}
            <div className="p-4 bg-morandi-sidebar text-white font-bold flex items-center border-r border-gray-200 sticky left-0 z-40 h-[50px]">
              <Users className="w-4 h-4 mr-2" /> 名冊 / 技能
            </div>
            
            {/* Header: Date Cells */}
            {days.map((date, i) => {
              const dateStrFull = format(date, "yyyy-MM-dd");
              const isHoliday = state.nationalHolidays.includes(dateStrFull);
              const isWk = isWeekend(date);
              const dateStr = format(date, "MM/dd");
              const dayStr = format(date, "EEEEEE", { locale: zhTW });
              return (
                <div 
                  key={i} 
                  className={cn(
                    "table-header-cell border-b border-gray-200 sticky top-0 z-30 h-[50px]",
                    isHoliday ? "bg-red-50 text-red-600 font-black" : (isWk ? "weekend-bg text-morandi-text" : "bg-morandi-purple-light text-morandi-purple-dark")
                  )}
                >
                  <span className="text-[10px] opacity-70">{dateStr}</span>
                  <span className="text-[12px] uppercase">{dayStr}</span>
                </div>
              );
            })}

            {/* Body: Employee Rows */}
            {dept.groups.map(group => (
              <div key={group.id} className="contents font-mono">
                <div className="col-span-full bg-gray-50/80 p-2.5 px-6 text-xs font-black text-morandi-sidebar border-b border-gray-200/50 flex items-center sticky left-0 z-20 backdrop-blur-sm shadow-[0_2px_5px_rgba(0,0,0,0.02)]">
                  <div className="w-2 h-2 bg-morandi-sidebar rounded-full mr-3 animate-pulse"></div>
                  {group.name} 
                  <span className="mx-3 opacity-30">|</span>
                  <span className="text-gray-400 font-bold">上限 {group.maxLeavePerDay} 人</span>
                </div>
                {state.employees.filter(e => e.groupId === group.id).map(emp => (
                  <div key={emp.id} className="contents group">
                    <div className="p-3 text-sm font-medium border-r border-gray-200 border-b border-gray-100 bg-white sticky left-0 z-30 flex flex-col items-start justify-center shadow-lg shadow-black/5">
                      <span className="text-morandi-sidebar font-black leading-tight truncate w-full">{emp.name}</span>
                      <div className="flex gap-1 mt-1">
                        {emp.skills.slice(0, 1).map((s, i) => (
                          <span key={i} className="px-1.5 py-0.5 bg-gray-100 text-gray-500 rounded text-[9px] font-black uppercase tracking-tighter">{s}</span>
                        ))}
                      </div>
                    </div>
                    {days.map((date, i) => {
                      const dateStr = format(date, "yyyy-MM-dd");
                      const record = state.records[`${emp.id}_${dateStr}`];
                      const isLeave = record?.type === "leave";
                      const isWk = isWeekend(date);

                      return (
                        <div 
                          key={dateStr} 
                          className={cn(
                            "table-body-cell border-b border-gray-100 relative group/cell transition-all min-h-[56px] flex flex-col items-center justify-center overflow-hidden select-none",
                            isWk ? "weekend-bg hover:bg-morandi-sidebar/5" : "hover:bg-gray-50",
                            record?.isManual && "bg-amber-50/10",
                            isLeave && record.code === "自休" && "bg-green-50/60",
                            activeMenu?.empId === emp.id && activeMenu?.dateStr === dateStr && "ring-2 ring-inset ring-morandi-sidebar bg-white z-20"
                          )}
                          // 點擊：直接切換自休
                          onClick={(e) => {
                            const nextVal = (isLeave && record.code === "自休") ? "無" : "自休";
                            handleTimeChange(emp.id, dateStr, nextVal);
                          }}
                          // 右鍵：開啟選單
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setActiveMenu({ empId: emp.id, dateStr, x: e.clientX, y: e.clientY });
                          }}
                        >
                            <div className={cn(
                              "text-center font-black transition-all z-10 pointer-events-none",
                              isLeave ? (
                                record.code === "自休" ? "text-green-600 text-xs scale-110" :
                                record.code === "系休" ? "text-blue-600 text-[10px]" : "text-red-500 text-[10px]"
                              ) : (
                                record?.time ? "text-gray-900 text-[11px]" : "text-gray-200 text-[9px] opacity-0 group-hover/cell:opacity-50"
                              )
                            )}>
                              {isLeave ? record.code : (record?.time || "自動")}
                            </div>

                            {/* 僅點擊下方微小按鈕才開啟選單 (防誤觸) */}
                            <button 
                              onClick={(e) => {
                                e.stopPropagation(); 
                                const rect = e.currentTarget.getBoundingClientRect();
                                setActiveMenu({ empId: emp.id, dateStr, x: rect.left, y: rect.bottom });
                              }}
                              className="absolute bottom-0.5 right-0.5 p-1 rounded bg-white/80 opacity-0 group-hover/cell:opacity-100 hover:bg-white hover:shadow-sm transition-all z-20 text-gray-400 hover:text-morandi-sidebar"
                              title="右鍵或點擊以選擇班別"
                            >
                              <MoreHorizontal className="w-2.5 h-2.5" />
                            </button>
                            
                            {/* 視覺裝飾 */}
                            {isLeave && record.code === "自休" && (
                              <div className="absolute top-1.5 left-1.5 w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse shadow-[0_0_8px_rgba(34,197,94,0.5)]"></div>
                            )}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>

        {/* --- 自定義浮動選單 (解決點選不便問題) --- */}
        <AnimatePresence>
          {activeMenu && (
            <>
              {/* 背景遮罩，點擊關閉 */}
              <div 
                className="fixed inset-0 z-50 pointer-events-auto" 
                onClick={() => setActiveMenu(null)}
              />
              <motion.div
                initial={{ opacity: 0, scale: 0.95, y: -10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: -10 }}
                style={{ 
                    position: 'fixed',
                    left: Math.min(activeMenu.x, window.innerWidth - 220), 
                    top: Math.min(activeMenu.y + 8, window.innerHeight - 300) 
                }}
                className="z-[60] w-52 bg-white rounded-2xl shadow-2xl border border-gray-100 overflow-hidden"
                onClick={(e) => e.stopPropagation()}
              >
                {/* 班別選擇區 */}
                <div className="p-2 border-b border-gray-50">
                  <div className="px-3 py-1.5 text-[10px] font-black text-gray-400 uppercase tracking-widest flex items-center">
                    <Clock className="w-3 h-3 mr-1.5" /> 選擇出勤班別
                  </div>
                  <div className="grid grid-cols-1 gap-1 mt-1">
                    {getAvailableShiftsForEmp(activeMenu.empId).map(t => (
                      <button
                        key={t}
                        onClick={() => {
                          handleTimeChange(activeMenu.empId, activeMenu.dateStr, t);
                          setActiveMenu(null);
                        }}
                        className="w-full text-left px-4 py-2 text-xs font-black text-morandi-sidebar hover:bg-gray-50 rounded-xl transition-colors flex justify-between items-center group"
                      >
                        {t}
                        <Plus className="w-3 h-3 opacity-0 group-hover:opacity-100 text-morandi-purple-mid" />
                      </button>
                    ))}
                  </div>
                </div>

                {/* 特殊休假區 (與班別分開) */}
                <div className="p-2 bg-gray-50/50">
                  <div className="px-3 py-1.5 text-[10px] font-black text-gray-400 uppercase tracking-widest flex items-center">
                    <ShieldCheck className="w-3 h-3 mr-1.5" /> 系統與預設
                  </div>
                  <div className="grid grid-cols-1 gap-1 mt-1">
                    <button
                        onClick={() => {
                            handleTimeChange(activeMenu.empId, activeMenu.dateStr, "無");
                            setActiveMenu(null);
                        }}
                        className="w-full text-left px-4 py-2 text-[10px] font-black text-gray-500 hover:bg-white rounded-xl transition-colors border border-dashed border-gray-200 flex justify-between items-center"
                    >
                        還原為系統自動排列
                        <RotateCcw className="w-3 h-3 opacity-50" />
                    </button>
                  </div>
                  <div className="grid grid-cols-2 gap-1 mt-2">
                    {[
                      { code: "系休", color: "text-blue-600 bg-white hover:bg-blue-50 border-blue-100" },
                      { code: "國定", color: "text-red-500 bg-white hover:bg-red-50 border-red-100" }
                    ].map(item => (
                      <button
                        key={item.code}
                        onClick={() => {
                          handleTimeChange(activeMenu.empId, activeMenu.dateStr, item.code);
                          setActiveMenu(null);
                        }}
                        className={cn(
                          "py-2 rounded-lg text-[10px] font-black transition-all border",
                          item.color
                        )}
                      >
                        {item.code}
                      </button>
                    ))}
                  </div>
                </div>
              </motion.div>
            </>
          )}
        </AnimatePresence>
      </div>
    );
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-morandi-bg flex items-center justify-center">
        <div className="flex flex-col items-center">
          <div className="w-12 h-12 border-4 border-morandi-sidebar border-t-transparent rounded-full animate-spin mb-4"></div>
          <p className="text-morandi-sidebar font-bold animate-pulse">雲端資料載入中...</p>
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen bg-morandi-bg flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-md w-full bg-white rounded-3xl shadow-2xl p-8 text-center border border-morandi-purple-light"
        >
          <div className="w-20 h-20 bg-morandi-purple-light/30 rounded-2xl flex items-center justify-center mx-auto mb-6">
            <LayoutGrid className="w-10 h-10 text-morandi-sidebar" />
          </div>
          <h1 className="text-2xl font-black text-morandi-sidebar mb-2 tracking-tight">物流智能排班系統</h1>
          <p className="text-gray-500 text-sm mb-8">請先登入 Google 帳號以啟用雲端同步功能</p>
          <button 
            onClick={handleLogin}
            className="w-full flex items-center justify-center py-4 bg-morandi-sidebar text-white rounded-2xl font-black shadow-xl shadow-morandi-sidebar/20 hover:scale-[1.02] transition-all"
          >
            <Play className="w-5 h-5 mr-3" /> 使用 Google 登入系統
          </button>
          <p className="mt-6 text-[10px] text-gray-400">大溪理貨/倉儲/運務課 專用後台</p>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-morandi-bg overflow-hidden text-morandi-text">
      {/* Sidebar */}
      <aside className="w-64 bg-morandi-sidebar text-white flex flex-col shadow-xl z-30">
        <div className="p-8 border-b border-white/10">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-white/10 rounded-2xl backdrop-blur-sm shadow-inner">
              <Clock className="w-6 h-6 text-morandi-purple-light" />
            </div>
            <div>
              <h1 className="font-black text-lg tracking-tighter leading-none">智能排班系統</h1>
              <p className="text-[10px] opacity-60 font-bold tracking-widest mt-1">ENTERPRISE EDITION</p>
            </div>
          </div>
        </div>

        <nav className="flex-1 p-4 space-y-8 overflow-y-auto">
          <div className="space-y-1">
            <div className="px-3 mb-4 text-[10px] font-black text-white/40 uppercase tracking-[0.2em]">系統模式</div>
            <button 
              onClick={() => setViewMode("front")}
              className={cn(
                "w-full flex items-center px-4 py-3 rounded-2xl transition-all font-bold text-sm",
                viewMode === "front" ? "bg-white text-morandi-sidebar shadow-lg" : "hover:bg-white/10 text-white/70 hover:text-white"
              )}
            >
              <LayoutGrid className="w-4 h-4 mr-3" /> 前台排班主面板
            </button>
            <button 
              onClick={() => setViewMode("back")}
              className={cn(
                "w-full flex items-center px-4 py-3 rounded-2xl transition-all font-bold text-sm",
                viewMode === "back" ? "bg-white text-morandi-sidebar shadow-lg" : "hover:bg-white/10 text-white/70 hover:text-white"
              )}
            >
              <Settings className="w-4 h-4 mr-3" /> 後台邏輯管理
            </button>
          </div>

          <div className="space-y-1">
            <div className="px-3 mb-4 text-[10px] font-black text-white/40 uppercase tracking-[0.2em]">組織層級</div>
            {state.departments.map(dept => (
              <div key={dept.id}>
                <div 
                  onClick={() => setActiveDept(dept.id)}
                  className={cn(
                    "px-4 py-3 text-sm font-bold cursor-pointer rounded-2xl flex items-center justify-between transition-all",
                    activeDept === dept.id ? "bg-black/20 text-white" : "text-white/50 hover:bg-white/5 hover:text-white"
                  )}
                >
                  <span className="truncate">{dept.name}</span>
                  <ChevronRight className={cn("w-3 h-3 transition-transform", activeDept === dept.id && "rotate-90")} />
                </div>
                {activeDept === dept.id && (
                  <div className="ml-6 mt-1 space-y-1 border-l border-white/10 pl-4 py-1">
                    {dept.groups.map(g => (
                      <div key={g.id} className="text-[11px] py-1.5 text-white/60 hover:text-white cursor-pointer transition-colors">
                        • {g.name}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </div>
        </nav>

        <div className="p-6 border-t border-white/5 bg-black/5">
          <div className="flex items-center p-3 bg-white/10 rounded-2xl mb-4">
            <div className="w-10 h-10 rounded-xl bg-white flex items-center justify-center text-morandi-sidebar font-black overflow-hidden shrink-0">
              {user.photoURL ? <img src={user.photoURL} alt="" referrerPolicy="no-referrer" /> : user.displayName?.charAt(0)}
            </div>
            <div className="ml-3 flex-1 overflow-hidden">
              <p className="text-xs font-black truncate">{user.displayName}</p>
              <button onClick={handleLogout} className="text-[10px] text-red-300 hover:text-red-500 block">登出系統</button>
            </div>
          </div>
          <div className="text-[9px] opacity-40 font-mono text-center">v2.5.1 STABLE — AUTO-SCHEDULER</div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        {/* Toolbar */}
        <header className="shrink-0 pt-8 px-8 pb-4 flex justify-between items-end gap-4 bg-morandi-bg/80 backdrop-blur-sm sticky top-0 z-20">
          <div className="flex-1">
            <h2 className="text-3xl font-black tracking-tighter text-morandi-text">
              {viewMode === "front" ? "28天智能自動排班面板" : "系統全域設定"}
            </h2>
            <div className="flex items-center gap-4 mt-2">
              <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-gray-200 shadow-sm">
                <Calendar className="w-4 h-4 text-morandi-sidebar" />
                <span className="text-xs font-bold text-gray-500">起始日期：</span>
                <input 
                  type="date" 
                  value={startDate}
                  onChange={(e) => handleStartDateChange(e.target.value)}
                  className="text-xs font-bold text-morandi-text bg-transparent border-none outline-none cursor-pointer"
                />
              </div>
              {viewMode === "front" && (
                <div className="text-[10px] text-morandi-purple-dark font-black tracking-widest uppercase">
                  ACTIVE PERIOD: {startDate} ~ {format(addDays(parseISO(startDate), 27), "yyyy-MM-dd")}
                </div>
              )}
            </div>
          </div>

          <div className="flex gap-2 mb-1">
            {viewMode === "front" ? (
              <>
                <label className="flex items-center px-4 py-2 bg-white text-morandi-sidebar border border-morandi-purple-light rounded-xl text-xs font-bold hover:bg-morandi-purple-light/10 transition-all cursor-pointer">
                  <Upload className="w-4 h-4 mr-2" /> 匯入 .xlsx
                  <input type="file" accept=".xlsx" onChange={importFromExcel} className="hidden" />
                </label>
                <button 
                  onClick={exportToExcel}
                  className="flex items-center px-4 py-2 bg-white text-morandi-sidebar border border-morandi-purple-light rounded-xl text-xs font-bold hover:bg-morandi-purple-light/10 transition-all"
                >
                  <Download className="w-4 h-4 mr-2" /> 匯出 .xlsx
                </button>
                <button 
                  onClick={runAutoSchedule}
                  className="flex items-center px-6 py-2 bg-morandi-sidebar text-white rounded-xl text-xs font-black shadow-lg shadow-morandi-sidebar/20 hover:scale-105 transition-all"
                >
                  <Play className="w-4 h-4 mr-2" /> 執行智能自動排班
                </button>
              </>
            ) : (
              <div className="flex gap-2">
                <button 
                  onClick={downloadEmployeeTemplate}
                  className="flex items-center px-4 py-2 bg-amber-50 text-amber-600 border border-amber-200 rounded-xl text-xs font-black hover:bg-amber-100 transition-all shadow-sm"
                  title="下載 Excel 人事公版格式"
                >
                  <FileSpreadsheet className="w-4 h-4 mr-2" /> 下載公版範本
                </button>
                
                <label className="flex items-center px-4 py-2 bg-morandi-sidebar text-white rounded-xl text-xs font-black shadow-lg shadow-morandi-sidebar/20 hover:scale-105 transition-all cursor-pointer">
                  <Upload className="w-4 h-4 mr-2" /> 全域人事匯入 (自動分配)
                  <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleImportEmployees(e)} />
                </label>

                <button className="flex items-center px-6 py-2 bg-white text-morandi-sidebar border border-gray-200 rounded-xl text-xs font-black hover:bg-gray-50 transition-all">
                  <CheckCircle2 className="w-4 h-4 mr-2" /> 儲存設定
                </button>
              </div>
            )}
          </div>
        </header>

        {/* Dynamic Content */}
        <div className="flex-1 overflow-auto p-8 pt-4">
          <AnimatePresence mode="wait">
            <motion.div
              key={viewMode + activeDept}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.15 }}
            >
              {viewMode === "front" ? <ScheduleTable /> : (
                <div className="max-w-4xl p-8 bg-white rounded-3xl border border-gray-200 shadow-sm space-y-8">
                    <section>
                      <div className="flex justify-between items-center mb-4 border-b pb-2">
                        <h3 className="text-sm font-black uppercase tracking-[0.2em] text-morandi-sidebar">課別技能池管理</h3>
                        <p className="text-[10px] text-gray-400 font-bold italic">* 定義該課別下所有組別通用的技能項目</p>
                      </div>
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {state.departments.map(dept => (
                          <div key={dept.id} className="p-4 bg-gray-50/50 rounded-2xl border border-gray-100 space-y-3">
                            <div className="flex justify-between items-center px-1">
                              <span className="text-xs font-black text-morandi-sidebar">{dept.name}</span>
                              <div className="flex gap-2">
                                <input 
                                  type="text" 
                                  id={`new-skill-${dept.id}`}
                                  placeholder="新增技能名稱"
                                  className="w-24 bg-white border border-gray-200 rounded-lg px-2 py-1 text-[10px] font-bold outline-none focus:ring-1 focus:ring-morandi-sidebar"
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') {
                                      const el = e.target as HTMLInputElement;
                                      if (el.value.trim()) {
                                        updateDepartmentSkills(dept.id, 'add', el.value.trim());
                                        el.value = "";
                                      }
                                    }
                                  }}
                                />
                                <button
                                  onClick={() => {
                                    const el = document.getElementById(`new-skill-${dept.id}`) as HTMLInputElement;
                                    if (el.value.trim()) {
                                      updateDepartmentSkills(dept.id, 'add', el.value.trim());
                                      el.value = "";
                                    }
                                  }}
                                  className="p-1 bg-morandi-sidebar text-white rounded-lg hover:scale-105 transition-all"
                                >
                                  <Plus className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </div>
                            <div className="flex flex-wrap gap-2 min-h-[32px]">
                              {dept.skills.length > 0 ? dept.skills.map(skill => (
                                <span 
                                  key={skill} 
                                  className="flex items-center gap-1.5 px-2.5 py-1 bg-white border border-gray-200 rounded-lg text-[10px] font-bold text-gray-600 group hover:border-red-200 hover:text-red-500 transition-all cursor-default"
                                >
                                  {skill}
                                  <button 
                                    onClick={() => updateDepartmentSkills(dept.id, 'remove', skill)}
                                    className="opacity-0 group-hover:opacity-100 transition-opacity"
                                  >
                                    <X className="w-2.5 h-2.5" />
                                  </button>
                                </span>
                              )) : (
                                <span className="text-[10px] text-gray-300 italic px-1">暫無定義技能</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </section>


                   <section className="grid grid-cols-2 gap-8">
                      <div>
                        <div className="flex justify-between items-end mb-4 border-b pb-2">
                          <h3 className="text-sm font-black uppercase tracking-[0.2em] text-morandi-sidebar">國定假日設定</h3>
                          <div className="text-[10px] font-bold text-amber-600 bg-amber-50 px-2 py-0.5 rounded-lg border border-amber-100 italic">
                            本週期內共有 {periodHolidays.length} 天國休
                          </div>
                        </div>
                        <div className="space-y-4">
                           <div className="flex gap-2">
                             <button 
                               onClick={fetchGovHolidays}
                               className="flex-1 flex items-center justify-center gap-2 bg-amber-50 text-amber-700 border border-amber-200 py-2 rounded-xl text-[10px] font-black hover:bg-amber-100 transition-all"
                             >
                               <Upload className="w-3 h-3" /> 自動搜集節日 ({parseISO(startDate).getFullYear()})
                             </button>
                             <button 
                               onClick={applyHolidayQuotaToAll}
                               className="flex-1 flex items-center justify-center gap-2 bg-morandi-purple-light/20 text-morandi-sidebar border border-morandi-purple-light py-2 rounded-xl text-[10px] font-black hover:bg-morandi-purple-light/40 transition-all"
                               title="將本週期內偵測到的國慶天數自動填入所有員工的配額中"
                             >
                               <CheckCircle2 className="w-3 h-3" /> 套用週期天數至全員
                             </button>
                           </div>
                          
                          {periodHolidays.length > 0 && (
                            <div className="p-3 bg-morandi-sidebar/5 rounded-2xl border border-morandi-sidebar/10 mb-4">
                              <p className="text-[10px] font-black text-morandi-sidebar mb-2 uppercase tracking-widest">本週期涵蓋假日：</p>
                              <div className="flex flex-wrap gap-2">
                                {periodHolidays.map(h => (
                                  <span key={h} className="px-2 py-1 bg-white border border-morandi-sidebar/20 rounded-lg text-[10px] font-mono font-bold text-morandi-sidebar shadow-sm">
                                    {format(parseISO(h), "MM/dd")} {h.split("-")[1] === "10" && h.split("-")[2] === "10" ? "國慶日" : ""}
                                  </span>
                                ))}
                              </div>
                            </div>
                          )}

                          <div className="flex gap-2">
                            <input 
                              type="date" 
                              id="new-holiday"
                              className="flex-1 bg-gray-50 border border-gray-200 rounded-xl px-3 py-2 text-xs" 
                            />
                            <button 
                              onClick={() => {
                                const el = document.getElementById("new-holiday") as HTMLInputElement;
                                if (el.value) toggleHoliday(el.value);
                              }}
                              className="bg-morandi-sidebar text-white px-4 py-2 rounded-xl text-xs font-bold"
                            >
                              新增
                            </button>
                          </div>
                          <div className="max-h-40 overflow-y-auto space-y-2 pr-2">
                            {state.nationalHolidays.sort().map(d => {
                              const isInPeriod = periodHolidays.includes(d);
                              return (
                                <div key={d} className={cn(
                                  "flex justify-between items-center p-2 px-3 rounded-xl border shrink-0 transition-colors",
                                  isInPeriod ? "bg-amber-50 border-amber-200 shadow-sm" : "bg-gray-50 border-gray-100"
                                )}>
                                  <div className="flex items-center gap-2">
                                    <span className="text-xs font-mono font-bold text-morandi-text">{d}</span>
                                    {isInPeriod && <span className="text-[8px] bg-amber-500 text-white px-1 rounded font-black animate-pulse">當季</span>}
                                  </div>
                                  <button onClick={() => toggleHoliday(d)} className="text-red-300 hover:text-red-500 transition-colors">
                                    <X className="w-3 h-3" />
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>

                      <div>
                        <h3 className="text-sm font-black uppercase tracking-[0.2em] mb-4 text-morandi-sidebar border-b pb-2">可休假區間限制</h3>
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest pl-1">區間開始日期</label>
                            <input 
                              type="date" 
                              value={state.leaveInterval.start}
                              onChange={(e) => updateLeaveInterval("start", e.target.value)}
                              className="w-full bg-gray-50 border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest pl-1">區間結束日期</label>
                            <input 
                              type="date" 
                              value={state.leaveInterval.end}
                              onChange={(e) => updateLeaveInterval("end", e.target.value)}
                              className="w-full bg-gray-50 border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold"
                            />
                          </div>
                          <p className="text-[10px] text-amber-500 font-bold leading-relaxed px-1">
                            * 注意：設定後，僅區間內日期可點選自休、系休、國定。
                          </p>
                        </div>
                      </div>
                   </section>

                     <section>
                       <div className="flex justify-between items-end mb-2 border-b pb-2">
                         <h3 className="text-sm font-black uppercase tracking-[0.2em] text-morandi-sidebar">組別編制與人力規則</h3>
                         <p className="text-[10px] text-gray-400 font-bold italic">* 提示：此處設定為「基本預設值」，可點擊「日期細項設定」針對特定日期微調</p>
                       </div>
                       <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                         {state.departments.map(dept => (
                           <div key={dept.id} className="space-y-4">
                             <div className="space-y-3">
                               {dept.groups.map(group => (
                                 <div key={group.id} className={cn(
                                   "p-4 bg-white border rounded-[2rem] shadow-sm space-y-4 transition-all duration-300",
                                   selectedGroupForDaily === group.id ? "border-morandi-sidebar ring-2 ring-morandi-sidebar/10 transform scale-[1.02]" : "border-gray-100 hover:border-morandi-purple-light"
                                 )}>
                                   <div className="flex justify-between items-center">
                                     <div className="flex items-center gap-2">
                                       <div className="flex items-center gap-2">
                                         <span className="text-sm font-black text-morandi-sidebar tracking-tight">{group.name}</span>
                                         <label className="flex items-center gap-1.5 px-2 py-1 bg-morandi-sidebar/10 text-morandi-sidebar border border-morandi-sidebar/20 rounded-lg hover:bg-morandi-sidebar text-white transition-all cursor-pointer group/import shadow-sm" title="快速匯入此組人員">
                                           <UserPlus className="w-3.5 h-3.5" />
                                           <span className="text-[9px] font-black">一鍵匯入</span>
                                           <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleImportEmployees(e, group.id)} />
                                         </label>
                                       </div>
                                     </div>
                                     <button 
                                       onClick={() => {
                                         const isSame = selectedGroupForDaily === group.id;
                                         setSelectedGroupForDaily(isSame ? "" : group.id);
                                         if (!isSame) {
                                           setTimeout(() => {
                                             document.getElementById('daily-scroll-anchor')?.scrollIntoView({ behavior: 'smooth' });
                                           }, 100);
                                         }
                                       }}
                                       className={cn(
                                         "px-4 py-1.5 rounded-xl text-[10px] font-black transition-all shadow-sm",
                                         selectedGroupForDaily === group.id ? "bg-morandi-sidebar text-white" : "bg-gray-100 text-gray-400 hover:bg-morandi-purple-light/20 hover:text-morandi-sidebar"
                                       )}
                                     >
                                       {selectedGroupForDaily === group.id ? "設定中..." : "日期細項設定"}
                                     </button>
                                   </div>
                                   <div className="grid grid-cols-2 gap-4">
                                     <div className="space-y-1.5 col-span-2">
                                       <label className="text-[9px] font-black text-gray-400 uppercase tracking-widest flex items-center gap-1">
                                         <ShieldCheck className="w-3 h-3" /> 基本代休上限 (人/日)
                                       </label>
                                       <input 
                                         type="number"
                                         value={group.maxLeavePerDay}
                                         onChange={(e) => updateGroupConfig(dept.id, group.id, 'maxLeavePerDay', parseInt(e.target.value) || 0)}
                                         className="w-full bg-gray-50 border border-gray-100 rounded-xl px-4 py-2 text-xs font-black focus:ring-1 focus:ring-morandi-sidebar outline-none shadow-inner"
                                       />
                                     </div>
                                     
                                     <div className="col-span-2 space-y-3 pt-3 border-t border-gray-50">
                                       <div className="flex justify-between items-center">
                                         <label className="text-[9px] font-black text-gray-400 uppercase tracking-widest flex items-center gap-1">
                                            <Award className="w-3 h-3" /> 各項技能人力需求 (每日)
                                         </label>
                                         <div className="flex gap-1 items-center">
                                            <select 
                                              onChange={(e) => {
                                                if (e.target.value) {
                                                  updateGroupSkills(dept.id, group.id, 'add', e.target.value);
                                                  e.target.value = "";
                                                }
                                              }}
                                              className="text-[9px] w-32 bg-gray-50 border border-gray-100 rounded-lg px-2 py-0.5 font-bold focus:outline-none focus:ring-1 focus:ring-morandi-sidebar"
                                            >
                                              <option value="">＋從池中選技能</option>
                                              {dept.skills.map(s => (
                                                <option key={s} value={s}>{s}</option>
                                              ))}
                                            </select>
                                         </div>
                                       </div>

                                       <div className="grid grid-cols-2 gap-2 text-[10px]">
                                         {group.skillRequirements.map(req => (
                                           <div key={req.skill} className="flex items-center justify-between p-2 bg-gray-50/50 rounded-xl border border-gray-100 group/item">
                                             <div className="flex items-center gap-2">
                                               <button 
                                                 onClick={() => updateGroupSkills(dept.id, group.id, 'remove', req.skill)}
                                                 className="p-1 text-red-300 hover:text-red-500 hover:bg-red-50 rounded"
                                               >
                                                 <Trash2 className="w-2.5 h-2.5" />
                                               </button>
                                               <span className="font-bold text-gray-500">{req.skill}</span>
                                             </div>
                                             <input 
                                               type="number"
                                               value={req.count}
                                               onChange={(e) => {
                                                 const val = parseInt(e.target.value) || 0;
                                                 const newSkills = [...group.skillRequirements];
                                                 const idx = newSkills.findIndex(s => s.skill === req.skill);
                                                 if (idx > -1) newSkills[idx] = { ...newSkills[idx], count: val };
                                                 updateGroupConfig(dept.id, group.id, 'skillRequirements', newSkills);
                                               }}
                                               className="w-12 bg-white border border-gray-200 rounded-lg text-center py-1 font-black focus:ring-1 focus:ring-morandi-sidebar outline-none shadow-sm"
                                             />
                                           </div>
                                         ))}
                                         {group.skillRequirements.length === 0 && (
                                           <p className="col-span-2 text-center py-2 text-[9px] text-gray-300 italic font-medium">該組別目前無特定技能需求</p>
                                         )}
                                       </div>
                                     </div>
                                   </div>
                                 </div>
                               ))}
                             </div>
                           </div>
                         ))}
                       </div>

                       <div id="daily-scroll-anchor" className="pt-2" />

                       {selectedGroupForDaily && (
                         <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4 duration-500 mb-8 pb-8">
                           <div className="flex justify-between items-center bg-white p-6 rounded-[2.5rem] border border-gray-100 shadow-xl shadow-morandi-sidebar/5">
                             <div className="flex items-center gap-4">
                               <div className="p-4 bg-morandi-sidebar/5 rounded-2xl">
                                 <CalendarDays className="w-6 h-6 text-morandi-sidebar" />
                               </div>
                               <div>
                                 <div className="flex items-center gap-2">
                                   <h4 className="text-lg font-black text-morandi-sidebar">
                                     {state.departments.flatMap(d => d.groups).find(g => g.id === selectedGroupForDaily)?.name} — 特殊日期微調
                                   </h4>
                                   <span className="px-2 py-0.5 bg-amber-100 text-amber-700 rounded text-[9px] font-black uppercase">本週期 28 天</span>
                                 </div>
                                 <p className="text-[11px] text-gray-400 font-medium mt-1">此處設定將覆蓋左側的「基本預設值」，適用於週末、連假或大促旺季</p>
                               </div>
                             </div>
                             <div className="flex gap-3">
                               <button 
                                 onClick={() => applyGroupDefaultsToAllDays(selectedGroupForDaily)}
                                 className="flex items-center gap-2 px-6 py-2.5 bg-morandi-sidebar text-white rounded-2xl text-xs font-black hover:bg-morandi-sidebar/90 transition-all shadow-lg shadow-morandi-sidebar/20"
                               >
                                 <RotateCcw className="w-3.5 h-3.5" /> 恢復為預設值
                               </button>
                               <button 
                                 onClick={() => setSelectedGroupForDaily("")}
                                 className="p-3 bg-gray-50 text-gray-400 border border-gray-200 rounded-2xl hover:bg-gray-100 hover:text-gray-600 transition-all"
                                 title="關閉設定"
                               >
                                 <X className="w-4 h-4" />
                               </button>
                             </div>
                           </div>

                          <div className="max-h-[500px] overflow-auto border border-gray-100 rounded-[2rem] shadow-xl bg-white">
                            <table className="w-full text-[10px] text-left border-collapse">
                              <thead className="sticky top-0 bg-gray-50 z-10 shadow-sm border-b border-gray-100">
                                <tr>
                                  <th className="p-4 w-32 font-black uppercase tracking-widest text-gray-400">日期</th>
                                  <th className="p-4 text-center text-red-500 font-black uppercase tracking-widest">休假上限</th>
                                  {(state.departments.flatMap(d => d.groups).find(g => g.id === selectedGroupForDaily)?.skillRequirements || []).map(reqDef => (
                                    <th key={reqDef.skill} className="p-4 text-center font-black uppercase tracking-widest text-morandi-sidebar">{reqDef.skill} 人數</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody className="bg-white">
                                {days.map(day => {
                                  const dateStr = format(day, "yyyy-MM-dd");
                                  const isWk = isWeekend(day);
                                  const isHoliday = state.nationalHolidays.includes(dateStr);
                                  const req = state.dailyRequirements[`${selectedGroupForDaily}_${dateStr}`];
                                  const group = state.departments.flatMap(d => d.groups).find(g => g.id === selectedGroupForDaily);
                                  const dept = state.departments.find(d => d.groups.some(g => g.id === selectedGroupForDaily));
                                  if (!group || !dept) return null;

                                  return (
                                    <tr key={dateStr} className={cn("hover:bg-morandi-bg/50 transition-colors", (isWk || isHoliday) && "bg-amber-50/20")}>
                                      <td className="p-4 border-b border-gray-50 font-mono font-bold">
                                        <div className="flex flex-col">
                                          <span className={cn("text-morandi-text", isHoliday && "text-red-600")}>
                                            {format(day, "MM/dd (EEEEE)", { locale: zhTW })}
                                          </span>
                                          {(isWk || isHoliday) && (
                                            <span className="text-[8px] text-amber-500 font-black tracking-tighter">
                                              {isHoliday ? "HOLIDAY" : "WEEKEND"}
                                            </span>
                                          )}
                                        </div>
                                      </td>
                                      <td className="p-4 border-b border-gray-50 text-center">
                                        <div className="flex flex-col items-center gap-1">
                                          <input 
                                            type="number"
                                            disabled={isBalancing}
                                            className="w-14 bg-gray-50 border border-gray-200 rounded-xl py-2 text-center font-black focus:ring-1 focus:ring-morandi-sidebar outline-none disabled:opacity-50"
                                            value={req ? req.maxLeave : group.maxLeavePerDay}
                                            onChange={(e) => updateDailyReq(selectedGroupForDaily, dateStr, 'maxLeave', undefined, parseInt(e.target.value) || 0)}
                                          />
                                          <button 
                                            onClick={() => rebalanceGroupLeaveLimits(selectedGroupForDaily, dateStr, req ? req.maxLeave : group.maxLeavePerDay)}
                                            disabled={isBalancing}
                                            className="text-[8px] font-black text-morandi-sidebar hover:bg-morandi-purple-light/20 px-1.5 py-0.5 rounded transition-all flex items-center gap-0.5 disabled:opacity-30"
                                            title="點擊後，本週期其餘日期的休假上限將自動平均增減以平衡總量"
                                          >
                                            <RotateCcw className="w-2 h-2" /> 平衡
                                          </button>
                                        </div>
                                      </td>
                                      {group.skillRequirements.map(reqDef => { const skill = reqDef.skill;
                                        const sReq = (req ? req.skillRequirements : group.skillRequirements).find(s => s.skill === skill);
                                        return (
                                          <td key={skill} className="p-4 border-b border-gray-50 text-center">
                                            <input 
                                              type="number"
                                              className="w-14 bg-gray-50 border border-gray-200 rounded-xl py-2 text-center font-black focus:ring-1 focus:ring-morandi-sidebar outline-none"
                                              value={sReq ? sReq.count : 0}
                                              onChange={(e) => updateDailyReq(selectedGroupForDaily, dateStr, 'skillCount', skill, parseInt(e.target.value) || 0)}
                                            />
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                    </section>

                   
                    <section>
                      {importErrors.length > 0 && (
                        <div id="import-error-anchor" className="bg-red-50 border border-red-200 rounded-2xl p-4 mb-6 animate-in fade-in slide-in-from-top-2">
                          <div className="flex items-center justify-between mb-3 px-1">
                            <div className="flex items-center gap-2 text-red-600">
                              <AlertCircle className="w-4 h-4" />
                              <span className="text-xs font-black">匯入異常診斷記錄 ({importErrors.length} 筆)</span>
                            </div>
                            <button 
                              onClick={() => setImportErrors([])} 
                              className="p-1 text-red-400 hover:text-red-600 hover:bg-red-100 rounded-lg transition-all"
                              title="關閉診斷記錄"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>
                          <div className="max-h-48 overflow-y-auto space-y-1.5 pr-2 custom-scrollbar">
                            {importErrors.map((err, idx) => (
                              <div key={idx} className="flex gap-4 text-[10px] text-red-500 font-black bg-white/60 p-2.5 rounded-xl border border-red-100/50 shadow-sm">
                                <div className="bg-red-500 text-white px-2 py-0.5 rounded-md text-[9px] shrink-0">Excel 第 {err.row} 行</div>
                                <div className="w-20 truncate border-r border-red-100 pr-2">{err.name}</div>
                                <div className="flex-1 text-red-400 italic">{err.reason}</div>
                              </div>
                            ))}
                          </div>
                          <p className="mt-3 text-[9px] text-red-400 pl-1 font-bold">
                            * 請根據上方提示檢查 Excel 檔案中的「課別」與「組別」名稱是否完全一致。
                          </p>
                        </div>
                      )}
                      <div className="flex justify-between items-center mb-4 border-b pb-2">
                         <h3 className="text-sm font-black uppercase tracking-[0.2em] text-morandi-sidebar">人事名冊管理</h3>
                         <div className="flex gap-2 text-morandi-sidebar">
                           <label className="flex items-center gap-1.5 px-3 py-1 bg-morandi-purple-light/20 text-morandi-sidebar border border-morandi-purple-light rounded-lg text-[10px] font-black hover:bg-morandi-purple-light/40 transition-all cursor-pointer">
                             <Upload className="w-3 h-3" /> 匯入 Excel
                             <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImportEmployees} />
                           </label>
                           <button 
                             onClick={async () => {
                               if (!showClearConfirm) {
                                 setShowClearConfirm(true);
                                 setTimeout(() => setShowClearConfirm(false), 3000);
                                 return;
                               }
                               setIsClearingAll(true);
                               try {
                                 const batch = writeBatch(db);
                                 state.employees.forEach(e => batch.delete(doc(db, "employees", e.id)));
                                 await batch.commit();
                                 setState(prev => ({ ...prev, employees: [] }));
                                 showNotice("已清空所有員工資料");
                                 setShowClearConfirm(false);
                               } catch (err) {
                                 console.error(err);
                                 showNotice("清空失敗", "error");
                               } finally {
                                 setIsClearingAll(false);
                               }
                             }}
                             disabled={isClearingAll || state.employees.length === 0}
                             className={cn(
                               "flex items-center gap-1.5 px-3 py-1 rounded-lg text-[10px] font-black transition-all border shrink-0",
                               showClearConfirm 
                                 ? "bg-red-500 text-white border-red-600 animate-pulse" 
                                 : "bg-red-50 text-red-500 border-red-100 hover:bg-red-100",
                               isClearingAll && "opacity-50 cursor-not-allowed"
                             )}
                           >
                             <Trash2 className="w-3 h-3" /> 
                             {isClearingAll ? "處理中..." : (showClearConfirm ? "再點擊一次確認清空" : "清空全員")}
                           </button>
                         </div>
                      </div>
                      <div className="table-wrapper border border-gray-100 rounded-2xl overflow-hidden shadow-sm">
                        <table className="w-full text-xs text-left border-collapse">
                          <thead>
                            <tr className="bg-gray-50 border-b border-gray-100">
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest">姓名</th>
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest">組別</th>
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest">具備技能</th>
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest text-center">系休</th>
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest text-center">國休</th>
                              <th className="p-3 font-black text-gray-400 uppercase tracking-widest text-right">操作</th>
                            </tr>
                          </thead>
                          <tbody className="bg-white">
                            {state.employees.map(emp => (
                              <tr key={emp.id} className="border-b border-gray-100 hover:bg-gray-50/50 transition-colors">
                                <td className="p-3 font-bold text-morandi-sidebar">{emp.name}</td>
                                <td className="p-3 text-[10px] text-morandi-sidebar font-medium">{state.departments.flatMap(d => d.groups).find(g => g.id === emp.groupId)?.name || "--"}</td>
                                <td className="p-3 text-[10px] text-gray-400 font-medium whitespace-nowrap overflow-hidden text-ellipsis max-w-[150px]" title={emp.skills.join(", ")}>{emp.skills.join(", ") || "--"}</td>
                                <td className="p-3 text-center">
                                  <input 
                                    type="number" 
                                    className="w-14 bg-gray-50 border border-gray-100 rounded-lg px-2 py-1 text-[11px] font-black text-center focus:ring-1 focus:ring-morandi-sidebar outline-none shadow-inner"
                                    value={emp.systemLeaveQuota}
                                    onChange={(e) => {
                                      const val = parseInt(e.target.value) || 0;
                                      const updatedEmps = state.employees.map(e2 => e2.id === emp.id ? { ...e2, systemLeaveQuota: val } : e2);
                                      setState(prev => ({ ...prev, employees: updatedEmps }));
                                      if (user) setDoc(doc(db, "employees", emp.id), { systemLeaveQuota: val }, { merge: true });
                                    }}
                                  />
                                </td>
                                <td className="p-3 text-center">
                                  <input 
                                    type="number" 
                                    className="w-14 bg-gray-50 border border-gray-100 rounded-lg px-2 py-1 text-[11px] font-black text-center focus:ring-1 focus:ring-morandi-sidebar outline-none shadow-inner"
                                    value={emp.nationalHolidayQuota}
                                    onChange={(e) => {
                                      const val = parseInt(e.target.value) || 0;
                                      const updatedEmps = state.employees.map(e2 => e2.id === emp.id ? { ...e2, nationalHolidayQuota: val } : e2);
                                      setState(prev => ({ ...prev, employees: updatedEmps }));
                                      if (user) setDoc(doc(db, "employees", emp.id), { nationalHolidayQuota: val }, { merge: true });
                                    }}
                                  />
                                </td>
                                <td className="p-3 text-right">
                                  <button 
                                    onClick={() => deleteEmployee(emp.id, emp.name)}
                                    className="p-2 text-gray-300 hover:text-red-500 hover:bg-red-50 rounded-xl transition-all"
                                    title="移除此員工"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                  </button>
                                </td>
                              </tr>
                            ))}
                            {state.employees.length === 0 && (
                              <tr>
                                <td colSpan={6} className="p-12 text-center text-gray-300 italic font-medium">尚未匯入或新增任何員工</td>
                              </tr>
                            )}
                            <tr>
                              <td className="p-3" colSpan={6}>
                                <div className="flex gap-2 items-center bg-gray-50/50 p-2 rounded-xl border border-dashed border-gray-200">
                                  <input id="new-emp-name" placeholder="姓名" className="flex-1 bg-white border border-gray-200 rounded-lg px-3 py-1.5 text-[11px] font-bold outline-none focus:ring-1 focus:ring-morandi-sidebar" />
                                  <select id="new-emp-group" className="bg-white border border-gray-200 rounded-lg px-3 py-1.5 text-[11px] font-bold outline-none focus:ring-1 focus:ring-morandi-sidebar">
                                    <option value="">選擇組別</option>
                                    {state.departments.map(d => d.groups.map(g => (
                                      <option key={g.id} value={g.id}>{d.name} - {g.name}</option>
                                    )))}
                                  </select>
                                  <input id="new-emp-skills" placeholder="技能 (逗號/空白分隔)" className="flex-1 bg-white border border-gray-200 rounded-lg px-3 py-1.5 text-[11px] font-bold outline-none focus:ring-1 focus:ring-morandi-sidebar" />
                                  <button onClick={async () => {
                                    const n = (document.getElementById("new-emp-name") as HTMLInputElement).value;
                                    const g = (document.getElementById("new-emp-group") as HTMLSelectElement).value;
                                    const s = (document.getElementById("new-emp-skills") as HTMLInputElement).value;
                                    if(!n || !g) { showNotice("請輸入姓名並選擇組別", "error"); return; }
                                    const dId = state.departments.find(d => d.groups.some(gr => gr.id === g))?.id || "";
                                    const empId = `emp_${Date.now()}`;
                                    const newEmp = { id: empId, name: n, deptId: dId, groupId: g, skills: s.split(/[,，\s/]+/).filter(Boolean), systemLeaveQuota: 0, nationalHolidayQuota: 0 };
                                    if (user) await setDoc(doc(db, "employees", empId), newEmp);
                                    setState(prev => ({ ...prev, employees: [...prev.employees, newEmp] }));
                                    (document.getElementById("new-emp-name") as HTMLInputElement).value = "";
                                    (document.getElementById("new-emp-skills") as HTMLInputElement).value = "";
                                    showNotice(`成功手動新增員工：${n}`);
                                  }} className="bg-morandi-sidebar text-white px-4 py-1.5 rounded-lg text-[10px] font-black shrink-0 shadow-lg shadow-morandi-sidebar/20 hover:scale-[1.02] transition-all">+ 手動新增人員</button>
                                </div>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </section>

                    <section>
                       <h3 className="text-sm font-black uppercase tracking-[0.2em] mb-4 text-morandi-sidebar border-b pb-2">組織技能定義</h3>
                       <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                         {state.departments.map(d => (
                           <div key={d.id} className="p-4 bg-gray-50 border border-gray-100 rounded-2xl flex justify-between items-center group/dept">
                              <div>
                                <p className="font-bold text-sm tracking-tight text-morandi-sidebar">{d.name}</p>
                                <p className="text-[10px] text-gray-400 mt-1 uppercase tracking-wider font-bold">技能池: {d.skills.join(", ") || "無定義"}</p>
                              </div>
                              <button 
                                onClick={() => {
                                  if (confirm("確定要刪除「" + d.name + "」課別定義嗎？\n(注意：組別與員工的歸屬將受影響，建議謹慎操作)")) {
                                    deleteDepartment(d.id, d.name);
                                  }
                                }}
                                className="p-2.5 bg-white text-gray-300 hover:text-red-500 border border-transparent hover:border-red-100 hover:shadow-sm rounded-xl transition-all"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                           </div>
                         ))}
                       </div>
                    </section>

                </div>
              )}
            </motion.div>
          </AnimatePresence>
        </div>

        {/* Global Notifications */}
        <AnimatePresence>
          {notification && (
            <motion.div 
              initial={{ opacity: 0, bottom: -20, x: "-50%" }}
              animate={{ opacity: 1, bottom: 40, x: "-50%" }}
              exit={{ opacity: 0, bottom: -20, x: "-50%" }}
              className={cn(
                "fixed left-1/2 -translate-x-1/2 z-50 flex items-center px-6 py-4 rounded-3xl shadow-2xl border",
                notification.type === "success" 
                  ? "bg-white border-green-100 text-green-800" 
                  : "bg-white border-red-100 text-red-800 shadow-red-200/50"
              )}
            >
              {notification.type === "success" ? (
                <CheckCircle2 className="w-6 h-6 mr-3 text-green-500" />
              ) : (
                <AlertCircle className="w-6 h-6 mr-3 text-red-500" />
              )}
              <span className="font-black text-sm tracking-tight">{notification.message}</span>
              <button onClick={() => setNotification(null)} className="ml-6 p-1 hover:bg-gray-100 rounded-full transition-colors">
                <X className="w-4 h-4 text-gray-400" />
              </button>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
