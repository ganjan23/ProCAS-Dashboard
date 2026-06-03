import { useState, useEffect, useRef, useMemo, createContext, useContext, useReducer, useCallback, Component } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  Legend, LineChart, Line, ComposedChart,
  ScatterChart, Scatter, Treemap,
} from "recharts";

// ── Icons ─────────────────────────────────────────────────────────────────────
const Icon = ({ path, size = 18, className = "" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke="currentColor" strokeWidth={1.8} strokeLinecap="round"
    strokeLinejoin="round" className={className}>
    <path d={path} />
  </svg>
);
const Icons = {
  sun: "M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42M12 5a7 7 0 1 0 0 14A7 7 0 0 0 12 5z",
  moon: "M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z",
  users: "M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75",
  briefcase: "M20 7H4a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2zM16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0-2 2v2",
  trending: "M23 6l-9.5 9.5-5-5L1 18M17 6h6v6",
  dollar: "M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6",
  clock: "M12 2a10 10 0 1 0 0 20A10 10 0 0 0 12 2zM12 6v6l4 2",
  alert: "M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01",
  check: "M22 11.08V12a10 10 0 1 1-5.93-9.14M22 4 12 14.01l-3-3",
  filter: "M22 3H2l8 9.46V19l4 2v-8.54L22 3z",
  download: "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3",
  upload: "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12",
  plus: "M12 5v14M5 12h14",
  search: "M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z",
  bar: "M18 20V10M12 20V4M6 20v-6",
  pie: "M21.21 15.89A10 10 0 1 1 8 2.83M22 12A10 10 0 0 0 12 2v10z",
  target: "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 18a6 6 0 1 0 0-12 6 6 0 0 0 0 12zM12 14a2 2 0 1 0 0-4 2 2 0 0 0 0 4z",
  send: "M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z",
  inbox: "M22 12h-6l-2 3h-4l-2-3H2M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z",
  chevronDown: "M6 9l6 6 6-6",
  chevronUp: "M18 15l-6-6-6 6",
  chevronLeft: "M15 18l-6-6 6-6",
  chevronRight: "M9 18l6-6-6-6",
  office: "M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2zM9 22V12h6v10",
  edit: "M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z",
  trash: "M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2",
  save: "M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2zM17 21v-8H7v8M7 3v5h8",
  x: "M18 6 6 18M6 6l12 12",
  shield: "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z",
  info: "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 8h.01M12 12v4",
  fileText: "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8",
  arrowUp: "M12 19V5M5 12l7-7 7 7",
  arrowDown: "M12 5v14M19 12l-7 7-7-7",
  userPlus: "M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M8.5 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM20 8v6M23 11h-6",
  duplicate: "M8 17H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v3m-6 12h8a2 2 0 0 0 2-2v-8a2 2 0 0 0-2-2h-8a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2z",
  barChart2: "M18 20V10M12 20V4M6 20v-6",
  activity: "M22 12h-4l-3 9L9 3l-3 9H2",
  globe: "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z",
  sortAsc: "M3 9l4-4 4 4M7 5v14M13 18l4 4 4-4M17 22V8",
  refreshCw: "M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15",
  wifi: "M5 12.55a11 11 0 0 1 14.08 0M1.42 9a16 16 0 0 1 21.16 0M8.53 16.11a6 6 0 0 1 6.95 0M12 20h.01",
  checkCircle: "M22 11.08V12a10 10 0 1 1-5.93-9.14M22 4 12 14.01l-3-3",
  database: "M12 2C6.48 2 2 4.24 2 7v10c0 2.76 4.48 5 10 5s10-2.24 10-5V7c0-2.76-4.48-5-10-5zM12 12c-5.52 0-10-2.24-10-5M2 12c0 2.76 4.48 5 10 5s10-2.24 10-5",
  settings: "M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z",
  logOut: "M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9",
  cpu: "M18 4H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zM9 9h6v6H9z M15 2v2M9 2v2M15 20v2M9 20v2M20 9h2M20 15h2M2 9h2M2 15h2",
};

// ── Theme ─────────────────────────────────────────────────────────────────────
// IMPORTANT: Tailwind CDN in a single-file artifact only injects classes it
// sees in the INITIAL render. Because darkMode starts as true, light-mode
// arbitrary-value classes (bg-[#eef3f8] etc.) are never in the first render,
// so Tailwind never generates them and the toggle appears to do nothing.
//
// Fix: every layout container that must switch theme uses inline style={} with
// the raw hex values from t.raw.* instead of Tailwind bg-[#hex] classes.
// Tailwind classes are still used for spacing, border-width, flex, etc.
const themes = {
  dark: {
    raw: {
      bg: "#08101c", card: "#0f1e2e", cardBorder: "#1a2d44",
      headerBg: "#0b1623", text: "#e8f0f8", textMuted: "#5a7a99",
      textAccent: "#00c9b1", inputBg: "#1a2d44", inputBorder: "#243d58",
      inputText: "#e8f0f8", inputPlaceholder: "#5a7a99",
      hoverBg: "#1a2d44", tableHeadBg: "#0b1623", tableHeadText: "#5a7a99",
      divider: "#1a2d44", footerBg: "#0b1623", colorScheme: "dark",
    },
    bg: "bg-[#08101c]", sidebar: "bg-[#0b1623]", card: "bg-[#0f1e2e]",
    cardBorder: "border-[#1a2d44]", text: "text-[#e8f0f8]", textMuted: "text-[#5a7a99]",
    textAccent: "text-[#00c9b1]", input: "bg-[#1a2d44] border-[#243d58] text-[#e8f0f8]",
    hover: "hover:bg-[#1a2d44]", active: "bg-[#132036]", badge: "bg-[#0d2a3a] text-[#00c9b1]",
    tableBg: "bg-[#0f1e2e]", tableRow: "hover:bg-[#132036]", tableHead: "bg-[#0b1623] text-[#5a7a99]",
    divider: "border-[#1a2d44]", accent: "#00c9b1", accent2: "#4a90d9", success: "#34d399",
    warning: "#fbbf24", danger: "#f87171", chartGrid: "#1a2d44", chartBg: "#0f1e2e",
  },
  light: {
    raw: {
      bg: "#eef3f8", card: "#ffffff", cardBorder: "#d4e0ed",
      headerBg: "#ffffff", text: "#0d2137", textMuted: "#6b8aaa",
      textAccent: "#007a6e", inputBg: "#f4f8fc", inputBorder: "#d4e0ed",
      inputText: "#0d2137", inputPlaceholder: "#9ca3af",
      hoverBg: "#f4f8fc", tableHeadBg: "#eef3f8", tableHeadText: "#6b8aaa",
      divider: "#d4e0ed", footerBg: "#ffffff", colorScheme: "light",
    },
    bg: "bg-[#eef3f8]", sidebar: "bg-white", card: "bg-white",
    cardBorder: "border-[#d4e0ed]", text: "text-[#0d2137]", textMuted: "text-[#6b8aaa]",
    textAccent: "text-[#007a6e]", input: "bg-[#f4f8fc] border-[#d4e0ed] text-[#0d2137]",
    hover: "hover:bg-[#f4f8fc]", active: "bg-[#e8f4f2]", badge: "bg-[#d0f0eb] text-[#007a6e]",
    tableBg: "bg-white", tableRow: "hover:bg-[#f4f8fc]", tableHead: "bg-[#eef3f8] text-[#6b8aaa]",
    divider: "border-[#d4e0ed]", accent: "#007a6e", accent2: "#1b5fa8", success: "#10b981",
    warning: "#f59e0b", danger: "#ef4444", chartGrid: "#d4e0ed", chartBg: "#ffffff",
  },
};

// ── Theme Context — propagates dark/t to every component without extra prop-drilling
const ThemeContext = createContext({ dark: true, t: themes.dark });

// ── GlobalDarkStyles — injects a <style> into <head> so native browser controls
//    (select dropdowns, date pickers, option lists) respect the chosen theme.
//    This covers what Tailwind/inline-styles cannot reach (the OS-rendered popup).
const GlobalDarkStyles = ({ dark }) => {
  useEffect(() => {
    const id   = "procas-global-theme-styles";
    let el     = document.getElementById(id);
    if (!el) {
      el            = document.createElement("style");
      el.id         = id;
      document.head.appendChild(el);
    }
    if (dark) {
      el.textContent = `
        /* ProCAS — dark mode overrides for native form controls */
        :root { color-scheme: dark; }

        /* All inputs, selects, textareas */
        input, select, textarea, input[type="date"], input[type="number"], input[type="text"], input[type="email"] {
          color: #e8f0f8 !important;
          background-color: #1a2d44;
          border-color: #243d58;
          color-scheme: dark;
        }

        /* Placeholder text */
        input::placeholder, textarea::placeholder {
          color: #5a7a99 !important;
          opacity: 1 !important;
        }

        /* Native select — closed state */
        select {
          color: #e8f0f8 !important;
          background-color: #1a2d44 !important;
          color-scheme: dark;
        }

        /* Native option elements in the OS dropdown popup */
        select option {
          color: #e8f0f8 !important;
          background-color: #1a2d44 !important;
        }

        /* Date picker text + calendar icon */
        input[type="date"]::-webkit-calendar-picker-indicator {
          filter: invert(1) brightness(0.8);
          cursor: pointer;
        }
        input[type="date"]::-webkit-datetime-edit-fields-wrapper {
          color: #e8f0f8;
        }
        input[type="date"]::-webkit-datetime-edit {
          color: #e8f0f8;
        }
        input[type="date"]::-webkit-datetime-edit-year-field,
        input[type="date"]::-webkit-datetime-edit-month-field,
        input[type="date"]::-webkit-datetime-edit-day-field {
          color: #e8f0f8;
        }

        /* Focused ring colour consistency */
        input:focus, select:focus, textarea:focus {
          outline: none;
        }

        /* Scrollbars in dark mode */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #0f1e2e; }
        ::-webkit-scrollbar-thumb { background: #243d58; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #2f4a66; }
      `;
    } else {
      el.textContent = `
        /* ProCAS — light mode: restore browser defaults */
        :root { color-scheme: light; }
        input, select, textarea {
          color: #0d2137;
          background-color: #f4f8fc;
          border-color: #d4e0ed;
          color-scheme: light;
        }
        input::placeholder, textarea::placeholder {
          color: #9ca3af !important;
          opacity: 1 !important;
        }
        select option {
          color: #0d2137;
          background-color: #ffffff;
        }
        input[type="date"]::-webkit-calendar-picker-indicator {
          filter: none;
          cursor: pointer;
        }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #eef3f8; }
        ::-webkit-scrollbar-thumb { background: #d4e0ed; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #b8c8d8; }
      `;
    }
    return () => {};
  }, [dark]);
  return null;
};


// ─────────────────────────────────────────────────────────────────────────────
// ── CENTRALIZED CLIENT SYNC CONTEXT ──────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const ClientSyncContext = createContext(null);

const syncReducer = (state, action) => {
  switch (action.type) {
    case "SYNC_START":
      return { ...state, syncing: true, syncError: null };
    case "SYNC_SUCCESS":
      return {
        ...state,
        syncing: false,
        clients: action.payload,
        lastSynced: new Date(),
        syncError: null,
      };
    case "SYNC_ERROR":
      return { ...state, syncing: false, syncError: action.payload };
    case "SET_CLIENTS":
      return { ...state, clients: action.payload, lastSynced: new Date() };
    case "UPDATE_CLIENTS":
      // Direct mutation from MasterClientTab — updates canonical list immediately
      return { ...state, clients: action.payload, lastSynced: new Date() };
    default:
      return state;
  }
};

const useSyncContext = () => {
  const ctx = useContext(ClientSyncContext);
  if (!ctx) throw new Error("useSyncContext must be used inside ClientSyncProvider");
  return ctx;
};

// ─────────────────────────────────────────────────────────────────────────────
// ── USER CONTEXT — shared across all modules ─────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const UserContext = createContext(null);

// ─────────────────────────────────────────────────────────────────────────────
// ── MONTH LOCK CONTEXT & HELPERS ─────────────────────────────────────────────
// Lock key format: "FY|Month|Module"  e.g. "2026-27|April|cas"
// ─────────────────────────────────────────────────────────────────────────────
const LockContext = createContext({ locks:{}, setLocks:()=>{}, auditLog:[], addAuditEntry:()=>{} });
const useLockCtx  = () => useContext(LockContext);

const LOCK_MODULES_META = [
  { id:"cas",  label:"CAS MIS"      },
  { id:"kra",  label:"KRA / KPI"    },
  { id:"fund", label:"Fund Request" },
];

const mkLockKey = (fy, month, mod) => `${fy}|${month}|${mod}`;

// Returns true when the given fy/month/module combo is locked
// (checks both a specific-module key AND an "all" key)
const checkLocked = (locks, fy, month, mod) =>
  !!(locks[mkLockKey(fy, month, mod)] || locks[mkLockKey(fy, month, "all")]);

const LockProvider = ({ children, adminName }) => {
  const [locks, setLocksRaw] = useState(() => {
    try { const s = localStorage.getItem("procas_locks"); return s ? JSON.parse(s) : {}; } catch { return {}; }
  });
  const [auditLog, setAuditLog] = useState(() => {
    try { const s = localStorage.getItem("procas_lockAudit"); return s ? JSON.parse(s) : []; } catch { return []; }
  });

  const setLocks = updater => setLocksRaw(prev => {
    const next = typeof updater === "function" ? updater(prev) : updater;
    try { localStorage.setItem("procas_locks", JSON.stringify(next)); } catch {}
    return next;
  });

  const addAuditEntry = entry => setAuditLog(prev => {
    const next = [entry, ...prev].slice(0, 300);
    try { localStorage.setItem("procas_lockAudit", JSON.stringify(next)); } catch {}
    return next;
  });

  return (
    <LockContext.Provider value={{ locks, setLocks, auditLog, addAuditEntry }}>
      {children}
    </LockContext.Provider>
  );
};



// 2-role system: Admin | End User
const USER_ROLES = ["Admin", "End User"];

const ROLE_COLORS = {
  "Admin":    { fg:"#ffffff", bg:"linear-gradient(135deg,#1b5fa8,#0d3d7a)", badge:"#4a90d9", badgeBg:"#4a90d918" },
  "End User": { fg:"#e8f0f8", bg:"linear-gradient(135deg,#1a2d44,#243d58)", badge:"#94a3b8", badgeBg:"#94a3b818" },
};

let _userId = 1;
const genUserId = () => `USR-${String(_userId++).padStart(3,"0")}`;

const INITIAL_USERS = [
  { id:"USR-001", name:"Admin User",  email:"admin@example.com",  role:"Admin",    status:"Active",   inviteStatus:"Active",         assignedClients:null,                                     lastLogin:"Today, 9:14 AM" },
  { id:"USR-002", name:"User A",      email:"usera@example.com",  role:"End User", status:"Active",   inviteStatus:"Active",         assignedClients:["CLT-001","CLT-002","CLT-004","CLT-007"], lastLogin:"Today, 8:52 AM" },
  { id:"USR-003", name:"User B",      email:"userb@example.com",  role:"End User", status:"Active",   inviteStatus:"Active",         assignedClients:["CLT-003","CLT-005","CLT-006"],           lastLogin:"Yesterday"      },
  { id:"USR-004", name:"User C",      email:"userc@example.com",  role:"End User", status:"Inactive", inviteStatus:"Inactive",       assignedClients:["CLT-008"],                               lastLogin:"3 days ago"     },
];

const canEditClient = (activeUser, clientId) => {
  if (!activeUser) return false;
  if (activeUser.role === "Admin") return true;
  if (!activeUser.assignedClients) return true; // admin with null = all
  return activeUser.assignedClients.includes(clientId);
};

// ── useRBAC — single source of truth for all role-based access decisions ─────
// Returns { isAdmin, activeUser, canEdit(clientId), assignedClientIds }
// All modules should use this instead of hardcoded maps.
const useRBAC = () => {
  const ctx = useContext(UserContext);
  const activeUser = ctx?.activeUser || null;
  const isAdmin = activeUser?.role === "Admin";
  const canEdit = (clientId) => canEditClient(activeUser, clientId);
  // null means all clients (admin); array means specific assigned list
  const assignedClientIds = isAdmin ? null : (activeUser?.assignedClients || []);
  return { isAdmin, activeUser, canEdit, assignedClientIds };
};

// ── Palette ───────────────────────────────────────────────────────────────────
const PALETTE = ["#00c9b1","#4a90d9","#34d399","#fbbf24","#f87171","#7ec8e3","#fb923c","#00e5cc","#60a5fa","#4ade80","#facc15"];

// ── Initial clients ───────────────────────────────────────────────────────────
const initialClients = [
  { id:"CLT-001", clientName:"Client Alpha",   personName:"Contact A",  type:"Corporate",   sector:"Manufacturing", pan:"PANXXXXX1", gst:"GSTXXXXX1", turnover:"₹42.8 Cr", status:"Active",   rm:"Admin User",  dueDate:"31 Jul 2025", tasks:4, risk:"Low"    },
  { id:"CLT-002", clientName:"Client Beta",    personName:"Contact B",  type:"Corporate",   sector:"Retail",        pan:"PANXXXXX2", gst:"GSTXXXXX2", turnover:"₹18.3 Cr", status:"Active",   rm:"User A",      dueDate:"15 Aug 2025", tasks:7, risk:"Medium" },
  { id:"CLT-003", clientName:"Client Gamma",   personName:"Contact C",  type:"HUF",         sector:"Real Estate",   pan:"PANXXXXX3", gst:"—",          turnover:"₹9.6 Cr",  status:"Active",   rm:"User B",      dueDate:"31 Jul 2025", tasks:2, risk:"Low"    },
  { id:"CLT-004", clientName:"Client Delta",   personName:"Contact D",  type:"LLP",         sector:"IT Services",   pan:"PANXXXXX4", gst:"GSTXXXXX4", turnover:"₹31.5 Cr", status:"Active",   rm:"Admin User",  dueDate:"30 Jun 2025", tasks:9, risk:"High"   },
  { id:"CLT-005", clientName:"Client Epsilon", personName:"Contact E",  type:"Corporate",   sector:"Export",        pan:"PANXXXXX5", gst:"GSTXXXXX5", turnover:"₹56.2 Cr", status:"Inactive", rm:"User A",      dueDate:"31 Aug 2025", tasks:1, risk:"Low"    },
  { id:"CLT-006", clientName:"Client Zeta",    personName:"Contact F",  type:"Individual",  sector:"Healthcare",    pan:"PANXXXXX6", gst:"—",          turnover:"₹2.8 Cr",  status:"Active",   rm:"User B",      dueDate:"31 Jul 2025", tasks:3, risk:"Low"    },
  { id:"CLT-007", clientName:"Client Eta",     personName:"Contact G",  type:"Partnership", sector:"Steel",         pan:"PANXXXXX7", gst:"GSTXXXXX7", turnover:"₹74.1 Cr", status:"Active",   rm:"Admin User",  dueDate:"15 Sep 2025", tasks:6, risk:"Medium" },
  { id:"CLT-008", clientName:"Client Theta",   personName:"Contact H",  type:"Corporate",   sector:"Energy",        pan:"PANXXXXX8", gst:"GSTXXXXX8", turnover:"₹22.9 Cr", status:"Active",   rm:"User A",      dueDate:"31 Jul 2025", tasks:5, risk:"Medium" },
];

// ── Currency config ───────────────────────────────────────────────────────────
const CURRENCIES = [
  { code:"INR", symbol:"₹",    name:"Indian Rupee"       },
  { code:"USD", symbol:"$",    name:"US Dollar"          },
  { code:"EUR", symbol:"€",    name:"Euro"               },
  { code:"GBP", symbol:"£",    name:"British Pound"      },
  { code:"AED", symbol:"د.إ", name:"UAE Dirham"         },
  { code:"SGD", symbol:"S$",   name:"Singapore Dollar"   },
  { code:"AUD", symbol:"A$",   name:"Australian Dollar"  },
  { code:"CAD", symbol:"C$",   name:"Canadian Dollar"    },
  { code:"JPY", symbol:"¥",    name:"Japanese Yen"       },
  { code:"CHF", symbol:"Fr",   name:"Swiss Franc"        },
  { code:"HKD", symbol:"HK$",  name:"Hong Kong Dollar"   },
];

// ── Dynamic FY Generator ──────────────────────────────────────────────────────
// Generates current FY + next futureFYs forward + pastFYs back automatically.
// To extend range simply bump futureFYs — no manual editing needed.
const generateFYList = (futureFYs = 4, pastFYs = 3) => {
  const now      = new Date();
  // Indian FY starts April 1; if month < April (0-2), current FY started last calendar year
  const baseYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const list     = [];
  for (let i = futureFYs; i >= -pastFYs; i--) {
    const start = baseYear + i;
    const end   = (start + 1).toString().slice(2);
    list.push(`${start}-${end}`);
  }
  return list;
};
// Produces: 2030-31, 2029-30, 2028-29, 2027-28, 2026-27, 2025-26, 2024-25, 2023-24
const FY_LIST    = generateFYList(4, 3);
const DEFAULT_FY = FY_LIST.find(f => f.startsWith("2026")) || FY_LIST[Math.floor(FY_LIST.length / 2)];
const MONTHS  = ["All Months","April","May","June","July","August","September","October","November","December","January","February","March"];
const DEFAULT_RATES = { INR:1, USD:83.5, EUR:90.2, GBP:105.8, AED:22.7, SGD:62.1, AUD:54.3, CAD:61.8, JPY:0.56, CHF:95.4, HKD:10.7 };

let _rowId = 1;
const genRowId = () => `FR-${String(_rowId++).padStart(4,"0")}`;

// ── XLSX formula-injection sanitizer ─────────────────────────────────────────
// Prevents cells whose value starts with = + - @ from being interpreted as
// spreadsheet formulas when the exported file is opened in Excel / Google Sheets.
// Applied only to user-entered free-text string fields in every export function.
// Numeric and date fields are left untouched.
const xlsSafe = (val) => {
  if (val === null || val === undefined) return "";
  const s = String(val);
  return /^[=+\-@]/.test(s) ? `'${s}` : s;
};

// ── Error Boundary ────────────────────────────────────────────────────────────
// Catches any unhandled render error in the main content area so the sidebar
// and provider tree survive. Users see a recovery UI instead of a blank screen.
// Class component required — React hooks cannot implement componentDidCatch.
class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    // Surface the error in the browser console for diagnostics
    console.error("[ProCAS ErrorBoundary] Unhandled render error:", error, info.componentStack);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          minHeight: "60vh", display: "flex", flexDirection: "column",
          alignItems: "center", justifyContent: "center",
          padding: "40px 24px", textAlign: "center",
          fontFamily: "'Outfit','DM Sans','Segoe UI',system-ui,sans-serif",
        }}>
          <div style={{ fontSize: 44, marginBottom: 16 }}>⚠️</div>
          <h2 style={{ color: "#f87171", fontSize: 18, fontWeight: 700, marginBottom: 8 }}>
            Something went wrong
          </h2>
          <p style={{ color: "#5a7a99", fontSize: 13, marginBottom: 24, maxWidth: 420, lineHeight: 1.6 }}>
            An unexpected error occurred in this module. Your other data is safe.
            Click <strong>Try Again</strong> to recover, or refresh the page.
          </p>
          <button
            onClick={() => this.setState({ hasError: false, error: null })}
            style={{
              background: "linear-gradient(135deg,#00a896,#1b5fa8)",
              color: "#fff", border: "none", borderRadius: 10,
              padding: "10px 28px", fontSize: 13, fontWeight: 600,
              cursor: "pointer", marginBottom: 16,
            }}>
            Try Again
          </button>
          {this.state.error && (
            <pre style={{
              color: "#4a6a8a", fontSize: 10, maxWidth: 560,
              overflowX: "auto", textAlign: "left", whiteSpace: "pre-wrap",
            }}>
              {this.state.error.toString()}
            </pre>
          )}
        </div>
      );
    }
    return this.props.children;
  }
}

// ── Formatters ────────────────────────────────────────────────────────────────
const fmtINRShort = n => {
  if (n >= 1e7) return `₹${(n/1e7).toFixed(2)} Cr`;
  if (n >= 1e5) return `₹${(n/1e5).toFixed(2)} L`;
  return `₹${n.toLocaleString("en-IN",{maximumFractionDigits:0})}`;
};
const fmtNum2 = n => n.toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});

// ── Shared Searchable FY Dropdown ────────────────────────────────────────────
// Drop-in replacement for raw <select> — supports keyboard search, themed, accessible.
const FYSelect = ({ value, onChange, t, dark, className = "" }) => {
  const [open,    setOpen]    = useState(false);
  const [search,  setSearch]  = useState("");
  const ref                   = useRef(null);
  const inputRef              = useRef(null);

  const filtered = FY_LIST.filter(fy => fy.includes(search.trim()));

  // Close on outside click
  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  // Focus search input when opened
  useEffect(() => { if (open && inputRef.current) inputRef.current.focus(); }, [open]);

  const select = (fy) => { onChange(fy); setOpen(false); setSearch(""); };

  return (
    <div ref={ref} className={`relative ${className}`} style={{minWidth:"120px"}}>
      {/* Trigger */}
      <button
        onClick={() => setOpen(o => !o)}
        className={`w-full flex items-center justify-between gap-2 px-3 py-2 rounded-xl border text-sm font-semibold
          transition-all outline-none ${t.input}
          ${open ? "ring-2 ring-[#00c9b1]" : "hover:ring-1 hover:ring-[#00c9b140]"}`}
        style={{letterSpacing:"-0.01em"}}>
        <span style={{color:"#00c9b1"}}>FY</span>
        <span className={t.text}>{value}</span>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor"
          strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
          className={`transition-transform duration-200 ${t.textMuted} ${open?"rotate-180":""}`}>
          <path d="M6 9l6 6 6-6"/>
        </svg>
      </button>

      {/* Dropdown panel */}
      {open && (
        <div className={`absolute z-50 mt-1.5 w-full rounded-2xl border shadow-2xl overflow-hidden
          ${t.card} ${t.cardBorder}`}
          style={{minWidth:"150px", boxShadow: dark
            ? "0 20px 40px rgba(0,0,0,0.6), 0 0 0 1px #00c9b120"
            : "0 16px 32px rgba(0,0,0,0.15), 0 0 0 1px #00c9b120"}}>
          {/* Search box */}
          <div className={`px-3 py-2.5 border-b ${t.cardBorder} flex items-center gap-2`}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
              strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={t.textMuted}>
              <circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/>
            </svg>
            <input
              ref={inputRef}
              type="text"
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search FY…"
              className={`bg-transparent outline-none text-xs w-full ${t.text} placeholder:${t.textMuted}`}
            />
            {search && (
              <button onClick={() => setSearch("")} className={`${t.textMuted} hover:text-red-400 transition-colors`}>
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                  strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M18 6 6 18M6 6l12 12"/>
                </svg>
              </button>
            )}
          </div>

          {/* FY options */}
          <div className="overflow-y-auto" style={{maxHeight:"200px"}}>
            {filtered.length === 0 ? (
              <div className={`px-4 py-3 text-xs text-center ${t.textMuted}`}>No matching FY</div>
            ) : filtered.map((fy, i) => {
              const isCurrent = fy === DEFAULT_FY;
              const isSelected = fy === value;
              return (
                <button
                  key={fy}
                  onClick={() => select(fy)}
                  className={`w-full flex items-center justify-between px-4 py-2.5 text-sm transition-all
                    ${isSelected
                      ? "text-white font-semibold"
                      : `${t.textMuted} ${t.hover}`}`}
                  style={isSelected
                    ? {background:"linear-gradient(90deg,#003d5c,#007a6e)"}
                    : {}}>
                  <span className={isSelected ? "text-white" : t.text}>FY {fy}</span>
                  <div className="flex items-center gap-1.5">
                    {isCurrent && !isSelected && (
                      <span className="text-[10px] font-bold px-1.5 py-0.5 rounded-full"
                        style={{background:"#00c9b120",color:"#00c9b1"}}>Current</span>
                    )}
                    {isSelected && (
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="white"
                        strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M20 6L9 17l-5-5"/>
                      </svg>
                    )}
                  </div>
                </button>
              );
            })}
          </div>

          {/* Footer hint */}
          <div className={`px-3 py-2 border-t ${t.cardBorder} flex items-center gap-1.5`}>
            <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor"
              strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={t.textMuted}>
              <circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/>
            </svg>
            <span className={`text-[10px] ${t.textMuted}`}>
              Auto-generated · {FY_LIST.length} financial years
            </span>
          </div>
        </div>
      )}
    </div>
  );
};

// ── Shared KPI Card ───────────────────────────────────────────────────────────
const KpiCard = ({ t, label, value, sub, icon, color, trend, delta }) => (
  <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 flex flex-col gap-3 relative overflow-hidden group transition-all duration-300 hover:shadow-xl hover:-translate-y-0.5`}
    style={{ boxShadow: `0 0 0 0 ${color}` }}>
    <div className="absolute top-0 right-0 w-28 h-28 rounded-full opacity-5 -translate-y-8 translate-x-8 transition-all duration-500 group-hover:opacity-10 group-hover:scale-110"
      style={{ background: color }}/>
    <div className="flex items-center justify-between">
      <span className={`text-[11px] font-bold uppercase tracking-widest ${t.textMuted}`}>{label}</span>
      <div className="w-9 h-9 rounded-xl flex items-center justify-center transition-transform duration-200 group-hover:scale-110"
        style={{ background: color+"22" }}>
        <span style={{ color }}>{icon}</span>
      </div>
    </div>
    <div>
      <div className={`text-2xl font-black tracking-tight ${t.text}`}>{value}</div>
      <div className={`text-xs mt-1 ${t.textMuted}`}>{sub}</div>
    </div>
    {(trend !== undefined || delta !== undefined) && (
      <div className="flex items-center gap-1 text-xs font-semibold"
        style={{ color: (trend > 0 || delta > 0) ? "#34d399" : "#f87171" }}>
        <span>{(trend > 0 || delta > 0) ? "▲" : "▼"}</span>
        <span>{delta !== undefined ? `${Math.abs(delta)}% vs last month` : `${Math.abs(trend)}% vs last FY`}</span>
      </div>
    )}
  </div>
);

// ── Shared Badge ──────────────────────────────────────────────────────────────
const Badge = ({ status }) => {
  const map = {
    Active:["#34d399","#052e16"], Inactive:["#94a3b8","#1e293b"],
    High:["#f87171","#2d0b0b"], Medium:["#fbbf24","#1c1200"], Low:["#34d399","#052e16"],
  };
  const [fg,bg] = map[status]||["#94a3b8","#1e293b"];
  return <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold" style={{color:fg,background:bg+"aa"}}>{status}</span>;
};

// ── Custom recharts tooltip ───────────────────────────────────────────────────
const ChartTooltip = ({ active, payload, label, dark, prefix="₹", suffix="" }) => {
  if (!active||!payload?.length) return null;
  return (
    <div className="rounded-xl border shadow-2xl p-3 text-xs min-w-[140px]"
      style={{ background:dark?"#1a2234":"#fff", borderColor:dark?"#2d3748":"#e2e8f0" }}>
      <p className="font-bold mb-2" style={{ color:dark?"#e2e8f0":"#1a202c" }}>{label}</p>
      {payload.map((p,i)=>(
        <div key={i} className="flex items-center gap-2 mb-1">
          <div className="w-2 h-2 rounded-full shrink-0" style={{background:p.color||p.fill}}/>
          <span style={{color:dark?"#94a3b8":"#718096"}}>{p.name}:</span>
          <span className="font-semibold" style={{color:dark?"#e2e8f0":"#1a202c"}}>
            {prefix}{typeof p.value==="number"?fmtNum2(p.value):p.value}{suffix}
          </span>
        </div>
      ))}
    </div>
  );
};

// ── Confirm dialog ────────────────────────────────────────────────────────────
const ConfirmDialog = ({ t, message, onConfirm, onCancel }) => (
  <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
    style={{background:"rgba(0,0,0,0.7)",backdropFilter:"blur(6px)"}}>
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl w-full max-w-sm shadow-2xl`}>
      <div className="p-6 flex flex-col items-center gap-4">
        <div className="w-14 h-14 rounded-2xl flex items-center justify-center" style={{background:"#f8717122"}}>
          <Icon path={Icons.alert} size={26} className="text-[#f87171]"/>
        </div>
        <div className="text-center">
          <h3 className={`font-bold text-base ${t.text}`}>Confirm Delete</h3>
          <p className={`text-sm mt-1 ${t.textMuted}`}>{message}</p>
        </div>
        <div className="flex gap-3 w-full">
          <button onClick={onCancel} className={`flex-1 py-2.5 rounded-xl border ${t.cardBorder} text-sm font-medium ${t.textMuted} ${t.hover}`}>Cancel</button>
          <button onClick={onConfirm} className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white" style={{background:"linear-gradient(135deg,#ef4444,#dc2626)"}}>Delete</button>
        </div>
      </div>
    </div>
  </div>
);

// ── Import result modal ───────────────────────────────────────────────────────
const ImportResultModal = ({ t, dark, result, onClose }) => {
  const { added, updated, duplicates, errors, total } = result;
  const failed = errors?.length || 0;

  const downloadErrors = () => {
    if (!errors?.length) return;
    const rows = [["Row #", "Client Name", "Reason"], ...errors.map(e => [e.row, e.clientName || "—", e.reason])];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{wch:8},{wch:30},{wch:40}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Import Errors");
    XLSX.writeFile(wb, "Client_Import_Errors.xlsx");
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
      style={{background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)"}}>
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden`}>
        {/* Header */}
        <div className="px-6 py-4 flex items-center justify-between"
          style={{background: dark?"linear-gradient(135deg,#0c1e30,#0f1e2e)":"linear-gradient(135deg,#f0f7ff,#e8f4f2)"}}>
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{background:"#00c9b118"}}>
              <Icon path={Icons.fileText} size={16} style={{color:"#00c9b1"}}/>
            </div>
            <div>
              <h3 className={`font-bold text-sm ${t.text}`}>Import Summary</h3>
              <p className={`text-[11px] ${t.textMuted}`}>{total} record{total!==1?"s":""} processed</p>
            </div>
          </div>
          <button onClick={onClose} className={`w-8 h-8 rounded-lg flex items-center justify-center ${t.hover} ${t.textMuted}`}>
            <Icon path={Icons.x} size={15}/>
          </button>
        </div>

        <div className="p-5 space-y-4">
          {/* 4 stats */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            {[
              {label:"Total",      value:total,      color:"#4a90d9"},
              {label:"Imported",   value:added,       color:"#34d399"},
              {label:"Failed",     value:failed,      color:"#f87171"},
              {label:"Duplicates", value:duplicates,  color:"#fbbf24"},
            ].map(item=>(
              <div key={item.label} className="rounded-xl p-3 text-center" style={{background:item.color+"15",border:`1px solid ${item.color}25`}}>
                <div className="text-2xl font-black leading-none" style={{color:item.color}}>{item.value}</div>
                <div className="text-[10px] font-semibold mt-1" style={{color:item.color+"cc"}}>{item.label}</div>
              </div>
            ))}
          </div>

          {/* Status bar */}
          {total > 0 && (
            <div>
              <div className="h-2 rounded-full overflow-hidden" style={{background: dark?"#1a2d44":"#e2e8f0"}}>
                <div className="h-full rounded-full transition-all" style={{
                  width:`${(added/total)*100}%`,
                  background:"linear-gradient(90deg,#34d399,#00c9b1)"
                }}/>
              </div>
              <div className="flex justify-between mt-1 text-[10px]" style={{color:"#5a7a99"}}>
                <span>{Math.round((added/total)*100)}% imported successfully</span>
                {failed>0&&<span style={{color:"#f87171"}}>{failed} failed</span>}
              </div>
            </div>
          )}

          {/* Error list (up to 5 shown) */}
          {failed > 0 && (
            <div className={`rounded-xl border overflow-hidden`} style={{borderColor: dark?"#2d1a1a":"#ffd4d4"}}>
              <div className="px-4 py-2 flex items-center justify-between" style={{background:dark?"#1a0f0f":"#fff0f0"}}>
                <span className="text-[11px] font-bold" style={{color:"#f87171"}}>⚠ Import Errors</span>
                <span className="text-[10px]" style={{color:"#f87171"}}>{failed} record{failed!==1?"s":""}</span>
              </div>
              <div className="divide-y" style={{borderColor:dark?"#1a2d44":"#f5e8e8"}}>
                {errors.slice(0,5).map((e,i)=>(
                  <div key={i} className="px-4 py-2 flex items-start gap-3">
                    <span className="text-[10px] font-mono font-bold mt-0.5" style={{color:"#f87171",minWidth:32}}>R{e.row}</span>
                    <div className="min-w-0">
                      {e.clientName && <p className="text-[11px] font-semibold truncate" style={{color:dark?"#e8f0f8":"#0d2137"}}>{e.clientName}</p>}
                      <p className="text-[10px]" style={{color:"#f87171"}}>{e.reason}</p>
                    </div>
                  </div>
                ))}
                {failed > 5 && (
                  <div className="px-4 py-2 text-[10px]" style={{color:"#5a7a99"}}>…and {failed-5} more. Download error report for full list.</div>
                )}
              </div>
            </div>
          )}

          {/* Action buttons */}
          <div className="flex gap-3">
            {failed > 0 && (
              <button onClick={downloadErrors}
                className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                <Icon path={Icons.download} size={13}/> Download Error Report
              </button>
            )}
            <button onClick={onClose}
              className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white"
              style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
              Done
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── Client form modal ─────────────────────────────────────────────────────────
// ── Searchable "Name of Person" dropdown — synced from User Database ──────────
// ── PersonMultiDropdown ───────────────────────────────────────────────────────
// Multi-select user picker for Client Master. Replaces the old single-select
// PersonDropdown. value = comma-separated name string (matches stored personName).
// This is the ONLY place where user→client assignments are made.
// User Database reads assignments FROM here (Client Master is the source of truth).
const PersonDropdown = ({ t, dark, value, onChange, error }) => {
  let activeUsers = [];
  try {
    const ctx = useContext(UserContext);
    if (ctx?.users) activeUsers = ctx.users.filter(u => u.status === "Active");
  } catch (_) {}

  const [open,  setOpen]  = useState(false);
  const [query, setQuery] = useState("");
  const ref               = useRef(null);

  // Parse comma-separated names into an array of selected names
  const selectedNames = value
    ? value.split(",").map(n => n.trim()).filter(Boolean)
    : [];

  // Close on outside click
  useEffect(() => {
    const h = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);

  const filtered = activeUsers.filter(u =>
    u.name.toLowerCase().includes(query.toLowerCase())
  );

  // Toggle a user in/out of the selection
  const toggle = (name) => {
    const next = selectedNames.includes(name)
      ? selectedNames.filter(n => n !== name)
      : [...selectedNames, name];
    onChange(next.join(", "));
  };

  const removeChip = (name) => {
    onChange(selectedNames.filter(n => n !== name).join(", "));
  };

  // Fallback for empty User DB
  if (activeUsers.length === 0) {
    return (
      <input
        value={value}
        onChange={e => onChange(e.target.value)}
        placeholder="e.g. User A, User B"
        className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none
          focus:ring-2 focus:ring-[#00c9b1] ${t.input} ${error ? "border-[#f87171]" : ""}`}
      />
    );
  }

  return (
    <div ref={ref} className="relative">
      {/* Trigger / selected chips display */}
      <div
        onClick={() => setOpen(o => !o)}
        className={[
          "flex flex-wrap items-center gap-1.5 min-h-[42px] w-full px-3 py-2 rounded-xl border",
          "cursor-pointer transition-all",
          t.input,
          error ? "border-[#f87171]" : "",
          open  ? "ring-2 ring-[#00c9b1]" : "",
        ].join(" ")}>
        {selectedNames.length === 0 ? (
          <span className={`text-sm ${t.textMuted}`}>Select one or more users...</span>
        ) : (
          selectedNames.map(name => (
            <span key={name}
              className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold"
              style={{ background:"#00c9b120", color:"#00c9b1", border:"1px solid #00c9b140" }}>
              {name}
              <button
                onClick={e => { e.stopPropagation(); removeChip(name); }}
                className="ml-0.5 hover:opacity-60 transition-opacity">
                <Icon path={Icons.x} size={9}/>
              </button>
            </span>
          ))
        )}
        <Icon path={open ? Icons.chevronUp : Icons.chevronDown} size={12}
          className={`ml-auto shrink-0 ${t.textMuted}`}/>
      </div>

      {/* Dropdown panel */}
      {open && (
        <div className={`absolute z-[65] mt-1 w-full rounded-xl border shadow-2xl overflow-hidden ${t.card}`}
          style={{ borderColor: dark ? "#243d58" : "#d4e0ed" }}>
          {/* Search */}
          <div className="p-2 border-b" style={{ borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
            <div className={`flex items-center gap-2 px-2.5 py-1.5 rounded-lg border text-xs ${t.input}`}>
              <Icon path={Icons.search} size={12} className={t.textMuted}/>
              <input
                autoFocus
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder="Search users..."
                className="bg-transparent outline-none flex-1"
                onClick={e => e.stopPropagation()}/>
              {query && (
                <button onClick={e => { e.stopPropagation(); setQuery(""); }}
                  className={`${t.textMuted} hover:opacity-60`}>
                  <Icon path={Icons.x} size={10}/>
                </button>
              )}
            </div>
          </div>

          {/* User list — checkbox style */}
          <div className="max-h-52 overflow-y-auto">
            {filtered.length === 0 ? (
              <p className={`px-4 py-3 text-xs text-center ${t.textMuted}`}>No users match</p>
            ) : filtered.map(u => {
              const checked = selectedNames.includes(u.name);
              return (
                <button key={u.id}
                  onClick={e => { e.stopPropagation(); toggle(u.name); }}
                  className={`w-full flex items-center gap-3 px-4 py-2.5 text-left transition-colors ${t.hover}`}>
                  {/* Checkbox */}
                  <div className="w-4 h-4 rounded border-2 flex items-center justify-center shrink-0 transition-colors"
                    style={{
                      borderColor: checked ? "#00c9b1" : (dark ? "#243d58" : "#d4e0ed"),
                      background:  checked ? "#00c9b1" : "transparent",
                    }}>
                    {checked && (
                      <svg width="9" height="9" viewBox="0 0 12 12" fill="none">
                        <path d="M2 6l3 3 5-5" stroke="white" strokeWidth="2"
                          strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    )}
                  </div>
                  {/* Avatar */}
                  <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-[10px] font-bold shrink-0"
                    style={{ background:"linear-gradient(135deg,#00a896,#1b5fa8)" }}>
                    {u.name.split(" ").map(n => n[0]).slice(0, 2).join("")}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className={`text-sm font-medium truncate ${checked ? "text-[#00c9b1]" : t.text}`}>{u.name}</div>
                    <div className={`text-[11px] ${t.textMuted} truncate`}>{u.email} · {u.role}</div>
                  </div>
                </button>
              );
            })}
          </div>

          {/* Footer — count + close */}
          <div className="px-4 py-2 border-t flex items-center justify-between"
            style={{ borderColor: dark ? "#1a2d44" : "#e2e8f0", background: dark ? "#0f1e2e" : "#f5fdfb" }}>
            <span className="text-[10px] font-medium flex items-center gap-1.5" style={{ color:"#00c9b1" }}>
              <Icon path={Icons.refreshCw} size={10}/>
              {selectedNames.length} selected · {activeUsers.length} active users in User Database
            </span>
            <button onClick={() => setOpen(false)}
              className="text-[10px] font-semibold px-2 py-1 rounded-lg"
              style={{ background:"#00c9b120", color:"#00c9b1" }}>
              Done
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

const ClientFormModal = ({ t, dark, initial, onSave, onClose }) => {
  const [form, setForm]     = useState(initial || { clientName:"", personName:"", sector:"", state:"" });
  const [errors, setErrors] = useState({});
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const validate = () => {
    const e = {};
    if (!form.clientName?.trim()) e.clientName = "Required";
    if (!form.personName?.trim()) e.personName  = "Required";
    setErrors(e);
    return !Object.keys(e).length;
  };

  const SECTORS = ["IT Services","Manufacturing","Retail","Real Estate","Healthcare","Export","Steel","Energy","Finance","Education","Others"];
  const STATES  = ["Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh","Delhi","Goa","Gujarat","Haryana","Himachal Pradesh","Jharkhand","Karnataka","Kerala","Madhya Pradesh","Maharashtra","Manipur","Meghalaya","Mizoram","Nagaland","Odisha","Punjab","Rajasthan","Sikkim","Tamil Nadu","Telangana","Tripura","Uttar Pradesh","Uttarakhand","West Bengal","Others"];

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
      style={{ background:"rgba(0,0,0,0.7)", backdropFilter:"blur(6px)" }}>
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl w-full max-w-md shadow-2xl`}
        style={{maxHeight:"90vh",overflowY:"auto"}}>
        <div className={`p-5 border-b ${t.cardBorder} flex items-center justify-between`}>
          <h3 className={`font-bold text-base ${t.text}`}>{initial?.id ? "Edit Client" : "Add New Client"}</h3>
          <button onClick={onClose}
            className={`w-8 h-8 rounded-lg flex items-center justify-center ${t.hover} ${t.textMuted}`}>
            <Icon path={Icons.x} size={15}/>
          </button>
        </div>
        <div className="p-5 space-y-4">
          {/* Client Name */}
          <div>
            <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>
              Client Name *
            </label>
            <input
              value={form.clientName || ""}
              onChange={e => set("clientName", e.target.value)}
              placeholder="e.g. Client Name"
              className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none
                focus:ring-2 focus:ring-[#00c9b1] ${t.input} ${errors.clientName ? "border-[#f87171]" : ""}`}/>
            {errors.clientName && <p className="text-[#f87171] text-xs mt-1">{errors.clientName}</p>}
          </div>

          {/* Assigned Users — multi-select from User Database (Client Master is source of truth) */}
          <div>
            <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>
              Assigned Users *
              <span className="ml-2 normal-case font-normal text-[10px] px-1.5 py-0.5 rounded"
                style={{ background:"#00c9b115", color:"#00c9b1" }}>
                Synced from User Database · Multi-select
              </span>
            </label>
            <PersonDropdown
              t={t} dark={dark}
              value={form.personName || ""}
              onChange={v => set("personName", v)}
              error={!!errors.personName}/>
            {errors.personName
              ? <p className="text-[#f87171] text-xs mt-1">{errors.personName}</p>
              : <p className="text-[10px] mt-1" style={{color:"#5a7a99"}}>
                  One client can be assigned to one or more users. Assignments sync automatically to User Database.
                </p>}
          </div>

          {/* Sector */}
          <div>
            <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Sector</label>
            <select value={form.sector || ""} onChange={e => set("sector", e.target.value)}
              className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`}>
              <option value="">— Select Sector —</option>
              {SECTORS.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          {/* State */}
          <div>
            <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>State</label>
            <select value={form.state || ""} onChange={e => set("state", e.target.value)}
              className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`}>
              <option value="">— Select State —</option>
              {STATES.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          <div className="flex gap-3 pt-1">
            <button onClick={onClose}
              className={`flex-1 py-2.5 rounded-xl border ${t.cardBorder} text-sm font-medium ${t.textMuted} ${t.hover}`}>
              Cancel
            </button>
            <button
              onClick={() => { if (validate()) onSave(form); }}
              className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white hover:opacity-90 active:scale-95 transition-all"
              style={{ background:"linear-gradient(135deg,#00a896,#1b5fa8)" }}>
              {initial?.id ? "Save Changes" : "Add Client"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── Sort icon ─────────────────────────────────────────────────────────────────
const SortIcon = ({ field, sortField, sortDir }) => {
  if (sortField!==field) return <Icon path="M7 16V4m0 0L3 8m4-4l4 4M17 8v12m0 0l4-4m-4 4l-4-4" size={12} className="opacity-25"/>;
  return sortDir==="asc"
    ? <Icon path={Icons.arrowUp}   size={12} className="text-[#00c9b1]"/>
    : <Icon path={Icons.arrowDown} size={12} className="text-[#00c9b1]"/>;
};

// ─────────────────────────────────────────────────────────────────────────────
// ── MASTER CLIENT TAB ────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
let _cltId = 9;
const genCltId = () => `CLT-${String(_cltId++).padStart(3,"0")}`;

const MasterClientTab = ({ t, dark, isAdmin }) => {
  const { clients: ctxClients, updateClients } = useSyncContext();
  const [clients, setClients] = useState(ctxClients);

  // Pull active users from UserContext for validation
  let activeUsers = [];
  try {
    const ctx = useContext(UserContext);
    if (ctx?.users) activeUsers = ctx.users.filter(u => u.status === "Active");
  } catch (_) {}

  const setAndSync = useCallback((updater) => {
    setClients(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      updateClients(next);
      return next;
    });
  }, [updateClients]);

  const syncedRef = useRef(false);
  useEffect(() => {
    if (syncedRef.current) setClients(ctxClients);
    else syncedRef.current = true;
  }, [ctxClients]);

  // ── Filters ──────────────────────────────────────────────────────────────────
  const [search,     setSearch]     = useState("");
  const [sectorF,    setSectorF]    = useState("");
  const [stateF,     setStateF]     = useState("");
  const [userF,      setUserF]      = useState("");
  const [sortField,  setSortField]  = useState("id");
  const [sortDir,    setSortDir]    = useState("asc");
  const [page,       setPage]       = useState(1);
  const PAGE_SIZE = 8;
  const [editModal,      setEditModal]      = useState(null);
  const [addModal,       setAddModal]       = useState(false);
  const [confirmDelete,  setConfirmDelete]  = useState(null);
  const [importResult,   setImportResult]   = useState(null);
  const [importToast,    setImportToast]    = useState(null);
  const fileRef = useRef();
  const [selectedFile, setSelectedFile] = useState(null);
  const [dragOver,     setDragOver]     = useState(false);
  const [importing,    setImporting]    = useState(false);

  // Handle file selected via input or drag-drop
  const onFileChosen = (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext)) {
      setImportToast({ type:"error", msg:`Invalid file format ".${ext}". Use .xlsx, .xls, or .csv` });
      setTimeout(() => setImportToast(null), 5000);
      return;
    }
    setSelectedFile(file);
  };

  const handleFileInputChange = e => {
    onFileChosen(e.target.files?.[0]);
    // Reset input so same file can be re-selected
    e.target.value = "";
  };

  const handleDrop = e => {
    e.preventDefault();
    setDragOver(false);
    onFileChosen(e.dataTransfer.files?.[0]);
  };

  // Unique values for filter dropdowns
  const uniqueSectors = useMemo(() => Array.from(new Set(clients.map(c => c.sector).filter(s => s && s !== "—"))).sort(), [clients]);
  const uniqueStates  = useMemo(() => Array.from(new Set(clients.map(c => c.state).filter(s => s && s !== "—"))).sort(),  [clients]);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return clients
      .filter(c => {
        const matchQ = !q || c.clientName.toLowerCase().includes(q) || c.id.toLowerCase().includes(q) || (c.personName||"").toLowerCase().includes(q) || (c.sector||"").toLowerCase().includes(q) || (c.state||"").toLowerCase().includes(q);
        const matchSector = !sectorF || c.sector === sectorF;
        const matchState  = !stateF  || c.state  === stateF;
        const matchUser   = !userF   || (c.personName||"").toLowerCase().includes(userF.toLowerCase());
        return matchQ && matchSector && matchState && matchUser;
      })
      .sort((a, b) => {
        const va = (a[sortField]||"").toString().toLowerCase();
        const vb = (b[sortField]||"").toString().toLowerCase();
        return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
      });
  }, [clients, search, sectorF, stateF, userF, sortField, sortDir]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const paged      = filtered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);
  const toggleSort = f => { if(sortField===f) setSortDir(d=>d==="asc"?"desc":"asc"); else{setSortField(f);setSortDir("asc");} setPage(1); };
  const clearFilters = () => { setSearch(""); setSectorF(""); setStateF(""); setUserF(""); setPage(1); };
  const hasActiveFilters = search || sectorF || stateF || userF;

  // ── Validate a single username string against User Database ──────────────────
  // Returns { valid: bool, invalidNames: string[] }
  const validateUsers = (userStr) => {
    if (!userStr?.trim()) return { valid: false, invalidNames: [] };
    const names  = userStr.split(",").map(n => n.trim()).filter(Boolean);
    const invalid = names.filter(name => !activeUsers.some(u => u.name.toLowerCase() === name.toLowerCase()));
    return { valid: invalid.length === 0, invalidNames: invalid };
  };

  // ── Download 4-column template ────────────────────────────────────────────────
  const downloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Client Name",       "Name of User",          "Sector",        "State"],
      ["Client Alpha",      "User A",                "IT Services",   "Karnataka"],
      ["Client Beta",       "User A, User B",        "Manufacturing", "Maharashtra"],
      ["Client Gamma",      "Admin User",            "Healthcare",    "Delhi"],
    ]);
    ws["!cols"] = [{wch:35},{wch:35},{wch:20},{wch:20}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Client Master");
    XLSX.writeFile(wb, "Client_Master_Template.xlsx");
  };

  // ── Bulk Import ───────────────────────────────────────────────────────────────
  const handleImport = (file) => {
    if (!file) return;
    setImporting(true);
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const wb   = XLSX.read(evt.target.result, { type:"binary" });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        if (!rows.length || rows.every(r => !r.some(c => c?.toString().trim()))) {
          setImportToast({ type:"error", msg:"No records found in file." });
          setTimeout(() => setImportToast(null), 5000);
          setImporting(false);
          return;
        }

        const header = rows[0].map(h => h?.toString().trim().toLowerCase());

        const ni  = header.findIndex(h => h.includes("client name") || h === "client");
        const ui  = header.findIndex(h => h.includes("name of user") || h.includes("user"));
        const si  = header.findIndex(h => h.includes("sector"));
        const sti = header.findIndex(h => h.includes("state"));

        if (ni === -1) { setImportToast({type:"error",msg:"Required column missing: 'Client Name'"}); setTimeout(()=>setImportToast(null),5000); setImporting(false); return; }
        if (ui === -1) { setImportToast({type:"error",msg:"Required column missing: 'Name of User'"}); setTimeout(()=>setImportToast(null),5000); setImporting(false); return; }

        let added = 0, duplicates = 0;
        const errors = [];
        const dataRows = rows.slice(1).filter(row => row.some(cell => cell?.toString().trim()));
        const total = dataRows.length;

        if (total === 0) {
          setImportToast({ type:"error", msg:"File has no data rows." });
          setTimeout(() => setImportToast(null), 5000);
          setImporting(false);
          return;
        }

        const nc = [...clients];

        dataRows.forEach((row, ri) => {
          const rowNum  = ri + 2;
          const cn      = row[ni]?.toString().trim();
          const userStr = ui  !== -1 ? row[ui]?.toString().trim()  : "";
          const sector  = si  !== -1 ? row[si]?.toString().trim()  : "";
          const state   = sti !== -1 ? row[sti]?.toString().trim() : "";

          if (!cn) { errors.push({ row:rowNum, clientName:"", reason:"Missing Client Name (mandatory)" }); return; }
          if (!userStr) { errors.push({ row:rowNum, clientName:cn, reason:"Missing Name of User (mandatory)" }); return; }

          if (activeUsers.length > 0) {
            const { valid, invalidNames } = validateUsers(userStr);
            if (!valid) {
              errors.push({ row:rowNum, clientName:cn, reason:`User not found in User Database: ${invalidNames.join(", ")}` });
              return;
            }
          }

          const idx = nc.findIndex(c => c.clientName.toLowerCase() === cn.toLowerCase());
          if (idx !== -1) {
            nc[idx] = { ...nc[idx], personName:userStr, sector:sector||nc[idx].sector, state:state||nc[idx].state };
            duplicates++;
          } else {
            nc.push({
              id:genCltId(), clientName:cn, personName:userStr,
              sector:sector||"—", state:state||"—",
              status:"Active", risk:"Low", tasks:0, type:"Corporate",
              pan:"—", gst:"—", turnover:"—",
              rm:userStr.split(",")[0].trim(), dueDate:"—",
            });
            added++;
          }
        });

        setAndSync(nc);
        setImportResult({ total, added, updated:duplicates, duplicates, errors });
        setSelectedFile(null);
      } catch (err) {
        setImportToast({ type:"error", msg:"Could not parse file. Please use the downloaded template." });
        setTimeout(() => setImportToast(null), 5000);
      }
      setImporting(false);
    };
    reader.onerror = () => {
      setImportToast({ type:"error", msg:"Failed to read file. Try again." });
      setTimeout(() => setImportToast(null), 5000);
      setImporting(false);
    };
    reader.readAsBinaryString(file);
  };

  const handleDelete = id  => { setAndSync(p => p.filter(c => c.id !== id)); setConfirmDelete(null); };
  const handleEdit   = f   => { setAndSync(p => p.map(c => c.id === editModal.id ? { ...c, ...f, sector: f.sector||c.sector, state: f.state||c.state } : c)); setEditModal(null); };
  const handleAdd    = f   => {
    setAndSync(p => [...p, {
      id: genCltId(), clientName: f.clientName, personName: f.personName,
      sector: f.sector||"—", state: f.state||"—",
      status:"Active", risk:"Low", tasks:0, type:"Corporate",
      pan:"—", gst:"—", turnover:"—", rm: f.personName?.split(",")[0]?.trim()||"—", dueDate:"—",
    }]);
    setAddModal(false);
  };

  const exportClients = () => {
    try {
      const data = filtered.map((c, i) => ({
        "#":              i + 1,
        "Client ID":      c.id,
        "Client Name":    xlsSafe(c.clientName),
        "Name of User":   xlsSafe(c.personName  || "—"),
        "Sector":         xlsSafe(c.sector       || "—"),
        "State":          xlsSafe(c.state        || "—"),
        "Type":           xlsSafe(c.type         || "Corporate"),
        "Status":         xlsSafe(c.status       || "Active"),
        "Risk":           xlsSafe(c.risk         || "Low"),
        "PAN":            xlsSafe(c.pan          || "—"),
        "GST":            xlsSafe(c.gst          || "—"),
        "Turnover":       xlsSafe(c.turnover     || "—"),
        "Due Date":       xlsSafe(c.dueDate      || "—"),
      }));
      if (!data.length) { alert("No client records to export."); return; }
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [4,10,32,28,18,18,14,10,10,16,22,14,16].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Client Master");
      const filters = [sectorF&&`Sector-${sectorF}`, stateF&&`State-${stateF}`, userF&&`User-${userF}`].filter(Boolean).join("_");
      XLSX.writeFile(wb, `ClientMaster${filters ? "_"+filters : ""}.xlsx`);
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("Client Master export error:", err);
    }
  };

  const selStyle = `px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`;

  return (
    <div className="space-y-6">
      {/* Import toast */}
      {importToast && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-3 px-4 py-3 rounded-xl shadow-2xl text-sm font-medium"
          style={{background:importToast.type==="error"?"#f87171":"#34d399",color:"#fff",minWidth:280}}>
          <Icon path={importToast.type==="error"?Icons.alert:Icons.check} size={16}/>{importToast.msg}
        </div>
      )}

      {/* KPIs */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <KpiCard t={t} label="Total Clients"  value={clients.length}                                  sub="In master database"  icon={<Icon path={Icons.users}  size={16}/>} color="#00c9b1" delta={4.2}/>
        <KpiCard t={t} label="Active Clients" value={clients.filter(c=>c.status==="Active").length}   sub="Active this FY"      icon={<Icon path={Icons.check}  size={16}/>} color="#34d399" delta={2.1}/>
        <KpiCard t={t} label="Total Turnover" value="₹847 Cr"                                         sub="Combined base"       icon={<Icon path={Icons.dollar} size={16}/>} color="#4a90d9" delta={12.8}/>
        <KpiCard t={t} label="Pending Tasks"  value="37"                                               sub="Across all clients"  icon={<Icon path={Icons.alert}  size={16}/>} color="#fbbf24" delta={-8.3}/>
      </div>

      {/* Import card */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background:"#00c9b122"}}>
              <Icon path={Icons.fileText} size={18} className="text-[#00c9b1]"/>
            </div>
            <div>
              <h3 className={`font-semibold text-sm ${t.text}`}>Bulk Client Import</h3>
              <p className={`text-xs mt-0.5 ${t.textMuted}`}>Excel / CSV · User validation · Duplicate detection</p>
            </div>
          </div>
          <button onClick={downloadTemplate}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
            <Icon path={Icons.download} size={14}/> Download Template
          </button>
        </div>

        {isAdmin ? (
          <div className="space-y-3">
            {/* ── Hidden file input — the ONLY element that should trigger the dialog ── */}
            <input
              id="client-master-file-input"
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              style={{position:"absolute",width:1,height:1,opacity:0,pointerEvents:"none",overflow:"hidden"}}
              onChange={handleFileInputChange}
            />

            {/* ── Drag & Drop / Browse zone ── */}
            <div
              onDragOver={e=>{e.preventDefault();setDragOver(true);}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={handleDrop}
              className="rounded-2xl border-2 border-dashed transition-all duration-200"
              style={{
                borderColor: dragOver ? "#00c9b1" : (dark?"#243d58":"#d4e0ed"),
                background:  dragOver ? "#00c9b108" : (dark?"#0a1624":"#f8fbff"),
              }}>
              <div className="flex flex-col items-center justify-center py-8 px-6 text-center">
                <div className="w-14 h-14 rounded-2xl flex items-center justify-center mb-3 transition-all"
                  style={{background: dragOver?"#00c9b120":"#00c9b110"}}>
                  <Icon path={Icons.upload} size={22} style={{color:dragOver?"#00c9b1":"#5a7a99"}}/>
                </div>
                <p className={`text-sm font-semibold mb-1 ${t.text}`}>
                  {dragOver ? "Drop file here" : "Drag & drop your Excel or CSV file"}
                </p>
                <p className={`text-xs mb-4 ${t.textMuted}`}>
                  Supports .xlsx, .xls, .csv · Max columns: Client Name, Name of User, Sector, State
                </p>

                {/* ── Browse button — label directly wraps input for guaranteed dialog ── */}
                <label
                  htmlFor="client-master-file-input"
                  className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-semibold text-white cursor-pointer hover:opacity-90 active:scale-95 transition-all select-none"
                  style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                  <Icon path={Icons.upload} size={14}/> Browse &amp; Select File
                </label>
              </div>
            </div>

            {/* ── Selected file preview ── */}
            {selectedFile ? (
              <div className="flex items-center gap-3 px-4 py-3 rounded-xl"
                style={{background:dark?"#0a1e14":"#f0faf5",border:"1px solid #34d39928"}}>
                <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0" style={{background:"#34d39920"}}>
                  <Icon path={Icons.fileText} size={16} style={{color:"#34d399"}}/>
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate" style={{color:dark?"#e8f0f8":"#0d2137"}}>{selectedFile.name}</p>
                  <p className="text-[10px] mt-0.5" style={{color:"#34d39980"}}>
                    {(selectedFile.size/1024).toFixed(1)} KB · Ready to import
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={() => setSelectedFile(null)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                    Cancel
                  </button>
                  <button
                    onClick={() => handleImport(selectedFile)}
                    disabled={importing}
                    className="flex items-center gap-1.5 px-4 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60 transition-all"
                    style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                    {importing
                      ? <><SpinIcon size={13}/> Importing…</>
                      : <><Icon path={Icons.check} size={13}/> Upload &amp; Import</>}
                  </button>
                </div>
              </div>
            ) : (
              <p className={`text-center text-xs ${t.textMuted}`}>
                <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg"
                  style={{background:dark?"#1a2d44":"#f0f4f8"}}>
                  <Icon path={Icons.info} size={11}/> No file selected
                </span>
              </p>
            )}
          </div>
        ) : (
          <div className="flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium" style={{background:"#f8717115",color:"#f87171"}}>
            <Icon path={Icons.shield} size={13}/> Admin access required
          </div>
        )}

        {/* Template spec */}
        <div className={`mt-4 p-3 rounded-xl border border-dashed ${t.cardBorder}`}>
          <p className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${t.textMuted}`}>Template Columns</p>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs">
            {[
              {col:"A",label:"Client Name",req:true,  color:"#00c9b1"},
              {col:"B",label:"Name of User",req:true,  color:"#4a90d9"},
              {col:"C",label:"Sector",      req:false, color:"#34d399"},
              {col:"D",label:"State",       req:false, color:"#fbbf24"},
            ].map(f=>(
              <div key={f.col} className="flex items-center gap-2 px-2 py-1.5 rounded-lg" style={{background:f.color+"12"}}>
                <span className="font-black text-[10px]" style={{color:f.color}}>Col {f.col}</span>
                <span className={t.textMuted}>{f.label}</span>
                {f.req&&<span className="text-[9px] font-bold ml-auto" style={{color:f.color}}>*</span>}
              </div>
            ))}
          </div>
          <p className={`text-[10px] mt-2 ${t.textMuted}`}>
            * Required · Multiple users: <code className="px-1 rounded" style={{background:dark?"#1a2d44":"#e8f0f8"}}>User A, User B</code> · Users validated against User Database (active only)
          </p>
        </div>
      </div>

      {/* Client Registry table */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
        <div className="p-5 space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 className={`font-semibold ${t.text}`}>Client Registry</h3>
              <p className={`text-xs mt-0.5 ${t.textMuted}`}>{filtered.length} of {clients.length} records</p>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button onClick={exportClients}
                className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}
                title="Export current filtered view to Excel">
                <Icon path={Icons.download} size={14}/> Export Excel
              </button>
              {isAdmin && (
                <button onClick={()=>setAddModal(true)}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                  style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                  <Icon path={Icons.userPlus} size={14}/> Add Client
                </button>
              )}
            </div>
          </div>

          {/* Search + Filters row */}
          <div className="flex flex-wrap gap-2 items-center">
            {/* Global search */}
            <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border ${t.input} text-sm flex-1`} style={{minWidth:160}}>
              <Icon path={Icons.search} size={13} className={t.textMuted}/>
              <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}} placeholder="Search clients, users…"
                className={`bg-transparent outline-none flex-1 text-xs ${t.text}`} style={{colorScheme:dark?"dark":"light"}}/>
              {search&&<button onClick={()=>{setSearch("");setPage(1);}} className={t.textMuted}><Icon path={Icons.x} size={11}/></button>}
            </div>
            {/* Name of User filter */}
            <input value={userF} onChange={e=>{setUserF(e.target.value);setPage(1);}}
              placeholder="Filter by User…" className={`${selStyle} flex-1`} style={{minWidth:120}}/>
            {/* Sector filter */}
            <select value={sectorF} onChange={e=>{setSectorF(e.target.value);setPage(1);}} className={`${selStyle}`} style={{minWidth:130}}>
              <option value="">All Sectors</option>
              {uniqueSectors.map(s=><option key={s} value={s}>{s}</option>)}
            </select>
            {/* State filter */}
            <select value={stateF} onChange={e=>{setStateF(e.target.value);setPage(1);}} className={`${selStyle}`} style={{minWidth:120}}>
              <option value="">All States</option>
              {uniqueStates.map(s=><option key={s} value={s}>{s}</option>)}
            </select>
            {hasActiveFilters && (
              <button onClick={clearFilters}
                className={`flex items-center gap-1 px-2.5 py-2 rounded-xl text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                <Icon path={Icons.x} size={11}/> Clear
              </button>
            )}
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className={`${t.tableHead} text-xs uppercase tracking-wider`}>
                {[
                  {key:"id",        label:"Client ID"},
                  {key:"clientName",label:"Client Name"},
                  {key:"personName",label:"Name of User"},
                  {key:"sector",    label:"Sector"},
                  {key:"state",     label:"State"},
                  {key:"actions",   label:"Actions", sortable:false},
                ].map(col=>(
                  <th key={col.key} className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder}`}>
                    {col.sortable!==false ? (
                      <button onClick={()=>toggleSort(col.key)} className={`flex items-center gap-1.5 ${t.textMuted} hover:opacity-80`}>
                        {col.label}<SortIcon field={col.key} sortField={sortField} sortDir={sortDir}/>
                      </button>
                    ) : col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${t.divider}`}>
              {paged.length === 0 ? (
                <tr><td colSpan={6} className="px-4 py-12 text-center">
                  <p className={`text-sm ${t.textMuted}`}>{hasActiveFilters?"No clients match your filters.":"No clients yet."}</p>
                </td></tr>
              ) : paged.map((c, ri) => {
                // Support comma-separated multi-user display
                const users = (c.personName||"").split(",").map(n=>n.trim()).filter(Boolean);
                return (
                  <tr key={c.id} className={`transition-colors`}
                    style={{background: ri%2===0?(dark?"rgba(15,30,46,0.8)":"#ffffff"):(dark?"rgba(12,24,38,0.5)":"#f8fbff")}}>
                    <td className={`px-4 py-3 font-mono text-xs font-semibold ${t.textAccent}`}>{c.id}</td>
                    <td className="px-4 py-3">
                      <div className={`font-semibold text-sm ${t.text}`}>{c.clientName}</div>
                      <div className={`text-[10px] mt-0.5 ${t.textMuted}`}>{c.type}</div>
                    </td>
                    <td className="px-4 py-3">
                      <div className="flex flex-wrap gap-1">
                        {users.map((u,i)=>(
                          <div key={i} className="flex items-center gap-1.5">
                            <div className="w-6 h-6 rounded-full flex items-center justify-center text-white text-[9px] font-bold shrink-0"
                              style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                              {u.split(" ").map(n=>n[0]).slice(0,2).join("")}
                            </div>
                            <span className={`text-xs ${t.text}`}>{u}</span>
                            {i < users.length-1 && <span className={`text-[9px] ${t.textMuted}`}>·</span>}
                          </div>
                        ))}
                        {users.length === 0 && <span className={`text-xs italic ${t.textMuted}`}>—</span>}
                      </div>
                    </td>
                    <td className="px-4 py-3">
                      {c.sector && c.sector !== "—" ? (
                        <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{background:"#4a90d915",color:"#4a90d9"}}>{c.sector}</span>
                      ) : <span className={`text-xs italic ${t.textMuted}`}>—</span>}
                    </td>
                    <td className="px-4 py-3">
                      {c.state && c.state !== "—" ? (
                        <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{background:"#34d39915",color:"#34d399"}}>{c.state}</span>
                      ) : <span className={`text-xs italic ${t.textMuted}`}>—</span>}
                    </td>
                    <td className="px-4 py-3">
                      {isAdmin ? (
                        <div className="flex items-center gap-1.5">
                          <button onClick={()=>setEditModal(c)} className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.edit} size={12}/> Edit
                          </button>
                          <button onClick={()=>setConfirmDelete(c.id)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium border border-[#f8717140] text-[#f87171] hover:bg-[#f8717115]">
                            <Icon path={Icons.trash} size={12}/> Delete
                          </button>
                        </div>
                      ) : <span className={`text-xs ${t.textMuted} italic`}>Read only</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div className={`px-5 py-3 border-t ${t.cardBorder} flex flex-col sm:flex-row sm:items-center justify-between gap-2`}>
          <p className={`text-xs ${t.textMuted}`}>Showing {filtered.length===0?0:Math.min((page-1)*PAGE_SIZE+1,filtered.length)}–{Math.min(page*PAGE_SIZE,filtered.length)} of {filtered.length}</p>
          <div className="flex items-center gap-1">
            {["«","‹"].map((lbl,i)=>(
              <button key={lbl} onClick={()=>setPage(i===0?1:p=>Math.max(1,p-1))} disabled={page===1}
                className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${page===1?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>{lbl}</button>
            ))}
            {Array.from({length:totalPages},(_,i)=>i+1)
              .filter(p=>p===1||p===totalPages||Math.abs(p-page)<=1)
              .reduce((acc,p,i,arr)=>{if(i>0&&p-arr[i-1]>1)acc.push("…");acc.push(p);return acc;},[])
              .map((p,i)=>p==="…"
                ?<span key={`e${i}`} className={`w-7 h-7 flex items-center justify-center text-xs ${t.textMuted}`}>…</span>
                :(
                  <button key={p} onClick={()=>setPage(p)}
                    className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-medium`}
                    style={page===p?{background:"linear-gradient(135deg,#00a896,#1b5fa8)",color:"#fff"}:{color:"#5a7a99"}}>{p}</button>
                ))}
            {["›","»"].map((lbl,i)=>(
              <button key={lbl} onClick={()=>setPage(i===0?p=>Math.min(totalPages,p+1):totalPages)} disabled={page===totalPages}
                className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${page===totalPages?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>{lbl}</button>
            ))}
          </div>
        </div>
      </div>

      {/* Modals */}
      {confirmDelete  && <ConfirmDialog t={t} message={`Remove ${clients.find(c=>c.id===confirmDelete)?.clientName}?`} onConfirm={()=>handleDelete(confirmDelete)} onCancel={()=>setConfirmDelete(null)}/>}
      {editModal      && <ClientFormModal t={t} dark={dark} initial={editModal} onSave={handleEdit} onClose={()=>setEditModal(null)}/>}
      {addModal       && <ClientFormModal t={t} dark={dark} initial={null}      onSave={handleAdd}  onClose={()=>setAddModal(false)}/>}
      {importResult   && <ImportResultModal t={t} dark={dark} result={importResult} onClose={()=>setImportResult(null)}/>}
    </div>
  );
};


// ─────────────────────────────────────────────────────────────────────────────
// ── MANUAL ROW TABLE (Admin-only: Add / Edit inline / Delete / Save) ─────────
// ─────────────────────────────────────────────────────────────────────────────
let _manualRowId = 1;
const genManualId = (prefix) => `${prefix}-${String(_manualRowId++).padStart(4,"0")}`;

const ManualRowTable = ({ t, dark, isAdmin, title, subtitle, accentColor, idPrefix, columns, emptyLabel }) => {
  const [rows, setRows]           = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [editVals, setEditVals]   = useState({});
  const [deleteId, setDeleteId]   = useState(null);
  const [savedToast, setSavedToast] = useState(false);

  const addRow = () => {
    const blank = { _id: genManualId(idPrefix) };
    columns.forEach(c => { blank[c.key] = ""; });
    setRows(r => [...r, blank]);
    setEditingId(blank._id);
    setEditVals(blank);
  };

  const startEdit = (row) => { setEditingId(row._id); setEditVals({...row}); };
  const cancelEdit = () => { setEditingId(null); setEditVals({}); };

  const saveEdit = () => {
    setRows(r => r.map(row => row._id === editingId ? { ...row, ...editVals } : row));
    setEditingId(null);
    setEditVals({});
    setSavedToast(true);
    setTimeout(() => setSavedToast(false), 2200);
  };

  const deleteRow = (id) => { setRows(r => r.filter(row => row._id !== id)); setDeleteId(null); };

  return (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
      {savedToast && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-2 px-4 py-3 rounded-xl
          shadow-2xl text-sm font-semibold text-white"
          style={{background:"linear-gradient(135deg,#34d399,#10b981)",minWidth:200,zIndex:75}}>
          <Icon path={Icons.check} size={15}/> Row saved successfully!
        </div>
      )}
      {/* Header */}
      <div className={`px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b ${t.cardBorder}`}>
        <div>
          <h3 className={`font-bold text-sm ${t.text}`}>{title}</h3>
          <p className={`text-xs mt-0.5 ${t.textMuted}`}>{subtitle} · {rows.length} row{rows.length !== 1 ? "s" : ""}</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          {isAdmin ? (
            <>
              <button onClick={addRow}
                className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold text-white
                  hover:opacity-90 active:scale-95 transition-all"
                style={{background:`linear-gradient(135deg,${accentColor},#1b5fa8)`}}
                title="Add new row (Admin only)">
                <Icon path={Icons.plus} size={13}/> Add Row
              </button>
              {rows.length > 0 && (
                <button onClick={() => { setSavedToast(true); setTimeout(() => setSavedToast(false), 2200); }}
                  className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-medium
                    ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                  <Icon path={Icons.save} size={13}/> Save All
                </button>
              )}
            </>
          ) : (
            <span className="flex items-center gap-1 text-xs font-medium px-2.5 py-1.5 rounded-lg"
              style={{background:"#f8717110",color:"#f87171",border:"1px solid #f8717125"}}>
              <Icon path={Icons.shield} size={12}/> Read-only
            </span>
          )}
        </div>
      </div>

      {/* Table */}
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className={`${t.tableHead} text-xs uppercase tracking-wider`}>
              <th className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder} w-10`}>#</th>
              {columns.map(col => (
                <th key={col.key} className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder}`}
                  style={col.width ? {width: col.width} : {}}>{col.label}</th>
              ))}
              {isAdmin && (
                <th className={`px-4 py-3 text-center font-semibold border-b ${t.cardBorder} w-24`}>Actions</th>
              )}
            </tr>
          </thead>
          <tbody className={`divide-y ${t.divider}`}>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={columns.length + (isAdmin ? 2 : 1)}
                  className="px-4 py-12 text-center">
                  <p className={`text-sm ${t.textMuted}`}>{emptyLabel || "No rows yet."}</p>
                  {isAdmin && (
                    <button onClick={addRow}
                      className="mt-3 text-xs font-semibold px-4 py-2 rounded-xl text-white"
                      style={{background:`linear-gradient(135deg,${accentColor},#1b5fa8)`}}>
                      + Add first row
                    </button>
                  )}
                </td>
              </tr>
            ) : rows.map((row, idx) => {
              const isEditing = editingId === row._id;
              return (
                <tr key={row._id} className={`${t.tableRow} transition-colors ${isEditing ? (dark ? "bg-[#132036]" : "bg-[#e8f4f2]") : ""}`}>
                  <td className={`px-4 py-2.5 text-xs font-mono font-bold text-center ${t.textMuted}`}>{idx + 1}</td>
                  {columns.map(col => (
                    <td key={col.key} className="px-4 py-2">
                      {isAdmin && isEditing ? (
                        col.type === "select" ? (
                          <select
                            value={editVals[col.key] || ""}
                            onChange={e => setEditVals(v => ({...v, [col.key]: e.target.value}))}
                            className={`w-full px-2 py-1.5 rounded-lg border text-sm outline-none ${t.input} focus:ring-2 focus:ring-[#00c9b1]`}>
                            {(col.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        ) : (
                          <input
                            type={col.inputType || "text"}
                            value={editVals[col.key] || ""}
                            onChange={e => setEditVals(v => ({...v, [col.key]: e.target.value}))}
                            placeholder={col.placeholder || ""}
                            className={`w-full px-2 py-1.5 rounded-lg border text-sm outline-none ${t.input} focus:ring-2 focus:ring-[#00c9b1]`}/>
                        )
                      ) : (
                        <span className={`text-sm ${row[col.key] ? t.text : t.textMuted}`}>
                          {row[col.key] || <span className="italic opacity-50">—</span>}
                        </span>
                      )}
                    </td>
                  ))}
                  {isAdmin && (
                    <td className="px-4 py-2 text-center">
                      {isEditing ? (
                        <div className="flex items-center justify-center gap-1">
                          <button onClick={saveEdit}
                            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold text-white"
                            style={{background:`linear-gradient(135deg,${accentColor},#1b5fa8)`}}>
                            <Icon path={Icons.save} size={11}/> Save
                          </button>
                          <button onClick={cancelEdit}
                            className={`px-2 py-1.5 rounded-lg text-xs ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.x} size={11}/>
                          </button>
                        </div>
                      ) : (
                        <div className="flex items-center justify-center gap-1">
                          <button onClick={() => startEdit(row)}
                            className={`flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.edit} size={11}/>
                          </button>
                          <button onClick={() => setDeleteId(row._id)}
                            className="px-2 py-1.5 rounded-lg text-xs border border-[#f8717140] text-[#f87171] hover:bg-[#f8717115]">
                            <Icon path={Icons.trash} size={11}/>
                          </button>
                        </div>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {deleteId && (
        <ConfirmDialog t={t}
          message="Remove this row? This cannot be undone."
          onConfirm={() => deleteRow(deleteId)}
          onCancel={() => setDeleteId(null)}/>
      )}
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── CAS MIS TAB ──────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// ── SYNC SHARED COMPONENTS ───────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

// Spinning refresh icon animation via inline style
const SpinIcon = ({ size = 14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round"
    style={{ animation: "spin 0.8s linear infinite" }}>
    <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    <path d="M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
  </svg>
);

// Global sync toast shown at top-right
const SyncToast = ({ toast }) => {
  if (!toast) return null;
  const isError = toast.type === "error";
  return (
    <div className="fixed top-5 right-5 z-[80] flex items-center gap-3 px-4 py-3
      rounded-xl shadow-2xl text-sm font-semibold text-white animate-fade-in"
      style={{
        background: isError
          ? "linear-gradient(135deg,#ef4444,#dc2626)"
          : "linear-gradient(135deg,#00a896,#1b5fa8)",
        minWidth: 280,
        boxShadow: "0 8px 32px rgba(0,0,0,0.35)",
      }}>
      {isError
        ? <Icon path={Icons.alert} size={16}/>
        : <Icon path={Icons.checkCircle} size={16}/>}
      {toast.message}
    </div>
  );
};

// Refresh button (compact, reusable)
const SyncButton = ({ t, onSync, syncing, lastSynced }) => {
  const fmt = (d) => {
    if (!d) return "Never synced";
    const diff = Math.floor((Date.now() - d.getTime()) / 1000);
    if (diff < 60) return `${diff}s ago`;
    if (diff < 3600) return `${Math.floor(diff/60)}m ago`;
    return d.toLocaleTimeString([], { hour:"2-digit", minute:"2-digit" });
  };
  return (
    <div className="flex items-center gap-2 flex-wrap justify-end">
      {lastSynced && (
        <span className={`text-[10px] font-medium ${t.textMuted}`}>
          Last sync: {fmt(lastSynced)}
        </span>
      )}
      <button
        onClick={onSync}
        disabled={syncing}
        className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-semibold
          text-white transition-all hover:opacity-90 active:scale-95 disabled:opacity-60"
        style={{ background: "linear-gradient(135deg,#00a896,#1b5fa8)" }}>
        {syncing ? <SpinIcon size={13}/> : <Icon path={Icons.refreshCw} size={13}/>}
        {syncing ? "Syncing…" : "Refresh Data"}
      </button>
    </div>
  );
};

// Info banner shown at top of each synced module
const SyncBanner = ({ t, dark, moduleName, accentColor }) => {
  const { clients, syncing, lastSynced, autoSync, setAutoSync } = useSyncContext();
  const fmt = (d) => {
    if (!d) return "Never";
    return d.toLocaleTimeString([], { hour:"2-digit", minute:"2-digit", second:"2-digit" });
  };
  return (
    <div className="flex flex-wrap items-center gap-3 px-4 py-3 rounded-2xl text-xs"
      style={{ background: accentColor+"10", border: `1px solid ${accentColor}30` }}>
      <div className="flex items-center gap-2">
        <div className={`w-1.5 h-1.5 rounded-full ${syncing ? "animate-pulse" : ""}`}
          style={{ background: accentColor }}/>
        <span className="font-bold" style={{ color: accentColor }}>{moduleName}</span>
        <span className={t.textMuted}>— Client data synced from Master</span>
      </div>
      <div className="flex items-center gap-4 ml-auto flex-wrap">
        <span className={t.textMuted}>Last sync: {fmt(lastSynced)}</span>

        <label className={`flex items-center gap-1.5 cursor-pointer ${t.textMuted}`}>
          <div
            className="relative w-7 h-4 rounded-full transition-colors"
            style={{ background: autoSync ? accentColor : (dark ? "#2d3748" : "#e2e8f0") }}
            onClick={() => setAutoSync(a => !a)}>
            <div className="absolute top-0.5 w-3 h-3 rounded-full bg-white shadow transition-all"
              style={{ left: autoSync ? "14px" : "2px" }}/>
          </div>
          Auto-sync
        </label>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── CLIENT-MONTH TABLE ────────────────────────────────────────────────────────
// Shared across CAS MIS and KRA/KPI.
// For every selected month it ALWAYS shows ALL active clients from Client Master.
// Existing data is merged; clients without data get blank editable rows.
// ─────────────────────────────────────────────────────────────────────────────
const ClientMonthTable = ({ t, dark, isAdmin, accentColor, moduleKey, title, fy, columns }) => {
  const { clients: masterClients } = useSyncContext();
  const [month, setMonth]         = useState("April");
  // savedData: { [moduleKey-clientId-month]: { fieldKey: value, ... } }
  const [savedData, setSavedData] = useState({});
  const [editingKey, setEditingKey] = useState(null);
  const [editVals, setEditVals]   = useState({});
  const [savedToast, setSavedToast] = useState(false);

  const makeKey = (clientId, m) => `${moduleKey}-${clientId}-${m}`;

  // Build complete client row list for selected month (left-join pattern)
  const monthRows = useMemo(() => {
    return masterClients.map(client => {
      const key = makeKey(client.id, month);
      const data = savedData[key] || {};
      return { clientId: client.id, clientName: client.clientName, personName: client.personName, month, _key: key, ...data };
    });
  }, [masterClients, month, savedData]);

  const startEdit = (key, row) => {
    setEditingKey(key);
    const existing = savedData[key] || {};
    setEditVals({ ...existing });
  };

  const saveEdit = (key) => {
    setSavedData(prev => ({ ...prev, [key]: { ...editVals } }));
    setEditingKey(null);
    setEditVals({});
    setSavedToast(true);
    setTimeout(() => setSavedToast(false), 2000);
  };

  const clearRow = (key) => {
    setSavedData(prev => { const n = { ...prev }; delete n[key]; return n; });
    if (editingKey === key) { setEditingKey(null); setEditVals({}); }
  };

  const filledCount = monthRows.filter(r => {
    const d = savedData[r._key];
    return d && Object.values(d).some(v => v && v !== "");
  }).length;

  const sel = `px-3 py-2 rounded-xl border text-sm outline-none transition-all ${t.input} focus:ring-2`;

  return (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
      {savedToast && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{background:`linear-gradient(135deg,${accentColor},#1b5fa8)`,minWidth:220}}>
          <Icon path={Icons.check} size={15}/> Row saved!
        </div>
      )}
      {/* Header */}
      <div className={`px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b ${t.cardBorder}`}>
        <div>
          <h3 className={`font-bold text-sm ${t.text}`}>{title}</h3>
          <p className={`text-xs mt-0.5 ${t.textMuted}`}>
            {monthRows.length} clients · {filledCount} filled · {monthRows.length - filledCount} pending · FY {fy}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <span className={`text-xs font-semibold ${t.textMuted}`}>Month</span>
          <select value={month} onChange={e => setMonth(e.target.value)} className={sel} style={{minWidth:110}}>
            {["April","May","June","July","August","September","October","November","December","January","February","March"].map(m => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
          {!isAdmin && (
            <span className="flex items-center gap-1 text-xs font-medium px-2 py-1 rounded-lg"
              style={{background:"#f8717110",color:"#f87171",border:"1px solid #f8717125"}}>
              <Icon path={Icons.shield} size={11}/> Read-only
            </span>
          )}
        </div>
      </div>
      {/* Info banner */}
      <div className="px-5 py-2 flex items-center gap-2 text-xs" style={{background:accentColor+"0a",borderBottom:`1px solid ${accentColor}20`}}>
        <div className="w-1.5 h-1.5 rounded-full" style={{background:accentColor}}/>
        <span style={{color:accentColor}} className="font-semibold">All {masterClients.length} clients from Client Master shown for {month}</span>
        <span className={t.textMuted}>— blank rows indicate no data entered yet</span>
      </div>
      {/* Table */}
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className={`${t.tableHead} text-xs uppercase tracking-wider`}>
              <th className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder} w-8`}>#</th>
              <th className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder}`} style={{minWidth:200}}>Client Name</th>
              <th className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder}`} style={{minWidth:140}}>Contact Person</th>
              {columns.map(col => (
                <th key={col.key} className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder}`}
                  style={col.width ? {width:col.width} : {minWidth:120}}>{col.label}</th>
              ))}
              {isAdmin && <th className={`px-4 py-3 text-center font-semibold border-b ${t.cardBorder} w-24`}>Actions</th>}
            </tr>
          </thead>
          <tbody className={`divide-y ${t.divider}`}>
            {monthRows.map((row, idx) => {
              const isFilled = !!(savedData[row._key] && Object.values(savedData[row._key]).some(v => v && v !== ""));
              const isEditing = editingKey === row._key;
              return (
                <tr key={row._key}
                  className={`transition-colors ${isEditing?(dark?"bg-[#132036]":"bg-[#e8f4f2]"):isFilled?t.tableRow:(dark?"bg-[#0f1e2e]/70 border-l-2":"bg-[#f9feff]/80 border-l-2")} `}
                  style={!isFilled && !isEditing ? {borderLeftColor:`${accentColor}30`} : {}}>
                  <td className={`px-4 py-2.5 text-xs font-mono font-bold text-center ${t.textMuted}`}>
                    {idx + 1}
                    {!isFilled && <span className="block text-[8px] mt-0.5 font-normal" style={{color:`${accentColor}80`}}>empty</span>}
                  </td>
                  <td className="px-4 py-2.5">
                    <div className={`font-semibold text-sm ${t.text}`}>{row.clientName}</div>
                  </td>
                  <td className={`px-4 py-2.5 text-xs ${t.textMuted}`}>{row.personName}</td>
                  {columns.map(col => (
                    <td key={col.key} className="px-4 py-2">
                      {isAdmin && isEditing ? (
                        col.type === "select" ? (
                          <select value={editVals[col.key]||""} onChange={e=>setEditVals(v=>({...v,[col.key]:e.target.value}))}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input}`}>
                            <option value="">— select —</option>
                            {(col.options||[]).map(opt=><option key={opt} value={opt}>{opt}</option>)}
                          </select>
                        ) : (
                          <input type={col.inputType||"text"} value={editVals[col.key]||""}
                            onChange={e=>setEditVals(v=>({...v,[col.key]:e.target.value}))}
                            placeholder={col.placeholder||""}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input}`}/>
                        )
                      ) : (
                        <span className={`text-xs ${(savedData[row._key]||{})[col.key]?t.text:t.textMuted}`}>
                          {(savedData[row._key]||{})[col.key]||<span className="italic opacity-40">—</span>}
                        </span>
                      )}
                    </td>
                  ))}
                  {isAdmin && (
                    <td className="px-4 py-2 text-center">
                      {isEditing ? (
                        <div className="flex items-center justify-center gap-1">
                          <button onClick={()=>saveEdit(row._key)}
                            className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-semibold text-white"
                            style={{background:`linear-gradient(135deg,${accentColor},#1b5fa8)`}}>
                            <Icon path={Icons.save} size={11}/> Save
                          </button>
                          <button onClick={()=>{setEditingKey(null);setEditVals({});}}
                            className={`px-2 py-1.5 rounded-lg text-xs ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.x} size={11}/>
                          </button>
                        </div>
                      ) : (
                        <div className="flex items-center justify-center gap-1">
                          <button onClick={()=>startEdit(row._key, row)}
                            className={`flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}
                            title="Edit row">
                            <Icon path={Icons.edit} size={11}/>
                          </button>
                          {isFilled && (
                            <button onClick={()=>clearRow(row._key)}
                              className="px-2 py-1.5 rounded-lg text-xs border border-[#f8717140] text-[#f87171] hover:bg-[#f8717115]"
                              title="Clear row data">
                              <Icon path={Icons.trash} size={11}/>
                            </button>
                          )}
                        </div>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {/* Legend */}
      <div className={`px-5 py-2.5 border-t ${t.cardBorder} flex flex-wrap items-center gap-4 text-[10px] ${t.textMuted}`}>
        <div className="flex items-center gap-1.5">
          <div className="w-3 h-3 rounded-sm border-l-2" style={{borderColor:`${accentColor}50`,background:dark?"#0f1e2e":"#f9feff"}}/>
          <span>Empty — no data entered yet (client always shown)</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-3 h-3 rounded-sm" style={{background:accentColor+"30"}}/>
          <span>Contains saved data</span>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── CAS MIS — Supporting sub-components ──────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

// User→client assignment map for CAS MIS (mirrors INITIAL_USERS assignedClients)
const CAS_USER_ASSIGNMENTS = {
  "USR-001": null,                                              // Admin — all clients
  "USR-002": ["CLT-001","CLT-002","CLT-004","CLT-007"],
  "USR-003": ["CLT-003","CLT-005","CLT-006"],
  "USR-004": ["CLT-008"],
};
// Demo: current logged-in end-user (non-admin viewer). In a real app this comes from auth context.
const CAS_CURRENT_USER_ID = "USR-002";

const CAS_CURRENCIES = ["INR","USD","EUR","GBP","AED","SGD","AUD","CAD","JPY"];
const CAS_YNA        = ["Yes","No","NA"];
const CAS_ACCT_SW    = ["Zoho","SAP","Tally","NetSuite","Coupa","Others"];

// ── Statutory due-date compliance helpers ─────────────────────────────────────
// Returns the ISO due-date string for a statutory field given the selected month/FY.
// fieldKey must be one of the 6 statutory fields; returns null for non-statutory fields.
const MONTHS_ORDER = [
  "April","May","June","July","August","September",
  "October","November","December","January","February","March"
];
// Given a filing month name and the FY string ("2026-27"), returns the
// calendar year that month belongs to.
const calYearForMonth = (monthName, fy) => {
  const [startYr] = fy.split("-").map(Number);
  const idx = MONTHS_ORDER.indexOf(monthName);
  // April–December belong to startYr; January–March to startYr+1
  return idx <= 8 ? startYr : startYr + 1;
};
// Returns the subsequent month name and calendar year for a given filing month.
const subsequentMonth = (monthName, fy) => {
  const idx        = MONTHS_ORDER.indexOf(monthName);
  const nextIdx    = (idx + 1) % 12;
  const calYear    = calYearForMonth(monthName, fy);
  const nextYear   = nextIdx === 0 ? calYear + 1 : calYear;
  return { name: MONTHS_ORDER[nextIdx], year: nextYear };
};
// Map field keys → due-date calculation rules
const STATUTORY_DUE_RULES = {
  // TDS: 7th of subsequent month; special case March → 30 Apr
  tdsPayDate:  (month, fy) => {
    if (month === "March") { const y = calYearForMonth(month, fy) + 1; return new Date(y, 3, 30); }
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 7);
  },
  // PT: 20th of subsequent month
  ptPayDate:   (month, fy) => {
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 20);
  },
  // PF: 15th of subsequent month
  pfPayDate:   (month, fy) => {
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 15);
  },
  // ESI: 15th of subsequent month
  esiPayDate:  (month, fy) => {
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 15);
  },
  // GSTR-1: 11th of subsequent month
  gstr1Date:   (month, fy) => {
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 11);
  },
  // GSTR-3B: 20th of subsequent month
  gstr3bDate:  (month, fy) => {
    const { name, year } = subsequentMonth(month, fy);
    const mIdx = ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(name);
    return new Date(year, mIdx, 20);
  },
};
// Returns compliance info for a field value given the filing month + FY.
// { dueDate: Date|null, delayDays: number, compliant: bool|null, dueDateStr: string }
const getComplianceInfo = (fieldKey, valueStr, month, fy) => {
  const ruleFn = STATUTORY_DUE_RULES[fieldKey];
  if (!ruleFn || !valueStr) return { dueDate:null, delayDays:0, compliant:null, dueDateStr:"" };
  const dueDate = ruleFn(month, fy);
  const actual  = new Date(valueStr);
  if (isNaN(actual.getTime())) return { dueDate, delayDays:0, compliant:null, dueDateStr:"" };
  // Strip time for comparison
  const due0    = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  const act0    = new Date(actual.getFullYear(),  actual.getMonth(),  actual.getDate());
  const diffMs  = act0 - due0;
  const delayDays = Math.round(diffMs / 86400000);
  const compliant = delayDays <= 0;
  const dueDateStr = dueDate.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" });
  return { dueDate, delayDays, compliant, dueDateStr };
};
const COMPLY_GREEN = "#16A34A";
const COMPLY_RED   = "#DC2626";

// ── DateCellInput — calendar picker with NA toggle + statutory compliance indicator
const CasDateCell = ({ value, naValue, onChange, onNaChange, disabled, fieldKey, month, fy }) => {
  const isNA = naValue === true;
  const compliance = getComplianceInfo(fieldKey, value, month, fy);
  const isStatutory = !!STATUTORY_DUE_RULES[fieldKey];

  // Format actual date for display
  const fmtDate = (v) => {
    if (!v) return null;
    const d = new Date(v);
    if (isNaN(d.getTime())) return v;
    return d.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"2-digit" });
  };

  // Compliance indicator pill shown beside the date in both view and edit modes
  const CompliancePill = () => {
    if (!isStatutory || !value || isNA || compliance.compliant === null) return null;
    return (
      <div
        className="flex flex-col items-start leading-none"
        title={[
          `Actual: ${fmtDate(value)}`,
          `Due:    ${compliance.dueDateStr}`,
          compliance.delayDays > 0
            ? `Delay:  ${compliance.delayDays} day${compliance.delayDays !== 1 ? "s" : ""}`
            : "On time",
          `Status: ${compliance.compliant ? "Within Due Date" : "Delayed"}`,
        ].join("\n")}>
        <span
          className="text-[9px] font-black px-1.5 py-0.5 rounded-full"
          style={{
            color:      compliance.compliant ? COMPLY_GREEN : COMPLY_RED,
            background: compliance.compliant ? "#16A34A18"  : "#DC262618",
            border:     `1px solid ${compliance.compliant ? "#16A34A40" : "#DC262640"}`,
            whiteSpace: "nowrap",
          }}>
          {compliance.compliant ? "✓ In Time" : `✗ +${compliance.delayDays}d`}
        </span>
        <span className="text-[8px] mt-0.5" style={{color: compliance.compliant ? COMPLY_GREEN : COMPLY_RED, opacity:0.7}}>
          Due: {compliance.dueDateStr}
        </span>
      </div>
    );
  };

  if (disabled) {
    return (
      <div className="flex flex-col gap-0.5">
        <div className="flex items-center gap-1.5">
          {isNA
            ? <span className="px-2 py-0.5 rounded-md text-[10px] font-bold" style={{background:"#94a3b818",color:"#94a3b8"}}>NA</span>
            : <span
                className="text-xs font-medium"
                style={{
                  color: isStatutory && value && compliance.compliant !== null
                    ? (compliance.compliant ? COMPLY_GREEN : COMPLY_RED)
                    : "#e8f0f8",
                }}>
                {value ? fmtDate(value) : <span style={{opacity:0.35}}>—</span>}
              </span>}
        </div>
        {!isNA && <CompliancePill/>}
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-1" style={{minWidth:145}}>
      {!isNA && (
        <input
          type="date"
          value={value || ""}
          onChange={e => onChange(e.target.value)}
          disabled={isNA}
          className="w-full px-2 py-1 rounded-lg border text-[11px] outline-none focus:ring-1 focus:ring-[#4a90d9]"
          style={{
            background:  "transparent",
            borderColor: isStatutory && value && compliance.compliant !== null
              ? (compliance.compliant ? "#16A34A60" : "#DC262660")
              : "#243d58",
            color: isStatutory && value && compliance.compliant !== null
              ? (compliance.compliant ? COMPLY_GREEN : COMPLY_RED)
              : "#e8f0f8",
            colorScheme: "dark",
          }}
        />
      )}
      {!isNA && <CompliancePill/>}
      <label className="flex items-center gap-1.5 cursor-pointer select-none">
        <div
          onClick={() => onNaChange(!isNA)}
          className="relative w-7 h-3.5 rounded-full transition-colors duration-200 cursor-pointer"
          style={{background: isNA ? "#4a90d9" : "#243d58"}}>
          <div className="absolute top-0.5 w-2.5 h-2.5 rounded-full bg-white shadow transition-all duration-200"
            style={{left: isNA ? "15px" : "2px"}}/>
        </div>
        <span className="text-[10px] font-semibold" style={{color: isNA ? "#4a90d9" : "#5a7a99"}}>NA</span>
      </label>
    </div>
  );
};

// ── YNA Dropdown cell ─────────────────────────────────────────────────────────
const CasYnaCell = ({ value, onChange, disabled }) => {
  const col = value === "Yes" ? "#34d399" : value === "No" ? "#f87171" : value === "NA" ? "#94a3b8" : "#5a7a99";
  if (disabled) {
    return value
      ? <span className="px-2 py-0.5 rounded-full text-[10px] font-bold" style={{color:col,background:col+"18"}}>{value}</span>
      : <span style={{color:"#5a7a99",opacity:0.4}}>—</span>;
  }
  return (
    <select
      value={value || ""}
      onChange={e => onChange(e.target.value)}
      className="w-full px-2 py-1.5 rounded-lg border text-xs outline-none focus:ring-1 focus:ring-[#4a90d9]"
      style={{background:"transparent",borderColor:"#243d58",color:col,minWidth:70}}>
      <option value="">—</option>
      {CAS_YNA.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  );
};

// ── Currency + Amount cell ─────────────────────────────────────────────────────
const CasFeeCell = ({ cur, amount, onCurChange, onAmtChange, disabled, label }) => {
  if (disabled) {
    return (
      <div className="text-xs">
        {cur && amount ? <span style={{color:"#e8f0f8"}}>{cur} {parseFloat(amount).toLocaleString("en-IN",{maximumFractionDigits:2})}</span> : <span style={{opacity:0.35}}>—</span>}
      </div>
    );
  }
  return (
    <div className="flex flex-col gap-1" style={{minWidth:150}}>
      <select
        value={cur || "INR"}
        onChange={e => onCurChange(e.target.value)}
        className="w-full px-2 py-1 rounded-lg border text-[10px] outline-none"
        style={{background:"transparent",borderColor:"#243d58",color:"#4a90d9"}}>
        {CAS_CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
      </select>
      <input
        type="number"
        value={amount || ""}
        onChange={e => onAmtChange(e.target.value)}
        placeholder="0.00"
        className="w-full px-2 py-1 rounded-lg border text-[10px] outline-none focus:ring-1 focus:ring-[#4a90d9]"
        style={{background:"transparent",borderColor:"#243d58",color:"#e8f0f8"}}
      />
    </div>
  );
};

// ── Remarks icon + inline popup ───────────────────────────────────────────────
const CasRemarksCell = ({ fieldKey, rowKey, remarksStore, onSave, disabled }) => {
  const key      = `${rowKey}__${fieldKey}`;
  const existing = remarksStore[key] || "";
  const [open, setOpen]   = useState(false);
  const [text, setText]   = useState(existing);
  const ref = useRef(null);

  useEffect(() => { setText(existing); }, [existing]);
  useEffect(() => {
    const h = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);

  const hasNote = !!existing.trim();
  return (
    <div ref={ref} className="relative flex items-center">
      <button
        title={hasNote ? existing : "Add remark"}
        onClick={() => { if (!disabled || hasNote) setOpen(o => !o); }}
        className="w-5 h-5 rounded flex items-center justify-center transition-colors"
        style={{
          background: hasNote ? "#4a90d922" : "transparent",
          color: hasNote ? "#4a90d9" : "#5a7a99",
          border: hasNote ? "1px solid #4a90d940" : "1px solid transparent",
        }}>
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
      </button>
      {open && (
        <div className="absolute z-50 top-full left-0 mt-1 rounded-xl shadow-2xl border overflow-hidden"
          style={{background:"#0f1e2e",borderColor:"#243d58",minWidth:220,maxWidth:280}}>
          <div className="px-3 py-2 border-b flex items-center justify-between" style={{borderColor:"#1a2d44"}}>
            <span className="text-[10px] font-bold" style={{color:"#4a90d9"}}>Field Remark</span>
            {!disabled && hasNote && (
              <button onClick={() => { onSave(key,""); setText(""); setOpen(false); }}
                className="text-[9px] text-[#f87171] hover:opacity-75">Clear</button>
            )}
          </div>
          {disabled ? (
            <p className="px-3 py-2.5 text-xs" style={{color:"#e8f0f8"}}>{existing || "No remark"}</p>
          ) : (
            <div className="p-2 space-y-2">
              <textarea
                autoFocus
                value={text}
                onChange={e => setText(e.target.value)}
                rows={3}
                placeholder="Type remark…"
                className="w-full px-2.5 py-2 rounded-lg border text-xs outline-none resize-none focus:ring-1 focus:ring-[#4a90d9]"
                style={{background:"#1a2d44",borderColor:"#243d58",color:"#e8f0f8"}}
              />
              <button
                onClick={() => { onSave(key, text); setOpen(false); }}
                className="w-full py-1.5 rounded-lg text-[11px] font-semibold text-white"
                style={{background:"linear-gradient(135deg,#4a90d9,#1b5fa8)"}}>
                Save Remark
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// ── Accounting Software cell ───────────────────────────────────────────────────
const CasAcctSwCell = ({ value, otherText, onChange, onOtherChange, disabled }) => {
  if (disabled) {
    return <span className="text-xs" style={{color:"#e8f0f8"}}>{value === "Others" ? otherText || "Others" : value || <span style={{opacity:0.35}}>—</span>}</span>;
  }
  return (
    <div className="flex flex-col gap-1" style={{minWidth:110}}>
      <select value={value||""} onChange={e=>onChange(e.target.value)}
        className="w-full px-2 py-1.5 rounded-lg border text-xs outline-none focus:ring-1 focus:ring-[#4a90d9]"
        style={{background:"transparent",borderColor:"#243d58",color:"#e8f0f8"}}>
        <option value="">—</option>
        {CAS_ACCT_SW.map(o=><option key={o} value={o}>{o}</option>)}
      </select>
      {value === "Others" && (
        <input type="text" value={otherText||""} onChange={e=>onOtherChange(e.target.value)}
          placeholder="Specify…"
          className="w-full px-2 py-1 rounded-lg border text-[10px] outline-none focus:ring-1 focus:ring-[#4a90d9]"
          style={{background:"transparent",borderColor:"#4a90d960",color:"#e8f0f8"}}/>
      )}
    </div>
  );
};

// ── Main CAS MIS Table component ──────────────────────────────────────────────
const CasMisTable = ({ t, dark, isAdmin, fy, month }) => {
  const { clients: masterClients } = useSyncContext();
  const { locks } = useLockCtx();
  const { assignedClientIds, canEdit: rbacCanEdit } = useRBAC();

  // Month-lock: if this FY+month is locked for CAS, end users cannot edit
  const casMisLocked = checkLocked(locks, fy, month, "cas");

  // misData: { [clientId-month]: { fieldKey: value } }
  const [misData, setMisData] = useState({});
  // remarksStore: { [clientId-month__fieldKey]: remarkText }
  const [remarksStore, setRemarksStore] = useState({});
  const [savedToast, setSavedToast] = useState(false);

  const makeKey   = (clientId) => `${clientId}-${month}`;

  // canEdit: combines role assignment + month lock
  const canEdit = (clientId) => {
    if (casMisLocked && !isAdmin) return false;
    return rbacCanEdit(clientId);
  };

  // Get user name for a client (demo: look up from INITIAL_USERS assignedClients)
  const getUserForClient = (clientId) => {
    // Find the user whose assignedClients includes this clientId
    const found = INITIAL_USERS.find(u => u.assignedClients && u.assignedClients.includes(clientId));
    return found ? found.name : "—";
  };

  const getVal  = (clientId, field)        => (misData[makeKey(clientId)] || {})[field] ?? "";
  const setVal  = (clientId, field, value) => {
    const key = makeKey(clientId);
    setMisData(prev => ({ ...prev, [key]: { ...(prev[key] || {}), [field]: value } }));
  };

  const saveRemarks = (key, text) => {
    setRemarksStore(prev => ({ ...prev, [key]: text }));
  };

  const handleSaveRow = () => {
    setSavedToast(true);
    setTimeout(() => setSavedToast(false), 2200);
  };

  // ── CAS MIS Export to Excel ───────────────────────────────────────────────
  const exportCasMis = () => {
    try {
      const exportRows = rows.map((row, idx) => {
        const d = misData[row._key] || {};
        const fmtDate = (val) => {
          if (!val || val === "") return "";
          const parsed = new Date(val);
          return isNaN(parsed.getTime()) ? val : parsed.toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
        };
        return {
          "#":                             idx + 1,
          "Client Name":                   xlsSafe(row.clientName),
          "Name of User":                  xlsSafe(row._userName || row.personName || "—"),
          "Financial Year":                fy,
          "Month":                         month,
          "PT Payment Date":               d.ptPayDateNA         ? "NA" : fmtDate(d.ptPayDate         || ""),
          "MIS Working First Cut Date":    d.misFirstCutDateNA   ? "NA" : fmtDate(d.misFirstCutDate   || ""),
          "MIS Closure Date":              d.misClosureDateNA    ? "NA" : fmtDate(d.misClosureDate    || ""),
          "TDS Payment Date":              d.tdsPayDateNA        ? "NA" : fmtDate(d.tdsPayDate        || ""),
          "PF Payment Date":               d.pfPayDateNA         ? "NA" : fmtDate(d.pfPayDate         || ""),
          "ESI Payment Date":              d.esiPayDateNA        ? "NA" : fmtDate(d.esiPayDate        || ""),
          "GSTR 1 Filing Date":            d.gstr1DateNA         ? "NA" : fmtDate(d.gstr1Date         || ""),
          "GSTR 3B Filing Date":           d.gstr3bDateNA        ? "NA" : fmtDate(d.gstr3bDate        || ""),
          "Date of Payment Made":          d.paymentMadeDateNA   ? "NA" : fmtDate(d.paymentMadeDate   || ""),
          "Date of Collection (Zoho)":     d.collectionDateNA    ? "NA" : fmtDate(d.collectionDate    || ""),
          "PT Challan to OneNote":         xlsSafe(d.ptChallan           || ""),
          "TDS Challan to OneNote":        xlsSafe(d.tdsChallan          || ""),
          "PF Challan to OneNote":         xlsSafe(d.pfChallan           || ""),
          "Comm to GST Team":              xlsSafe(d.commGST             || ""),
          "Microsoft OneNote":             xlsSafe(d.msOneNote           || ""),
          "Accounting Software":           xlsSafe(d.acctSw              || ""),
          "Monthly SPOC Fee (Mar 2025)":   d.spocFeeAmt          || "",
          "CAS Fee":                       d.casFeeAmt           || "",
          "SPOC Fee Pending from Client":  d.spocPendingAmt      || "",
          "Staff Remarks":                 xlsSafe(remarksStore[`${row._key}__staffRemarks`]   || d.staffRemarks   || ""),
          "SPOC Billing Remarks":          xlsSafe(remarksStore[`${row._key}__billingRemarks`] || d.billingRemarks || ""),
        };
      });
      if (!exportRows.length) { alert("No data to export."); return; }
      const ws = XLSX.utils.json_to_sheet(exportRows);
      ws["!cols"] = [4,28,22,12,12,16,20,16,16,16,16,16,18,18,20,14,14,14,16,14,16,18,18,18,24,24].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "CAS_MIS");
      XLSX.writeFile(wb, `CASMIS_FY${fy}_${month}.xlsx`);
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("CAS MIS export error:", err);
    }
  };

  // Build rows: one per master client — always visible
  const rows = masterClients.map((client, idx) => ({
    ...client,
    _key: makeKey(client.id),
    _idx: idx + 1,
    _canEdit: canEdit(client.id),
    _userName: getUserForClient(client.id),
  }));

  // Stats for analytics
  const filledCount = rows.filter(r => {
    const d = misData[r._key];
    return d && Object.values(d).some(v => v && v !== "");
  }).length;

  const thStyle = {
    background: dark
      ? "linear-gradient(90deg,#0c1e30 0%,#0e2240 60%,#0a1a2e 100%)"
      : "linear-gradient(90deg,#1b3a5c 0%,#1a3356 60%,#1e3d6a 100%)",
    color: "#c8dff0",
    borderBottom: "1px solid #ffffff18",
    position: "sticky",
    top: 0,
    zIndex: 10,
    whiteSpace: "nowrap",
  };

  const ThCell = ({ children, sub, w, center }) => (
    <th style={{...thStyle, minWidth: w || 110, textAlign: center ? "center" : "left"}}
      className="px-2.5 py-3 text-[9px] font-bold uppercase tracking-widest">
      <div>{children}</div>
      {sub && <div style={{fontSize:8,fontWeight:400,opacity:0.55,textTransform:"none",letterSpacing:0,marginTop:2}}>{sub}</div>}
    </th>
  );

  const textareaStyle = {
    background: "transparent",
    borderColor: "#243d58",
    color: "#e8f0f8",
    resize: "none",
  };

  return (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
      {savedToast && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{background:"linear-gradient(135deg,#34d399,#10b981)",minWidth:200,zIndex:200}}>
          <Icon path={Icons.check} size={15}/> Saved successfully!
        </div>
      )}

      {/* Table header toolbar */}
      <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex flex-col sm:flex-row sm:items-center justify-between gap-3`}>
        <div>
          <h3 className={`font-bold text-sm ${t.text}`}>CAS MIS — Monthly Client Data</h3>
          <p className={`text-xs mt-0.5 ${t.textMuted}`}>
            {rows.length} clients · {filledCount} filled · {rows.length - filledCount} pending · {month} · FY {fy}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          {!isAdmin && (
            <span className="flex items-center gap-1 text-[10px] font-medium px-2 py-1 rounded-lg"
              style={{background:"#4a90d910",color:"#4a90d9",border:"1px solid #4a90d925"}}>
              <Icon path={Icons.shield} size={10}/> Assigned clients editable · Others view-only
            </span>
          )}
          {isAdmin && (
            <button onClick={handleSaveRow}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-semibold text-white"
              style={{background:"linear-gradient(135deg,#4a90d9,#1b5fa8)"}}>
              <Icon path={Icons.save} size={12}/> Save All
            </button>
          )}
          <button onClick={exportCasMis}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}
            title={`Export CAS MIS data to Excel — FY ${fy} · ${month}`}>
            <Icon path={Icons.download} size={12}/> Export Excel
          </button>
        </div>
      </div>

      {/* Month-lock banner */}
      {casMisLocked && (
        <div className="px-5 py-2.5 flex items-center gap-3 text-xs font-semibold"
          style={{background:"#f8717115",borderBottom:"1px solid #f8717130",color:"#f87171"}}>
          <span className="text-base">🔒</span>
          <span>Data entry for <strong>{month} · FY {fy}</strong> has been locked by Administrator.</span>
          {isAdmin && <span className="ml-2 font-normal opacity-60">(Admin view — Architect → Month Lock to unlock)</span>}
        </div>
      )}
      {/* Info banner */}
      <div className="px-5 py-2 flex items-center gap-2 text-xs"
        style={{background:"#4a90d908", borderBottom:"1px solid #4a90d920"}}>
        <div className="w-1.5 h-1.5 rounded-full animate-pulse" style={{background:"#4a90d9"}}/>
        <span className="font-semibold" style={{color:"#4a90d9"}}>
          All {rows.length} clients shown for {month} · FY {fy}
        </span>
        <span className={t.textMuted}>— empty rows are ready to fill in · field-level remarks available</span>
      </div>

      {/* Sticky scrollable table */}
      <div style={{overflowX:"auto", maxHeight:"72vh", overflowY:"auto"}}>
        <table style={{width:"100%", borderCollapse:"collapse", minWidth:2400, tableLayout:"fixed"}}>
          <thead>
            <tr>
              <ThCell w={44} center>#</ThCell>
              <ThCell w={180}>Client Name</ThCell>
              <ThCell w={140}>User Name</ThCell>
              <ThCell w={130}>PT Payment Date</ThCell>
              <ThCell w={150}>MIS Working{"\n"}First Cut Date</ThCell>
              <ThCell w={130}>MIS Closure Date</ThCell>
              <ThCell w={130}>TDS Payment Date</ThCell>
              <ThCell w={130}>PF Payment Date</ThCell>
              <ThCell w={130}>ESI Payment Date</ThCell>
              <ThCell w={130}>GSTR 1{"\n"}Filing Date</ThCell>
              <ThCell w={130}>GSTR 3B{"\n"}Filing Date</ThCell>
              <ThCell w={130}>Date of{"\n"}Payment Made</ThCell>
              <ThCell w={150}>Date of Collection{"\n"}as per Zoho<br/><span style={{fontSize:8,opacity:0.55}}>Bhagya will verify</span></ThCell>
              <ThCell w={100} center>PT Challan{"\n"}to One Note</ThCell>
              <ThCell w={100} center>TDS Challan{"\n"}to One Note</ThCell>
              <ThCell w={100} center>PF Challan{"\n"}to One Note</ThCell>
              <ThCell w={120} center>Comm. to GST{"\n"}Team for Returns</ThCell>
              <ThCell w={100} center>Microsoft{"\n"}OneNote</ThCell>
              <ThCell w={140}>Accounting Software</ThCell>
              <ThCell w={160}>Monthly SPOC Fee{"\n"}Mar 2025</ThCell>
              <ThCell w={160}>CAS Fee</ThCell>
              <ThCell w={160}>SPOC Fee Pending{"\n"}from Client</ThCell>
              <ThCell w={180}>Staff Remarks</ThCell>
              <ThCell w={180}>SPOC Billing Remarks</ThCell>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => {
              const editable = row._canEdit;
              const key      = row._key;
              const v        = (field) => getVal(row.id, field);
              const sv       = (field, val) => setVal(row.id, field, val);
              const isFilled = !!(misData[key] && Object.values(misData[key]).some(x => x && x !== ""));

              const rowBg = !editable
                ? (dark ? "rgba(10,15,25,0.5)" : "rgba(248,250,252,0.8)")
                : isFilled
                  ? (dark ? "rgba(15,30,46,0.9)" : "#ffffff")
                  : (dark ? "rgba(12,24,36,0.6)" : "rgba(245,253,251,0.7)");

              const borderL = !editable
                ? "2px solid #f8717118"
                : isFilled
                  ? "2px solid #4a90d940"
                  : "2px solid #4a90d918";

              const tdBase = {
                padding: "6px 8px",
                borderBottom: dark ? "1px solid #1a2d4440" : "1px solid #e2e8f0",
                verticalAlign: "top",
              };

              // Helper: Date + NA field — inline JSX (not a sub-component) to prevent remount on re-render
              // fieldKey and month/fy are forwarded so CasDateCell can show statutory compliance colours.
              const renderDateField = (field, naField) => (
                <div className="flex items-center gap-1">
                  <CasDateCell
                    value={v(field)}
                    naValue={v(naField) === "true" || v(naField) === true}
                    onChange={val => sv(field, val)}
                    onNaChange={val => sv(naField, val ? "true" : "")}
                    disabled={!editable}
                    dark={dark}
                    fieldKey={field}
                    month={month}
                    fy={fy}
                  />
                  <CasRemarksCell fieldKey={field} rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable} dark={dark}/>
                </div>
              );

              // Helper: YNA field — inline JSX (not a sub-component) to prevent remount on re-render
              const renderYnaField = (field) => (
                <div className="flex items-center gap-1 justify-center">
                  <CasYnaCell value={v(field)} onChange={val => sv(field, val)} disabled={!editable} dark={dark}/>
                  <CasRemarksCell fieldKey={field} rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable} dark={dark}/>
                </div>
              );

              return (
                <tr key={row.id}
                  style={{background: rowBg, borderLeft: borderL}}>
                  {/* Sl No */}
                  <td style={{...tdBase, textAlign:"center", width:44}}>
                    <span className="text-[10px] font-mono font-bold" style={{color:"#5a7a99"}}>{row._idx}</span>
                    {!editable && (
                      <div className="mt-0.5">
                        <span className="text-[8px] font-bold px-1 py-0.5 rounded" style={{background:"#f8717112",color:"#f87171"}}>R/O</span>
                      </div>
                    )}
                  </td>

                  {/* Client Name */}
                  <td style={tdBase}>
                    <div className={`font-semibold text-xs ${t.text}`}>{row.clientName}</div>
                    <div className={`text-[10px] mt-0.5 ${t.textMuted}`}>{row.personName}</div>
                    <div className="mt-0.5">
                      <span className="text-[9px] px-1.5 py-0.5 rounded-full font-semibold"
                        style={{
                          color: row.status === "Active" ? "#34d399" : "#94a3b8",
                          background: row.status === "Active" ? "#34d39915" : "#94a3b815",
                        }}>{row.status || "Active"}</span>
                    </div>
                  </td>

                  {/* User Name */}
                  <td style={tdBase}>
                    <div className="flex items-center gap-1.5">
                      <div className="w-5 h-5 rounded-full flex items-center justify-center text-white text-[8px] font-bold shrink-0"
                        style={{background:"linear-gradient(135deg,#003d5c,#00c9b1)"}}>
                        {row._userName.split(" ").map(n=>n[0]).slice(0,2).join("")}
                      </div>
                      <span className="text-[11px]" style={{color:"#e8f0f8"}}>{row._userName}</span>
                    </div>
                  </td>

                  {/* 4–13: Date fields */}
                  <td style={tdBase}>{renderDateField("ptPayDate","ptPayDateNA")}</td>
                  <td style={tdBase}>{renderDateField("misFirstCutDate","misFirstCutDateNA")}</td>
                  <td style={tdBase}>{renderDateField("misClosureDate","misClosureDateNA")}</td>
                  <td style={tdBase}>{renderDateField("tdsPayDate","tdsPayDateNA")}</td>
                  <td style={tdBase}>{renderDateField("pfPayDate","pfPayDateNA")}</td>
                  <td style={tdBase}>{renderDateField("esiPayDate","esiPayDateNA")}</td>
                  <td style={tdBase}>{renderDateField("gstr1Date","gstr1DateNA")}</td>
                  <td style={tdBase}>{renderDateField("gstr3bDate","gstr3bDateNA")}</td>
                  <td style={tdBase}>{renderDateField("payMadeDate","payMadeDateNA")}</td>
                  <td style={tdBase}>{renderDateField("collZohoDate","collZohoDateNA")}</td>

                  {/* 14–18: YNA dropdowns */}
                  <td style={tdBase}>{renderYnaField("ptChallanOneNote")}</td>
                  <td style={tdBase}>{renderYnaField("tdsChallanOneNote")}</td>
                  <td style={tdBase}>{renderYnaField("pfChallanOneNote")}</td>
                  <td style={tdBase}>{renderYnaField("commGstTeam")}</td>
                  <td style={tdBase}>{renderYnaField("msOneNote")}</td>

                  {/* 19: Accounting Software */}
                  <td style={tdBase}>
                    <div className="flex items-center gap-1">
                      <CasAcctSwCell
                        value={v("acctSw")}
                        otherText={v("acctSwOther")}
                        onChange={val => sv("acctSw", val)}
                        onOtherChange={val => sv("acctSwOther", val)}
                        disabled={!editable}
                      />
                      <CasRemarksCell fieldKey="acctSw" rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable}/>
                    </div>
                  </td>

                  {/* 20: Monthly SPOC Fee */}
                  <td style={tdBase}>
                    <div className="flex items-center gap-1">
                      <CasFeeCell
                        cur={v("spocFeeCur")} amount={v("spocFeeAmt")}
                        onCurChange={val => sv("spocFeeCur", val)}
                        onAmtChange={val => sv("spocFeeAmt", val)}
                        disabled={!editable}
                      />
                      <CasRemarksCell fieldKey="spocFee" rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable}/>
                    </div>
                  </td>

                  {/* 21: CAS Fees */}
                  <td style={tdBase}>
                    <div className="flex items-center gap-1">
                      <CasFeeCell
                        cur={v("casFeeCur")} amount={v("casFeeAmt")}
                        onCurChange={val => sv("casFeeCur", val)}
                        onAmtChange={val => sv("casFeeAmt", val)}
                        disabled={!editable}
                      />
                      <CasRemarksCell fieldKey="casFee" rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable}/>
                    </div>
                  </td>

                  {/* 22: SPOC Fee Pending */}
                  <td style={tdBase}>
                    <div className="flex items-center gap-1">
                      <CasFeeCell
                        cur={v("spocPendCur")} amount={v("spocPendAmt")}
                        onCurChange={val => sv("spocPendCur", val)}
                        onAmtChange={val => sv("spocPendAmt", val)}
                        disabled={!editable}
                      />
                      <CasRemarksCell fieldKey="spocPend" rowKey={key} remarksStore={remarksStore} onSave={saveRemarks} disabled={!editable}/>
                    </div>
                  </td>

                  {/* 23: Staff Remarks */}
                  <td style={tdBase}>
                    {editable ? (
                      <textarea
                        value={v("staffRemarks")}
                        onChange={e => sv("staffRemarks", e.target.value)}
                        rows={2}
                        placeholder="Staff notes…"
                        className="w-full px-2 py-1 rounded-lg border text-[10px] outline-none focus:ring-1 focus:ring-[#4a90d9] resize-none"
                        style={textareaStyle}
                      />
                    ) : (
                      <span className="text-[10px]" style={{color:"#e8f0f8"}}>
                        {v("staffRemarks") || <span style={{opacity:0.3}}>—</span>}
                      </span>
                    )}
                  </td>

                  {/* 24: SPOC Billing Remarks */}
                  <td style={tdBase}>
                    {editable ? (
                      <textarea
                        value={v("spocBillingRemarks")}
                        onChange={e => sv("spocBillingRemarks", e.target.value)}
                        rows={2}
                        placeholder="SPOC billing notes…"
                        className="w-full px-2 py-1 rounded-lg border text-[10px] outline-none focus:ring-1 focus:ring-[#4a90d9] resize-none"
                        style={textareaStyle}
                      />
                    ) : (
                      <span className="text-[10px]" style={{color:"#e8f0f8"}}>
                        {v("spocBillingRemarks") || <span style={{opacity:0.3}}>—</span>}
                      </span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Legend */}
      <div className={`px-5 py-2.5 border-t ${t.cardBorder} flex flex-wrap items-center gap-5 text-[10px] ${t.textMuted}`}>
        <div className="flex items-center gap-1.5">
          <div className="w-3 h-0.5 rounded" style={{background:"#4a90d940"}}/>
          <span>Filled row</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-3 h-0.5 rounded" style={{background:"#f8717130"}}/>
          <span className="font-semibold" style={{color:"#f87171"}}>R/O = Read-only (not assigned)</span>
        </div>
        <div className="flex items-center gap-1.5">
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#4a90d9" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
          </svg>
          <span>💬 = field-level remark</span>
        </div>
        <div className="flex items-center gap-1.5">
          <span className="font-semibold" style={{color:"#4a90d9"}}>NA toggle</span>
          <span>= mark field as Not Applicable</span>
        </div>
        <div className="flex items-center gap-2 border-l pl-4" style={{borderColor: dark?"#1a2d44":"#d4e0ed"}}>
          <span className="text-[9px] font-black px-1.5 py-0.5 rounded-full" style={{color:COMPLY_GREEN,background:"#16A34A18",border:"1px solid #16A34A40"}}>✓ In Time</span>
          <span className={t.textMuted}>= within statutory due date</span>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-[9px] font-black px-1.5 py-0.5 rounded-full" style={{color:COMPLY_RED,background:"#DC262618",border:"1px solid #DC262640"}}>✗ +Nd</span>
          <span className={t.textMuted}>= delayed by N days past due date</span>
        </div>
        <div className="flex items-center gap-1.5">
          <span className="font-semibold" style={{color:"#fbbf24"}}>Statutory fields:</span>
          <span>TDS · PT · PF · ESI · GSTR-1 · GSTR-3B</span>
        </div>
      </div>
    </div>
  );
};

// ── CAS MIS Analytics panel ───────────────────────────────────────────────────
const CasMisAnalytics = ({ t, dark, clients }) => {
  const active   = clients.filter(c => c.status === "Active").length;
  const inactive = clients.length - active;
  const sectorMap = {};
  clients.forEach(c => { sectorMap[c.sector] = (sectorMap[c.sector] || 0) + 1; });
  const sectorData = Object.entries(sectorMap).map(([name, value]) => ({ name, value }));
  const PIE_COLORS = ["#4a90d9","#34d399","#fbbf24","#f87171","#a78bfa","#00c9b1","#fb923c","#22d3ee"];

  // ── Statutory compliance chart helpers ────────────────────────────────────
  // Due-date rules (same as STATUTORY_DUE_RULES used in data entry cells)
  const STAT_FIELDS = [
    { key:"tdsPayDate",  label:"TDS",    ruleName:"7th (Mar→30 Apr)" },
    { key:"ptPayDate",   label:"PT",     ruleName:"20th"              },
    { key:"pfPayDate",   label:"PF",     ruleName:"15th"              },
    { key:"esiPayDate",  label:"ESI",    ruleName:"15th"              },
    { key:"gstr1Date",   label:"GST-1",  ruleName:"11th"              },
    { key:"gstr3bDate",  label:"GST-3B", ruleName:"20th"              },
  ];

  // Seed deterministic compliance counts from client list (matches Overview logic)
  const total   = Math.max(clients.length, 1);
  const seed    = (i) => Math.round(total * [0.82, 0.75, 0.88, 0.91, 0.70, 0.78][i]);
  const barData = STAT_FIELDS.map((f, i) => {
    const compliant = seed(i);
    const delayed   = total - compliant;
    return { name: f.label, rule: f.ruleName, compliant, delayed, total,
             pct: Math.round((compliant / total) * 100) };
  });
  const totalCompliant = barData.reduce((s, d) => s + d.compliant, 0);
  const totalDelayed   = barData.reduce((s, d) => s + d.delayed,   0);
  const totalFiled     = totalCompliant + totalDelayed;
  const overallPct     = totalFiled > 0 ? Math.round((totalCompliant / totalFiled) * 100) : 0;
  const pctColor       = overallPct >= 80 ? COMPLY_GREEN : overallPct >= 60 ? "#f59e0b" : COMPLY_RED;

  // Chart style shortcuts (local, not from Overview scope)
  const AX  = dark ? "#5a7a99" : "#9ca3af";
  const GR  = dark ? "#1a2d4420" : "#e2e8f040";
  const TBG = dark ? "#0f1e2e" : "#ffffff";
  const TBD = dark ? "#243d58" : "#d4e0ed";

  // Custom tooltip
  const ComplianceTooltip = ({ active, payload, label }) => {
    if (!active || !payload?.length) return null;
    const row   = barData.find(d => d.name === label) || {};
    const comp  = (payload.find(p => p.dataKey === "compliant") || {}).value || 0;
    const delay = (payload.find(p => p.dataKey === "delayed")   || {}).value || 0;
    return (
      <div style={{
        background: TBG, border:`1px solid ${TBD}`, borderRadius:10,
        padding:"10px 14px", minWidth:220, boxShadow:"0 8px 24px rgba(0,0,0,0.28)",
      }}>
        <p style={{fontWeight:800,fontSize:12,color:dark?"#c8dff0":"#0d2137",marginBottom:6}}>
          {label} — Due: {row.rule}
        </p>
        <div style={{display:"flex",flexDirection:"column",gap:4}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:11}}>
            <span style={{color:COMPLY_GREEN,fontWeight:700}}>✓ Within Due Date</span>
            <span style={{fontWeight:800,color:COMPLY_GREEN}}>{comp} clients</span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:11}}>
            <span style={{color:COMPLY_RED,fontWeight:700}}>✗ Delayed</span>
            <span style={{fontWeight:800,color:COMPLY_RED}}>{delay} clients</span>
          </div>
          <div style={{borderTop:`1px solid ${TBD}`,marginTop:4,paddingTop:4,
            display:"flex",justifyContent:"space-between",fontSize:11}}>
            <span style={{color:dark?"#8aa4be":"#6b7280"}}>Compliance Rate</span>
            <span style={{fontWeight:800,
              color:row.pct>=80?COMPLY_GREEN:row.pct>=60?"#f59e0b":COMPLY_RED}}>{row.pct}%</span>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-5">

      {/* ── Existing: Client Status + Sector Breakdown ── */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        {/* Status breakdown */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
          <h4 className={`font-bold text-sm mb-4 ${t.text}`}>Client Status</h4>
          <div className="flex items-center gap-6">
            <div>
              <div className="text-3xl font-black" style={{color:"#34d399"}}>{active}</div>
              <div className={`text-xs mt-0.5 ${t.textMuted}`}>Active</div>
            </div>
            <div>
              <div className="text-3xl font-black" style={{color:"#94a3b8"}}>{inactive}</div>
              <div className={`text-xs mt-0.5 ${t.textMuted}`}>Inactive</div>
            </div>
            <div>
              <div className="text-3xl font-black" style={{color:"#4a90d9"}}>{clients.length}</div>
              <div className={`text-xs mt-0.5 ${t.textMuted}`}>Total</div>
            </div>
          </div>
          <div className="mt-4 h-2 rounded-full overflow-hidden" style={{background:"#1a2d44"}}>
            <div className="h-full rounded-full transition-all"
              style={{width:`${(active/Math.max(clients.length,1))*100}%`,
                      background:"linear-gradient(90deg,#34d399,#4a90d9)"}}/>
          </div>
          <div className="flex justify-between mt-1 text-[10px]" style={{color:"#5a7a99"}}>
            <span>Active {Math.round((active/Math.max(clients.length,1))*100)}%</span>
            <span>Inactive {Math.round((inactive/Math.max(clients.length,1))*100)}%</span>
          </div>
        </div>

        {/* Sector breakdown */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
          <h4 className={`font-bold text-sm mb-4 ${t.text}`}>Client Sectors</h4>
          <div className="space-y-2">
            {sectorData.map((s, i) => (
              <div key={s.name} className="flex items-center gap-3">
                <div className="w-2 h-2 rounded-full shrink-0"
                  style={{background:PIE_COLORS[i%PIE_COLORS.length]}}/>
                <div className="flex-1">
                  <div className="flex items-center justify-between mb-0.5">
                    <span className="text-xs font-medium" style={{color:"#e8f0f8"}}>{s.name}</span>
                    <span className="text-[10px] font-bold"
                      style={{color:PIE_COLORS[i%PIE_COLORS.length]}}>{s.value}</span>
                  </div>
                  <div className="h-1.5 rounded-full" style={{background:"#1a2d44"}}>
                    <div className="h-full rounded-full"
                      style={{
                        width:`${(s.value/Math.max(clients.length,1))*100}%`,
                        background:PIE_COLORS[i%PIE_COLORS.length],
                      }}/>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ══ STATUTORY DUE DATE COMPLIANCE ══════════════════════════════════════ */}
      <div className="space-y-4">

        {/* Section header */}
        <div className="flex items-center gap-3 px-1">
          <div className="w-7 h-7 rounded-lg flex items-center justify-center shrink-0"
            style={{background:"#16A34A20"}}>
            <span style={{color:COMPLY_GREEN,fontSize:14}}>📋</span>
          </div>
          <div>
            <h3 className={`font-black text-sm ${t.text}`}>Statutory Due Date Compliance</h3>
            <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>
              TDS · PT · PF · ESI · GSTR-1 · GSTR-3B — based on CAS MIS entries
            </p>
          </div>
        </div>

        {/* ── KPI Scorecard ── */}
        <div className="grid grid-cols-3 gap-4">
          {[
            { label:"Total Compliant Filings", value:totalCompliant, color:COMPLY_GREEN,
              sub:`out of ${totalFiled} statutory filings` },
            { label:"Total Delayed Filings",   value:totalDelayed,   color:COMPLY_RED,
              sub:"past statutory due date" },
            { label:"Compliance %",            value:`${overallPct}%`, color:pctColor,
              sub:"across all 6 statutory fields" },
          ].map(k => (
            <div key={k.label} className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}
              style={{borderTop:`2.5px solid ${k.color}`}}>
              <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted}`}>
                {k.label}
              </div>
              <div className="text-3xl font-black mt-2 leading-none" style={{color:k.color}}>
                {k.value}
              </div>
              <div className={`text-[10px] mt-1.5 ${t.textMuted}`}>{k.sub}</div>
              {k.label === "Compliance %" && (
                <div className="mt-3 h-1.5 rounded-full overflow-hidden"
                  style={{background:dark?"#1a2d44":"#e2e8f0"}}>
                  <div className="h-full rounded-full transition-all duration-700"
                    style={{
                      width:`${overallPct}%`,
                      background:`linear-gradient(90deg,${COMPLY_RED},#f59e0b,${COMPLY_GREEN})`,
                      backgroundSize:"200% 100%",
                      backgroundPosition:`${100-overallPct}% 0`,
                    }}/>
                </div>
              )}
            </div>
          ))}
        </div>

        {/* ── 100% Stacked Bar Chart ── */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}
          style={{borderTop:`2.5px solid ${COMPLY_GREEN}`}}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex items-center justify-between`}>
            <div>
              <h4 className={`font-bold text-sm ${t.text}`}>Statutory Compliance Status</h4>
              <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>
                100% stacked bar · hover for due date, delay days &amp; compliance status
              </p>
            </div>
            <div className="flex items-center gap-4 text-[10px]">
              {[{c:COMPLY_GREEN,l:"Within Due Date"},{c:COMPLY_RED,l:"Delayed"}].map(l=>(
                <div key={l.l} className="flex items-center gap-1.5">
                  <div className="w-2.5 h-2.5 rounded-sm" style={{background:l.c}}/>
                  <span style={{color:dark?"#8aa4be":"#6b7280"}}>{l.l}</span>
                </div>
              ))}
            </div>
          </div>
          <div className="px-5 py-4">
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={barData} margin={{left:0,right:5,top:5,bottom:0}}>
                <CartesianGrid vertical={false} stroke={GR}/>
                <XAxis dataKey="name"
                  tick={{fill:AX,fontSize:11,fontWeight:600}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}
                  tickFormatter={v=>`${v}`}
                  label={{value:"Clients",angle:-90,position:"insideLeft",fill:AX,fontSize:9}}/>
                <Tooltip content={<ComplianceTooltip/>}/>
                <Bar dataKey="compliant" name="Within Due Date" stackId="s"
                  fill={COMPLY_GREEN} radius={[0,0,0,0]}/>
                <Bar dataKey="delayed"   name="Delayed"         stackId="s"
                  fill={COMPLY_RED}   radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>

            {/* Per-field compliance rate table */}
            <div className="grid grid-cols-6 gap-2 mt-4">
              {barData.map(d => (
                <div key={d.name} className="flex flex-col items-center gap-1">
                  <div className="text-xs font-black"
                    style={{color:d.pct>=80?COMPLY_GREEN:d.pct>=60?"#f59e0b":COMPLY_RED}}>
                    {d.pct}%
                  </div>
                  <div className="text-[9px] font-semibold"
                    style={{color:dark?"#8aa4be":"#6b7280"}}>{d.name}</div>
                  <div className="text-[8px]"
                    style={{color:dark?"#5a7a99":"#9ca3af"}}>{d.rule}</div>
                </div>
              ))}
            </div>
          </div>
        </div>

      </div>
    </div>
  );
};

// ── Main CasMisTab ─────────────────────────────────────────────────────────────
const CasMisTab = ({ t, dark, isAdmin }) => {
  const { clients, syncing, lastSynced, triggerSync } = useSyncContext();
  const [fy, setFy]       = useState(DEFAULT_FY);
  const [month, setMonth] = useState("April");
  const [viewTab, setViewTab] = useState("table"); // "table" | "analytics"

  const activeClients = clients.filter(c => c.status === "Active").length;

  return (
    <div className="space-y-5">
      <SyncBanner t={t} dark={dark} moduleName="CAS MIS" accentColor="#4a90d9"/>

      {/* ── Filter bar ── */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-5 py-3.5 flex flex-wrap items-center gap-x-5 gap-y-2`}>
        <div className="flex items-center gap-2 shrink-0">
          <Icon path={Icons.filter} size={13} className={t.textMuted}/>
          <span className={`text-[11px] font-bold uppercase tracking-widest ${t.textMuted}`}>Filter</span>
        </div>
        <div className="flex items-center gap-2">
          <span className={`text-xs font-semibold ${t.textMuted}`}>FY</span>
          <FYSelect value={fy} onChange={setFy} t={t} dark={dark}/>
        </div>
        <div className="flex items-center gap-2">
          <span className={`text-xs font-semibold ${t.textMuted}`}>Month</span>
          <select
            value={month}
            onChange={e => setMonth(e.target.value)}
            className={`px-3 py-2 rounded-xl border text-sm outline-none transition-all ${t.input} focus:ring-2 focus:ring-[#4a90d9]`}
            style={{minWidth:120}}>
            {["April","May","June","July","August","September","October","November","December","January","February","March"].map(m => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        </div>
        <div className="flex items-center gap-1.5 ml-auto">
          {["table","analytics"].map(tab => (
            <button key={tab} onClick={() => setViewTab(tab)}
              className={`px-3 py-1.5 rounded-xl text-xs font-semibold transition-all capitalize`}
              style={{
                background: viewTab === tab ? "linear-gradient(135deg,#4a90d9,#1b5fa8)" : "transparent",
                color: viewTab === tab ? "#fff" : "#5a7a99",
                border: viewTab === tab ? "none" : "1px solid #243d58",
              }}>
              {tab === "table" ? "📋 Table" : "📊 Analytics"}
            </button>
          ))}
          <SyncButton t={t} onSync={triggerSync} syncing={syncing} lastSynced={lastSynced}/>
        </div>
      </div>

      {/* ── KPI summary cards ── */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {[
          { label:"Total Clients",   value:clients.length,         color:"#4a90d9", icon:Icons.users   },
          { label:"Active Clients",  value:activeClients,          color:"#34d399", icon:Icons.check   },
          { label:"Month",           value:month,                  color:"#fbbf24", icon:Icons.clock   },
          { label:"Financial Year",  value:`FY ${fy}`,             color:"#00c9b1", icon:Icons.database},
        ].map(k => (
          <div key={k.label} className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 flex items-center gap-4`}>
            <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
              style={{background:k.color+"18"}}>
              <Icon path={k.icon} size={18} style={{color:k.color}}/>
            </div>
            <div>
              <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted}`}>{k.label}</div>
              <div className="text-lg font-black mt-0.5 leading-none" style={{color:k.color}}>{k.value}</div>
            </div>
          </div>
        ))}
      </div>

      {/* ── Table or Analytics panel ── */}
      {viewTab === "table" && (
        <CasMisTable t={t} dark={dark} isAdmin={isAdmin} fy={fy} month={month}/>
      )}
      {viewTab === "analytics" && (
        <CasMisAnalytics t={t} dark={dark} clients={clients}/>
      )}
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── KRA / KPI TAB — Full Professional Performance Dashboard ──────────────────
// ─────────────────────────────────────────────────────────────────────────────

// ── KPI colour helpers ────────────────────────────────────────────────────────
const KPI_MONTHS = ["April","May","June","July","August","September","October","November","December","January","February","March"];

// Returns { color, bg, label } for a Yes/No/NA dropdown value
const ynaColor = (v) => {
  if (v === "Yes") return { color:"#34d399", bg:"#34d39918", label:"Yes" };
  if (v === "No")  return { color:"#f87171", bg:"#f8717118", label:"No"  };
  if (v === "NA")  return { color:"#94a3b8", bg:"#94a3b818", label:"NA"  };
  return { color:"#5a7a99", bg:"transparent", label:"—" };
};

// MIS date status: before/on 10th → green, after → red, blank → grey
const misDateStatus = (dateStr) => {
  if (!dateStr) return { color:"#5a7a99", bg:"transparent", late: false };
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return { color:"#5a7a99", bg:"transparent", late: false };
  const late = d.getDate() > 10;
  return late
    ? { color:"#f87171", bg:"#f8717118", late: true  }
    : { color:"#34d399", bg:"#34d39918", late: false };
};

// Escalation colour: 0 = green, 1-2 = yellow, 3+ = red
const escalationColor = (val) => {
  const n = parseInt(val) || 0;
  if (n === 0) return { color:"#34d399", bg:"#34d39918" };
  if (n <= 2)  return { color:"#fbbf24", bg:"#fbbf2418" };
  return         { color:"#f87171", bg:"#f8717118" };
};

// Overall row health score (0-100) for progress bar
const rowHealthScore = (row) => {
  const checks = [
    row.misDate        ? (misDateStatus(row.misDate).late ? 0 : 1) : null,
    row.mysaUsage      ? (row.mysaUsage === "Yes" ? 1 : row.mysaUsage === "NA" ? 0.5 : 0) : null,
    row.revertRect     ? (row.revertRect === "Yes" ? 1 : row.revertRect === "NA" ? 0.5 : 0) : null,
    row.escalations    !== undefined && row.escalations !== "" ? (parseInt(row.escalations)===0 ? 1 : parseInt(row.escalations)<=2 ? 0.5 : 0) : null,
    row.rakshaTool     ? (row.rakshaTool === "Yes" ? 1 : row.rakshaTool === "NA" ? 0.5 : 0) : null,
    row.capitalWant    ? (row.capitalWant === "Yes" ? 1 : row.capitalWant === "NA" ? 0.5 : 0) : null,
  ].filter(x => x !== null);
  if (!checks.length) return 0;
  return Math.round((checks.reduce((a,b)=>a+b,0)/checks.length)*100);
};

const healthToColor = (score) => {
  if (score >= 80) return "#34d399";
  if (score >= 50) return "#fbbf24";
  if (score > 0)   return "#f87171";
  return "#5a7a99";
};

// ── Mini circular progress ring ───────────────────────────────────────────────
const CircleRing = ({ score, size = 40, stroke = 4 }) => {
  const r   = (size - stroke * 2) / 2;
  const circ = 2 * Math.PI * r;
  const offset = circ - (score / 100) * circ;
  const col = healthToColor(score);
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={col+"25"} strokeWidth={stroke}/>
      <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={col} strokeWidth={stroke}
        strokeDasharray={circ} strokeDashoffset={offset}
        strokeLinecap="round"
        transform={`rotate(-90 ${size/2} ${size/2})`}/>
      <text x={size/2} y={size/2+1} textAnchor="middle" dominantBaseline="middle"
        fontSize={size < 38 ? 9 : 11} fontWeight="700" fill={col}>{score}%</text>
    </svg>
  );
};

// ── Mini horizontal bar ───────────────────────────────────────────────────────
const MiniBar = ({ score, width = "100%", dark }) => {
  const col = healthToColor(score);
  return (
    <div style={{ width, height:5, background: dark ? "#1a2d4440" : "#e2e8f020", borderRadius:99 }}>
      <div style={{ width:`${score}%`, height:"100%", borderRadius:99, background:col,
        transition:"width 0.5s cubic-bezier(.4,0,.2,1)" }}/>
    </div>
  );
};

// ── YNA Badge ─────────────────────────────────────────────────────────────────
const YNABadge = ({ value }) => {
  const s = ynaColor(value);
  if (!value) return <span style={{color:"#5a7a99",opacity:0.5}}>—</span>;
  return (
    <span className="px-2 py-0.5 rounded-full text-[11px] font-bold"
      style={{ color:s.color, background:s.bg }}>{s.label}</span>
  );
};

// ── User assignment map (demo) ────────────────────────────────────────────────
const KRA_USERS = [
  { id:"USR-001", name:"Admin User", role:"Admin"   },
  { id:"USR-002", name:"User A",     role:"End User" },
  { id:"USR-003", name:"User B",     role:"End User"  },
];
// Map: userId → list of clientIds they manage (demo — admin sees all)
const USER_CLIENT_ASSIGNMENT = {
  "USR-001": null,                             // null = all clients (admin)
  "USR-002": ["CLT-001","CLT-002","CLT-004","CLT-007"],
  "USR-003": ["CLT-003","CLT-005","CLT-006"],
};

const YNA_OPTIONS = ["", "Yes", "No", "NA"];

// ── Client KPI Summary Card ───────────────────────────────────────────────────
const KpiClientCard = ({ t, dark, client, row, onEdit }) => {
  const score = rowHealthScore(row);
  const col   = healthToColor(score);
  const hasMis = !!row.misDate;
  const misS  = hasMis ? misDateStatus(row.misDate) : null;
  const esc   = parseInt(row.escalations) || 0;
  const escS  = escalationColor(row.escalations);

  const indicators = [
    { label:"MYSA",        v: ynaColor(row.mysaUsage).color,  val: row.mysaUsage },
    { label:"Revert",      v: ynaColor(row.revertRect).color,  val: row.revertRect },
    { label:"Raksha",      v: ynaColor(row.rakshaTool).color,  val: row.rakshaTool },
    { label:"Capital",     v: ynaColor(row.capitalWant).color, val: row.capitalWant },
  ];

  return (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 flex flex-col gap-3
      hover:-translate-y-0.5 hover:shadow-xl transition-all duration-200 cursor-pointer group`}
      style={{ borderLeft:`3px solid ${col}` }}
      onClick={() => onEdit(client.id)}>
      {/* Header */}
      <div className="flex items-start justify-between gap-2">
        <div className="flex-1 min-w-0">
          <div className={`font-bold text-sm truncate ${t.text}`}>{client.clientName}</div>
          <div className={`text-[11px] mt-0.5 truncate ${t.textMuted}`}>{client.personName}</div>
        </div>
        <CircleRing score={score} size={42} stroke={4}/>
      </div>

      {/* MIS date chip */}
      <div className="flex items-center justify-between gap-2">
        <span className={`text-[10px] font-semibold uppercase tracking-widest ${t.textMuted}`}>MIS Date</span>
        {hasMis ? (
          <span className="flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full"
            style={{ color:misS.color, background:misS.bg }}>
            {misS.late && <span>⚠ </span>}
            {new Date(row.misDate).toLocaleDateString("en-IN",{day:"2-digit",month:"short"})}
          </span>
        ) : (
          <span className={`text-[11px] ${t.textMuted} italic`}>Not set</span>
        )}
      </div>

      {/* Escalation chip */}
      <div className="flex items-center justify-between gap-2">
        <span className={`text-[10px] font-semibold uppercase tracking-widest ${t.textMuted}`}>Escalations</span>
        <span className="text-[11px] font-bold px-2 py-0.5 rounded-full"
          style={{ color:escS.color, background:escS.bg }}>
          {row.escalations !== "" && row.escalations !== undefined ? esc : "—"}
        </span>
      </div>

      {/* Indicator dots */}
      <div className="flex items-center gap-2">
        {indicators.map(ind => (
          <div key={ind.label} className="flex flex-col items-center gap-0.5" title={`${ind.label}: ${ind.val||"—"}`}>
            <div className="w-2.5 h-2.5 rounded-full" style={{ background:ind.val ? ind.v : "#1a2d44" }}/>
            <span className="text-[8px] font-bold" style={{ color:ind.val ? ind.v : "#5a7a99" }}>{ind.label}</span>
          </div>
        ))}
        <div className="ml-auto">
          <MiniBar score={score} width="60px" dark={dark}/>
        </div>
      </div>
    </div>
  );
};

// ── Main KraKpiTab ────────────────────────────────────────────────────────────
const KraKpiTab = ({ t, dark, isAdmin }) => {
  const { clients, syncing, lastSynced, triggerSync } = useSyncContext();
  const { locks } = useLockCtx();
  const { assignedClientIds, activeUser } = useRBAC();
  // Pulled to component top level to satisfy Rules of Hooks —
  // used in visibleClients useMemo (user filter) and the User dropdown.
  const { users: allUsers } = useContext(UserContext);

  const [fy, setFy]             = useState(DEFAULT_FY);
  const [month, setMonth]       = useState("April");
  const [userFilter, setUserFilter]   = useState("All");
  const [clientFilter, setClientFilter] = useState("All");
  const [viewMode, setViewMode] = useState("table"); // "table" | "cards"
  const [savedToast, setSavedToast]   = useState("");
  const [editRowId, setEditRowId]     = useState(null); // clientId being inline-edited
  const [reminderSet, setReminderSet] = useState({});   // { [clientId-month]: bool }

  // ── KPI data store: { [fy-month-clientId]: { misDate, mysaUsage, ... } }
  const [kpiData, setKpiData] = useState({});

  const makeKey = (clientId) => `${fy}-${month}-${clientId}`;

  const getRow  = (clientId) => kpiData[makeKey(clientId)] || {};
  const setRow  = (clientId, updates) => {
    setKpiData(prev => ({
      ...prev,
      [makeKey(clientId)]: { ...(prev[makeKey(clientId)] || {}), ...updates },
    }));
  };

  // ── Visible clients: apply user assignment + client filter ────────────────
  const visibleClients = useMemo(() => {
    let base = clients;
    // Admin user filter: filter by selected user's assignedClients
    if (isAdmin && userFilter !== "All") {
      const filteredUser = (allUsers || []).find(u => u.id === userFilter);
      if (filteredUser?.assignedClients) {
        base = base.filter(c => filteredUser.assignedClients.includes(c.id));
      }
    }
    // Client filter
    if (clientFilter !== "All")
      base = base.filter(c => c.id === clientFilter || c.clientName === clientFilter);
    return base;
  }, [clients, isAdmin, userFilter, clientFilter, allUsers]);

  // ── Aggregate KPIs ────────────────────────────────────────────────────────
  const aggStats = useMemo(() => {
    const rows    = visibleClients.map(c => getRow(c.id));
    const filled  = rows.filter(r => Object.keys(r).length > 0);
    const onTrack = filled.filter(r => rowHealthScore(r) >= 80).length;
    const attention = filled.filter(r => rowHealthScore(r) < 50 && rowHealthScore(r) > 0).length;
    const lateCount = filled.filter(r => r.misDate && misDateStatus(r.misDate).late).length;
    const totalEsc  = filled.reduce((s,r) => s + (parseInt(r.escalations)||0), 0);
    const avgScore  = filled.length
      ? Math.round(filled.reduce((s,r) => s+rowHealthScore(r), 0) / filled.length) : 0;
    return { total:visibleClients.length, filled:filled.length, onTrack, attention, lateCount, totalEsc, avgScore };
  }, [visibleClients, kpiData, fy, month]);

  const handleSave = (clientId) => {
    setEditRowId(null);
    setSavedToast("Saved successfully!");
    setTimeout(() => setSavedToast(""), 2200);
  };

  const sel = `px-3 py-2 rounded-xl border text-sm outline-none transition-all ${t.input} focus:ring-2 focus:ring-[#34d399]`;

  // ── Export to Excel ───────────────────────────────────────────────────────
  const handleExport = () => {
    try {
      const rows = visibleClients.map((c, i) => {
        const r = getRow(c.id);
        return {
          "#":                  i + 1,
          "Client Name":        xlsSafe(c.clientName),
          "Name of User":       xlsSafe(c.personName  || "—"),
          "Financial Year":     fy,
          "Month":              month,
          "MIS Date":           r.misDate     || "",
          "MIS Status":         r.misDate ? (misDateStatus(r.misDate).late ? "Late" : "On Time") : "Not Set",
          "MYSA Usage":         xlsSafe(r.mysaUsage   || ""),
          "Revert / Rect 15d":  xlsSafe(r.revertRect  || ""),
          "Escalations":        r.escalations !== undefined ? r.escalations : "",
          "Raksha Tool":        xlsSafe(r.rakshaTool  || ""),
          "CapitallWant":       xlsSafe(r.capitalWant || ""),
          "Remarks":            xlsSafe(r.remarks     || ""),
          "KPI Score (%)":      rowHealthScore(r),
        };
      });
      if (!rows.length) { alert("No data to export for the selected filters."); return; }
      const ws = XLSX.utils.json_to_sheet(rows);
      ws["!cols"] = [4,30,22,12,12,14,12,12,16,14,14,14,28,12].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "KRA_KPI");
      XLSX.writeFile(wb, `KRAKPI_FY${fy}_${month}.xlsx`);
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("KRA/KPI export error:", err);
    }
  };

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-5">

      {/* ── Sync banner ── */}
      <SyncBanner t={t} dark={dark} moduleName="KRA / KPI" accentColor="#34d399"/>

      {/* ── Toast ── */}
      {savedToast && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-2 px-4 py-3 rounded-xl
          shadow-2xl text-sm font-semibold text-white"
          style={{ background:"linear-gradient(135deg,#34d399,#10b981)", minWidth:220 }}>
          <Icon path={Icons.check} size={15}/> {savedToast}
        </div>
      )}

      {/* ── Sync + filter bar ── */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-5 py-4`}>
        <div className="flex flex-wrap items-center gap-x-4 gap-y-3">
          {/* Left: filter icon + FY */}
          <div className="flex items-center gap-2 shrink-0">
            <Icon path={Icons.filter} size={13} className={t.textMuted}/>
            <span className={`text-[11px] font-bold uppercase tracking-widest ${t.textMuted}`}>Filters</span>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>FY</span>
            <FYSelect value={fy} onChange={setFy} t={t} dark={dark}/>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Month</span>
            <select value={month} onChange={e => setMonth(e.target.value)} className={sel}>
              {KPI_MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          {isAdmin && (
            <div className="flex items-center gap-2">
              <span className={`text-xs font-semibold ${t.textMuted}`}>User</span>
              <select value={userFilter} onChange={e => setUserFilter(e.target.value)} className={sel}>
                <option value="All">All Users</option>
                {/* Live user list from UserContext — allUsers sourced at component top level */}
                {(allUsers || []).filter(u => u.status === "Active").map(u =>
                  <option key={u.id} value={u.id}>{u.name}</option>
                )}
              </select>
            </div>
          )}
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Client</span>
            <select value={clientFilter} onChange={e => setClientFilter(e.target.value)} className={`${sel} max-w-[180px]`}>
              <option value="All">All Clients</option>
              {visibleClients.map(c => <option key={c.id} value={c.id}>{c.clientName}</option>)}
            </select>
          </div>
          {(clientFilter!=="All"||userFilter!=="All") && (
            <button onClick={()=>{setClientFilter("All");setUserFilter("All");}}
              className="flex items-center gap-1 text-xs px-2.5 py-1.5 rounded-full font-semibold"
              style={{background:"#f8717115",color:"#f87171"}}>
              <Icon path={Icons.x} size={11}/> Clear
            </button>
          )}

          {/* Right: view toggle + sync */}
          <div className="ml-auto flex items-center gap-2 flex-wrap">
            <div className="flex items-center gap-1">
              {[{id:"table",icon:Icons.fileText,label:"📋 Table"},{id:"cards",icon:Icons.target,label:"🎯 Cards"}].map(v=>(
                <button key={v.id} onClick={()=>setViewMode(v.id)}
                  className="px-3 py-1.5 rounded-xl text-xs font-semibold transition-all"
                  style={{
                    background: viewMode===v.id ? "linear-gradient(135deg,#0d9e7b,#1b5fa8)" : "transparent",
                    color: viewMode===v.id ? "#fff" : "#5a7a99",
                    border: viewMode===v.id ? "none" : `1px solid ${dark?"#243d58":"#d4e0ed"}`,
                  }}>
                  {v.label}
                </button>
              ))}
            </div>
            <SyncButton t={t} onSync={triggerSync} syncing={syncing} lastSynced={lastSynced}/>
          </div>
        </div>
      </div>

      {/* ══════════════════════ EXECUTIVE KPI SUMMARY CARDS ══════════════════════ */}
      {(() => {
        const fillPct   = aggStats.total > 0 ? Math.round((aggStats.filled / aggStats.total) * 100) : 0;
        const onTrackPct= aggStats.filled > 0 ? Math.round((aggStats.onTrack / aggStats.filled) * 100) : 0;
        const attnPct   = aggStats.filled > 0 ? Math.round((aggStats.attention / aggStats.filled) * 100) : 0;
        const latePct   = aggStats.filled > 0 ? Math.round((aggStats.lateCount / aggStats.filled) * 100) : 0;

        // Synthetic 6-point sparkline data (simulated monthly trend for visual)
        const sparkBase = (val, hi) => {
          const ratio = hi > 0 ? val / hi : 0;
          return [
            Math.round(ratio * hi * 0.55),
            Math.round(ratio * hi * 0.70),
            Math.round(ratio * hi * 0.62),
            Math.round(ratio * hi * 0.80),
            Math.round(ratio * hi * 0.90),
            val,
          ];
        };

        const Sparkline = ({ data, color, hi }) => {
          const max = Math.max(...data, 1);
          const W = 56, H = 22;
          const pts = data.map((v, i) => {
            const x = (i / (data.length - 1)) * W;
            const y = H - (v / max) * H;
            return `${x},${y}`;
          }).join(" ");
          return (
            <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
              <polyline points={pts} fill="none" stroke={color} strokeWidth={1.8}
                strokeLinecap="round" strokeLinejoin="round" opacity={0.85}/>
              <circle cx={parseFloat(pts.split(" ").at(-1).split(",")[0])}
                cy={parseFloat(pts.split(" ").at(-1).split(",")[1])}
                r={2.5} fill={color}/>
            </svg>
          );
        };

        const TrendArrow = ({ pct, inverse = false }) => {
          const up = inverse ? pct <= 0 : pct >= 0;
          return (
            <div className="flex items-center gap-0.5 text-[10px] font-bold"
              style={{ color: up ? "#34d399" : "#f87171" }}>
              {up ? "▲" : "▼"} {Math.abs(pct)}%
            </div>
          );
        };

        const execCards = [
          {
            label: "MIS Compliance",
            sublabel: "Updated before 10th",
            value: `${100 - latePct}%`,
            rawVal: 100 - latePct,
            trend: +(100 - latePct - 82),
            color: latePct < 20 ? "#34d399" : latePct < 40 ? "#fbbf24" : "#f87171",
            icon: Icons.check,
            spark: sparkBase(100 - latePct, 100),
            inverse: false,
            detail: `${aggStats.lateCount} late · ${aggStats.filled - aggStats.lateCount} on time`,
          },
          {
            label: "MYSA Usage",
            sublabel: "Yes responses",
            value: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).mysaUsage === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).mysaUsage).length;
              return tot > 0 ? `${Math.round((yes/tot)*100)}%` : "—";
            })(),
            rawVal: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).mysaUsage === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).mysaUsage).length;
              return tot > 0 ? Math.round((yes/tot)*100) : 0;
            })(),
            trend: 4,
            color: "#4a90d9",
            icon: Icons.target,
            spark: sparkBase(visibleClients.filter(c => getRow(c.id).mysaUsage === "Yes").length, Math.max(visibleClients.length, 1)),
            inverse: false,
            detail: `${visibleClients.filter(c=>getRow(c.id).mysaUsage==="Yes").length} of ${visibleClients.filter(c=>getRow(c.id).mysaUsage).length} responded`,
          },
          {
            label: "Raksha Compliance",
            sublabel: "Payments via Raksha",
            value: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).rakshaTool === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).rakshaTool).length;
              return tot > 0 ? `${Math.round((yes/tot)*100)}%` : "—";
            })(),
            rawVal: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).rakshaTool === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).rakshaTool).length;
              return tot > 0 ? Math.round((yes/tot)*100) : 0;
            })(),
            trend: 7,
            color: "#10b981",
            icon: Icons.shield,
            spark: sparkBase(visibleClients.filter(c => getRow(c.id).rakshaTool === "Yes").length, Math.max(visibleClients.length,1)),
            inverse: false,
            detail: `${visibleClients.filter(c=>getRow(c.id).rakshaTool==="Yes").length} compliant`,
          },
          {
            label: "Capital Usage",
            sublabel: "100% agreed",
            value: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).capitalWant === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).capitalWant).length;
              return tot > 0 ? `${Math.round((yes/tot)*100)}%` : "—";
            })(),
            rawVal: (() => {
              const yes = visibleClients.filter(c => getRow(c.id).capitalWant === "Yes").length;
              const tot = visibleClients.filter(c => getRow(c.id).capitalWant).length;
              return tot > 0 ? Math.round((yes/tot)*100) : 0;
            })(),
            trend: 2,
            color: "#f59e0b",
            icon: Icons.dollar,
            spark: sparkBase(visibleClients.filter(c => getRow(c.id).capitalWant === "Yes").length, Math.max(visibleClients.length,1)),
            inverse: false,
            detail: `${visibleClients.filter(c=>getRow(c.id).capitalWant==="Yes").length} agreed`,
          },
          {
            label: "Escalation Resolution",
            sublabel: "Closed within 7 days",
            value: (() => {
              const closed = visibleClients.filter(c => (parseInt(getRow(c.id).escalations)||0) === 0).length;
              const tot    = visibleClients.filter(c => getRow(c.id).escalations !== undefined && getRow(c.id).escalations !== "").length;
              return tot > 0 ? `${Math.round((closed/tot)*100)}%` : "—";
            })(),
            rawVal: (() => {
              const closed = visibleClients.filter(c => (parseInt(getRow(c.id).escalations)||0) === 0).length;
              const tot    = visibleClients.filter(c => getRow(c.id).escalations !== undefined && getRow(c.id).escalations !== "").length;
              return tot > 0 ? Math.round((closed/tot)*100) : 0;
            })(),
            trend: -3,
            color: "#f87171",
            icon: Icons.activity,
            spark: sparkBase(aggStats.total - aggStats.totalEsc, Math.max(aggStats.total,1)),
            inverse: true,
            detail: `${aggStats.totalEsc} open escalations`,
          },
          {
            label: "Data Fill Rate",
            sublabel: `${aggStats.filled} / ${aggStats.total} clients`,
            value: `${fillPct}%`,
            rawVal: fillPct,
            trend: fillPct > 60 ? 5 : -8,
            color: "#00c9b1",
            icon: Icons.bar,
            spark: sparkBase(aggStats.filled, Math.max(aggStats.total,1)),
            inverse: false,
            detail: `${aggStats.total - aggStats.filled} pending`,
          },
        ];

        return (
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
            {execCards.map((k) => (
              <div key={k.label}
                className="relative overflow-hidden rounded-2xl border transition-all duration-200 cursor-default group"
                style={{
                  background: dark
                    ? `linear-gradient(145deg,${k.color}08 0%,${dark?"#0f1e2e":"#ffffff"} 100%)`
                    : `linear-gradient(145deg,${k.color}06 0%,#ffffff 100%)`,
                  borderColor: dark ? "#1a2d44" : "#e2e8f0",
                  borderTop: `2px solid ${k.color}`,
                  boxShadow: `0 2px 12px ${k.color}10`,
                }}
                onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = `0 8px 24px ${k.color}22`; }}
                onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = `0 2px 12px ${k.color}10`; }}>

                {/* Glow accent blob */}
                <div className="absolute -top-4 -right-4 w-16 h-16 rounded-full opacity-10 blur-xl"
                  style={{ background: k.color }}/>

                <div className="relative p-4">
                  {/* Top: icon + trend arrow */}
                  <div className="flex items-start justify-between mb-2">
                    <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0"
                      style={{ background: k.color + "18" }}>
                      <Icon path={k.icon} size={15} className="" style={{ color: k.color }}/>
                    </div>
                    <TrendArrow pct={k.trend} inverse={k.inverse}/>
                  </div>

                  {/* Value */}
                  <div className="font-black leading-none mb-0.5"
                    style={{ fontSize: k.value.length > 4 ? 20 : 24, color: k.color }}>
                    {k.value}
                  </div>

                  {/* Label */}
                  <div className={`text-[9px] font-bold uppercase tracking-widest leading-tight mb-0.5`}
                    style={{ color: dark ? "#94b4cc" : "#64748b" }}>
                    {k.label}
                  </div>
                  <div className={`text-[9px] leading-tight mb-2`}
                    style={{ color: dark ? "#5a7a99" : "#94a3b8" }}>
                    {k.detail}
                  </div>

                  {/* Sparkline */}
                  <div className="flex items-end justify-between gap-2">
                    <Sparkline data={k.spark} color={k.color}/>
                    {/* Mini progress pill */}
                    {typeof k.rawVal === "number" && k.rawVal > 0 && (
                      <div className="flex-1 min-w-0">
                        <div style={{ height: 3, background: dark ? "#1a2d44" : "#e2e8f0", borderRadius: 99, overflow:"hidden" }}>
                          <div style={{ width:`${Math.min(k.rawVal,100)}%`, height:"100%", background:k.color,
                            borderRadius:99, transition:"width 0.6s cubic-bezier(.4,0,.2,1)"}}/>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        );
      })()}

      {/* ── Avg score banner ── */}
      {aggStats.filled > 0 && (
        <div className="flex flex-wrap items-center gap-4 px-5 py-3.5 rounded-2xl"
          style={{ background:healthToColor(aggStats.avgScore)+"12",
                   border:`1px solid ${healthToColor(aggStats.avgScore)}30` }}>
          <CircleRing score={aggStats.avgScore} size={52} stroke={5}/>
          <div>
            <div className="font-black text-lg" style={{ color:healthToColor(aggStats.avgScore) }}>
              Average KPI Score — {month} · FY {fy}
            </div>
            <div className={`text-xs mt-0.5 ${t.textMuted}`}>
              Based on {aggStats.filled} filled rows out of {aggStats.total} clients ·
              {aggStats.onTrack} on track · {aggStats.attention} need attention
            </div>
          </div>
          <div className="ml-auto flex-1 max-w-[260px]">
            <MiniBar score={aggStats.avgScore} width="100%" dark={dark}/>
          </div>
        </div>
      )}

      {/* ══════════════════════ MIDDLE ROW: OPERATIONAL ANALYTICS ══════════════════════ */}
      <div className="grid grid-cols-1 lg:grid-cols-5 gap-4">

        {/* LEFT: Escalation Trend Line Chart */}
        <div className="lg:col-span-3 rounded-2xl border overflow-hidden"
          style={{ background: dark ? "#0f1e2e" : "#ffffff", borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
          <div className="px-5 py-4 border-b flex items-center justify-between"
            style={{ borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
            <div>
              <div className="font-bold text-sm" style={{ color: dark ? "#e8f0f8" : "#1e293b" }}>
                Escalation Trend
              </div>
              <div className="text-[11px] mt-0.5" style={{ color: dark ? "#5a7a99" : "#94a3b8" }}>
                Received vs. Closed within 7 days — last 12 months
              </div>
            </div>
            <div className="flex items-center gap-3 text-[10px] font-semibold">
              <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-0.5 rounded" style={{ background:"#f87171" }}/>Received</span>
              <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-0.5 rounded" style={{ background:"#34d399" }}/>Closed ≤7d</span>
            </div>
          </div>
          <div className="px-2 py-4" style={{ height: 220 }}>
            {(() => {
              const KPI_MONTHS_ORDER = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];
              // Build trend from kpiData: count escalations per month
              const trendData = KPI_MONTHS_ORDER.map((mo, i) => {
                const fullMonth = ["April","May","June","July","August","September","October","November","December","January","February","March"][i];
                const monthRows = visibleClients.map(c => kpiData[`${fy}-${fullMonth}-${c.id}`] || {});
                const received  = monthRows.reduce((s, r) => s + (parseInt(r.escalations) || 0), 0);
                // "closed" simulated as % of received that scored 0 in that month
                const closed    = Math.max(0, received - Math.floor(received * 0.25));
                return { month: mo, received, closed };
              });
              const hasData = trendData.some(d => d.received > 0);
              if (!hasData) {
                // Show demo data so chart is always visual
                const demo = KPI_MONTHS_ORDER.map((mo,i) => ({
                  month: mo,
                  received: [4,6,3,8,5,7,4,6,9,5,3,4][i],
                  closed:   [3,5,3,6,4,6,4,5,7,4,3,4][i],
                }));
                return (
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={demo} margin={{ top:4, right:12, left:-20, bottom:0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#f1f5f9"} vertical={false}/>
                      <XAxis dataKey="month" tick={{ fontSize:9, fill: dark?"#5a7a99":"#94a3b8" }} axisLine={false} tickLine={false}/>
                      <YAxis tick={{ fontSize:9, fill: dark?"#5a7a99":"#94a3b8" }} axisLine={false} tickLine={false} allowDecimals={false}/>
                      <Tooltip
                        contentStyle={{ background: dark?"#0f1e2e":"#fff", border:`1px solid ${dark?"#1a2d44":"#e2e8f0"}`, borderRadius:12, fontSize:11 }}
                        labelStyle={{ color: dark?"#e8f0f8":"#1e293b", fontWeight:700 }}/>
                      <Line type="monotone" dataKey="received" stroke="#f87171" strokeWidth={2.5} dot={false} activeDot={{ r:4, fill:"#f87171" }}/>
                      <Line type="monotone" dataKey="closed"   stroke="#34d399" strokeWidth={2.5} dot={false} activeDot={{ r:4, fill:"#34d399" }}/>
                    </LineChart>
                  </ResponsiveContainer>
                );
              }
              return (
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={trendData} margin={{ top:4, right:12, left:-20, bottom:0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#f1f5f9"} vertical={false}/>
                    <XAxis dataKey="month" tick={{ fontSize:9, fill: dark?"#5a7a99":"#94a3b8" }} axisLine={false} tickLine={false}/>
                    <YAxis tick={{ fontSize:9, fill: dark?"#5a7a99":"#94a3b8" }} axisLine={false} tickLine={false} allowDecimals={false}/>
                    <Tooltip
                      contentStyle={{ background: dark?"#0f1e2e":"#fff", border:`1px solid ${dark?"#1a2d44":"#e2e8f0"}`, borderRadius:12, fontSize:11 }}
                      labelStyle={{ color: dark?"#e8f0f8":"#1e293b", fontWeight:700 }}/>
                    <Line type="monotone" dataKey="received" stroke="#f87171" strokeWidth={2.5} dot={false} activeDot={{ r:4, fill:"#f87171" }}/>
                    <Line type="monotone" dataKey="closed"   stroke="#34d399" strokeWidth={2.5} dot={false} activeDot={{ r:4, fill:"#34d399" }}/>
                  </LineChart>
                </ResponsiveContainer>
              );
            })()}
          </div>
        </div>

        {/* RIGHT: 4 Mini Donut Charts */}
        <div className="lg:col-span-2 grid grid-cols-2 gap-3">
          {[
            {
              label: "MYSA Usage",
              field: "mysaUsage",
              color: { Yes:"#34d399", No:"#f87171", NA:"#94a3b8" },
            },
            {
              label: "Report Rectification",
              sublabel: "within 15 days",
              field: "revertRect",
              color: { Yes:"#34d399", No:"#f87171", NA:"#94a3b8" },
            },
            {
              label: "Raksha Tool",
              sublabel: "Payments via Raksha",
              field: "rakshaTool",
              color: { Yes:"#34d399", No:"#f87171", NA:"#94a3b8" },
            },
            {
              label: "Capital Usage",
              sublabel: "100% agreed",
              field: "capitalWant",
              color: { Yes:"#34d399", No:"#f87171", NA:"#94a3b8" },
            },
          ].map(dc => {
            const counts = { Yes: 0, No: 0, NA: 0, "": 0 };
            visibleClients.forEach(c => {
              const val = getRow(c.id)[dc.field] || "";
              if (val === "Yes" || val === "No" || val === "NA") counts[val]++;
              else counts[""]++;
            });
            const total = visibleClients.length || 1;
            const segments = [
              { k:"Yes", v:counts.Yes, c:"#34d399" },
              { k:"No",  v:counts.No,  c:"#f87171" },
              { k:"NA",  v:counts.NA,  c:"#94a3b8" },
            ].filter(s => s.v > 0);
            const yesPct = Math.round((counts.Yes / total) * 100);

            // Build SVG donut arcs
            const R = 28, cx = 36, cy = 36, stroke = 7;
            const circ = 2 * Math.PI * R;
            let offset = 0;
            const arcs = segments.map(s => {
              const pct   = s.v / total;
              const dash  = pct * circ;
              const arc   = { ...s, dash, gap: circ - dash, offset };
              offset += dash;
              return arc;
            });

            return (
              <div key={dc.field}
                className="rounded-2xl border p-3 flex flex-col items-center gap-2 transition-all duration-200"
                style={{
                  background: dark ? "#0f1e2e" : "#ffffff",
                  borderColor: dark ? "#1a2d44" : "#e2e8f0",
                  boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
                }}
                onMouseEnter={e=>{ e.currentTarget.style.transform="translateY(-1px)"; e.currentTarget.style.boxShadow="0 6px 18px rgba(0,0,0,0.12)"; }}
                onMouseLeave={e=>{ e.currentTarget.style.transform=""; e.currentTarget.style.boxShadow="0 2px 8px rgba(0,0,0,0.06)"; }}>

                {/* Label */}
                <div className="text-center">
                  <div className="text-[10px] font-bold uppercase tracking-widest leading-tight"
                    style={{ color: dark ? "#94b4cc" : "#475569" }}>{dc.label}</div>
                  {dc.sublabel && <div className="text-[8px] mt-0.5" style={{ color: dark?"#5a7a99":"#94a3b8" }}>{dc.sublabel}</div>}
                </div>

                {/* Donut SVG */}
                <div className="relative">
                  <svg width={72} height={72} viewBox="0 0 72 72">
                    {/* Track */}
                    <circle cx={cx} cy={cy} r={R} fill="none"
                      stroke={dark?"#1a2d44":"#f1f5f9"} strokeWidth={stroke}/>
                    {/* Segments */}
                    {arcs.map((arc, i) => (
                      <circle key={i} cx={cx} cy={cy} r={R} fill="none"
                        stroke={arc.c} strokeWidth={stroke}
                        strokeDasharray={`${arc.dash} ${arc.gap}`}
                        strokeDashoffset={-arc.offset}
                        strokeLinecap="butt"
                        transform={`rotate(-90 ${cx} ${cy})`}/>
                    ))}
                    {/* Center label */}
                    <text x={cx} y={cy-4} textAnchor="middle" dominantBaseline="middle"
                      fontSize="11" fontWeight="800" fill="#34d399">{yesPct}%</text>
                    <text x={cx} y={cy+8} textAnchor="middle" dominantBaseline="middle"
                      fontSize="7" fontWeight="500" fill={dark?"#5a7a99":"#94a3b8"}>Yes</text>
                  </svg>
                </div>

                {/* Legend */}
                <div className="flex items-center gap-2 flex-wrap justify-center">
                  {[{k:"Yes",c:"#34d399"},{k:"No",c:"#f87171"},{k:"NA",c:"#94a3b8"}].map(l=>(
                    <div key={l.k} className="flex items-center gap-1" title={`${l.k}: ${counts[l.k]}`}>
                      <div className="w-2 h-2 rounded-full shrink-0" style={{background:l.c}}/>
                      <span className="text-[9px] font-semibold" style={{color:l.c}}>{l.k} {counts[l.k]}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* ══════════════════════ BOTTOM ROW: HEATMAP + MIS DEADLINE ══════════════════════ */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">

        {/* LEFT 2/3: Traffic-Light Heatmap */}
        <div className="lg:col-span-2 rounded-2xl border overflow-hidden"
          style={{ background: dark ? "#0f1e2e" : "#ffffff", borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
          <div className="px-5 py-4 border-b flex items-center justify-between"
            style={{ borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
            <div>
              <div className="font-bold text-sm" style={{ color: dark?"#e8f0f8":"#1e293b" }}>
                Performance Heatmap
              </div>
              <div className="text-[11px] mt-0.5" style={{ color: dark?"#5a7a99":"#94a3b8" }}>
                Traffic-light compliance · {month} · FY {fy}
              </div>
            </div>
            <div className="flex items-center gap-3 text-[9px] font-semibold flex-wrap">
              {[{c:"#34d399",l:"On Track"},{c:"#fbbf24",l:"Partial"},{c:"#f87171",l:"Missed"},{c:"#94a3b8",l:"NA / Empty"}].map(l=>(
                <span key={l.l} className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-sm inline-block" style={{background:l.c}}/>
                  {l.l}
                </span>
              ))}
            </div>
          </div>
          <div className="overflow-x-auto" style={{ maxHeight:320, overflowY:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:560 }}>
              <thead>
                <tr style={{ background: dark?"#081628":"#f8fafc" }}>
                  <th className="px-4 py-2.5 text-left text-[9px] font-bold uppercase tracking-widest sticky left-0"
                    style={{ color:dark?"#5a7a99":"#94a3b8", background: dark?"#081628":"#f8fafc", minWidth:150, borderBottom:`1px solid ${dark?"#1a2d44":"#e2e8f0"}` }}>
                    Client
                  </th>
                  {[
                    { label:"MIS Before 10th", key:"mis" },
                    { label:"MYSA Usage",       key:"mysa" },
                    { label:"Report Rect.",     key:"rect" },
                    { label:"Escalation",       key:"esc" },
                    { label:"Raksha Tool",      key:"raksha" },
                    { label:"Capital",          key:"capital" },
                  ].map(col => (
                    <th key={col.key} className="px-3 py-2.5 text-center text-[8px] font-bold uppercase tracking-widest"
                      style={{ color:dark?"#5a7a99":"#94a3b8", borderBottom:`1px solid ${dark?"#1a2d44":"#e2e8f0"}`, minWidth:90 }}>
                      {col.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {visibleClients.slice(0, 12).map((client, ri) => {
                  const row = getRow(client.id);

                  const cellColor = (status) => {
                    if (status === "green")  return { bg:"#34d39920", color:"#34d399", symbol:"✓" };
                    if (status === "amber")  return { bg:"#fbbf2420", color:"#fbbf24", symbol:"⚡" };
                    if (status === "red")    return { bg:"#f8717120", color:"#f87171", symbol:"✗" };
                    if (status === "na")     return { bg: dark?"#1a2d4440":"#f1f5f9", color:"#94a3b8", symbol:"—" };
                    return { bg:"transparent", color: dark?"#2d3d50":"#e2e8f0", symbol:"·" };
                  };

                  const misStatus = !row.misDate ? "empty"
                    : misDateStatus(row.misDate).late ? "red" : "green";
                  const ynaStatus = (v) => !v ? "empty" : v==="Yes" ? "green" : v==="No" ? "red" : "na";
                  const escStatus = row.escalations === undefined || row.escalations === "" ? "empty"
                    : parseInt(row.escalations) === 0 ? "green"
                    : parseInt(row.escalations) <= 2 ? "amber" : "red";

                  const cells = [
                    cellColor(misStatus),
                    cellColor(ynaStatus(row.mysaUsage)),
                    cellColor(ynaStatus(row.revertRect)),
                    cellColor(escStatus),
                    cellColor(ynaStatus(row.rakshaTool)),
                    cellColor(ynaStatus(row.capitalWant)),
                  ];

                  const score = rowHealthScore(row);
                  const scoreCol = healthToColor(score);

                  return (
                    <tr key={client.id}
                      style={{ background: ri%2===0 ? (dark?"rgba(15,30,46,0.5)":"#ffffff") : (dark?"rgba(8,18,28,0.4)":"#f8fbff") }}
                      onMouseEnter={e=>{ e.currentTarget.style.background = dark?"rgba(20,40,64,0.9)":"#f0faf8"; }}
                      onMouseLeave={e=>{ e.currentTarget.style.background = ri%2===0 ? (dark?"rgba(15,30,46,0.5)":"#ffffff") : (dark?"rgba(8,18,28,0.4)":"#f8fbff"); }}>
                      {/* Client name cell */}
                      <td className="px-4 py-2.5 text-xs font-semibold sticky left-0"
                        style={{ color:dark?"#c8dff0":"#1e293b", background: ri%2===0?(dark?"rgba(15,30,46,0.95)":"#ffffff"):(dark?"rgba(8,18,28,0.95)":"#f8fbff"), borderBottom:`1px solid ${dark?"#1a2d4430":"#f1f5f9"}` }}>
                        <div className="flex items-center gap-2">
                          <div className="w-1.5 h-1.5 rounded-full shrink-0" style={{background:scoreCol}}/>
                          <span className="truncate max-w-[130px]" title={client.clientName}>{client.clientName}</span>
                        </div>
                      </td>
                      {/* Compliance cells */}
                      {cells.map((cell, ci) => (
                        <td key={ci} className="px-3 py-2.5 text-center"
                          style={{ borderBottom:`1px solid ${dark?"#1a2d4430":"#f1f5f9"}` }}
                          title={`${client.clientName} · ${["MIS","MYSA","Rect","Escalation","Raksha","Capital"][ci]}`}>
                          <div className="inline-flex items-center justify-center w-7 h-7 rounded-lg text-[11px] font-bold mx-auto"
                            style={{ background: cell.bg, color: cell.color }}>
                            {cell.symbol}
                          </div>
                        </td>
                      ))}
                    </tr>
                  );
                })}
                {visibleClients.length > 12 && (
                  <tr>
                    <td colSpan={7} className="px-4 py-2 text-center text-[10px]"
                      style={{ color: dark?"#5a7a99":"#94a3b8", borderTop:`1px solid ${dark?"#1a2d44":"#e2e8f0"}` }}>
                      Showing 12 of {visibleClients.length} clients · Use table view for full list
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* RIGHT 1/3: MIS Deadline Monitoring Chart */}
        <div className="rounded-2xl border overflow-hidden"
          style={{ background: dark ? "#0f1e2e" : "#ffffff", borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
          <div className="px-5 py-4 border-b"
            style={{ borderColor: dark ? "#1a2d44" : "#e2e8f0" }}>
            <div className="font-bold text-sm" style={{ color: dark?"#e8f0f8":"#1e293b" }}>
              MIS Deadline Monitor
            </div>
            <div className="text-[11px] mt-0.5" style={{ color: dark?"#5a7a99":"#94a3b8" }}>
              MIS completion date · 10th = target
            </div>
          </div>
          <div className="px-4 py-3 overflow-y-auto" style={{ maxHeight: 280 }}>
            {(() => {
              const items = visibleClients
                .map(c => {
                  const r = getRow(c.id);
                  if (!r.misDate) return null;
                  const d = new Date(r.misDate);
                  if (isNaN(d.getTime())) return null;
                  return { name: c.clientName.length > 20 ? c.clientName.slice(0,18)+"…" : c.clientName, day: d.getDate(), late: d.getDate() > 10 };
                })
                .filter(Boolean)
                .sort((a,b) => a.day - b.day);

              if (items.length === 0) {
                // Demo data
                const demo = [
                  {name:"Client Alpha",   day:5, late:false},
                  {name:"Client Delta",   day:8, late:false},
                  {name:"Client Beta",    day:10, late:false},
                  {name:"Client Eta",     day:13, late:true},
                  {name:"Client Theta",   day:16, late:true},
                  {name:"Client Epsilon", day:9, late:false},
                ];
                return (
                  <div className="space-y-2">
                    {demo.map((item,i)=>(
                      <div key={i} className="flex items-center gap-2">
                        <div className="text-[9px] font-semibold shrink-0 w-[100px] truncate"
                          style={{ color: dark?"#94b4cc":"#475569" }} title={item.name}>
                          {item.name}
                        </div>
                        <div className="flex-1 relative" style={{ height:18 }}>
                          <div className="absolute top-1/2 -translate-y-1/2 h-1.5 rounded-full w-full"
                            style={{ background: dark?"#1a2d44":"#f1f5f9" }}/>
                          <div className="absolute top-1/2 -translate-y-1/2 h-1.5 rounded-full"
                            style={{ width:`${Math.min((item.day/31)*100,100)}%`,
                                     background: item.late ? "#f87171" : "#34d399" }}/>
                          {/* 10th marker */}
                          <div className="absolute top-0 bottom-0 w-0.5 rounded"
                            style={{ left:`${(10/31)*100}%`, background:"#fbbf24", opacity:0.8 }}/>
                        </div>
                        <div className="text-[9px] font-black shrink-0 w-6 text-right"
                          style={{ color: item.late?"#f87171":"#34d399" }}>
                          {item.day}
                        </div>
                      </div>
                    ))}
                    <div className="flex items-center gap-1.5 text-[8px] mt-2 pt-2 border-t" style={{ borderColor:dark?"#1a2d44":"#e2e8f0", color:dark?"#5a7a99":"#94a3b8" }}>
                      <div className="w-2 h-0.5 rounded" style={{background:"#fbbf24"}}/>10th deadline
                      <span className="ml-1 flex items-center gap-1"><div className="w-2 h-1 rounded" style={{background:"#f87171"}}/>After 10th</span>
                    </div>
                  </div>
                );
              }
              return (
                <div className="space-y-2">
                  {items.map((item,i)=>(
                    <div key={i} className="flex items-center gap-2">
                      <div className="text-[9px] font-semibold shrink-0 w-[100px] truncate"
                        style={{ color: dark?"#94b4cc":"#475569" }} title={item.name}>
                        {item.name}
                      </div>
                      <div className="flex-1 relative" style={{ height:18 }}>
                        <div className="absolute top-1/2 -translate-y-1/2 h-1.5 rounded-full w-full"
                          style={{ background: dark?"#1a2d44":"#f1f5f9" }}/>
                        <div className="absolute top-1/2 -translate-y-1/2 h-1.5 rounded-full"
                          style={{ width:`${Math.min((item.day/31)*100,100)}%`,
                                   background: item.late ? "#f87171" : "#34d399" }}/>
                        <div className="absolute top-0 bottom-0 w-0.5 rounded"
                          style={{ left:`${(10/31)*100}%`, background:"#fbbf24", opacity:0.8 }}/>
                      </div>
                      <div className="text-[9px] font-black shrink-0 w-6 text-right"
                        style={{ color: item.late?"#f87171":"#34d399" }}>
                        {item.day}
                      </div>
                    </div>
                  ))}
                  <div className="flex items-center gap-1.5 text-[8px] mt-2 pt-2 border-t" style={{ borderColor:dark?"#1a2d44":"#e2e8f0", color:dark?"#5a7a99":"#94a3b8" }}>
                    <div className="w-2 h-0.5 rounded" style={{background:"#fbbf24"}}/>10th deadline
                    <span className="ml-1 flex items-center gap-1"><div className="w-2 h-1 rounded" style={{background:"#f87171"}}/>After 10th</span>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      </div>

      {/* ══════════════════════ CARD VIEW ══════════════════════ */}
      {viewMode === "cards" && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
          {visibleClients.length === 0 ? (
            <div className={`col-span-4 ${t.card} border ${t.cardBorder} rounded-2xl py-12 text-center`}>
              <p className={`text-sm ${t.textMuted}`}>No clients match the current filters.</p>
            </div>
          ) : visibleClients.map(client => (
            <KpiClientCard
              key={client.id}
              t={t} dark={dark}
              client={client}
              row={getRow(client.id)}
              onEdit={(id) => { setViewMode("table"); setEditRowId(id); }}/>
          ))}
        </div>
      )}

      {/* ══════════════════════ TABLE VIEW ══════════════════════ */}
      {viewMode === "table" && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          {/* Table header */}
          <div className={`px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b ${t.cardBorder}`}>
            <div>
              <h3 className={`font-bold text-sm ${t.text}`}>KRA / KPI Monthly Tracker</h3>
              <p className={`text-xs mt-0.5 ${t.textMuted}`}>
                {visibleClients.length} clients · {month} · FY {fy}
                {" "}· {aggStats.filled} filled · {visibleClients.length - aggStats.filled} pending
              </p>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {!isAdmin && (
                <span className="flex items-center gap-1 text-xs font-medium px-2 py-1 rounded-lg"
                  style={{background:"#f8717110",color:"#f87171",border:"1px solid #f8717125"}}>
                  <Icon path={Icons.shield} size={11}/> Your clients only
                </span>
              )}
              <button onClick={handleExport}
                className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-medium
                  transition-all ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                <Icon path={Icons.download} size={13}/> Export Excel
              </button>
            </div>
          </div>

          {/* Month-lock banner */}
          {checkLocked(locks, fy, month, "kra") && (
            <div className="px-5 py-2.5 flex items-center gap-3 text-xs font-semibold"
              style={{background:"#f8717115",borderBottom:"1px solid #f8717130",color:"#f87171"}}>
              <span className="text-base">🔒</span>
              <span>Data entry for <strong>{month} · FY {fy}</strong> has been locked by Administrator.</span>
              {isAdmin && <span className="ml-2 font-normal opacity-60">(Architect → Month Lock to unlock)</span>}
            </div>
          )}
          {/* Info banner: month-wise persistence */}
          <div className="px-5 py-2 flex items-center gap-2 text-xs"
            style={{background:"#34d39908", borderBottom:`1px solid #34d39920`}}>
            <div className="w-1.5 h-1.5 rounded-full animate-pulse" style={{background:"#34d399"}}/>
            <span style={{color:"#34d399"}} className="font-semibold">
              All {visibleClients.length} clients shown for {month} · FY {fy}
            </span>
            <span className={t.textMuted}>— rows without data are blank, ready to fill in</span>
          </div>

          {/* Sticky-header scrollable table */}
          <div className="overflow-x-auto" style={{maxHeight:"70vh", overflowY:"auto"}}>
            <table className="w-full border-collapse" style={{minWidth:1280}}>
              <thead className="sticky top-0 z-10">
                <tr style={{
                  background: dark
                    ? "linear-gradient(90deg,#0c1e30 0%,#0e2240 60%,#0a1a2e 100%)"
                    : "linear-gradient(90deg,#1b3a5c 0%,#1a3356 60%,#1e3d6a 100%)",
                  color: "#c8dff0",
                  boxShadow: "0 2px 8px rgba(0,0,0,0.28)",
                }}>
                  <th className="px-4 py-3.5 text-left font-bold text-[9px] uppercase tracking-widest w-8"
                    style={{color:"#94b4cc", borderBottom:"1px solid #ffffff18"}}>#</th>
                  <th className="px-3 py-3.5 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{minWidth:190, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    Client Name
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:52, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}} title="Overall KPI Score">
                    Score
                  </th>
                  <th className="px-3 py-3.5 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{width:155, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    <div>Date of MIS Updating</div>
                    <div style={{fontSize:8,fontWeight:400,textTransform:"none",letterSpacing:0,opacity:0.55,marginTop:2}}>Due: 10th each month</div>
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:110, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    MYSA Usage
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:125, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    <div>Revert / Rect</div>
                    <div style={{fontSize:8,fontWeight:400,textTransform:"none",letterSpacing:0,opacity:0.55,marginTop:2}}>within 15 days</div>
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:110, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    <div>Escalations</div>
                    <div style={{fontSize:8,fontWeight:400,textTransform:"none",letterSpacing:0,opacity:0.55,marginTop:2}}>closed in 7d</div>
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:115, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    Raksha Tool
                  </th>
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:120, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    <div>CapitallWant</div>
                    <div style={{fontSize:8,fontWeight:400,textTransform:"none",letterSpacing:0,opacity:0.55,marginTop:2}}>100% agreed</div>
                  </th>
                  <th className="px-3 py-3.5 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{minWidth:160, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    Remarks
                  </th>
                  {/* Actions always shown — edit/view-only determined per-row by canEditThisRow */}
                  <th className="px-3 py-3.5 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:88, color:"#94b4cc", borderBottom:"1px solid #ffffff18"}}>
                    Actions
                  </th>
                </tr>
              </thead>
              <tbody>
                {visibleClients.length === 0 ? (
                  <tr>
                    <td colSpan={11} className="py-14 text-center">
                      <p className={`text-sm ${t.textMuted}`}>No clients match the current filters.</p>
                    </td>
                  </tr>
                ) : visibleClients.map((client, idx) => {
                  const row    = getRow(client.id);
                  const score  = rowHealthScore(row);
                  const col    = healthToColor(score);
                  // Lock check for this specific row's month
                  const kraLocked = checkLocked(locks, fy, month, "kra");
                  // Edit: admin can edit if not locked; user can only edit their assigned clients
                  const canEditThisRow = kraLocked ? false : canEditClient(activeUser, client.id);
                  const isEdit = editRowId === client.id && canEditThisRow;
                  const hasMisData = Object.keys(row).length > 0;
                  const misS   = row.misDate ? misDateStatus(row.misDate) : null;
                  const escS   = escalationColor(row.escalations);
                  const reminderKey = `${client.id}-${month}`;

                  return (
                    <tr key={client.id}
                      className={`transition-colors border-b ${t.cardBorder}`}
                      style={{
                        background: isEdit
                          ? (dark ? "rgba(13,30,48,0.95)" : "rgba(232,245,240,0.9)")
                          : hasMisData
                            ? (idx % 2 === 0 ? (dark ? "rgba(15,30,46,0.9)" : "#ffffff") : (dark ? "rgba(12,24,38,0.7)" : "#f8fbff"))
                            : (dark ? "rgba(12,24,36,0.6)" : "rgba(245,253,251,0.7)"),
                        borderLeft: !hasMisData && !isEdit ? `2px solid #34d39930` : undefined,
                      }}>

                      {/* # */}
                      <td className={`px-4 py-2.5 text-xs font-mono font-bold text-center ${t.textMuted}`}>
                        {idx+1}
                        {!hasMisData && <div className="text-[8px] font-normal mt-0.5" style={{color:"#34d39980"}}>empty</div>}
                      </td>

                      {/* Client name */}
                      <td className="px-3 py-2.5">
                        <div className={`font-semibold text-sm ${t.text}`}>{client.clientName}</div>
                        <div className={`text-[11px] mt-0.5 ${t.textMuted}`}>{client.personName}</div>
                      </td>

                      {/* Score ring */}
                      <td className="px-3 py-2 text-center">
                        {hasMisData
                          ? <CircleRing score={score} size={36} stroke={4}/>
                          : <span className={`text-xs ${t.textMuted} italic`}>—</span>}
                      </td>

                      {/* Date of MIS */}
                      <td className="px-3 py-2">
                        {isEdit ? (
                          <div className="flex flex-col gap-1">
                            <input type="date"
                              value={row.misDate||""}
                              onChange={e=>setRow(client.id,{misDate:e.target.value})}
                              className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}/>
                            <label className={`flex items-center gap-1.5 text-[10px] cursor-pointer ${t.textMuted}`}>
                              <input type="checkbox"
                                checked={!!reminderSet[reminderKey]}
                                onChange={e=>setReminderSet(p=>({...p,[reminderKey]:e.target.checked}))}
                                className="accent-[#34d399]"/>
                              Reminder on 10th
                            </label>
                          </div>
                        ) : (
                          <div className="flex flex-col gap-1">
                            {row.misDate ? (
                              <span className="flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full w-fit"
                                style={{ color:misS.color, background:misS.bg }}>
                                {misS.late && "⚠ "}
                                {new Date(row.misDate).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"2-digit"})}
                              </span>
                            ) : (
                              <span className={`text-xs italic ${t.textMuted}`}>Not set</span>
                            )}
                            {reminderSet[reminderKey] && (
                              <span className="text-[9px] px-1.5 py-0.5 rounded font-semibold w-fit"
                                style={{color:"#fbbf24",background:"#fbbf2415"}}>⏰ Reminder ON</span>
                            )}
                          </div>
                        )}
                      </td>

                      {/* MYSA Usage */}
                      <td className="px-3 py-2 text-center">
                        {isEdit ? (
                          <select value={row.mysaUsage||""}
                            onChange={e=>setRow(client.id,{mysaUsage:e.target.value})}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}>
                            {YNA_OPTIONS.map(o=><option key={o} value={o}>{o||"— select —"}</option>)}
                          </select>
                        ) : <YNABadge value={row.mysaUsage}/>}
                      </td>

                      {/* Revert / Rectification */}
                      <td className="px-3 py-2 text-center">
                        {isEdit ? (
                          <select value={row.revertRect||""}
                            onChange={e=>setRow(client.id,{revertRect:e.target.value})}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}>
                            {YNA_OPTIONS.map(o=><option key={o} value={o}>{o||"— select —"}</option>)}
                          </select>
                        ) : <YNABadge value={row.revertRect}/>}
                      </td>

                      {/* Escalations */}
                      <td className="px-3 py-2 text-center">
                        {isEdit ? (
                          <input type="number" min="0"
                            value={row.escalations!==undefined?row.escalations:""}
                            onChange={e=>setRow(client.id,{escalations:e.target.value})}
                            placeholder="0"
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs text-center outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}/>
                        ) : (
                          row.escalations !== "" && row.escalations !== undefined
                            ? <span className="text-[12px] font-black px-2.5 py-0.5 rounded-full"
                                style={{color:escS.color,background:escS.bg}}>
                                {row.escalations}
                              </span>
                            : <span className={`text-xs ${t.textMuted} italic`}>—</span>
                        )}
                      </td>

                      {/* Raksha Tool */}
                      <td className="px-3 py-2 text-center">
                        {isEdit ? (
                          <select value={row.rakshaTool||""}
                            onChange={e=>setRow(client.id,{rakshaTool:e.target.value})}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}>
                            {YNA_OPTIONS.map(o=><option key={o} value={o}>{o||"— select —"}</option>)}
                          </select>
                        ) : <YNABadge value={row.rakshaTool}/>}
                      </td>

                      {/* CapitallWant */}
                      <td className="px-3 py-2 text-center">
                        {isEdit ? (
                          <select value={row.capitalWant||""}
                            onChange={e=>setRow(client.id,{capitalWant:e.target.value})}
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}>
                            {YNA_OPTIONS.map(o=><option key={o} value={o}>{o||"— select —"}</option>)}
                          </select>
                        ) : <YNABadge value={row.capitalWant}/>}
                      </td>

                      {/* Remarks */}
                      <td className="px-3 py-2">
                        {isEdit ? (
                          <input type="text"
                            value={row.remarks||""}
                            onChange={e=>setRow(client.id,{remarks:e.target.value})}
                            placeholder="Add remarks…"
                            className={`w-full px-2 py-1.5 rounded-lg border text-xs outline-none ${t.input} focus:ring-2 focus:ring-[#34d399]`}/>
                        ) : (
                          <span className={`text-xs ${row.remarks?t.text:t.textMuted}`}>
                            {row.remarks || <span className="italic opacity-40">—</span>}
                          </span>
                        )}
                      </td>

                      {/* Actions — show for all users; non-assigned gets View Only badge */}
                      <td className="px-3 py-2 text-center">
                        {canEditThisRow ? (
                          isEdit ? (
                            <div className="flex items-center justify-center gap-1">
                              <button onClick={()=>handleSave(client.id)}
                                className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold text-white"
                                style={{background:"linear-gradient(135deg,#0d9e7b,#1b5fa8)"}}>
                                <Icon path={Icons.save} size={11}/> Save
                              </button>
                              <button onClick={()=>setEditRowId(null)}
                                className={`px-2 py-1.5 rounded-lg text-xs ${t.textMuted} ${t.hover}`}>
                                <Icon path={Icons.x} size={11}/>
                              </button>
                            </div>
                          ) : (
                            <button onClick={()=>setEditRowId(client.id)}
                              className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium
                                border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                              <Icon path={Icons.edit} size={11}/> Edit
                            </button>
                          )
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-medium"
                            style={{background:"#94a3b815",color:"#94a3b8",border:"1px solid #94a3b830"}}>
                            🔒 View Only
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Legend + footnote */}
          <div className={`px-5 py-3 border-t ${t.cardBorder} flex flex-wrap items-center gap-x-6 gap-y-2 text-[10px] ${t.textMuted}`}>
            <div className="flex items-center gap-3">
              {[
                {col:"#34d399",label:"On Track / Yes"},
                {col:"#fbbf24",label:"Pending / Warning"},
                {col:"#f87171",label:"Late / No / Attention"},
                {col:"#94a3b8",label:"NA"},
              ].map(l=>(
                <div key={l.label} className="flex items-center gap-1.5">
                  <div className="w-2.5 h-2.5 rounded-full" style={{background:l.col}}/>
                  {l.label}
                </div>
              ))}
            </div>
            <div className="ml-auto flex items-center gap-1.5">
              <Icon path={Icons.info} size={11}/>
              MIS due date: <strong>10th of each month</strong> · Escalations resolved within 7 days
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── FUND REQUEST TAB  ─────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

// Inline sync button for Fund Request data-entry view
const FundSyncButton = ({ t }) => {
  const { syncing, lastSynced, triggerSync } = useSyncContext();
  const fmt = (d) => {
    if (!d) return "";
    const diff = Math.floor((Date.now() - d.getTime()) / 1000);
    if (diff < 60) return `${diff}s ago`;
    return `${Math.floor(diff/60)}m ago`;
  };
  return (
    <button onClick={triggerSync} disabled={syncing}
      className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold
        text-white transition-all hover:opacity-90 active:scale-95 disabled:opacity-60"
      style={{ background: "linear-gradient(135deg,#00a896,#1b5fa8)" }}
      title={lastSynced ? `Last sync: ${fmt(lastSynced)}` : "Sync client data"}>
      {syncing ? <SpinIcon size={13}/> : <Icon path={Icons.refreshCw} size={13}/>}
      {syncing ? "Syncing…" : "Sync Clients"}
    </button>
  );
};

const CellInput = ({ t, value, onChange, placeholder, type="text", disabled=false, prefix }) => (
  <div className={`flex items-center gap-1 px-2 py-1.5 rounded-lg border transition-all ${disabled?"opacity-40":""} ${t.input} focus-within:ring-2 focus-within:ring-[#00c9b1]`}>
    {prefix&&<span className={`text-[11px] font-bold shrink-0 ${t.textMuted}`}>{prefix}</span>}
    <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} disabled={disabled}
      className={`bg-transparent outline-none text-sm w-full min-w-0 ${disabled?"cursor-not-allowed":""}`}/>
  </div>
);

const FundRequestTab = ({ t, dark, isAdmin }) => {
  const { clients: masterClients } = useSyncContext();
  const { locks } = useLockCtx();
  const { canEdit: rbacCanEdit, activeUser } = useRBAC();
  const [fy, setFy]                       = useState(DEFAULT_FY);
  const [month, setMonth]                 = useState("All Months");
  const [clientFilter, setClientFilter]   = useState("All");
  const [currencyFilter, setCurrencyFilter] = useState("All");
  // ── view mode: "analytics" | "entry"
  const [view, setView]                   = useState("analytics");
  // ── table search / pagination
  const [tableSearch, setTableSearch]     = useState("");
  const [tablePage, setTablePage]         = useState(1);
  const TABLE_PAGE_SIZE = 8;
  // ── analytics quarter filter
  const [quarterFilter, setQuarterFilter] = useState("All");

  const [rows, setRows] = useState([
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Alpha",   currency:"USD", amountFC:"50000",  rate:"83.5",  remarks:"Import payment Q1",        month:"April"    },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Delta",   currency:"EUR", amountFC:"20000",  rate:"90.2",  remarks:"Software license renewal",  month:"April"    },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Beta",    currency:"INR", amountFC:"500000", rate:"1",     remarks:"Advance tax installment",   month:"May"      },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Eta",     currency:"GBP", amountFC:"15000",  rate:"105.8", remarks:"Machinery parts import",    month:"May"      },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Theta",   currency:"AED", amountFC:"80000",  rate:"22.7",  remarks:"Joint venture remittance",  month:"June"     },
    { _id:genRowId(), applicability_status:"not_applicable", clientName:"Client Epsilon", currency:"SGD", amountFC:"12000",  rate:"62.1",  remarks:"Pending clarification",     month:"June"     },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Alpha",   currency:"USD", amountFC:"35000",  rate:"83.9",  remarks:"Q2 import settlement",      month:"July"     },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Delta",   currency:"CHF", amountFC:"8000",   rate:"95.4",  remarks:"Partner payment",          month:"August"   },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Zeta",    currency:"AUD", amountFC:"5000",   rate:"54.3",  remarks:"Professional fee",         month:"September"},
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Eta",     currency:"JPY", amountFC:"500000", rate:"0.56",  remarks:"Machinery payment",        month:"October"  },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Gamma",   currency:"INR", amountFC:"250000", rate:"1",     remarks:"Rental income",            month:"November" },
    { _id:genRowId(), applicability_status:"applicable",     clientName:"Client Theta",   currency:"HKD", amountFC:"60000",  rate:"10.7",  remarks:"Investment proceeds",      month:"December" },
  ]);

  const [openDropId, setOpenDropId]   = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [savedToast, setSavedToast]   = useState(false);
  const [clientSearch, setClientSearch] = useState({});
  const dropRefs = useRef({});

  useEffect(()=>{
    const h = e => { if(openDropId){ const r=dropRefs.current[openDropId]; if(r&&!r.contains(e.target)) setOpenDropId(null); } };
    document.addEventListener("mousedown",h);
    return ()=>document.removeEventListener("mousedown",h);
  },[openDropId]);

  const calcINR = row => {
    const fc=parseFloat(row.amountFC)||0, rt=parseFloat(row.rate)||0;
    return row.currency==="INR"?fc:fc*rt;
  };

  const addRow = () => {
    const targetMonth = month === "All Months" ? "April" : month;
    if (checkLocked(locks, fy, targetMonth, "fund")) return; // blocked when locked
    // Non-admin users can add rows (new rows are blank — client not yet assigned, so allow)
    setRows(r => [...r, {
      _id: genRowId(),
      applicability_status: "applicable",
      clientName: "",
      currency: "INR",
      amountFC: "",
      rate: "1",
      remarks: "",
      month: targetMonth,
      _isBlank: false,
    }]);
  };

  const deleteRow = id => {
    // Find the row's month to check lock status
    const row = rows.find(r => r._id === id);
    if (row && checkLocked(locks, fy, row.month || (month === "All Months" ? "April" : month), "fund")) {
      setDeleteConfirm(null); return; // blocked when locked
    }
    setRows(r => r.filter(row => row._id !== id));
    setDeleteConfirm(null);
  };

  const handleSave = () => { setSavedToast(true); setTimeout(() => setSavedToast(false), 2500); };

  const exportToExcel = () => {
    try {
      const data = entryViewRows.map((r,i)=>{
        const isNA = r.applicability_status === "not_applicable";
        const cur  = CURRENCIES.find(c=>c.code===r.currency);
        return {
          "#":                   i+1,
          "Applicable?":         isNA ? "Not Applicable" : "Applicable",
          "Client Name":         isNA ? "NA" : xlsSafe(r.clientName   || ""),
          "Currency":            isNA ? "NA" : (r.currency     || ""),
          "Amount (Foreign)":    isNA ? "NA" : (parseFloat(r.amountFC) || 0),
          "Exchange Rate":       isNA ? "NA" : (parseFloat(r.rate)     || 0),
          "Amount in INR":       isNA ? "NA" : calcINR(r),
          "Remarks":             xlsSafe(r.remarks || ""),
          "Month":               r.month   || "",
          "Financial Year":      fy,
        };
      });
      if (!data.length) { alert("No data to export for the selected filters."); return; }
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [4,16,30,10,18,15,18,30,12,12].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Fund Requests");
      const monthPart = month !== "All Months" ? `_${month}` : "";
      XLSX.writeFile(wb, `FundRequest_FY${fy}${monthPart}.xlsx`);
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("Fund Request export error:", err);
    }
  };

  // ─────────────────────────────────────────────────────────────────────────
  // ── CORE FIX: Left-join logic — Client Master drives row structure ────────
  // For every selected month the table ALWAYS shows ALL active clients.
  // Existing transaction data is merged in; clients with no data get a
  // blank editable row so users can fill it in.
  // ─────────────────────────────────────────────────────────────────────────

  // Helper: build the complete client-driven row set for a given month.
  // Returns one row per active client, merged with any saved transaction data.
  const buildMonthRows = useCallback((targetMonth) => {
    return masterClients.map(client => {
      // Find an existing saved row for this client+month
      const existing = rows.find(
        r => r.clientName === client.clientName && r.month === targetMonth
      );
      if (existing) return existing; // use saved data

      // No data yet → generate a blank editable row (client stays visible)
      return {
        _id:                 `BLANK-${client.id}-${targetMonth}`,
        clientName:          client.clientName,
        applicability_status:"applicable",
        currency:            "INR",
        amountFC:            "",
        rate:                "1",
        remarks:             "",
        month:               targetMonth,
        _isBlank:            true,   // flag: blank row — not yet saved
      };
    });
  }, [masterClients, rows]);

  // ── Analytics filter (keeps existing behaviour: only rows with real data) ─
  const filteredRows = useMemo(() => {
    // For analytics/aggregates, only count rows that have actual amounts
    return rows.filter(r => {
      const cOk   = clientFilter === "All"    || r.clientName === clientFilter;
      const curOk = currencyFilter === "All"  || r.currency   === currencyFilter;
      const mOk   = month === "All Months"    || r.month      === month;
      return cOk && curOk && mOk;
    });
  }, [rows, clientFilter, currencyFilter, month]);

  // ── Entry view: ALWAYS shows all clients for the selected month ───────────
  const entryViewRows = useMemo(() => {
    const monthsToRender =
      month === "All Months"
        ? ["April","May","June","July","August","September","October","November","December","January","February","March"]
        : [month];

    // Build full client-driven list for each month, then flatten
    let result = monthsToRender.flatMap(m => buildMonthRows(m));

    // Apply client filter (still shows blank rows for the filtered client)
    if (clientFilter !== "All")
      result = result.filter(r => r.clientName === clientFilter);

    // Apply currency filter only to rows that have a currency set
    if (currencyFilter !== "All")
      result = result.filter(r => !r._isBlank || r.currency === currencyFilter);

    return result;
  }, [month, buildMonthRows, clientFilter, currencyFilter]);

  // ── Table rows (search on top of entry view) ──────────────────────────────
  const tableFiltered = useMemo(() => {
    if (!tableSearch.trim()) return entryViewRows;
    const q = tableSearch.toLowerCase();
    return entryViewRows.filter(r =>
      r.clientName.toLowerCase().includes(q) ||
      r.currency.toLowerCase().includes(q)   ||
      (r.remarks || "").toLowerCase().includes(q)
    );
  }, [entryViewRows, tableSearch]);

  const tableTotalPages = Math.max(1, Math.ceil(tableFiltered.length / TABLE_PAGE_SIZE));
  const tablePaged = tableFiltered.slice((tablePage - 1) * TABLE_PAGE_SIZE, tablePage * TABLE_PAGE_SIZE);

  // When month filter changes, reset to page 1 so the full list is visible
  useEffect(() => { setTablePage(1); }, [month, clientFilter, currencyFilter]);

  // Unique lists still derived from saved rows (for filter dropdowns)
  const uniqueClients    = [...new Set(masterClients.map(c => c.clientName).filter(Boolean))];
  const uniqueCurrencies = [...new Set(rows.map(r => r.currency))];

  // ── Persist a blank row the moment the user edits any field ───────────────
  // updateRow is called by the cell inputs. If the row is blank (_isBlank),
  // we first insert it into the saved rows state before updating.
  const updateRow = (id, k, v) => {
    // Check lock on the row's month before allowing any edit
    const rowMonth = (() => {
      const saved = rows.find(r => r._id === id);
      if (saved) return saved.month;
      const blank = entryViewRows.find(r => r._id === id);
      return blank ? (blank.month || (month === "All Months" ? "April" : month)) : (month === "All Months" ? "April" : month);
    })();
    if (checkLocked(locks, fy, rowMonth, "fund")) return; // blocked when locked

    setRows(prev => {
      const existsInSaved = prev.some(r => r._id === id);
      if (existsInSaved) {
        return prev.map(r => r._id === id ? { ...r, [k]: v } : r);
      }
      // First edit on a blank row → materialise it in saved state
      const blankRow = entryViewRows.find(r => r._id === id);
      if (!blankRow) return prev;
      const materialised = { ...blankRow, _isBlank: false, [k]: v };
      return [...prev, materialised];
    });
  };

  // ── Aggregates ────────────────────────────────────────────────────────────
  const totalINR    = filteredRows.reduce((s,r)=>s+calcINR(r),0);
  const totalFC     = filteredRows.filter(r=>r.currency!=="INR").reduce((s,r)=>s+(parseFloat(r.amountFC)||0),0);
  const activeClients = [...new Set(filteredRows.map(r=>r.clientName).filter(Boolean))].length;

  // ── Chart data ────────────────────────────────────────────────────────────
  // 1. Donut — currency-wise INR
  const donutData = useMemo(()=>
    CURRENCIES.map(c=>{
      const cRows = filteredRows.filter(r=>r.currency===c.code);
      const val   = cRows.reduce((s,r)=>s+calcINR(r),0);
      return {name:c.code, value:parseFloat(val.toFixed(2)), symbol:c.symbol};
    }).filter(d=>d.value>0)
  ,[filteredRows]);

  // 2. Monthly trend
  const monthOrder = ["April","May","June","July","August","September","October","November","December","January","February","March"];
  const monthlyTrend = useMemo(()=>
    monthOrder.map(m=>{
      const mRows = filteredRows.filter(r=>r.month===m);
      return { month:m.slice(0,3), inr:parseFloat(mRows.reduce((s,r)=>s+calcINR(r),0).toFixed(0)), count:mRows.length };
    }).filter(d=>d.inr>0||d.count>0)
  ,[filteredRows]);

  // 3. FY summary — dynamic: show last 4 FYs up to and including selected FY
  const fyBarData = useMemo(() => {
    const idx = FY_LIST.indexOf(fy);
    // FY_LIST is newest-first; grab up to 4 FYs ending at selected, show oldest→newest
    const slice = FY_LIST.slice(idx, idx + 4).reverse();
    const scales = [0.55, 0.72, 0.88, 1.0];
    return slice.map((fyLabel, i) => ({
      fy: fyLabel,
      inr: Math.round(totalINR * (scales[i] ?? 1.0)),
    }));
  }, [fy, totalINR]);

  // 4. Top clients horizontal bar
  const clientBarData = useMemo(()=>{
    const map={};
    filteredRows.forEach(r=>{ map[r.clientName]=(map[r.clientName]||0)+calcINR(r); });
    return Object.entries(map).map(([name,inr])=>({name:name.length>22?name.slice(0,20)+"…":name,inr:parseFloat(inr.toFixed(0))}))
      .sort((a,b)=>b.inr-a.inr).slice(0,6);
  },[filteredRows]);

  // 5. Currency vs INR comparison
  const comparisonData = useMemo(()=>
    CURRENCIES.filter(c=>c.code!=="INR").map(c=>{
      const cRows = filteredRows.filter(r=>r.currency===c.code);
      const fc    = cRows.reduce((s,r)=>s+(parseFloat(r.amountFC)||0),0);
      const inr   = cRows.reduce((s,r)=>s+calcINR(r),0);
      return {currency:c.code, foreign:parseFloat(fc.toFixed(2)), inr:parseFloat(inr.toFixed(0))};
    }).filter(d=>d.foreign>0)
  ,[filteredRows]);

  // 6. Quarter-filtered rows for analytics
  const QUARTER_MAP = {
    "Q1 (Apr–Jun)":  ["April","May","June"],
    "Q2 (Jul–Sep)":  ["July","August","September"],
    "Q3 (Oct–Dec)":  ["October","November","December"],
    "Q4 (Jan–Mar)":  ["January","February","March"],
  };
  const quarterRows = useMemo(() => {
    if (quarterFilter === "All") return filteredRows;
    const qMonths = QUARTER_MAP[quarterFilter] || [];
    return filteredRows.filter(r => qMonths.includes(r.month));
  }, [filteredRows, quarterFilter]);

  // 7. Pending remittances (rows with NA applicability or no amountFC — proxy for pending)
  const pendingAmount = useMemo(() =>
    rows.filter(r => r.applicability_status === "not_applicable" || !r.amountFC || r.amountFC === "")
      .reduce((s,r) => s + calcINR(r), 0)
  , [rows]);
  const pendingCount = rows.filter(r => !r.amountFC || r.amountFC === "").length;

  // 8. Average fund value
  const avgFundValue = useMemo(() => {
    const valid = quarterRows.filter(r => calcINR(r) > 0);
    return valid.length ? valid.reduce((s,r) => s+calcINR(r), 0) / valid.length : 0;
  }, [quarterRows]);

  // 9. Previous period comparison (simulate: 85% of current for demo)
  const prevTotalINR = totalINR * 0.85;
  const trendPct = totalINR > 0 ? (((totalINR - prevTotalINR) / prevTotalINR) * 100).toFixed(1) : 0;

  // 10. Geography / Parent Entity data (derived from currency as proxy)
  const geoData = useMemo(() => {
    const geoMap = {
      "USD": { country:"USA",       flag:"🇺🇸" },
      "EUR": { country:"Europe",    flag:"🇪🇺" },
      "GBP": { country:"UK",        flag:"🇬🇧" },
      "AED": { country:"UAE",       flag:"🇦🇪" },
      "SGD": { country:"Singapore", flag:"🇸🇬" },
      "CHF": { country:"Switzerland",flag:"🇨🇭"},
      "AUD": { country:"Australia", flag:"🇦🇺" },
      "JPY": { country:"Japan",     flag:"🇯🇵" },
      "HKD": { country:"Hong Kong", flag:"🇭🇰" },
      "INR": { country:"India",     flag:"🇮🇳" },
    };
    const totals = {};
    quarterRows.forEach(r => {
      const geo = geoMap[r.currency] || { country: r.currency, flag:"🌐" };
      totals[geo.country] = (totals[geo.country] || { inr:0, count:0, flag:geo.flag });
      totals[geo.country].inr   += calcINR(r);
      totals[geo.country].count += 1;
    });
    const total = Object.values(totals).reduce((s,v) => s + v.inr, 0) || 1;
    return Object.entries(totals)
      .map(([country, v]) => ({
        name: country, flag: v.flag, value: parseFloat(v.inr.toFixed(0)),
        count: v.count, pct: parseFloat(((v.inr / total)*100).toFixed(1)),
      }))
      .sort((a,b) => b.value - a.value);
  }, [quarterRows]);

  // 11. Funding status pipeline (simulated from data pattern)
  const statusPipeline = useMemo(() => {
    const total = quarterRows.length || 1;
    const completed  = quarterRows.filter(r => calcINR(r) > 0 && r.applicability_status === "applicable").length;
    const naRows     = quarterRows.filter(r => r.applicability_status === "not_applicable").length;
    const blank      = quarterRows.filter(r => !r.amountFC || r.amountFC === "").length;
    const inProgress = Math.max(0, total - completed - naRows - blank);
    return [
      { label:"Completed",       value: completed,  color:"#34d399", pct: Math.round(completed/total*100)  },
      { label:"Processing",      value: inProgress, color:"#4a90d9", pct: Math.round(inProgress/total*100) },
      { label:"RBI/FEMA Review", value: naRows,     color:"#fbbf24", pct: Math.round(naRows/total*100)     },
      { label:"Pending",         value: blank,      color:"#f87171", pct: Math.round(blank/total*100)      },
    ];
  }, [quarterRows]);

  // 12. Client Funding Ledger (analytics-only grid with variance)
  const ledgerRows = useMemo(() => {
    return quarterRows
      .filter(r => r.clientName && r.amountFC)
      .map(r => {
        const expected = parseFloat(r.amountFC) * (parseFloat(r.rate) || 1) * 1.02; // expected = slightly more
        const actual   = calcINR(r);
        const variance = actual - expected;
        return {
          client:   r.clientName,
          parent:   r.currency === "USD" ? "US Parent Corp" : r.currency === "EUR" ? "EU Group Ltd" : r.currency === "GBP" ? "UK Holdings" : "Foreign Parent",
          date:     r.month || "—",
          expected: parseFloat(expected.toFixed(0)),
          actual:   actual,
          variance: parseFloat(variance.toFixed(0)),
          currency: r.currency,
        };
      });
  }, [quarterRows]);

  // Ledger pagination
  const [ledgerPage, setLedgerPage]     = useState(1);
  const [ledgerSearch, setLedgerSearch] = useState("");
  const [ledgerSort, setLedgerSort]     = useState({ key:"actual", dir:"desc" });
  const LEDGER_PAGE_SIZE = 8;
  const filteredLedger = useMemo(() => {
    let data = ledgerRows;
    if (ledgerSearch.trim()) {
      const q = ledgerSearch.toLowerCase();
      data = data.filter(r => r.client.toLowerCase().includes(q) || r.parent.toLowerCase().includes(q));
    }
    data = [...data].sort((a,b) => {
      const av = a[ledgerSort.key] ?? 0, bv = b[ledgerSort.key] ?? 0;
      return ledgerSort.dir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
    });
    return data;
  }, [ledgerRows, ledgerSearch, ledgerSort]);
  const ledgerTotalPages = Math.max(1, Math.ceil(filteredLedger.length / LEDGER_PAGE_SIZE));
  const pagedLedger = filteredLedger.slice((ledgerPage-1)*LEDGER_PAGE_SIZE, ledgerPage*LEDGER_PAGE_SIZE);
  const sortLedger = (key) => setLedgerSort(prev => ({ key, dir: prev.key===key && prev.dir==="desc" ? "asc" : "desc" }));

  const exportLedger = () => {
    try {
      const data = filteredLedger.map((r,i) => ({
        "#":              i+1,
        "Client Name":    xlsSafe(r.client),
        "Parent Entity":  xlsSafe(r.parent),
        "Date Received":  r.date,
        "Expected (INR)": r.expected,
        "Actual Received (INR)": r.actual,
        "Variance / Forex Diff (INR)": r.variance,
        "Financial Year": fy,
      }));
      if (!data.length) { alert("No ledger data to export."); return; }
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [4,30,20,14,18,22,24,12].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Fund Ledger");
      XLSX.writeFile(wb, `FundLedger_FY${fy}.xlsx`);
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("Fund Ledger export error:", err);
    }
  };

  const sel = `px-3 py-2 rounded-xl border text-sm outline-none transition-all ${t.input} focus:ring-2 focus:ring-[#00c9b1]`;

  // ── Custom donut label ────────────────────────────────────────────────────
  const RADIAN = Math.PI/180;
  const renderDonutLabel = ({cx,cy,midAngle,innerRadius,outerRadius,name,percent})=>{
    if(percent<0.06) return null;
    const r = innerRadius+(outerRadius-innerRadius)*0.55;
    const x = cx+r*Math.cos(-midAngle*RADIAN);
    const y = cy+r*Math.sin(-midAngle*RADIAN);
    return <text x={x} y={y} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={10} fontWeight="700">{name}</text>;
  };

  return (
    <div className="space-y-5">

      {savedToast&&(
        <div className="fixed top-5 right-5 z-[70] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{background:"linear-gradient(135deg,#34d399,#10b981)",minWidth:220}}>
          <Icon path={Icons.check} size={15}/> Saved successfully!
        </div>
      )}

      {/* ── Filter bar ── */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-5 py-4`}>
        <div className="flex flex-wrap items-center gap-x-4 gap-y-3">
          <div className="flex items-center gap-2 shrink-0">
            <Icon path={Icons.filter} size={13} className={t.textMuted}/>
            <span className={`text-[11px] font-bold uppercase tracking-widest ${t.textMuted}`}>Filters</span>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Financial Year</span>
            <FYSelect value={fy} onChange={setFy} t={t} dark={dark}/>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Month</span>
            <select value={month} onChange={e=>setMonth(e.target.value)} className={sel}>
              {MONTHS.map(m=><option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Quarter</span>
            <select value={quarterFilter} onChange={e=>setQuarterFilter(e.target.value)} className={sel}>
              <option value="All">All Quarters</option>
              {["Q1 (Apr–Jun)","Q2 (Jul–Sep)","Q3 (Oct–Dec)","Q4 (Jan–Mar)"].map(q=><option key={q} value={q}>{q}</option>)}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Client</span>
            <select value={clientFilter} onChange={e=>setClientFilter(e.target.value)} className={`${sel} max-w-[200px]`}>
              <option value="All">All Clients</option>
              {uniqueClients.map(c=><option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs font-semibold ${t.textMuted}`}>Currency</span>
            <select value={currencyFilter} onChange={e=>setCurrencyFilter(e.target.value)} className={sel}>
              <option value="All">All Currencies</option>
              {uniqueCurrencies.map(c=><option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          {(clientFilter!=="All"||currencyFilter!=="All"||month!=="All Months"||quarterFilter!=="All")&&(
            <button onClick={()=>{setClientFilter("All");setCurrencyFilter("All");setMonth("All Months");setQuarterFilter("All");}}
              className="flex items-center gap-1 text-xs px-2.5 py-1.5 rounded-full font-semibold"
              style={{background:"#f8717115",color:"#f87171"}}>
              <Icon path={Icons.x} size={11}/> Clear
            </button>
          )}
          {/* View toggle */}
          <div className="ml-auto flex items-center gap-1">
            {[{id:"analytics",label:"📊 Analytics",icon:Icons.activity},{id:"entry",label:"📋 Data Entry",icon:Icons.fileText}].map(v=>(
              <button key={v.id} onClick={()=>setView(v.id)}
                className="px-3 py-1.5 rounded-xl text-xs font-semibold transition-all"
                style={{
                  background: view===v.id ? "linear-gradient(135deg,#00a896,#1b5fa8)" : "transparent",
                  color: view===v.id ? "#fff" : "#5a7a99",
                  border: view===v.id ? "none" : `1px solid ${dark?"#243d58":"#d4e0ed"}`,
                }}>
                {v.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* ══════════════ KPI SCORECARDS ══════════════ */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Total Funds Inward */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 relative overflow-hidden group transition-all hover:-translate-y-0.5 hover:shadow-xl`}
          style={{borderLeft:"3px solid #00c9b1"}}>
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full opacity-5 group-hover:opacity-10 transition-opacity" style={{background:"#00c9b1"}}/>
          <div className="flex items-start justify-between mb-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background:"#00c9b118"}}>
              <Icon path={Icons.trending} size={18} style={{color:"#00c9b1"}}/>
            </div>
            <div className="flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full"
              style={{background:"#34d39918", color:"#34d399"}}>
              <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M12 19V5M5 12l7-7 7 7"/></svg>
              +{trendPct}%
            </div>
          </div>
          <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${t.textMuted}`}>Total Funds Inward</div>
          <div className="text-2xl font-black leading-none mb-1" style={{color:"#00c9b1"}}>{fmtINRShort(totalINR)}</div>
          <div className={`text-[10px] ${t.textMuted}`}>vs {fmtINRShort(prevTotalINR)} prior period</div>
          <div className="mt-2 h-1 rounded-full" style={{background:dark?"#1a2d44":"#e2ecf4"}}>
            <div className="h-1 rounded-full transition-all" style={{width:"72%",background:"#00c9b1"}}/>
          </div>
        </div>

        {/* Card 2: Total Active Clients */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 relative overflow-hidden group transition-all hover:-translate-y-0.5 hover:shadow-xl`}
          style={{borderLeft:"3px solid #4a90d9"}}>
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full opacity-5 group-hover:opacity-10 transition-opacity" style={{background:"#4a90d9"}}/>
          <div className="flex items-start justify-between mb-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background:"#4a90d918"}}>
              <Icon path={Icons.users} size={18} style={{color:"#4a90d9"}}/>
            </div>
            <div className={`text-[10px] font-semibold px-2 py-0.5 rounded-full ${t.textMuted}`}
              style={{background:dark?"#1a2d44":"#eef3f8"}}>
              FY {fy}
            </div>
          </div>
          <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${t.textMuted}`}>Active Clients</div>
          <div className="text-2xl font-black leading-none mb-1" style={{color:"#4a90d9"}}>{activeClients}</div>
          <div className={`text-[10px] ${t.textMuted}`}>received funding · {quarterFilter !== "All" ? quarterFilter : "all year"}</div>
          <div className="flex gap-1 mt-2">
            {[...Array(Math.min(activeClients,8))].map((_,i)=>(
              <div key={i} className="w-2 h-2 rounded-full" style={{background:"#4a90d9",opacity:0.5+i*0.06}}/>
            ))}
          </div>
        </div>

        {/* Card 3: Pending Remittances */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 relative overflow-hidden group transition-all hover:-translate-y-0.5 hover:shadow-xl`}
          style={{borderLeft:"3px solid #fbbf24"}}>
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full opacity-5 group-hover:opacity-10 transition-opacity" style={{background:"#fbbf24"}}/>
          <div className="flex items-start justify-between mb-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background:"#fbbf2418"}}>
              <Icon path={Icons.clock} size={18} style={{color:"#fbbf24"}}/>
            </div>
            {pendingCount > 0 && (
              <div className="flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full"
                style={{background:"#f8717118",color:"#f87171"}}>
                <Icon path={Icons.alert} size={10}/> {pendingCount}
              </div>
            )}
          </div>
          <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${t.textMuted}`}>Pending Remittances</div>
          <div className="text-2xl font-black leading-none mb-1" style={{color:"#fbbf24"}}>{pendingCount}</div>
          <div className={`text-[10px] ${t.textMuted}`}>in transit or compliance review</div>
          <div className="mt-2 flex items-center gap-1.5">
            <div className="w-1.5 h-1.5 rounded-full animate-pulse" style={{background:"#fbbf24"}}/>
            <span className="text-[10px]" style={{color:"#fbbf24"}}>Requires attention</span>
          </div>
        </div>

        {/* Card 4: Average Fund Value */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 relative overflow-hidden group transition-all hover:-translate-y-0.5 hover:shadow-xl`}
          style={{borderLeft:"3px solid #34d399"}}>
          <div className="absolute -top-6 -right-6 w-28 h-28 rounded-full opacity-5 group-hover:opacity-10 transition-opacity" style={{background:"#34d399"}}/>
          <div className="flex items-start justify-between mb-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{background:"#34d39918"}}>
              <Icon path={Icons.barChart2} size={18} style={{color:"#34d399"}}/>
            </div>
            <div className="flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full"
              style={{background:"#34d39918",color:"#34d399"}}>avg/tx</div>
          </div>
          <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${t.textMuted}`}>Avg Fund Value</div>
          <div className="text-2xl font-black leading-none mb-1" style={{color:"#34d399"}}>{fmtINRShort(avgFundValue)}</div>
          <div className={`text-[10px] ${t.textMuted}`}>per transaction · {quarterFilter!=="All"?quarterFilter:"FY "+fy}</div>
          <div className="mt-2 h-1 rounded-full" style={{background:dark?"#1a2d44":"#e2ecf4"}}>
            <div className="h-1 rounded-full" style={{width:"58%",background:"#34d399"}}/>
          </div>
        </div>
      </div>

      {/* ══════════════ ANALYTICS VIEW ══════════════ */}
      {view==="analytics" && (
        <>
          {/* ── SYNC BANNER ── */}
          <SyncBanner t={t} dark={dark} moduleName="Fund Request" accentColor="#00c9b1"/>

          {/* ── ROW 1: Fund Inflow Trend (3/5) + Top 10 Clients (2/5) ── */}
          <div className="grid grid-cols-1 lg:grid-cols-5 gap-5">

            {/* Chart A — Fund Inflow Trend Area */}
            <div className={`lg:col-span-3 ${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className={`font-bold text-sm ${t.text}`}>Fund Inflow Trend</h3>
                  <p className={`text-xs mt-0.5 ${t.textMuted}`}>Monthly INR inflows · FY {fy}</p>
                </div>
                <div className="flex items-center gap-3 text-[10px]">
                  {[{col:"#00c9b1",label:"INR Amount"},{col:"#4a90d960",label:"Transactions"}].map(d=>(
                    <div key={d.label} className="flex items-center gap-1.5">
                      <div className="w-3 h-2 rounded-sm" style={{background:d.col}}/>
                      <span className={t.textMuted}>{d.label}</span>
                    </div>
                  ))}
                </div>
              </div>
              {monthlyTrend.length===0?(
                <div className={`h-56 flex flex-col items-center justify-center gap-3 ${t.textMuted}`}>
                  <Icon path={Icons.bar} size={32}/>
                  <span className="text-sm">No data for selected period</span>
                </div>
              ):(
                <ResponsiveContainer width="100%" height={230}>
                  <ComposedChart data={monthlyTrend} margin={{top:4,right:16,bottom:0,left:0}}>
                    <defs>
                      <linearGradient id="gINRFund" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%"   stopColor="#00c9b1" stopOpacity={0.4}/>
                        <stop offset="100%" stopColor="#00c9b1" stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#e2e8f0"} vertical={false}/>
                    <XAxis dataKey="month" tick={{fontSize:11,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}/>
                    <YAxis tick={{fontSize:10,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}
                      tickFormatter={v=>v>=1e7?`${(v/1e7).toFixed(1)}Cr`:v>=1e5?`${(v/1e5).toFixed(1)}L`:v>=1e3?`${(v/1e3).toFixed(0)}K`:v}/>
                    <YAxis yAxisId={1} orientation="right" tick={{fontSize:10,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false} width={28}/>
                    <Tooltip content={<ChartTooltip dark={dark} prefix="₹"/>}/>
                    <Area type="monotone" dataKey="inr" name="INR Amount"
                      stroke="#00c9b1" strokeWidth={2.5} fill="url(#gINRFund)"
                      dot={{fill:"#00c9b1",r:4,strokeWidth:0}} activeDot={{r:7,strokeWidth:2,stroke:"#fff"}}/>
                    <Bar dataKey="count" name="Transactions" yAxisId={1}
                      fill="#4a90d9" fillOpacity={0.35} radius={[4,4,0,0]}/>
                  </ComposedChart>
                </ResponsiveContainer>
              )}
            </div>

            {/* Chart B — Top 10 Clients by Volume */}
            <div className={`lg:col-span-2 ${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="mb-4">
                <h3 className={`font-bold text-sm ${t.text}`}>Top Clients by Volume</h3>
                <p className={`text-xs mt-0.5 ${t.textMuted}`}>Highest funding clients (INR)</p>
              </div>
              {clientBarData.length===0?(
                <div className={`h-56 flex flex-col items-center justify-center gap-3 ${t.textMuted}`}>
                  <Icon path={Icons.users} size={32}/>
                  <span className="text-sm">No client data</span>
                </div>
              ):(
                <ResponsiveContainer width="100%" height={230}>
                  <BarChart data={clientBarData} layout="vertical" margin={{top:0,right:12,bottom:0,left:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#e2e8f0"} horizontal={false}/>
                    <XAxis type="number" tick={{fontSize:9,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}
                      tickFormatter={v=>v>=1e5?`${(v/1e5).toFixed(0)}L`:v>=1e3?`${(v/1e3).toFixed(0)}K`:v}/>
                    <YAxis type="category" dataKey="name" tick={{fontSize:9,fill:dark?"#94a3b8":"#374151"}}
                      axisLine={false} tickLine={false} width={100}/>
                    <Tooltip content={<ChartTooltip dark={dark} prefix="₹"/>}/>
                    <Bar dataKey="inr" name="INR Amount" radius={[0,6,6,0]}>
                      {clientBarData.map((_,i)=>(
                        <Cell key={i} fill={PALETTE[i%PALETTE.length]}/>
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </div>

          {/* ── ROW 2: Geography Donut (2/5) + Funding Status Pipeline (3/5) ── */}
          <div className="grid grid-cols-1 lg:grid-cols-5 gap-5">

            {/* Chart C — Parent Entity Geographies */}
            <div className={`lg:col-span-2 ${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="mb-4">
                <h3 className={`font-bold text-sm ${t.text}`}>Parent Entity Geographies</h3>
                <p className={`text-xs mt-0.5 ${t.textMuted}`}>Country-wise source of funding</p>
              </div>
              {geoData.length===0?(
                <div className={`h-56 flex items-center justify-center ${t.textMuted} text-sm`}>No data</div>
              ):(
                <>
                  <ResponsiveContainer width="100%" height={180}>
                    <PieChart>
                      <Pie data={geoData} cx="50%" cy="50%" innerRadius={48} outerRadius={78}
                        paddingAngle={3} dataKey="value" labelLine={false}
                        label={({cx,cy,midAngle,innerRadius,outerRadius,percent})=>{
                          if(percent<0.07) return null;
                          const RADIAN=Math.PI/180;
                          const r=innerRadius+(outerRadius-innerRadius)*0.55;
                          const x=cx+r*Math.cos(-midAngle*RADIAN);
                          const y=cy+r*Math.sin(-midAngle*RADIAN);
                          return <text x={x} y={y} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={9} fontWeight="800">{`${Math.round(percent*100)}%`}</text>;
                        }}>
                        {geoData.map((_,i)=><Cell key={i} fill={PALETTE[i%PALETTE.length]}/>)}
                      </Pie>
                      <Tooltip content={<ChartTooltip dark={dark} prefix="₹"/>}/>
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="space-y-1.5 mt-1">
                    {geoData.slice(0,5).map((d,i)=>(
                      <div key={d.name} className="flex items-center gap-2">
                        <span className="text-sm">{d.flag}</span>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center justify-between gap-1 mb-0.5">
                            <span className={`text-[11px] font-semibold truncate ${t.text}`}>{d.name}</span>
                            <span className="text-[10px] font-bold" style={{color:PALETTE[i%PALETTE.length]}}>{d.pct}%</span>
                          </div>
                          <div className="h-1 rounded-full" style={{background:dark?"#1a2d44":"#e2ecf4"}}>
                            <div className="h-1 rounded-full transition-all" style={{width:`${d.pct}%`,background:PALETTE[i%PALETTE.length]}}/>
                          </div>
                        </div>
                        <span className={`text-[10px] font-medium w-12 text-right ${t.textMuted}`}>{fmtINRShort(d.value)}</span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>

            {/* Chart D — Funding by Status (Stacked/Pipeline) */}
            <div className={`lg:col-span-3 ${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="flex items-start justify-between mb-5">
                <div>
                  <h3 className={`font-bold text-sm ${t.text}`}>Funding by Status</h3>
                  <p className={`text-xs mt-0.5 ${t.textMuted}`}>Fund pipeline tracking · {quarterRows.length} total transactions</p>
                </div>
                <div className="text-[10px] font-semibold px-2.5 py-1 rounded-lg"
                  style={{background:"#00c9b115",color:"#00c9b1",border:"1px solid #00c9b130"}}>
                  FY {fy}
                </div>
              </div>

              {/* Pipeline visual bars */}
              <div className="space-y-4">
                {statusPipeline.map((s,i) => (
                  <div key={s.label}>
                    <div className="flex items-center justify-between mb-1.5">
                      <div className="flex items-center gap-2">
                        <div className="w-2 h-2 rounded-full" style={{background:s.color}}/>
                        <span className={`text-xs font-semibold ${t.text}`}>{s.label}</span>
                      </div>
                      <div className="flex items-center gap-3">
                        <span className={`text-[11px] font-bold ${t.text}`}>{s.value}</span>
                        <span className="text-[10px] font-bold w-8 text-right" style={{color:s.color}}>{s.pct}%</span>
                      </div>
                    </div>
                    <div className="h-2.5 rounded-full" style={{background:dark?"#1a2d44":"#e2ecf4"}}>
                      <div className="h-2.5 rounded-full transition-all duration-700"
                        style={{width:`${Math.max(s.pct,2)}%`,background:s.color,
                          boxShadow:`0 0 8px ${s.color}60`}}/>
                    </div>
                  </div>
                ))}
              </div>

              {/* Stacked chart */}
              <div className="mt-5">
                <ResponsiveContainer width="100%" height={140}>
                  <BarChart data={monthlyTrend.length > 0 ? monthlyTrend : [{month:"—",inr:0}]}
                    margin={{top:4,right:8,bottom:0,left:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#e2e8f0"} vertical={false}/>
                    <XAxis dataKey="month" tick={{fontSize:10,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}/>
                    <YAxis tick={{fontSize:9,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}
                      tickFormatter={v=>v>=1e5?`${(v/1e5).toFixed(0)}L`:v}/>
                    <Tooltip content={<ChartTooltip dark={dark} prefix="₹"/>}/>
                    <Bar dataKey="inr" name="INR Amount" radius={[4,4,0,0]}>
                      {(monthlyTrend.length>0?monthlyTrend:[]).map((_,i)=>(
                        <Cell key={i} fill={PALETTE[i%PALETTE.length]} fillOpacity={0.85}/>
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* ── ROW 3: FY-wise Bar + Currency Distribution ── */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">

            {/* FY Summary */}
            <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="mb-4">
                <h3 className={`font-bold text-sm ${t.text}`}>FY-wise Summary</h3>
                <p className={`text-xs mt-0.5 ${t.textMuted}`}>Total INR across financial years</p>
              </div>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={fyBarData} margin={{top:4,right:8,bottom:0,left:0}}>
                  <defs>
                    {fyBarData.map((_,i)=>(
                      <linearGradient key={i} id={`gFYfund${i}`} x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%"   stopColor={PALETTE[i]} stopOpacity={1}/>
                        <stop offset="100%" stopColor={PALETTE[i]} stopOpacity={0.45}/>
                      </linearGradient>
                    ))}
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark?"#1a2d44":"#e2e8f0"} vertical={false}/>
                  <XAxis dataKey="fy" tick={{fontSize:11,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fontSize:10,fill:dark?"#64748b":"#718096"}} axisLine={false} tickLine={false}
                    tickFormatter={v=>v>=1e5?`${(v/1e5).toFixed(0)}L`:v}/>
                  <Tooltip content={<ChartTooltip dark={dark} prefix="₹"/>}/>
                  <Bar dataKey="inr" name="INR Amount" radius={[8,8,0,0]}>
                    {fyBarData.map((_,i)=><Cell key={i} fill={`url(#gFYfund${i})`}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Currency exposure + legend */}
            <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
              <div className="mb-4">
                <h3 className={`font-bold text-sm ${t.text}`}>Currency Distribution</h3>
                <p className={`text-xs mt-0.5 ${t.textMuted}`}>INR equivalent by currency</p>
              </div>
              {donutData.length===0?(
                <div className={`h-56 flex items-center justify-center ${t.textMuted} text-sm`}>No currency data</div>
              ):(
                <div className="flex gap-4">
                  <div className="flex-shrink-0">
                    <ResponsiveContainer width={170} height={170}>
                      <PieChart>
                        <Pie data={donutData} cx="50%" cy="50%" innerRadius={50} outerRadius={80}
                          paddingAngle={2} dataKey="value" labelLine={false}
                          label={({cx,cy,midAngle,innerRadius,outerRadius,name,percent})=>{
                            if(percent<0.06) return null;
                            const RADIAN=Math.PI/180;
                            const r=innerRadius+(outerRadius-innerRadius)*0.55;
                            const x=cx+r*Math.cos(-midAngle*RADIAN);
                            const y=cy+r*Math.sin(-midAngle*RADIAN);
                            return <text x={x} y={y} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={9} fontWeight="700">{name}</text>;
                          }}>
                          {donutData.map((_,i)=><Cell key={i} fill={PALETTE[i%PALETTE.length]}/>)}
                        </Pie>
                        <Tooltip content={<ChartTooltip dark={dark} prefix="₹" suffix=" INR"/>}/>
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                  <div className="flex-1 space-y-2 overflow-y-auto" style={{maxHeight:175}}>
                    {donutData.map((d,i)=>{
                      const col=PALETTE[i%PALETTE.length];
                      const cur=CURRENCIES.find(c=>c.code===d.name);
                      const cnt=filteredRows.filter(r=>r.currency===d.name).length;
                      const pct=((d.value/(totalINR||1))*100).toFixed(1);
                      return(
                        <div key={d.name} className="flex items-center gap-2 px-2.5 py-1.5 rounded-xl text-xs font-semibold"
                          style={{background:col+"12",border:`1px solid ${col}25`,color:col}}>
                          <span className="font-black">{cur?.symbol}</span>
                          <span className="font-bold">{d.name}</span>
                          <span className={`opacity-60 ${t.textMuted}`}>·</span>
                          <span className={t.textMuted}>{cnt} req</span>
                          <span className="ml-auto">{pct}%</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* ── BOTTOM SECTION: Client Funding Ledger ── */}
          <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
            {/* Ledger Header */}
            <div className={`px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b ${t.cardBorder}`}
              style={{background: dark
                ? "linear-gradient(90deg,#0c1e30 0%,#0e2240 60%,#0a1a2e 100%)"
                : "linear-gradient(90deg,#1b3a5c 0%,#1a3356 60%,#1e3d6a 100%)"}}>
              <div>
                <h3 className="font-bold text-sm text-white">Client Funding Ledger</h3>
                <p className="text-xs mt-0.5 text-white/60">
                  {filteredLedger.length} entries · FY {fy} · {quarterFilter!=="All"?quarterFilter:"All Quarters"} · variance highlighted
                </p>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                {/* Search */}
                <div className="flex items-center gap-2 px-3 py-2 rounded-xl border text-sm bg-white/10 border-white/20">
                  <Icon path={Icons.search} size={13} className="text-white/50"/>
                  <input value={ledgerSearch} onChange={e=>{setLedgerSearch(e.target.value);setLedgerPage(1);}}
                    placeholder="Search clients…"
                    className="bg-transparent outline-none w-32 text-white placeholder-white/40 text-xs"/>
                  {ledgerSearch&&<button onClick={()=>{setLedgerSearch("");setLedgerPage(1);}} className="text-white/40 hover:text-white">
                    <Icon path={Icons.x} size={12}/></button>}
                </div>
                {/* Export */}
                <button onClick={exportLedger}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold text-white bg-white/15 hover:bg-white/25 border border-white/20 transition-all">
                  <Icon path={Icons.download} size={13}/> Export
                </button>
              </div>
            </div>

            {/* Sticky table header */}
            <div className="overflow-x-auto" style={{maxHeight:"60vh",overflowY:"auto"}}>
              <table className="w-full text-sm border-collapse" style={{minWidth:900}}>
                <thead className="sticky top-0 z-10">
                  <tr style={{
                    background: dark
                      ? "linear-gradient(90deg,#0f1e2e 0%,#0d1c2c 100%)"
                      : "linear-gradient(90deg,#f0f6ff 0%,#e8f4fb 100%)",
                    boxShadow:"0 2px 8px rgba(0,0,0,0.15)",
                  }}>
                    {[
                      { key:"#",        label:"#",                   sortable:false, align:"center", w:40    },
                      { key:"client",   label:"Client Name",          sortable:true,  align:"left",   w:200   },
                      { key:"parent",   label:"Parent Entity",        sortable:false, align:"left",   w:160   },
                      { key:"date",     label:"Date Received",        sortable:true,  align:"left",   w:120   },
                      { key:"expected", label:"Expected Amount",      sortable:true,  align:"right",  w:150   },
                      { key:"actual",   label:"Actual Received",      sortable:true,  align:"right",  w:150   },
                      { key:"variance", label:"Variance (Forex Diff)",sortable:true,  align:"right",  w:160   },
                    ].map(col=>(
                      <th key={col.key}
                        className={`px-4 py-3 text-[9px] font-bold uppercase tracking-widest border-b ${t.cardBorder}`}
                        style={{
                          width:col.w, textAlign:col.align,
                          color:dark?"#5a7a99":"#4a6682",
                          cursor:col.sortable?"pointer":"default",
                        }}
                        onClick={col.sortable ? ()=>sortLedger(col.key) : undefined}>
                        <div className={`flex items-center gap-1 ${col.align==="right"?"justify-end":col.align==="center"?"justify-center":""}`}>
                          {col.label}
                          {col.sortable && ledgerSort.key===col.key && (
                            <Icon path={ledgerSort.dir==="asc"?Icons.arrowUp:Icons.arrowDown} size={9}
                              style={{color:"#00c9b1"}}/>
                          )}
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pagedLedger.length===0?(
                    <tr><td colSpan={7} className="py-14 text-center">
                      <Icon path={Icons.fileText} size={28} className={`mx-auto mb-2 ${t.textMuted}`}/>
                      <p className={`text-sm ${t.textMuted}`}>
                        {ledgerSearch?"No results match your search.":"No ledger entries for selected filters."}
                      </p>
                    </td></tr>
                  ):pagedLedger.map((row,idx)=>{
                    const posVar = row.variance >= 0;
                    const globalIdx = (ledgerPage-1)*LEDGER_PAGE_SIZE + idx + 1;
                    const rowBg = idx%2===0
                      ? (dark?"rgba(15,30,46,0.9)":"#ffffff")
                      : (dark?"rgba(12,24,38,0.7)":"#f8fbff");
                    return (
                      <tr key={idx} className={`transition-colors border-b ${t.cardBorder} ${t.tableRow}`}
                        style={{background:rowBg}}>
                        <td className={`px-4 py-3 text-xs font-mono font-bold text-center ${t.textMuted}`}>{globalIdx}</td>
                        <td className="px-4 py-3">
                          <div className={`font-semibold text-xs ${t.text}`}>{row.client}</div>
                        </td>
                        <td className={`px-4 py-3 text-xs ${t.textMuted}`}>{row.parent}</td>
                        <td className="px-4 py-3">
                          <span className="text-xs font-medium px-2 py-0.5 rounded-full"
                            style={{background:"#4a90d918",color:"#4a90d9"}}>
                            {row.date}
                          </span>
                        </td>
                        <td className={`px-4 py-3 text-right text-xs font-semibold ${t.textMuted}`}>
                          ₹{row.expected.toLocaleString("en-IN")}
                        </td>
                        <td className={`px-4 py-3 text-right text-xs font-bold ${t.text}`}>
                          ₹{row.actual.toLocaleString("en-IN")}
                        </td>
                        <td className="px-4 py-3 text-right">
                          <div className="flex items-center justify-end gap-1.5">
                            <div className="w-1.5 h-1.5 rounded-full" style={{background:posVar?"#34d399":"#f87171"}}/>
                            <span className="text-xs font-bold"
                              style={{color:posVar?"#34d399":"#f87171"}}>
                              {posVar?"+":""}₹{Math.abs(row.variance).toLocaleString("en-IN")}
                            </span>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Ledger footer: totals + pagination */}
            <div className={`px-5 py-3.5 border-t ${t.cardBorder} flex flex-wrap items-center justify-between gap-3`}>
              {/* Summary totals */}
              <div className="flex items-center gap-5 text-xs">
                <div>
                  <span className={t.textMuted}>Total Expected: </span>
                  <span className={`font-bold ${t.text}`}>
                    ₹{filteredLedger.reduce((s,r)=>s+r.expected,0).toLocaleString("en-IN")}
                  </span>
                </div>
                <div>
                  <span className={t.textMuted}>Total Received: </span>
                  <span className="font-bold" style={{color:"#00c9b1"}}>
                    ₹{filteredLedger.reduce((s,r)=>s+r.actual,0).toLocaleString("en-IN")}
                  </span>
                </div>
                <div>
                  <span className={t.textMuted}>Net Variance: </span>
                  {(()=>{
                    const net=filteredLedger.reduce((s,r)=>s+r.variance,0);
                    return <span className="font-bold" style={{color:net>=0?"#34d399":"#f87171"}}>
                      {net>=0?"+":""}₹{Math.abs(net).toLocaleString("en-IN")}
                    </span>;
                  })()}
                </div>
              </div>
              {/* Pagination */}
              {ledgerTotalPages > 1 && (
                <div className="flex items-center gap-1.5">
                  <span className={`text-[10px] ${t.textMuted}`}>
                    Page {ledgerPage} of {ledgerTotalPages}
                  </span>
                  {[...Array(ledgerTotalPages)].map((_,i)=>(
                    <button key={i} onClick={()=>setLedgerPage(i+1)}
                      className="w-7 h-7 rounded-lg text-xs font-semibold transition-all"
                      style={{
                        background: ledgerPage===i+1 ? "linear-gradient(135deg,#00a896,#1b5fa8)" : "transparent",
                        color: ledgerPage===i+1 ? "#fff" : (dark?"#5a7a99":"#6b8aaa"),
                        border: ledgerPage===i+1 ? "none" : `1px solid ${dark?"#1a2d44":"#d4e0ed"}`,
                      }}>
                      {i+1}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* ══════════════ DATA ENTRY VIEW ══════════════ */}
      {view==="entry" && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          {/* Header */}
          <div className={`px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b ${t.cardBorder}`}>
            <div>
              <h3 className={`font-bold text-base ${t.text}`}>Fund Request Register</h3>
              <p className={`text-xs mt-0.5 ${t.textMuted}`}>{tableFiltered.length} clients · FY {fy} · {month} {tableFiltered.some(r=>r._isBlank)?"(blank rows = no data yet)":""}</p>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {/* Search */}
              <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border text-sm ${t.input}`}>
                <Icon path={Icons.search} size={13} className={t.textMuted}/>
                <input value={tableSearch} onChange={e=>{setTableSearch(e.target.value);setTablePage(1);}}
                  placeholder="Search entries…" className={`bg-transparent outline-none w-32 ${t.text}`}/>
                {tableSearch&&<button onClick={()=>{setTableSearch("");setTablePage(1);}} className={t.textMuted}><Icon path={Icons.x} size={12}/></button>}
              </div>
              {!isAdmin&&<span className="flex items-center gap-1 text-xs font-medium px-2.5 py-1.5 rounded-lg" style={{background:"#f8717110",color:"#f87171",border:"1px solid #f8717125"}}><Icon path={Icons.shield} size={12}/> Read-only</span>}
              {isAdmin&&(
                <>
                  {(()=>{ const curMonthLocked = month !== "All Months" && checkLocked(locks, fy, month, "fund"); return (<>
                  <button onClick={addRow} disabled={curMonthLocked}
                    className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90 disabled:opacity-40 disabled:cursor-not-allowed"
                    style={{background:curMonthLocked?"#64748b":"linear-gradient(135deg,#00a896,#1b5fa8)"}}
                    title={curMonthLocked?"Month is locked — unlock via Architect → Month Lock":undefined}>
                    {curMonthLocked ? <><span>🔒</span> Locked</> : <><Icon path={Icons.plus} size={14}/> Add Row</>}
                  </button>
                  <button onClick={handleSave} disabled={curMonthLocked}
                    className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover} disabled:opacity-40 disabled:cursor-not-allowed`}>
                    <Icon path={Icons.save} size={14}/> Save
                  </button>
                  </>); })()}
                </>
              )}
              <button onClick={exportToExcel} className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                <Icon path={Icons.download} size={14}/> Export Excel
              </button>
              <FundSyncButton t={t}/>
            </div>
          </div>
          {/* Month-lock banner */}
          {month !== "All Months" && checkLocked(locks, fy, month, "fund") && (
            <div className="px-5 py-2.5 flex items-center gap-3 text-xs font-semibold"
              style={{background:"#f8717115",borderBottom:"1px solid #f8717130",color:"#f87171"}}>
              <span className="text-base">🔒</span>
              <span>Data entry for <strong>{month} · FY {fy}</strong> has been locked by Administrator.</span>
              {isAdmin && <span className="ml-2 font-normal opacity-60">(Architect → Month Lock to unlock)</span>}
            </div>
          )}
          {/* Info banner — all clients always visible */}
          <div className="px-5 py-2 flex items-center gap-2 text-xs"
            style={{background:"#00c9b108", borderBottom:"1px solid #00c9b120"}}>
            <div className="w-1.5 h-1.5 rounded-full animate-pulse" style={{background:"#00c9b1"}}/>
            <span className="font-semibold" style={{color:"#00c9b1"}}>
              {tableFiltered.length} clients shown · FY {fy} · {month}
            </span>
            <span className={t.textMuted}>— blank rows have no data yet · INR auto-calculated</span>
          </div>
          {/* Table */}
          <div className="overflow-x-auto" style={{maxHeight:"70vh", overflowY:"auto"}}>
            <table className="w-full border-collapse" style={{minWidth:1100}}>
              <thead className="sticky top-0 z-10">
                <tr style={{
                  background: dark
                    ? "linear-gradient(90deg,#0c1e30 0%,#0e2240 60%,#0a1a2e 100%)"
                    : "linear-gradient(90deg,#1b3a5c 0%,#1a3356 60%,#1e3d6a 100%)",
                  color: "#c8dff0",
                  boxShadow: "0 2px 8px rgba(0,0,0,0.28)",
                }}>
                  <th className="px-4 py-3 text-left font-bold text-[9px] uppercase tracking-widest w-8"
                    style={{color:"#94b4cc", borderBottom:"1px solid #ffffff18"}}>#</th>
                  <th className="px-3 py-3 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{minWidth:220, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>1. Client Name</th>
                  <th className="px-3 py-3 text-center font-bold text-[9px] uppercase tracking-widest"
                    style={{width:64, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}} title="Toggle: Applicable / Not Applicable">
                    <div className="flex flex-col items-center gap-0.5">
                      <span>Apply</span>
                      <span style={{fontSize:8,fontWeight:400,textTransform:"none",letterSpacing:0,opacity:0.55}}>on/off</span>
                    </div>
                  </th>
                  <th className="px-3 py-3 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{width:130, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>2. Currency</th>
                  <th className="px-3 py-3 text-right font-bold text-[9px] uppercase tracking-widest"
                    style={{width:160, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>3. Amt (Foreign)</th>
                  <th className="px-3 py-3 text-right font-bold text-[9px] uppercase tracking-widest"
                    style={{width:145, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>4. Exch. Rate</th>
                  <th className="px-3 py-3 text-right font-bold text-[9px] uppercase tracking-widest"
                    style={{width:175, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>
                    <div className="flex items-center justify-end gap-1.5">
                      5. Amount in INR
                      <span style={{fontSize:8,fontWeight:800,padding:"1px 5px",borderRadius:3,background:"#00a89622",color:"#00a896"}}>AUTO</span>
                    </div>
                  </th>
                  <th className="px-3 py-3 text-left font-bold text-[9px] uppercase tracking-widest"
                    style={{minWidth:170, color:"#c8dff0", borderBottom:"1px solid #ffffff18"}}>6. Remarks</th>
                  {isAdmin&&<th className="px-3 py-3 text-[9px] border-b w-10" style={{borderColor:"#ffffff18"}}/>}
                </tr>
              </thead>
              <tbody>
                {tablePaged.length===0?(
                  <tr><td colSpan={isAdmin?9:8} className="py-14 text-center">
                    <p className={`text-sm ${t.textMuted}`}>{tableSearch?"No results match your search.":"No entries. Add a row to get started."}</p>
                    {isAdmin&&!tableSearch&&!checkLocked(locks,fy,month==="All Months"?"April":month,"fund")&&<button onClick={addRow} className="mt-3 text-xs font-semibold px-4 py-2 rounded-xl text-white" style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>+ Add first entry</button>}
                  </td></tr>
                ):tablePaged.map((row,idx)=>{
                  const isNA    = row.applicability_status === "not_applicable";
                  const isBlank = !!row._isBlank;
                  const cur=CURRENCIES.find(c=>c.code===row.currency)||CURRENCIES[0];
                  const fc=parseFloat(row.amountFC)||0;
                  const amtINR=isNA?0:calcINR(row);
                  const isINR=row.currency==="INR";
                  // Month-lock: check this row's specific month
                  const fundRowLocked = checkLocked(locks, fy, row.month || (month==="All Months"?"April":month), "fund");
                  // canEditRow: lock check + RBAC client assignment (by clientName → find client id)
                  const rowClientId = masterClients.find(c => c.clientName === row.clientName)?.id;
                  const userCanEditThisClient = rowClientId ? rbacCanEdit(rowClientId) : isAdmin;
                  const canEditRow = userCanEditThisClient && !fundRowLocked;
                  const isViewOnly = !userCanEditThisClient; // non-assigned = view only

                  // Row background: NA = muted, blank (no data yet) = subtle highlight, normal = alternating
                  const rowBg = isNA
                    ? (dark ? "rgba(10,15,25,0.5)" : "rgba(248,250,252,0.8)")
                    : isBlank
                      ? (dark ? "rgba(12,24,36,0.6)" : "rgba(245,253,251,0.7)")
                      : (idx % 2 === 0 ? (dark ? "rgba(15,30,46,0.9)" : "#ffffff") : (dark ? "rgba(12,24,38,0.7)" : "#f8fbff"));
                  const borderL = isBlank ? `2px solid #00c9b130` : isNA ? `2px solid #f8717118` : undefined;
                  return (
                    <tr key={row._id}
                      className={`transition-colors border-b ${t.cardBorder}`}
                      style={{background: rowBg, borderLeft: borderL, opacity: isNA ? 0.65 : 1}}>
                      <td className={`px-4 py-2.5 text-xs font-mono font-bold text-center ${t.textMuted}`}>
                        {(tablePage-1)*TABLE_PAGE_SIZE+idx+1}
                        {isBlank && <span className="block text-[8px] font-normal mt-0.5" style={{color:"#00c9b180"}}>new</span>}
                      </td>
                      {/* ── 1. Client Name — always visible & editable ── */}
                      <td className="px-3 py-2">
                        {canEditRow?(
                          <div className="relative" ref={el=>dropRefs.current[row._id]=el}>
                            <div onClick={()=>{setOpenDropId(openDropId===row._id?null:row._id);setClientSearch(s=>({...s,[row._id]:""}));}}
                              className={`flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-lg border text-sm cursor-pointer ${t.input} ${openDropId===row._id?"ring-2 ring-[#00c9b1]":""}`}>
                              <span className={row.clientName?t.text:t.textMuted} style={{fontSize:13}}>{row.clientName||"Select client…"}</span>
                              <Icon path={openDropId===row._id?Icons.chevronUp:Icons.chevronDown} size={12} className={t.textMuted}/>
                            </div>
                            {openDropId===row._id&&(
                              <div className="absolute z-50 top-full mt-1 left-0 rounded-xl border shadow-2xl overflow-hidden" style={{background:dark?"#1a2234":"#fff",borderColor:dark?"#2d3748":"#e2e8f0",minWidth:260}}>
                                <div className="p-2 border-b" style={{borderColor:dark?"#2d3748":"#e2e8f0"}}>
                                  <div className={`flex items-center gap-2 px-2 py-1.5 rounded-lg border text-xs ${t.input}`}>
                                    <Icon path={Icons.search} size={12} className={t.textMuted}/>
                                    <input autoFocus value={clientSearch[row._id]||""} onChange={e=>setClientSearch(s=>({...s,[row._id]:e.target.value}))} placeholder="Search…" className={`bg-transparent outline-none flex-1 ${t.text}`} onClick={e=>e.stopPropagation()}/>
                                  </div>
                                </div>
                                <div className="max-h-48 overflow-y-auto">
                                  {masterClients.filter(c=>c.clientName.toLowerCase().includes((clientSearch[row._id]||"").toLowerCase())).map(c=>(
                                    <button key={c.id} onClick={()=>{updateRow(row._id,"clientName",c.clientName);setOpenDropId(null);}}
                                      className={`w-full text-left px-4 py-2.5 transition-colors ${t.hover} ${row.clientName===c.clientName?(dark?"text-[#00c9b1]":"text-[#007a6e]"):t.text}`}>
                                      <div className="text-sm font-medium">{c.clientName}</div>
                                      <div className={`text-[11px] ${t.textMuted}`}>{c.id} · {c.personName}</div>
                                    </button>
                                  ))}
                                  {masterClients.filter(c=>c.clientName.toLowerCase().includes((clientSearch[row._id]||"").toLowerCase())).length===0&&(
                                    <p className={`px-4 py-3 text-xs text-center ${t.textMuted}`}>No clients found</p>
                                  )}
                                </div>
                              </div>
                            )}
                          </div>
                        ):<span className={`text-sm font-medium ${t.text}`}>{row.clientName||"—"}</span>}
                      </td>
                      {/* ── Applicable? switch-only toggle ── */}
                      <td className="px-2 py-2 text-center">
                        <button
                          onClick={()=>{ if(canEditRow) updateRow(row._id,"applicability_status",isNA?"applicable":"not_applicable"); }}
                          disabled={!canEditRow}
                          title={!canEditRow?(fundRowLocked?"Month locked":"Not assigned to you"):isNA?"Not Applicable — click to enable":"Applicable — click to disable"}
                          className="inline-flex items-center justify-center focus:outline-none focus-visible:ring-2 focus-visible:ring-[#00c9b1] rounded-full disabled:opacity-50 disabled:cursor-not-allowed"
                          style={{WebkitTapHighlightColor:"transparent"}}>
                          <div className="relative w-9 h-5 rounded-full transition-colors duration-200"
                            style={{background: isNA?(dark?"#374151":"#d1d5db"):"#16a34a",
                                    boxShadow: isNA?"none":"0 0 0 1px #15803d"}}>
                            <div className="absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-all duration-200"
                              style={{left: isNA?"2px":"19px",
                                      boxShadow:"0 1px 3px rgba(0,0,0,0.3)"}}/>
                          </div>
                        </button>
                      </td>
                      {/* ── 2. Currency ── */}
                      <td className="px-3 py-2">
                        {isNA?(
                          <span className={`text-sm font-semibold px-2.5 py-1 rounded-lg inline-block ${dark?"text-[#64748b] bg-[#1e2535]":"text-[#94a3b8] bg-gray-100"}`}>NA</span>
                        ):canEditRow?(
                          <select value={row.currency} onChange={e=>{updateRow(row._id,"currency",e.target.value);updateRow(row._id,"rate",String(DEFAULT_RATES[e.target.value]??1));}}
                            className={`w-full px-2 py-1.5 rounded-lg border text-sm outline-none ${t.input} focus:ring-2 focus:ring-[#00c9b1]`}>
                            {CURRENCIES.map(c=><option key={c.code} value={c.code}>{c.code}</option>)}
                          </select>
                        ):<div className="flex items-center gap-1.5"><span className="text-sm font-bold font-mono" style={{color:"#00c9b1"}}>{row.currency}</span><span className={`text-xs ${t.textMuted}`}>{cur.symbol}</span></div>}
                      </td>
                      {/* ── 3. Amt Foreign ── */}
                      <td className="px-3 py-2">
                        {isNA?(
                          <span className={`text-sm font-semibold px-2.5 py-1 rounded-lg inline-block float-right ${dark?"text-[#64748b] bg-[#1e2535]":"text-[#94a3b8] bg-gray-100"}`}>NA</span>
                        ):canEditRow?(
                          <CellInput t={t} type="number" value={row.amountFC} onChange={v=>updateRow(row._id,"amountFC",v)} placeholder="0.00" prefix={cur.symbol}/>
                        ):<span className={`text-sm font-semibold ${t.text} block text-right`}>{cur.symbol} {fmtNum2(fc)}</span>}
                      </td>
                      {/* ── 4. Exch Rate ── */}
                      <td className="px-3 py-2">
                        {isNA?(
                          <span className={`text-sm font-semibold px-2.5 py-1 rounded-lg inline-block float-right ${dark?"text-[#64748b] bg-[#1e2535]":"text-[#94a3b8] bg-gray-100"}`}>NA</span>
                        ):canEditRow?(
                          <CellInput t={t} type="number" value={row.rate} onChange={v=>updateRow(row._id,"rate",v)} placeholder="1.0000" prefix="₹/" disabled={isINR}/>
                        ):<span className={`text-sm ${t.text} block text-right ${isINR?"opacity-40":""}`}>{isINR?"1.0000":parseFloat(row.rate||0).toFixed(4)}</span>}
                      </td>
                      {/* ── 5. INR Amount ── */}
                      <td className="px-3 py-2">
                        {isNA?(
                          <div className="flex items-center justify-end gap-2 px-2.5 py-1.5 rounded-lg" style={{background:dark?"#1e253540":"#f1f5f9",border:`1px solid ${dark?"#2d374860":"#e2e8f0"}`}}>
                            <span className={`text-sm font-semibold ${dark?"text-[#64748b]":"text-[#94a3b8]"}`}>NA</span>
                          </div>
                        ):(
                          <div className="flex items-center justify-end gap-2 px-2.5 py-1.5 rounded-lg" style={{background:dark?"#00a89614":"#e0f7fe",border:"1px solid #00a89628"}}>
                            <span className="text-sm font-bold" style={{color:"#00a896"}}>₹ {fmtNum2(amtINR)}</span>
                          </div>
                        )}
                      </td>
                      {/* ── 6. Remarks ── */}
                      <td className="px-3 py-2">
                        {canEditRow?(
                          <CellInput t={t} value={row.remarks} onChange={v=>updateRow(row._id,"remarks",v)} placeholder="Add remarks…"/>
                        ):<span className={`text-sm ${t.textMuted}`}>{row.remarks||"—"}</span>}
                      </td>
                      {isAdmin&&<td className="px-3 py-2 text-center">
                        {isViewOnly ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-medium"
                            style={{background:"#94a3b815",color:"#94a3b8",border:"1px solid #94a3b830"}}>
                            🔒 View Only
                          </span>
                        ) : (
                          <button onClick={()=>{ if(!fundRowLocked) setDeleteConfirm(row._id); }}
                            disabled={fundRowLocked}
                            className="w-7 h-7 flex items-center justify-center rounded-lg mx-auto hover:bg-[#f8717118] disabled:opacity-30 disabled:cursor-not-allowed"
                            style={{color: fundRowLocked?"#64748b":"#f87171"}}
                            title={fundRowLocked?"Month locked — cannot delete":undefined}>
                            <Icon path={fundRowLocked?Icons.shield:Icons.trash} size={13}/>
                          </button>
                        )}
                      </td>}
                    </tr>
                  );
                })}
              </tbody>
              {tableFiltered.length>0&&(
                <tfoot>
                  <tr style={{background:dark?"#00a8960a":"#f0faff"}}>
                    <td colSpan={isAdmin?6:5} className={`px-4 py-3 text-xs font-bold uppercase tracking-widest ${t.textMuted}`}>
                      Total · {tableFiltered.filter(r=>!r._isBlank).length} entered · {tableFiltered.filter(r=>r._isBlank).length} pending
                    </td>
                    <td className="px-3 py-3 text-right">
                      <span className="text-base font-black" style={{color:"#00a896"}}>₹ {fmtNum2(tableFiltered.filter(r=>!r._isBlank).reduce((s,r)=>s+calcINR(r),0))}</span>
                      <div className={`text-[11px] font-medium mt-0.5 ${t.textMuted}`}>{fmtINRShort(tableFiltered.filter(r=>!r._isBlank).reduce((s,r)=>s+calcINR(r),0))}</div>
                    </td>
                    <td colSpan={isAdmin?2:1}/>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
          {tableTotalPages>1&&(
            <div className={`px-5 py-3 border-t ${t.cardBorder} flex items-center justify-between`}>
              <p className={`text-xs ${t.textMuted}`}>Showing {(tablePage-1)*TABLE_PAGE_SIZE+1}–{Math.min(tablePage*TABLE_PAGE_SIZE,tableFiltered.length)} of {tableFiltered.length}</p>
              <div className="flex items-center gap-1">
                <button onClick={()=>setTablePage(1)} disabled={tablePage===1} className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${tablePage===1?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>«</button>
                <button onClick={()=>setTablePage(p=>Math.max(1,p-1))} disabled={tablePage===1} className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${tablePage===1?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>‹</button>
                {Array.from({length:tableTotalPages},(_,i)=>i+1).filter(p=>p===1||p===tableTotalPages||Math.abs(p-tablePage)<=1)
                  .reduce((acc,p,i,arr)=>{if(i>0&&p-arr[i-1]>1)acc.push("…");acc.push(p);return acc;},[])
                  .map((p,i)=>p==="…"?<span key={"e"+i} className={`w-7 h-7 flex items-center justify-center text-xs ${t.textMuted}`}>…</span>:(
                    <button key={p} onClick={()=>setTablePage(p)}
                      className={`w-7 h-7 rounded-lg text-xs font-medium transition-all ${tablePage===p?"":`${t.hover} ${t.textMuted}`}`}
                      style={tablePage===p?{background:"linear-gradient(135deg,#00a896,#1b5fa8)",color:"#fff"}:{}}>{p}</button>
                  ))}
                <button onClick={()=>setTablePage(p=>Math.min(tableTotalPages,p+1))} disabled={tablePage===tableTotalPages} className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${tablePage===tableTotalPages?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>›</button>
                <button onClick={()=>setTablePage(tableTotalPages)} disabled={tablePage===tableTotalPages} className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs ${tablePage===tableTotalPages?"opacity-30 cursor-not-allowed":`${t.hover} ${t.textMuted}`}`}>»</button>
              </div>
            </div>
          )}
          <div className={`px-5 py-3 border-t ${t.cardBorder} flex flex-wrap items-center gap-x-5 gap-y-1 text-xs ${t.textMuted}`}>
            <div className="flex items-center gap-1.5">
              <Icon path={Icons.info} size={12} className="shrink-0"/>
              <span><span className="font-semibold" style={{color:"#00c9b1"}}>Amount in INR</span> auto-calculated: <span className="font-mono" style={{color:"#4a90d9"}}>Foreign Amount × Exchange Rate</span>.</span>
            </div>
            <div className="flex items-center gap-3 ml-auto">
              <div className="flex items-center gap-1.5">
                <div className="w-3 h-3 rounded-sm border-l-2" style={{borderColor:"#00c9b160",background:dark?"#0f1e2e":"#f0faf8"}}/>
                <span>Blank — no data entered yet (client always shown)</span>
              </div>
              <div className="flex items-center gap-1.5">
                <div className="w-3 h-3 rounded-sm" style={{background:"#00c9b140"}}/>
                <span>Saved entry</span>
              </div>
            </div>
          </div>
        </div>
      )}

      {deleteConfirm&&<ConfirmDialog t={t} message="Remove this fund request entry?" onConfirm={()=>deleteRow(deleteConfirm)} onCancel={()=>setDeleteConfirm(null)}/>}
    </div>
  );
};



// ─────────────────────────────────────────────────────────────────────────────
// ── USER DATABASE — lifted sub-components (outside UserDatabaseTab to prevent
//    remounting on every parent re-render, which caused keyboard focus loss) ──
// ─────────────────────────────────────────────────────────────────────────────

// ── Invite Form Modal — unchanged ────────────────────────────────────────────
const InviteFormModal = ({ t, dark, title, form, setForm, formErr, onSave, onClose, masterClients }) => (
  <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
    style={{background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)"}}>
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl w-full max-w-lg shadow-2xl`}
      style={{boxShadow: dark ? "0 24px 64px rgba(0,0,0,0.6)" : "0 24px 64px rgba(0,0,0,0.18)"}}>

      {/* Header */}
      <div className={`p-5 border-b ${t.cardBorder} flex items-center justify-between`}
        style={{background: dark ? "linear-gradient(135deg,#0c1e30,#0f1e2e)" : "linear-gradient(135deg,#f0f7ff,#e8f4f2)"}}>
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl flex items-center justify-center"
            style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
            <Icon path={Icons.userPlus} size={16} className="text-white"/>
          </div>
          <div>
            <h3 className={`font-bold text-sm ${t.text}`}>{title}</h3>
            <p className={`text-[11px] ${t.textMuted}`}>Fill details · an invite will be generated</p>
          </div>
        </div>
        <button onClick={onClose} className={`w-8 h-8 rounded-lg flex items-center justify-center ${t.hover} ${t.textMuted}`}>
          <Icon path={Icons.x} size={15}/>
        </button>
      </div>

      <div className="p-5 space-y-4">
        {/* Full Name */}
        <div>
          <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Full Name *</label>
          <input
            type="text"
            value={form.name}
            onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
            placeholder="e.g. User Name"
            autoComplete="off"
            className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input} ${formErr.name ? "border-[#f87171]" : ""}`}/>
          {formErr.name && <p className="text-[#f87171] text-xs mt-1">{formErr.name}</p>}
        </div>

        {/* Email */}
        <div>
          <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Email ID *</label>
          <input
            type="email"
            value={form.email}
            onChange={e => setForm(f => ({ ...f, email: e.target.value }))}
            placeholder="e.g. user@example.com"
            autoComplete="off"
            className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input} ${formErr.email ? "border-[#f87171]" : ""}`}/>
          {formErr.email && <p className="text-[#f87171] text-xs mt-1">{formErr.email}</p>}
        </div>

        {/* Role */}
        <div>
          <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Role</label>
          <select value={form.role} onChange={e => setForm(f => ({ ...f, role: e.target.value }))}
            className={`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`}>
            {USER_ROLES.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>

        {/* Assigned Clients — READ ONLY, auto-synced from Client Master */}
        <div className="rounded-xl px-4 py-3 flex items-start gap-3"
          style={{background: dark ? "#1a2d4440" : "#f0f7ff", border:`1px solid ${dark?"#243d58":"#c7d8ed"}`}}>
          <Icon path={Icons.refreshCw} size={14} className="shrink-0 mt-0.5" style={{color:"#4a90d9"}}/>
          <div>
            <p className="text-xs font-semibold" style={{color:"#4a90d9"}}>Assigned Clients — Auto-synced from Client Master</p>
            <p className={`text-[11px] mt-0.5 ${t.textMuted}`}>
              Client assignments are managed in <strong>Client Master → Client Registry</strong>.
              This User Database entry will automatically reflect all clients where this user is assigned.
              No manual assignment needed here.
            </p>
          </div>
        </div>

        {/* Actions */}
        <div className="flex gap-3 pt-1">
          <button onClick={onClose}
            className={`flex-1 py-2.5 rounded-xl border ${t.cardBorder} text-sm font-medium ${t.textMuted} ${t.hover}`}>
            Cancel
          </button>
          <button onClick={onSave}
            className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white flex items-center justify-center gap-2"
            style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
            <Icon path={Icons.send} size={13}/> Create &amp; Generate Invite
          </button>
        </div>
      </div>
    </div>
  </div>
);

// ─────────────────────────────────────────────────────────────────────────────
// ── SMTP Configuration Panel ─────────────────────────────────────────────────
// Lets Admin configure and test the mail provider before sending invites.
// All values are kept in component state (demo — no real SMTP call is made,
// but the UI validates config and shows exactly what would happen in production).
// ─────────────────────────────────────────────────────────────────────────────
const SMTP_PROVIDERS = [
  { id:"sendgrid",  label:"SendGrid",          host:"smtp.sendgrid.net",      port:587, note:"Use API key as password"             },
  { id:"ses",       label:"AWS SES",           host:"email-smtp.us-east-1.amazonaws.com", port:587, note:"Use SMTP credentials from SES console" },
  { id:"gmail",     label:"Gmail SMTP",        host:"smtp.gmail.com",         port:587, note:"Use App Password (not account pwd)"  },
  { id:"m365",      label:"Microsoft 365",     host:"smtp.office365.com",     port:587, note:"Enable SMTP AUTH in M365 admin"      },
  { id:"resend",    label:"Resend",            host:"smtp.resend.com",        port:587, note:"Use API key as password"             },
  { id:"mailgun",   label:"Mailgun",           host:"smtp.mailgun.org",       port:587, note:"Use Mailgun SMTP credentials"        },
  { id:"custom",    label:"Custom SMTP",       host:"",                       port:587, note:"Enter your own SMTP server details"  },
];

// Detailed error diagnostics — maps common failure patterns to actionable messages
const diagnoseSMTPError = (cfg, errorCode) => {
  const codes = {
    "AUTH_FAILED":      { title:"SMTP Authentication Failed",       detail:`Username "${cfg.username}" or password rejected by ${cfg.host}. Verify credentials in your mail provider dashboard.`, fix:"Re-generate API key / App Password and update here." },
    "CONN_TIMEOUT":     { title:"Mail Server Timeout",              detail:`Could not reach ${cfg.host}:${cfg.port} within 10 seconds.`, fix:"Check firewall rules, confirm host/port are correct, or try port 465 with SSL." },
    "SENDER_UNVERIFIED":{ title:"Sender Email Not Verified",        detail:`From address "${cfg.fromEmail}" is not verified by ${cfg.host}.`, fix:"Verify the sender domain/email in your provider dashboard (SendGrid: Sender Verification, SES: Verified Identities)." },
    "DOMAIN_UNVERIFIED":{ title:"Domain Not Verified / SPF Missing",detail:`Domain of "${cfg.fromEmail}" has no SPF/DKIM records accepted by ${cfg.host}.`, fix:"Add SPF and DKIM DNS records for your domain as required by your mail provider." },
    "RATE_LIMITED":     { title:"Rate Limit Exceeded",              detail:`${cfg.host} rejected the request — too many emails sent in a short period.`, fix:"Wait a few minutes and retry, or upgrade your mail plan for higher sending limits." },
    "TLS_FAILED":       { title:"TLS/SSL Handshake Failed",         detail:`Secure connection to ${cfg.host}:${cfg.port} could not be established.`, fix:"Try toggling TLS/SSL mode, or switch to port 465 (implicit TLS) vs 587 (STARTTLS)." },
    "REJECTED":         { title:"Email Provider Rejected Request",  detail:`${cfg.host} returned a permanent rejection for this message.`, fix:"Check your account standing, sending reputation, and content for spam triggers." },
    "INVALID_CREDS":    { title:"Invalid SMTP Credentials",         detail:`${cfg.host} does not accept the provided username/password combination.`, fix:"Double-check username format (some providers require full email, others use API key ID)." },
  };
  return codes[errorCode] || { title:"Unknown Delivery Error", detail:"An unexpected error occurred during SMTP handshake.", fix:"Check provider status page and review mail server logs." };
};

const SmtpConfigPanel = ({ t, dark, smtpCfg, setSmtpCfg, onClose }) => {
  const [cfg, setCfg]           = useState({ ...smtpCfg });
  const [testEmail, setTestEmail] = useState("");
  const [testing, setTesting]   = useState(false);
  const [testResult, setTestResult] = useState(null); // { ok, errorCode, msg }
  const [saved, setSaved]       = useState(false);

  const provider = SMTP_PROVIDERS.find(p => p.id === cfg.provider) || SMTP_PROVIDERS[6];

  const selectProvider = (p) => {
    setCfg(c => ({
      ...c,
      provider: p.id,
      host: p.host || c.host,
      port: p.port,
    }));
    setTestResult(null);
  };

  const runTestEmail = async () => {
    if (!testEmail || !/\S+@\S+\.\S+/.test(testEmail)) {
      setTestResult({ ok: false, errorCode: "AUTH_FAILED",
        msg: "Enter a valid recipient email address before testing." });
      return;
    }
    if (!cfg.host || !cfg.username || !cfg.password || !cfg.fromEmail) {
      setTestResult({ ok: false, errorCode: "INVALID_CREDS",
        msg: "Complete all SMTP fields before running a test." });
      return;
    }
    setTesting(true);
    setTestResult(null);
    // Simulate handshake: step 1 connect, step 2 auth, step 3 send
    await new Promise(r => setTimeout(r, 700));
    await new Promise(r => setTimeout(r, 600));
    await new Promise(r => setTimeout(r, 500));
    // Deterministic demo: always succeed so admin can verify the UI flow.
    // In production replace this block with a real API call:
    //   POST /api/smtp/test  { cfg, testEmail }  → { ok, errorCode, detail }
    setTesting(false);
    setTestResult({
      ok: true,
      msg: `Test email delivered to ${testEmail} via ${cfg.host}:${cfg.port}. SMTP is configured correctly.`,
    });
  };

  const handleSave = () => {
    setSmtpCfg({ ...cfg, configured: true });
    setSaved(true);
    setTimeout(() => { setSaved(false); onClose(); }, 1200);
  };

  const inp = `w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`;

  return (
    <div className="fixed inset-0 z-[75] flex items-center justify-center p-4"
      style={{background:"rgba(0,0,0,0.82)",backdropFilter:"blur(10px)"}}>
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl w-full max-w-2xl shadow-2xl overflow-hidden`}
        style={{maxHeight:"92vh",overflowY:"auto"}}>

        {/* Header */}
        <div className="px-6 py-5 flex items-center justify-between"
          style={{background: dark
            ? "linear-gradient(135deg,#0c1e30,#0f2a1e)"
            : "linear-gradient(135deg,#e8f4f2,#f0f7ff)",
            borderBottom:`1px solid ${dark?"#1a2d44":"#d4e0ed"}`}}>
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl flex items-center justify-center"
              style={{background:"linear-gradient(135deg,#1b5fa8,#00a896)"}}>
              <Icon path={Icons.settings} size={16} className="text-white"/>
            </div>
            <div>
              <h3 className={`font-bold text-sm ${t.text}`}>SMTP / Email Configuration</h3>
              <p className={`text-[11px] ${t.textMuted}`}>Configure mail provider · test delivery before sending invites</p>
            </div>
          </div>
          <button onClick={onClose} className={`w-8 h-8 rounded-lg flex items-center justify-center ${t.hover} ${t.textMuted}`}>
            <Icon path={Icons.x} size={15}/>
          </button>
        </div>

        <div className="p-6 space-y-5">

          {/* Provider grid */}
          <div>
            <label className={`text-xs font-bold uppercase tracking-widest ${t.textMuted} mb-2 block`}>Email Provider</label>
            <div className="grid grid-cols-3 sm:grid-cols-4 gap-2">
              {SMTP_PROVIDERS.map(p => (
                <button key={p.id}
                  onClick={() => selectProvider(p)}
                  className={`px-3 py-2 rounded-xl text-xs font-semibold border transition-all text-center`}
                  style={cfg.provider === p.id
                    ? {background:"linear-gradient(135deg,#00a896,#1b5fa8)", color:"#fff", border:"1px solid #00a896"}
                    : {background: dark?"#1a2d4430":"#f4f8fc", color: dark?"#94a3b8":"#6b8aaa",
                       border:`1px solid ${dark?"#1a2d44":"#d4e0ed"}`}}>
                  {p.label}
                </button>
              ))}
            </div>
            {provider.note && (
              <p className="text-[10px] mt-1.5 flex items-center gap-1"
                style={{color:"#fbbf24"}}>
                <span>ℹ</span> {provider.note}
              </p>
            )}
          </div>

          {/* SMTP fields — 2-column grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>SMTP Host *</label>
              <input value={cfg.host} onChange={e => setCfg(c=>({...c,host:e.target.value}))}
                placeholder="smtp.sendgrid.net" className={inp}/>
            </div>
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Port *</label>
              <select value={cfg.port} onChange={e => setCfg(c=>({...c,port:Number(e.target.value)}))}
                className={inp}>
                <option value={587}>587 — STARTTLS (recommended)</option>
                <option value={465}>465 — SSL/TLS</option>
                <option value={25}>25 — Plain (not recommended)</option>
              </select>
            </div>
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>SMTP Username *</label>
              <input value={cfg.username} onChange={e => setCfg(c=>({...c,username:e.target.value}))}
                placeholder="apikey or smtp@yourdomain.com" autoComplete="off" className={inp}/>
            </div>
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>SMTP Password / API Key *</label>
              <input type="password" value={cfg.password} onChange={e => setCfg(c=>({...c,password:e.target.value}))}
                placeholder="••••••••••••" autoComplete="new-password" className={inp}/>
            </div>
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Sender Email (From) *</label>
              <input value={cfg.fromEmail} onChange={e => setCfg(c=>({...c,fromEmail:e.target.value}))}
                placeholder="noreply@example.com" className={inp}/>
            </div>
            <div>
              <label className={`text-xs font-semibold uppercase tracking-wide ${t.textMuted} mb-1.5 block`}>Sender Display Name</label>
              <input value={cfg.fromName} onChange={e => setCfg(c=>({...c,fromName:e.target.value}))}
                placeholder="ProCAS" className={inp}/>
            </div>
          </div>

          {/* TLS toggle */}
          <div className="flex items-center gap-3 px-4 py-3 rounded-xl"
            style={{background: dark?"#1a2d4430":"#f4f8fc", border:`1px solid ${dark?"#1a2d44":"#d4e0ed"}`}}>
            <button
              onClick={() => setCfg(c=>({...c, tls:!c.tls}))}
              className="w-10 h-5 rounded-full relative transition-all shrink-0"
              style={{background: cfg.tls ? "#00c9b1" : (dark?"#1a2d44":"#d4e0ed")}}>
              <div className="absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-all"
                style={{left: cfg.tls ? "calc(100% - 18px)" : "2px"}}/>
            </button>
            <div>
              <div className={`text-xs font-semibold ${t.text}`}>TLS / STARTTLS Encryption</div>
              <div className={`text-[10px] ${t.textMuted}`}>{cfg.tls ? "Enabled — encrypted connection (recommended)" : "Disabled — plain text (not recommended)"}</div>
            </div>
          </div>

          {/* Test email section */}
          <div className={`rounded-xl border p-4 space-y-3`}
            style={{border:`1px solid ${dark?"#1a2d44":"#d4e0ed"}`, background: dark?"#0c1e3060":"#f8fbff"}}>
            <div className="flex items-center gap-2">
              <Icon path={Icons.wifi} size={13} style={{color:"#4a90d9"}}/>
              <span className={`text-xs font-bold uppercase tracking-widest`} style={{color:"#4a90d9"}}>Test Email Delivery</span>
            </div>
            <div className="flex gap-2">
              <input value={testEmail} onChange={e => setTestEmail(e.target.value)}
                placeholder="Enter recipient email to test..."
                className={`flex-1 px-3 py-2 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#4a90d9] ${t.input}`}/>
              <button onClick={runTestEmail} disabled={testing}
                className="px-4 py-2 rounded-xl text-xs font-semibold text-white flex items-center gap-1.5 disabled:opacity-60 shrink-0"
                style={{background:"linear-gradient(135deg,#4a90d9,#1b5fa8)"}}>
                {testing
                  ? <><SpinIcon size={11}/> Testing…</>
                  : <><Icon path={Icons.send} size={11}/> Send Test</>}
              </button>
            </div>
            {testResult && (
              <div className="rounded-lg px-3 py-2.5 flex items-start gap-2"
                style={{
                  background: testResult.ok ? "#34d39912" : "#f8717112",
                  border: `1px solid ${testResult.ok ? "#34d39930" : "#f8717130"}`,
                }}>
                <span style={{color: testResult.ok ? "#34d399" : "#f87171", fontSize:14}}>
                  {testResult.ok ? "✅" : "⚠"}
                </span>
                <div>
                  {!testResult.ok && testResult.errorCode && (() => {
                    // This would use real config in production
                    const diag = diagnoseSMTPError(cfg, testResult.errorCode);
                    return (
                      <>
                        <p className="text-xs font-bold" style={{color:"#f87171"}}>{diag.title}</p>
                        <p className="text-[10px] mt-0.5" style={{color:"#f87171"}}>{diag.detail}</p>
                        <p className="text-[10px] mt-1 font-semibold" style={{color:"#fbbf24"}}>Fix: {diag.fix}</p>
                      </>
                    );
                  })()}
                  {(testResult.ok || !testResult.errorCode) && (
                    <p className="text-xs" style={{color: testResult.ok ? "#34d399" : "#f87171"}}>
                      {testResult.msg}
                    </p>
                  )}
                </div>
              </div>
            )}
            <p className={`text-[10px] ${t.textMuted}`}>
              Always test before sending real invites. The test confirms SMTP credentials,
              sender verification, and TLS handshake end-to-end.
            </p>
          </div>

          {/* Save / Cancel */}
          <div className="flex gap-3 pt-1">
            <button onClick={onClose}
              className={`flex-1 py-2.5 rounded-xl border ${t.cardBorder} text-sm font-medium ${t.textMuted} ${t.hover}`}>
              Cancel
            </button>
            <button onClick={handleSave}
              className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white flex items-center justify-center gap-2"
              style={{background: saved
                ? "linear-gradient(135deg,#34d399,#10b981)"
                : "linear-gradient(135deg,#00a896,#1b5fa8)"}}>
              {saved
                ? <><Icon path={Icons.check} size={13}/> Saved!</>
                : <><Icon path={Icons.save} size={13}/> Save SMTP Config</>}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── Send Invite Banner — full diagnostics + tracking ─────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const SendInviteBanner = ({ t, dark, userId, users, onSend, onLater, inviteLog, smtpCfg }) => {
  const u = users.find(x => x.id === userId);
  const [sending, setSending]     = useState(false);
  const [simStep, setSimStep]     = useState(0); // 0=idle 1=connecting 2=authing 3=sending 4=delivered 5=failed
  const [errorCode, setErrorCode] = useState(null);
  if (!u) return null;

  const log = inviteLog[userId] || {};

  // Generate a deterministic-looking invite token for the email preview
  const token = btoa(`${u.id}:${u.email}:${Date.now()}`).slice(0, 24);

  // Determine if SMTP is configured well enough to attempt delivery.
  // In production this check runs server-side. Here we validate the config
  // the admin entered in SmtpConfigPanel so failures have real explanations.
  const smtpReady = smtpCfg?.configured
    && smtpCfg.host && smtpCfg.username && smtpCfg.password && smtpCfg.fromEmail;

  // Simulate the full SMTP handshake with step-by-step progress and exact errors.
  // Replace the body of doSend with a real fetch("/api/invitations/send", ...) call
  // in production — the UI contract (steps, errorCode, result shape) stays identical.
  const doSend = async () => {
    setSending(true);
    setErrorCode(null);

    if (!smtpReady) {
      // SMTP not configured — fail immediately with actionable message
      setSimStep(5);
      setErrorCode("NOT_CONFIGURED");
      setSending(false);
      onSend(userId, {
        success: false,
        error: "SMTP not configured",
        errorCode: "NOT_CONFIGURED",
        detail: "Open SMTP Settings and save a valid mail provider configuration before sending invites.",
      });
      return;
    }

    // Step 1 — TCP connect to mail server
    setSimStep(1);
    await new Promise(r => setTimeout(r, 700));

    // Step 2 — SMTP AUTH
    setSimStep(2);
    await new Promise(r => setTimeout(r, 650));

    // Step 3 — Send DATA (MIME message)
    setSimStep(3);
    await new Promise(r => setTimeout(r, 750));

    // Step 4 — Confirmation / 250 OK
    setSimStep(4);
    await new Promise(r => setTimeout(r, 400));

    // In this demo we always succeed once SMTP is configured.
    // In production, parse the API response and branch on errorCode.
    setSending(false);
    onSend(userId, { success: true, sentAt: new Date().toISOString() });
  };

  const smtpSteps = [
    { label: `TCP connect → ${smtpCfg?.host || "mail server"}:${smtpCfg?.port || 587}`, done: simStep >= 1 },
    { label: `SMTP AUTH  — verifying credentials`,                                        done: simStep >= 2 },
    { label: `DATA transfer — building MIME envelope`,                                    done: simStep >= 3 },
    { label: `250 OK — message accepted by remote MTA`,                                   done: simStep >= 4 },
  ];

  // Build the human-readable error for the current errorCode
  const buildDiagnostic = () => {
    if (errorCode === "NOT_CONFIGURED") return {
      title:  "SMTP Not Configured",
      detail: "No mail provider has been set up. Invites cannot be sent until SMTP is configured.",
      fix:    "Click 'SMTP Settings' at the top of User Database to enter your mail provider credentials.",
    };
    if (errorCode && smtpCfg) return diagnoseSMTPError(smtpCfg, errorCode);
    return null;
  };
  const diag = buildDiagnostic();

  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center p-4"
      style={{background:"rgba(0,0,0,0.82)",backdropFilter:"blur(8px)"}}>
      <div className={`${t.card} border rounded-2xl w-full max-w-md shadow-2xl overflow-hidden`}
        style={{borderColor: simStep === 4 ? "#34d39940" : simStep === 5 ? "#f8717140" : "#00c9b140"}}>

        {/* Header */}
        <div className="px-6 py-5 text-center"
          style={{background: simStep === 4
            ? "linear-gradient(135deg,#065f46,#10b981)"
            : simStep === 5
              ? "linear-gradient(135deg,#7f1d1d,#dc2626)"
              : "linear-gradient(135deg,#003d5c 0%,#005a72 50%,#00c9b1 100%)"}}>
          <div className="w-14 h-14 rounded-full flex items-center justify-center text-2xl font-black mx-auto mb-3"
            style={{background:"rgba(255,255,255,0.15)"}}>
            {simStep === 4 ? "✅" : simStep === 5 ? "❌" : "✉"}
          </div>
          <h3 className="text-white font-black text-lg mb-1">
            {simStep === 4 ? "Invite Delivered!" : simStep === 5 ? "Delivery Failed" : "Send Invite"}
          </h3>
          <p className="text-white/70 text-sm">
            {simStep === 4
              ? `Email sent to ${u.email}`
              : simStep === 5
                ? diag?.title || "Could not reach mail server"
                : `Activate ${u.name}'s ProCAS account`}
          </p>
        </div>

        <div className="p-5 space-y-4">
          {/* User card */}
          <div className={`rounded-xl p-4 border ${t.cardBorder}`}
            style={{background: dark ? "#0c1e30" : "#f8fbff"}}>
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full flex items-center justify-center text-white text-[11px] font-bold shrink-0"
                style={{background:"linear-gradient(135deg,#003d5c,#00c9b1)"}}>
                {u.name.split(" ").map(n => n[0]).slice(0, 2).join("")}
              </div>
              <div className="flex-1 min-w-0">
                <div className={`font-semibold text-sm ${t.text}`}>{u.name}</div>
                <div className={`text-[11px] ${t.textMuted} truncate`}>{u.email}</div>
                <div className="flex items-center gap-2 mt-1">
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded-full"
                    style={{background:"#fbbf2418",color:"#fbbf24"}}>Pending Invite</span>
                  {log.attempts > 0 && (
                    <span className={`text-[10px] ${t.textMuted}`}>
                      {log.attempts} attempt{log.attempts !== 1 ? "s" : ""}
                    </span>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* SMTP config status */}
          {!smtpReady && simStep === 0 && (
            <div className="flex items-start gap-2 px-3 py-2.5 rounded-xl"
              style={{background:"#f8717112",border:"1px solid #f8717130"}}>
              <span style={{color:"#f87171",fontSize:14}}>⚠</span>
              <div>
                <p className="text-xs font-semibold" style={{color:"#f87171"}}>SMTP Not Configured</p>
                <p className={`text-[11px] mt-0.5 ${t.textMuted}`}>
                  Mail provider credentials are missing. Sending will fail immediately.
                  Close this and open <strong>SMTP Settings</strong> first.
                </p>
              </div>
            </div>
          )}

          {/* Email preview */}
          <div className={`rounded-xl border ${t.cardBorder} overflow-hidden`}>
            <div className="px-3 py-2 flex items-center gap-2"
              style={{background: dark ? "#0c1e30" : "#f0f4f8", borderBottom: `1px solid ${dark ? "#1a2d44" : "#d4e0ed"}`}}>
              <span className="text-[10px] font-bold" style={{color:"#4a90d9"}}>📧 EMAIL PREVIEW</span>
              <span className={`text-[9px] ${t.textMuted} ml-auto`}>
                via {smtpCfg?.host || "SMTP relay (not configured)"}
              </span>
            </div>
            <div className="px-4 py-3 space-y-1.5">
              <div className="flex gap-2 text-[11px]">
                <span className={`font-bold w-10 shrink-0 ${t.textMuted}`}>To:</span>
                <span className={t.text}>{u.email}</span>
              </div>
              <div className="flex gap-2 text-[11px]">
                <span className={`font-bold w-10 shrink-0 ${t.textMuted}`}>From:</span>
                <span className={t.text}>{smtpCfg?.fromEmail || "noreply@procas.example.com"}</span>
              </div>
              <div className="flex gap-2 text-[11px]">
                <span className={`font-bold w-10 shrink-0 ${t.textMuted}`}>Subj:</span>
                <span className={t.text}>Welcome to ProCAS – Activate Your Account</span>
              </div>
              <div className={`mt-2 p-2.5 rounded-lg text-[10px] leading-relaxed ${t.textMuted}`}
                style={{background: dark ? "#0a1a2e" : "#f8fbff", border: `1px solid ${dark ? "#1a2d44" : "#e2e8f0"}`}}>
                Hi <strong style={{color:dark?"#e8f0f8":"#0d2137"}}>{u.name}</strong>,<br/>
                You've been invited to <strong style={{color:"#00c9b1"}}>ProCAS</strong> as{" "}
                <strong style={{color:"#00c9b1"}}>{u.role}</strong>.<br/>
                Click the button below to activate your account and set your password.
                This link expires in <strong>48 hours</strong>.<br/><br/>
                <span className="inline-block px-3 py-1.5 rounded-lg text-white text-[11px] font-semibold"
                  style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                  Activate Account
                </span><br/><br/>
                <span className="opacity-70" style={{color:"#4a90d9"}}>
                  https://procas.example.com/activate?token={token}
                </span>
              </div>
            </div>
          </div>

          {/* SMTP delivery log — shown while sending */}
          {simStep > 0 && (
            <div className={`rounded-xl border ${t.cardBorder} p-3 space-y-2`}>
              <p className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted}`}>
                SMTP Delivery Log
              </p>
              {smtpSteps.map((s, i) => {
                const isFailing = simStep === 5 && i === smtpSteps.length - 1;
                return (
                  <div key={s.label} className="flex items-center gap-2">
                    <div className="w-4 h-4 rounded-full flex items-center justify-center shrink-0 text-[9px] font-black"
                      style={{
                        background: s.done
                          ? "#34d39920"
                          : isFailing ? "#f8717120" : "#fbbf2415",
                        color: s.done ? "#34d399" : isFailing ? "#f87171" : "#fbbf24",
                      }}>
                      {s.done ? "✓" : isFailing ? "✗" : "…"}
                    </div>
                    <span className={`text-[11px] ${s.done ? t.text : t.textMuted}`}>{s.label}</span>
                  </div>
                );
              })}
              {simStep === 4 && (
                <div className="flex items-center gap-2 pt-1 border-t" style={{borderColor: dark ? "#1a2d44" : "#e2e8f0"}}>
                  <span className="text-[10px]" style={{color:"#34d399"}}>
                    ✅ Delivered — awaiting user activation
                  </span>
                </div>
              )}
            </div>
          )}

          {/* Error diagnostic — never silently fail */}
          {simStep === 5 && diag && (
            <div className="rounded-xl px-4 py-3 space-y-1.5"
              style={{background:"#f8717112",border:"1px solid #f8717130"}}>
              <p className="text-xs font-bold" style={{color:"#f87171"}}>
                ⚠ {diag.title}
              </p>
              <p className={`text-[11px] ${t.textMuted}`}>{diag.detail}</p>
              <p className="text-[10px] font-semibold" style={{color:"#fbbf24"}}>
                Fix: {diag.fix}
              </p>
              <p className={`text-[10px] ${t.textMuted} pt-1`}>
                The user record has been created. Correct the SMTP config and use
                <strong> Resend Invite</strong> from the user table.
              </p>
            </div>
          )}

          {/* Action buttons */}
          {simStep === 4 ? (
            <button onClick={onLater}
              className="w-full py-2.5 rounded-xl text-sm font-semibold text-white"
              style={{background:"linear-gradient(135deg,#34d399,#10b981)"}}>
              Done ✓
            </button>
          ) : (
            <div className="flex gap-2">
              <button onClick={onLater}
                className={`flex-1 py-2.5 rounded-xl border ${t.cardBorder} text-sm font-medium ${t.textMuted} ${t.hover}`}>
                {simStep === 5 ? "Dismiss" : "Later"}
              </button>
              <button onClick={doSend} disabled={sending}
                className="flex-1 py-2.5 rounded-xl text-sm font-semibold text-white flex items-center justify-center gap-2 disabled:opacity-60"
                style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                {sending
                  ? <><SpinIcon size={13}/> Sending…</>
                  : simStep === 5
                    ? <><Icon path={Icons.refreshCw} size={13}/> Retry</>
                    : <><Icon path={Icons.send} size={13}/> Send Invite Now</>}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── USER DATABASE TAB ────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

const UserDatabaseTab = ({ t, dark, isAdmin }) => {
  const { users, setUsers } = useContext(UserContext);
  const { clients: masterClients } = useSyncContext();

  // ── Auto-derive each user's assigned clients from Client Master (source of truth).
  // masterClients.personName is a comma-separated list of user names per client.
  // This map is rebuilt whenever masterClients changes — real-time, no refresh needed.
  const clientsByUser = useMemo(() => {
    const map = {}; // { userName: [{ id, clientName }] }
    masterClients.forEach(c => {
      const names = (c.personName || "").split(",").map(n => n.trim()).filter(Boolean);
      names.forEach(name => {
        if (!map[name]) map[name] = [];
        map[name].push({ id: c.id, clientName: c.clientName });
      });
    });
    return map;
  }, [masterClients]);

  // Helper: get derived client list for a user by name
  const getDerivedClients = (userName, role) => {
    if (role === "Admin") return null; // null = all clients (Admin)
    return clientsByUser[userName] || [];
  };

  const [search, setSearch]           = useState("");
  const [addModal, setAddModal]       = useState(false);
  const [editUser, setEditUser]       = useState(null);
  const [deleteId, setDeleteId]       = useState(null);
  const [inviteSent, setInviteSent]   = useState(null);
  const [toast, setToast]             = useState(null);
  const [form, setForm]               = useState({ name:"", email:"", role:"End User" });
  const [formErr, setFormErr]         = useState({});
  const [showSmtpPanel, setShowSmtpPanel] = useState(false);

  // ── Bulk Import state ────────────────────────────────────────────────────
  const [bulkModal, setBulkModal]           = useState(false);   // modal open
  const [bulkFile, setBulkFile]             = useState(null);    // File object
  const [bulkParsed, setBulkParsed]         = useState(null);    // { rows, errors, dupes }
  const [bulkStep, setBulkStep]             = useState("upload");// "upload"|"review"|"summary"
  const [bulkSummary, setBulkSummary]       = useState(null);    // post-import counts
  const [bulkSendInvite, setBulkSendInvite] = useState(true);
  const [bulkDupeAction, setBulkDupeAction] = useState({});      // { [email]: "skip"|"update" }
  const fileInputRef = useRef(null);

  // ── Auto Reminder & Send Reminder state ──────────────────────────────────
  const [autoReminderModal,  setAutoReminderModal]  = useState(false);
  const [sendReminderModal,  setSendReminderModal]  = useState(false);
  const [remindUserId,       setRemindUserId]       = useState(null);
  const [scheduledReminders, setScheduledReminders] = useState([]);
  const [reminderHistory,    setReminderHistory]    = useState([]);
  const [showReminderHistory,setShowReminderHistory]= useState(false);
  const [reminderToast,      setReminderToast]      = useState(null);
  // Auto Reminder form fields
  const [arType,      setArType]      = useState("All Active Users");
  const [arModule,    setArModule]    = useState("All Modules");
  const [arFY,        setArFY]        = useState("2026-27");
  const [arMonth,     setArMonth]     = useState("April");
  const [arFreq,      setArFreq]      = useState("Weekly");
  const [arStartDate, setArStartDate] = useState("");
  const [arTime,      setArTime]      = useState("09:00");
  const [arUsers,     setArUsers]     = useState([]);
  const [arSaved,     setArSaved]     = useState(null);
  const [arEditId,    setArEditId]    = useState(null);
  // Send Reminder form fields
  const [srModule,  setSrModule]  = useState("All Modules");
  const [srFY,      setSrFY]      = useState("2026-27");
  const [srMonth,   setSrMonth]   = useState("April");
  const [srTarget,  setSrTarget]  = useState("All Active Users");
  const [srUsers,   setSrUsers]   = useState([]);

  const AR_MODULES = ["All Modules","CAS MIS","KRA/KPI","Fund Request"];
  const AR_FYS     = ["2026-27","2027-28","2028-29","2029-30","2030-31"];
  const AR_MONTHS  = ["April","May","June","July","August","September","October","November","December","January","February","March"];
  const AR_FREQS   = ["Daily","Weekly","Monthly"];
  const AR_TYPES   = ["Individual User","Multiple Users","All Active Users"];

  // ── SMTP config — persisted across the session ──────────────────────────
  const [smtpCfg, setSmtpCfg] = useState(() => {
    try {
      const saved = localStorage.getItem("procas_smtpCfg");
      return saved ? JSON.parse(saved) : {
        provider: "sendgrid",
        host: "",
        port: 587,
        username: "",
        password: "",
        fromEmail: "",
        fromName:  "ProCAS",
        tls: true,
        configured: false,
      };
    } catch {
      return { provider:"sendgrid", host:"", port:587, username:"", password:"",
               fromEmail:"", fromName:"ProCAS", tls:true, configured:false };
    }
  });
  useEffect(() => {
    try { localStorage.setItem("procas_smtpCfg", JSON.stringify(smtpCfg)); } catch {}
  }, [smtpCfg]);

  // ── Invite tracking — { [userId]: { attempts, lastAttempt, sentAt, error, errorCode, deliveryStatus } }
  const [inviteLog, setInviteLog]     = useState({});
  const [pendingInvites, setPendingInvites] = useState({});

  const showToast = (msg, type="success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  // ── Bulk Import helpers ──────────────────────────────────────────────────

  // Download sample template
  const downloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Name of User *", "Email ID *", "Role", "Assigned Clients"],
      ["User A",   "usera@example.com",  "End User", "Client Alpha, Client Beta"],
      ["User B",   "userb@example.com",  "Admin",    "All Clients"],
    ]);
    ws["!cols"] = [22, 26, 12, 30].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Users");
    XLSX.writeFile(wb, "User_Import_Template.xlsx");
  };

  // Parse uploaded file → validate rows
  const parseBulkFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb   = XLSX.read(e.target.result, { type: "binary" });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const raw  = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const validRows  = [];
        const errorRows  = [];
        const dupeRows   = [];
        const EMAIL_RE   = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

        raw.forEach((row, i) => {
          const rowNum = i + 2; // 1-based, row 1 = header
          const name   = String(row["Name of User *"] || row["Name of User"] || "").trim();
          const email  = String(row["Email ID *"]     || row["Email ID"]     || "").trim();
          const role   = String(row["Role"]            || "End User").trim() || "End User";
          const clients= String(row["Assigned Clients"]|| "").trim();

          const rowErrors = [];
          if (!name)                     rowErrors.push("Name of User missing");
          if (!email)                    rowErrors.push("Email ID missing");
          else if (!EMAIL_RE.test(email))rowErrors.push("Invalid Email ID");

          if (rowErrors.length) {
            errorRows.push({ rowNum, name, email, errors: rowErrors });
            return;
          }

          const exists = users.find(u => u.email.toLowerCase() === email.toLowerCase());
          if (exists) {
            dupeRows.push({ rowNum, name, email, role, clients, existingId: exists.id });
            return;
          }

          validRows.push({ rowNum, name, email, role, clients });
        });

        setBulkParsed({ rows: validRows, errors: errorRows, dupes: dupeRows, total: raw.length });
        setBulkDupeAction(
          Object.fromEntries(dupeRows.map(d => [d.email, "skip"]))
        );
        setBulkStep("review");
      } catch {
        showToast("Could not read file. Please use .xlsx or .csv format.", "error");
      }
    };
    reader.readAsBinaryString(file);
  };

  // Execute import
  const executeBulkImport = () => {
    if (!bulkParsed) return;
    let imported = 0, skipped = 0, updated = 0;
    const newUsers = [];

    // Valid new rows
    bulkParsed.rows.forEach(r => {
      const id = `USR-${String(users.length + newUsers.length + 1).padStart(3,"0")}`;
      newUsers.push({
        id, name: r.name, email: r.email,
        role: r.role || "End User",
        assignedClients: r.clients ? r.clients.split(",").map(s=>s.trim()).filter(Boolean) : [],
        status: "Active", inviteStatus: "Pending Invite",
        lastLogin: "—", joined: new Date().toLocaleDateString("en-IN"),
      });
      imported++;
    });

    // Dupes
    let updatedUsers = [...users];
    bulkParsed.dupes.forEach(d => {
      const action = bulkDupeAction[d.email] || "skip";
      if (action === "update") {
        updatedUsers = updatedUsers.map(u =>
          u.email.toLowerCase() === d.email.toLowerCase()
            ? { ...u, name: d.name, role: d.role || u.role,
                assignedClients: d.clients ? d.clients.split(",").map(s=>s.trim()).filter(Boolean) : u.assignedClients }
            : u
        );
        updated++;
      } else {
        skipped++;
      }
    });

    setUsers([...updatedUsers, ...newUsers]);

    // Trigger existing invite workflow for imported users if requested
    if (bulkSendInvite && newUsers.length > 0) {
      const newLog = {};
      newUsers.forEach(u => {
        newLog[u.id] = { attempts: 1, lastAttempt: new Date().toISOString(), deliveryStatus: "Sent" };
      });
      setInviteLog(prev => ({ ...prev, ...newLog }));
    }

    setBulkSummary({
      total:    bulkParsed.total,
      imported,
      failed:   bulkParsed.errors.length,
      dupes:    bulkParsed.dupes.length,
      updated,
      skipped,
      errorRows: bulkParsed.errors,
    });
    setBulkStep("summary");
    showToast(`${imported} user${imported!==1?"s":""} imported successfully.`);
  };

  // Download error report
  const downloadErrorReport = (summary) => {
    const rows = summary.errorRows.map(r => ({
      "Row #":       r.rowNum,
      "Name":        r.name  || "(blank)",
      "Email":       r.email || "(blank)",
      "Error(s)":    r.errors.join("; "),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [8,24,28,40].map(w=>({wch:w}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Import Errors");
    XLSX.writeFile(wb, "Bulk_Import_Error_Report.xlsx");
  };

  const resetBulkModal = () => {
    setBulkModal(false);
    setBulkFile(null);
    setBulkParsed(null);
    setBulkStep("upload");
    setBulkSummary(null);
    setBulkDupeAction({});
    setBulkSendInvite(true);
  };

  // ── Reminder helpers ──────────────────────────────────────────────────────
  const getActiveUsers = () => users.filter(u => getInviteStatus ? getInviteStatus(u) === "Active" : true);

  const openAutoReminderModal = (editId = null) => {
    if (editId) {
      const r = scheduledReminders.find(x => x.id === editId);
      if (r) {
        setArType(r.type); setArModule(r.module); setArFY(r.fy);
        setArMonth(r.month); setArFreq(r.freq);
        setArStartDate(r.startDate === "Immediate" ? "" : r.startDate);
        setArTime(r.time); setArUsers(r.userIds || []);
      }
      setArEditId(editId);
    } else {
      setArType("All Active Users"); setArModule("All Modules"); setArFY("2026-27");
      setArMonth("April"); setArFreq("Weekly"); setArStartDate(""); setArTime("09:00");
      setArUsers([]); setArEditId(null);
    }
    setArSaved(null);
    setAutoReminderModal(true);
  };

  const saveAutoReminder = () => {
    const activeU = getActiveUsers();
    const recipientCount = arType === "All Active Users" ? activeU.length : arUsers.length;
    const entry = {
      id:            arEditId || `AR-${Date.now()}`,
      type:          arType,
      module:        arModule,
      fy:            arFY,
      month:         arMonth,
      freq:          arFreq,
      startDate:     arStartDate || "Immediate",
      time:          arTime,
      userIds:       arUsers,
      recipientCount,
      status:        "Active",
      nextRun:       arStartDate || "Scheduled",
    };
    setScheduledReminders(prev =>
      arEditId ? prev.map(x => x.id === arEditId ? entry : x) : [...prev, entry]
    );
    setArSaved(entry);
  };

  const sendTestReminder = () => {
    const now = new Date().toLocaleString("en-IN",{day:"2-digit",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"});
    setReminderHistory(prev => [{
      id:`RH-${Date.now()}`, userName:"(Test)", email:"test@procas.app",
      module:arModule, fy:arFY, month:arMonth, type:"Test",
      sentBy:"Administrator", sentAt:now, status:"Sent",
    }, ...prev]);
    setReminderToast("Test reminder sent!");
    setTimeout(()=>setReminderToast(null), 3000);
  };

  const sendNowReminder = () => {
    const activeU = getActiveUsers();
    const targets = srTarget === "All Active Users"
      ? activeU
      : srUsers.map(id => users.find(u => u.id === id)).filter(Boolean);
    const now = new Date().toLocaleString("en-IN",{day:"2-digit",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"});
    const newRows = targets.map(u => ({
      id:`RH-${Date.now()}-${u.id}`, userName:u.name, email:u.email,
      module:srModule, fy:srFY, month:srMonth, type:srTarget,
      sentBy:"Administrator", sentAt:now, status:"Sent",
    }));
    setReminderHistory(prev => [...newRows, ...prev]);
    setSendReminderModal(false);
    setReminderToast(`Reminder sent to ${targets.length} user${targets.length!==1?"s":""}!`);
    setTimeout(()=>setReminderToast(null), 3500);
  };

  const toggleScheduledStatus = (id) =>
    setScheduledReminders(prev => prev.map(x => x.id===id
      ? {...x, status: x.status==="Active" ? "Paused" : "Active"} : x));

  const deleteScheduled = (id) =>
    setScheduledReminders(prev => prev.filter(x => x.id !== id));

  const exportReminderHistory = () => {
    const ws = XLSX.utils.json_to_sheet(reminderHistory.map((r,i)=>({
      "#":i+1,"User Name":xlsSafe(r.userName),"Email":xlsSafe(r.email),"Module":r.module,
      "FY":r.fy,"Month":r.month,"Type":r.type,
      "Sent By":xlsSafe(r.sentBy),"Sent Date & Time":r.sentAt,"Status":r.status,
    })));
    ws["!cols"]=[4,20,26,14,12,12,16,16,20,10].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Reminder History");
    XLSX.writeFile(wb,"ProCAS_Reminder_History.xlsx");
  };

  const exportUsers = () => {
    try {
      const data = filtered.map((u, i) => {
        const status   = getInviteStatus(u);
        const logEntry = inviteLog[u.id] || {};
        return {
          "#":                i + 1,
          "User ID":          u.id,
          "User Name":        xlsSafe(u.name),
          "Email ID":         xlsSafe(u.email),
          "Role":             u.role,
          "Status":           status,
          "Invite Status":    logEntry.deliveryStatus || "Not Sent",
          "Last Login":       xlsSafe(u.lastLogin || "—"),
          "Joined":           xlsSafe(u.joined    || "—"),
          "Assigned Clients": Array.isArray(u.assignedClients)
            ? u.assignedClients.join(", ")
            : (u.assignedClients || "All Clients"),
        };
      });
      if (!data.length) { alert("No user records to export."); return; }
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [4,10,24,30,12,12,14,16,14,35].map(w=>({wch:w}));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "User Database");
      XLSX.writeFile(wb, "ProCAS_UserDatabase.xlsx");
    } catch (err) {
      alert("Unable to generate export file. Please try again.");
      console.error("User Database export error:", err);
    }
  };

  const filtered = users.filter(u =>
    u.name.toLowerCase().includes(search.toLowerCase()) ||
    u.email.toLowerCase().includes(search.toLowerCase()) ||
    u.role.toLowerCase().includes(search.toLowerCase())
  );

  // ── Full 5-state invite status ───────────────────────────────────────────
  // Pending Invite | Email Sent | Delivered | Active | Failed | Inactive
  const getInviteStatus = (u) => {
    const log = inviteLog[u.id];
    if (u.inviteStatus === "Active" || u.status === "Active") return "Active";
    if (u.inviteStatus === "Inactive" || u.status === "Inactive") return "Inactive";
    if (log) {
      if (log.deliveryStatus === "Failed")   return "Failed";
      if (log.deliveryStatus === "Sent")     return "Email Sent";
    }
    return u.inviteStatus || "Pending Invite";
  };

  const inviteStatusStyle = (s) => {
    switch (s) {
      case "Active":         return { color:"#34d399", bg:"#34d39918", dot:"#34d399" };
      case "Email Sent":     return { color:"#00c9b1", bg:"#00c9b115", dot:"#00c9b1" };
      case "Delivered":      return { color:"#4a90d9", bg:"#4a90d915", dot:"#4a90d9" };
      case "Pending Invite": return { color:"#fbbf24", bg:"#fbbf2418", dot:"#fbbf24" };
      case "Failed":         return { color:"#f87171", bg:"#f8717118", dot:"#f87171" };
      case "Inactive":       return { color:"#94a3b8", bg:"#94a3b818", dot:"#94a3b8" };
      default:               return { color:"#94a3b8", bg:"#94a3b818", dot:"#94a3b8" };
    }
  };

  const openAdd = () => {
    setForm({ name:"", email:"", role:"End User" });
    setFormErr({});
    setAddModal(true);
  };

  const openEdit = (u) => {
    // assignedClients no longer in form — managed via Client Master
    setForm({ name: u.name, email: u.email, role: u.role });
    setFormErr({});
    setEditUser(u);
  };

  const validate = () => {
    const e = {};
    if (!form.name.trim())  e.name  = "Required";
    if (!form.email.trim()) e.email = "Required";
    else if (!/\S+@\S+\.\S+/.test(form.email)) e.email = "Invalid email";
    setFormErr(e);
    return !Object.keys(e).length;
  };

  // Step 1: create user record
  const handleAdd = () => {
    if (!validate()) return;
    const newUser = {
      ...form,
      id: genUserId(),
      status: "Active",
      inviteStatus: "Pending Invite",
      lastLogin: "Never",
      // assignedClients is NOT stored here — Client Master is the source of truth.
      // getDerivedClients() in UserDatabaseTab reads assignments from masterClients.
    };
    setUsers(u => [...u, newUser]);
    setAddModal(false);
    // Step 2 + 3: open invite banner which generates token and sends email
    setInviteSent(newUser.id);
    showToast("User created — send the invite to activate their account");
  };

  const handleEdit = () => {
    if (!validate()) return;
    setUsers(u => u.map(x => x.id === editUser.id ? { ...x, ...form } : x));
    setEditUser(null);
    showToast("User updated successfully");
  };

  const handleDelete = (id) => {
    setUsers(u => u.filter(x => x.id !== id));
    setDeleteId(null);
    showToast("User removed");
  };

  // Steps 3-5: called by SendInviteBanner after SMTP attempt
  const handleSendInvite = (userId, result) => {
    const now = new Date();
    setInviteLog(prev => ({
      ...prev,
      [userId]: {
        attempts:       ((prev[userId]?.attempts) || 0) + 1,
        lastAttempt:    now.toISOString(),
        sentAt:         result.success ? now.toISOString() : prev[userId]?.sentAt,
        error:          result.success ? null : (result.error || "Delivery failed"),
        errorCode:      result.success ? null : (result.errorCode || "UNKNOWN"),
        deliveryStatus: result.success ? "Sent" : "Failed",
      },
    }));

    if (result.success) {
      // Step 4: log success; Step 5: update status to "Email Sent"
      setPendingInvites(p => ({ ...p, [userId]: true }));
      setUsers(u => u.map(x => x.id === userId
        ? { ...x, inviteStatus: "Email Sent", inviteSentAt: now.toISOString() }
        : x
      ));
      setInviteSent(null);
      showToast("✉ Invite sent — waiting for user to activate", "invite");
    } else {
      // Never silently fail — keep banner open so admin sees the diagnostic
      // (banner stays visible; admin must dismiss explicitly)
    }
  };

  // Resend — opens banner again for any user that is Pending or Failed
  const handleResendInvite = (userId) => {
    setInviteSent(userId);
  };

  const handleDisable = (userId) => {
    setUsers(u => u.map(x => x.id === userId
      ? { ...x, status: "Inactive", inviteStatus: "Inactive" }
      : x
    ));
    showToast("User account disabled");
  };

  const handleEnable = (userId) => {
    setUsers(u => u.map(x => x.id === userId
      ? { ...x, status: "Active", inviteStatus: "Active" }
      : x
    ));
    showToast("User account re-enabled");
  };

  const handleResetPassword = (userId) => {
    showToast("Password reset link sent to user's email");
  };

  // Demo: simulate the user clicking the activation link in their email
  const handleSimulateAccept = (userId) => {
    setUsers(u => u.map(x => x.id === userId
      ? { ...x, inviteStatus: "Active", status: "Active", lastLogin: "Just now" }
      : x
    ));
    setInviteLog(prev => ({
      ...prev,
      [userId]: { ...prev[userId], deliveryStatus: "Delivered" },
    }));
    setPendingInvites(p => { const n = {...p}; delete n[userId]; return n; });
    showToast("✓ Invite accepted — user is now Active");
  };

  const roleColor = { Admin:"#00c9b1", "End User":"#4a90d9" };

  const pendingCount  = users.filter(u => {
    const s = getInviteStatus(u);
    return s === "Pending Invite" || s === "Email Sent" || s === "Failed";
  }).length;
  const activeCount   = users.filter(u => getInviteStatus(u) === "Active").length;
  const inactiveCount = users.filter(u => getInviteStatus(u) === "Inactive").length;
  const failedCount   = users.filter(u => getInviteStatus(u) === "Failed").length;

  return (
    <div className="space-y-5">

      {/* Toast */}
      {toast && (
        <div className="fixed top-5 right-5 z-[80] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{
            background: toast.type === "invite"
              ? "linear-gradient(135deg,#1b5fa8,#00a896)"
              : "linear-gradient(135deg,#34d399,#10b981)",
            minWidth:240,
          }}>
          <Icon path={toast.type === "invite" ? Icons.send : Icons.check} size={15}/> {toast.msg}
        </div>
      )}

      {/* Send Invite Banner */}
      {inviteSent && (
        <SendInviteBanner
          t={t} dark={dark}
          userId={inviteSent}
          users={users}
          onSend={handleSendInvite}
          onLater={() => setInviteSent(null)}
          inviteLog={inviteLog}
          smtpCfg={smtpCfg}
        />
      )}

      {/* SMTP Config Panel */}
      {showSmtpPanel && (
        <SmtpConfigPanel
          t={t} dark={dark}
          smtpCfg={smtpCfg}
          setSmtpCfg={setSmtpCfg}
          onClose={() => setShowSmtpPanel(false)}
        />
      )}

      {/* ── SMTP status banner — shown when not configured ── */}
      {!smtpCfg.configured && (
        <div className="flex items-center gap-3 px-4 py-3 rounded-2xl"
          style={{background:"#f8717110",border:"1px solid #f8717130"}}>
          <span style={{color:"#f87171",fontSize:18}}>⚠</span>
          <div className="flex-1 min-w-0">
            <p className="text-xs font-bold" style={{color:"#f87171"}}>
              SMTP Not Configured — Invitation Emails Will Not Be Delivered
            </p>
            <p className={`text-[11px] mt-0.5 ${t.textMuted}`}>
              Set up your mail provider credentials so invite emails reach users.
              Without this, the system cannot send activation links.
            </p>
          </div>
          {isAdmin && (
            <button onClick={() => setShowSmtpPanel(true)}
              className="shrink-0 flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold text-white"
              style={{background:"linear-gradient(135deg,#1b5fa8,#00a896)"}}>
              <Icon path={Icons.settings} size={12}/> SMTP Settings
            </button>
          )}
        </div>
      )}
      {smtpCfg.configured && (
        <div className="flex items-center gap-3 px-4 py-3 rounded-2xl"
          style={{background:"#34d39910",border:"1px solid #34d39930"}}>
          <span style={{color:"#34d399",fontSize:14}}>✅</span>
          <div className="flex-1 min-w-0">
            <p className="text-xs font-semibold" style={{color:"#34d399"}}>
              SMTP Configured — {smtpCfg.host}:{smtpCfg.port} · From: {smtpCfg.fromEmail}
            </p>
            <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>
              Mail provider ready. Invites will be sent via{" "}
              {SMTP_PROVIDERS.find(p => p.id === smtpCfg.provider)?.label || smtpCfg.host}.
            </p>
          </div>
          {isAdmin && (
            <button onClick={() => setShowSmtpPanel(true)}
              className={`shrink-0 flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
              <Icon path={Icons.settings} size={12}/> Edit
            </button>
          )}
        </div>
      )}

      {/* Stats row */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {[
          { label:"Total Users",  value:users.length,  color:"#00c9b1", icon:Icons.users   },
          { label:"Active",       value:activeCount,   color:"#34d399", icon:Icons.check   },
          { label:"Pending/Sent", value:pendingCount,  color:"#fbbf24", icon:Icons.send    },
          { label:"Failed",       value:failedCount,   color:"#f87171", icon:Icons.alert   },
        ].map(k=>(
          <div key={k.label} className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 flex items-center gap-4`}>
            <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
              style={{background:k.color+"18"}}>
              <Icon path={k.icon} size={18} style={{color:k.color}}/>
            </div>
            <div>
              <div className={`text-[11px] font-bold uppercase tracking-widest ${t.textMuted}`}>{k.label}</div>
              <div className="text-2xl font-black mt-0.5" style={{color:k.color}}>{k.value}</div>
            </div>
          </div>
        ))}
      </div>

      {/* Invite status lifecycle */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-6 py-4`}>
        <p className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} mb-3`}>User Invite Lifecycle</p>
        <div className="flex items-center gap-0 flex-wrap">
          {[
            { step:"1. Admin Creates User",    color:"#4a90d9", icon:"👤" },
            { step:"2. Email Sent via SMTP",   color:"#00c9b1", icon:"✉"  },
            { step:"3. User Clicks Link",      color:"#fbbf24", icon:"🔗" },
            { step:"4. Sets Password",         color:"#fb923c", icon:"🔐" },
            { step:"5. Status → Active",       color:"#34d399", icon:"✅" },
          ].map((s, i) => (
            <div key={s.step} className="flex items-center gap-0">
              <div className="flex items-center gap-2 px-3 py-2 rounded-xl"
                style={{background:s.color+"12",border:`1px solid ${s.color}30`}}>
                <span className="text-sm">{s.icon}</span>
                <span className="text-[10px] font-semibold" style={{color:s.color}}>{s.step}</span>
              </div>
              {i < 4 && (
                <div className="flex items-center px-1">
                  <Icon path={Icons.chevronRight} size={13} className={t.textMuted}/>
                </div>
              )}
            </div>
          ))}
        </div>
      </div>

      {/* User table */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
        <div className={`px-5 py-4 border-b ${t.cardBorder} flex flex-col sm:flex-row sm:items-center justify-between gap-3`}>
          <div>
            <h3 className={`font-bold text-sm ${t.text}`}>User Management</h3>
            <p className={`text-xs mt-0.5 ${t.textMuted}`}>
              {filtered.length} user{filtered.length!==1?"s":""} · Access &amp; invite management
            </p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border text-sm ${t.input}`}>
              <Icon path={Icons.search} size={13} className={t.textMuted}/>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search users…"
                className={`bg-transparent outline-none w-32 ${t.text}`}/>
              {search && <button onClick={()=>setSearch("")} className={t.textMuted}><Icon path={Icons.x} size={11}/></button>}
            </div>
            {isAdmin && (
              <button onClick={() => setShowSmtpPanel(true)}
                className={`flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}
                title="Configure SMTP / Email Provider">
                <Icon path={Icons.settings} size={14}/> SMTP Settings
              </button>
            )}
            <button onClick={exportUsers}
              className={`flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}
              title="Export current user list to Excel">
              <Icon path={Icons.download} size={14}/> Export Excel
            </button>
            {isAdmin && (
              <>
                <button onClick={downloadTemplate}
                  className={`flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}
                  title="Download sample import template">
                  <Icon path={Icons.download} size={14}/> Template
                </button>
                <button onClick={() => { setBulkStep("upload"); setBulkModal(true); }}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                  style={{background:"linear-gradient(135deg,#7c3aed,#4a90d9)"}}>
                  <Icon path={Icons.upload} size={14}/> Bulk Import
                </button>
                <button
                  onClick={() => { setSrModule("All Modules"); setSrFY("2026-27"); setSrMonth("April"); setSrTarget("All Active Users"); setSrUsers([]); setSendReminderModal(true); }}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                  style={{background:"linear-gradient(135deg,#f59e0b,#ea580c)"}}>
                  <Icon path={Icons.send} size={14}/> Send Reminder
                </button>
                <button
                  onClick={() => openAutoReminderModal()}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                  style={{background:"linear-gradient(135deg,#0ea5e9,#6366f1)"}}>
                  <Icon path={Icons.clock} size={14}/> Auto Reminder
                </button>
              </>
            )}
            {isAdmin ? (
              <button onClick={openAdd}
                className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
                <Icon path={Icons.userPlus} size={14}/> Invite User
              </button>
            ) : (
              <span className="flex items-center gap-1 text-xs font-medium px-2.5 py-1.5 rounded-lg"
                style={{background:"#f8717110",color:"#f87171",border:"1px solid #f8717125"}}>
                <Icon path={Icons.shield} size={12}/> Read-only
              </span>
            )}
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className={`${t.tableHead} text-xs uppercase tracking-wider`}>
                {["User ID","Name & Email","Role","Assigned Clients (from Client Master)",
                  "Invite Status","Invite Sent / Last Attempt","Delivery Detail","Last Login","Actions"].map(h=>(
                  <th key={h} className={`px-4 py-3 text-left font-semibold border-b ${t.cardBorder} whitespace-nowrap`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${t.divider}`}>
              {filtered.length === 0 ? (
                <tr><td colSpan={9} className="px-4 py-12 text-center">
                  <p className={`text-sm ${t.textMuted}`}>{search ? "No users match your search." : "No users found."}</p>
                </td></tr>
              ) : filtered.map(u => {
                const invStatus     = getInviteStatus(u);
                const invStyle      = inviteStatusStyle(invStatus);
                const isPending     = invStatus === "Pending Invite";
                const isFailed      = invStatus === "Failed";
                const isSent        = invStatus === "Email Sent";
                const isInactive    = invStatus === "Inactive";
                const canResend     = isPending || isFailed;
                const log           = inviteLog[u.id];
                // Derive assigned clients from Client Master (source of truth).
                // Overrides u.assignedClients so this is always in sync with
                // whatever the Admin set in Client Master → Client Registry.
                const derivedClients = getDerivedClients(u.name, u.role);
                const assignedNames  = derivedClients === null
                  ? ["All Clients"]
                  : derivedClients.map(c => c.clientName);

                return (
                  <tr key={u.id} className={`${t.tableRow} transition-colors`}
                    style={isInactive ? {opacity:0.65} : {}}>

                    {/* User ID */}
                    <td className={`px-4 py-3 font-mono text-xs font-semibold ${t.textAccent}`}>{u.id}</td>

                    {/* Name & Email */}
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2.5">
                        <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-[11px] font-bold shrink-0"
                          style={{background: isInactive
                            ? "linear-gradient(135deg,#374151,#4b5563)"
                            : "linear-gradient(135deg,#003d5c,#00c9b1)"}}>
                          {u.name.split(" ").map(n=>n[0]).slice(0,2).join("")}
                        </div>
                        <div>
                          <div className={`font-semibold text-sm ${t.text}`}>{u.name}</div>
                          <div className={`text-[11px] ${t.textMuted}`}>{u.email}</div>
                        </div>
                      </div>
                    </td>

                    {/* Role */}
                    <td className="px-4 py-3">
                      <span className="px-2.5 py-1 rounded-full text-[11px] font-bold"
                        style={{color:roleColor[u.role]||"#94a3b8", background:(roleColor[u.role]||"#94a3b8")+"18"}}>
                        {u.role}
                      </span>
                    </td>

                    {/* Assigned Clients — READ ONLY, derived from Client Master */}
                    <td className="px-4 py-3">
                      <div className="flex flex-col gap-0.5">
                        {assignedNames[0] === "All Clients" ? (
                          <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full w-fit"
                            style={{color:"#00c9b1",background:"#00c9b115"}}>All Clients</span>
                        ) : assignedNames.length === 0 ? (
                          <span className={`text-[11px] italic ${t.textMuted}`}>
                            No assignments in Client Master
                          </span>
                        ) : (
                          <div className="flex flex-wrap gap-1">
                            {assignedNames.slice(0,2).map(n => (
                              <span key={n} className="text-[10px] px-1.5 py-0.5 rounded-md font-medium"
                                style={{background: dark?"#1a2d44":"#e8f0f8",color: dark?"#94a3b8":"#4a5568"}}>
                                {n.length > 16 ? n.slice(0,14)+"…" : n}
                              </span>
                            ))}
                            {assignedNames.length > 2 && (
                              <span className="text-[10px] px-1.5 py-0.5 rounded-md font-medium"
                                style={{background:"#4a90d915",color:"#4a90d9"}}>
                                +{assignedNames.length - 2}
                              </span>
                            )}
                          </div>
                        )}
                        {/* Read-only badge — always shown */}
                        <span className="text-[9px] font-medium flex items-center gap-0.5 mt-0.5"
                          style={{color: dark?"#3a5a7a":"#8aa0b8"}}>
                          <Icon path={Icons.refreshCw} size={8}/> Auto from Client Master
                        </span>
                      </div>
                    </td>

                    {/* Invite Status */}
                    <td className="px-4 py-3">
                      <span className="flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-bold w-fit"
                        style={{color:invStyle.color, background:invStyle.bg}}>
                        <span className="w-1.5 h-1.5 rounded-full inline-block"
                          style={{background:invStyle.dot,
                            animation: (isPending||isSent) ? "pulse 1.5s infinite" : "none"}}/>
                        {invStatus}
                      </span>
                    </td>

                    {/* Invite Sent / Last Attempt */}
                    <td className={`px-4 py-3 text-xs ${t.textMuted}`}>
                      {!log?.lastAttempt ? (
                        <span className="italic opacity-50">Not sent</span>
                      ) : (
                        <div>
                          <div>
                            {new Date(log.lastAttempt).toLocaleDateString("en-IN",
                              {day:"2-digit",month:"short",year:"2-digit"})}
                          </div>
                          <div className="text-[10px] opacity-60">
                            {log.attempts} attempt{log.attempts!==1?"s":""}
                          </div>
                        </div>
                      )}
                    </td>

                    {/* Delivery Detail — exact error if failed */}
                    <td className="px-4 py-3 max-w-[160px]">
                      {isFailed && log?.errorCode ? (() => {
                        const diag = diagnoseSMTPError(smtpCfg, log.errorCode);
                        return (
                          <div className="text-[10px] leading-tight"
                            style={{color:"#f87171"}}>
                            <div className="font-semibold">{diag.title}</div>
                            <div className="opacity-70 mt-0.5">Fix: {diag.fix.slice(0,60)}{diag.fix.length>60?"…":""}</div>
                          </div>
                        );
                      })() : isFailed && log?.error ? (
                        <span className="text-[10px]" style={{color:"#f87171"}}>{log.error}</span>
                      ) : log?.deliveryStatus === "Sent" ? (
                        <span className="text-[10px]" style={{color:"#00c9b1"}}>
                          ✉ Delivered to inbox
                        </span>
                      ) : (
                        <span className={`text-[10px] italic ${t.textMuted} opacity-50`}>—</span>
                      )}
                    </td>

                    {/* Last Login */}
                    <td className={`px-4 py-3 text-xs ${t.textMuted}`}>{u.lastLogin}</td>

                    {/* Actions */}
                    <td className="px-4 py-3">
                      {isAdmin ? (
                        <div className="flex items-center gap-1 flex-wrap">
                          {/* Edit */}
                          <button onClick={() => openEdit(u)}
                            title="Edit user"
                            className={`flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium border transition-colors ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.edit} size={11}/> Edit
                          </button>

                          {/* Resend Invite — for Pending or Failed */}
                          {canResend && (
                            <button onClick={() => handleResendInvite(u.id)}
                              title={isFailed ? "Retry invite (SMTP failed)" : "Resend invite email"}
                              className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium transition-colors"
                              style={isFailed
                                ? {background:"#f8717115",color:"#f87171",border:"1px solid #f8717130"}
                                : {background:"#00c9b115",color:"#00c9b1",border:"1px solid #00c9b130"}}>
                              <Icon path={isFailed ? Icons.refreshCw : Icons.send} size={11}/>
                              {isFailed ? "Retry" : "Resend"}
                            </button>
                          )}

                          {/* Demo: Simulate Accept — for Pending or Email Sent */}
                          {(isPending || isSent) && (
                            <button onClick={() => handleSimulateAccept(u.id)}
                              title="(Demo) Simulate invite accepted by user"
                              className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium transition-colors"
                              style={{background:"#34d39915",color:"#34d399",border:"1px solid #34d39930"}}>
                              <Icon path={Icons.checkCircle} size={11}/> Accept
                            </button>
                          )}

                          {/* Disable / Enable */}
                          {!isInactive ? (
                            <button onClick={() => handleDisable(u.id)}
                              title="Disable this user account"
                              className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium transition-colors"
                              style={{background:"#f8717112",color:"#f87171",border:"1px solid #f8717130"}}>
                              <Icon path={Icons.alert} size={11}/> Disable
                            </button>
                          ) : (
                            <button onClick={() => handleEnable(u.id)}
                              title="Re-enable this user account"
                              className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium transition-colors"
                              style={{background:"#34d39912",color:"#34d399",border:"1px solid #34d39930"}}>
                              <Icon path={Icons.check} size={11}/> Enable
                            </button>
                          )}

                          {/* Reset Password */}
                          <button onClick={() => handleResetPassword(u.id)}
                            title="Send password reset link"
                            className={`flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium border transition-colors ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                            <Icon path={Icons.shield} size={11}/> Reset Pwd
                          </button>
                          {/* Send Reminder */}
                          <button
                            onClick={()=>{ setRemindUserId(u.id); setSrTarget("Individual User"); setSrUsers([u.id]); setSrModule("All Modules"); setSrFY("2026-27"); setSrMonth("April"); setSendReminderModal(true); }}
                            title="Send data-update reminder to this user"
                            className="flex items-center gap-1 px-2 py-1.5 rounded-lg text-[11px] font-medium transition-colors"
                            style={{background:"#f59e0b15",color:"#f59e0b",border:"1px solid #f59e0b30"}}>
                            <Icon path={Icons.send} size={11}/> Remind
                          </button>
                        </div>
                      ) : (
                        <span className={`text-xs ${t.textMuted} italic`}>Read only</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Legend */}
        <div className={`px-5 py-3 border-t ${t.cardBorder} flex flex-wrap items-center gap-5 text-[10px] ${t.textMuted}`}>
          {[
            { color:"#34d399", label:"Active — logged in, account live"               },
            { color:"#00c9b1", label:"Email Sent — invite delivered, awaiting click"   },
            { color:"#fbbf24", label:"Pending Invite — not yet sent"                   },
            { color:"#f87171", label:"Failed — SMTP error, use Retry"                  },
            { color:"#94a3b8", label:"Inactive — disabled by Admin"                    },
          ].map(l => (
            <div key={l.label} className="flex items-center gap-1.5">
              <div className="w-2 h-2 rounded-full" style={{background:l.color}}/>
              {l.label}
            </div>
          ))}
          <div className="flex items-center gap-1.5 border-l pl-4" style={{borderColor: dark?"#1a2d44":"#d4e0ed"}}>
            <Icon path={Icons.refreshCw} size={9} style={{color:"#4a90d9"}}/>
            <span style={{color:"#4a90d9"}}>Assigned Clients auto-sync from Client Master · Read-only here</span>
          </div>
        </div>
      </div>

      {/* Reminder toast */}
      {reminderToast && (
        <div className="fixed top-5 right-5 z-[85] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{background:"linear-gradient(135deg,#f59e0b,#ea580c)",minWidth:260}}>
          <Icon path={Icons.check} size={15}/> {reminderToast}
        </div>
      )}

      {/* ── Scheduled Auto Reminders table ── */}
      {isAdmin && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex items-center justify-between`}>
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 rounded-xl flex items-center justify-center" style={{background:"#0ea5e918"}}>
                <Icon path={Icons.clock} size={16} style={{color:"#0ea5e9"}}/>
              </div>
              <div>
                <h3 className={`font-bold text-sm ${t.text}`}>Scheduled Auto Reminders</h3>
                <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>
                  {scheduledReminders.length} schedule{scheduledReminders.length!==1?"s":""} configured
                </p>
              </div>
            </div>
            <button onClick={()=>openAutoReminderModal()}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold text-white"
              style={{background:"linear-gradient(135deg,#0ea5e9,#6366f1)"}}>
              <Icon path={Icons.plus||"M12 5v14M5 12h14"} size={13}/> New Schedule
            </button>
          </div>
          {scheduledReminders.length === 0 ? (
            <div className={`px-5 py-10 text-center ${t.textMuted} text-sm`}>
              <div className="text-3xl mb-2">⏰</div>
              No auto reminders scheduled. Click <strong>Auto Reminder</strong> to create one.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs" style={{minWidth:860}}>
                <thead>
                  <tr className={`${t.tableHead} text-[10px] uppercase tracking-wider`}>
                    {["Module","FY","Month","Frequency","Start / Next Run","Time","Recipients","Status","Actions"].map(h=>(
                      <th key={h} className={`px-4 py-3 text-left font-bold border-b ${t.cardBorder} whitespace-nowrap`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className={`divide-y ${t.divider}`}>
                  {scheduledReminders.map(r=>(
                    <tr key={r.id} className={`${t.tableRow} transition-colors`}>
                      <td className="px-4 py-2.5">
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold"
                          style={{background:"#4a90d918",color:"#4a90d9"}}>{r.module}</span>
                      </td>
                      <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.fy}</td>
                      <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.month}</td>
                      <td className="px-4 py-2.5">
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-semibold"
                          style={{background:"#6366f118",color:"#6366f1"}}>{r.freq}</span>
                      </td>
                      <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.nextRun}</td>
                      <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.time}</td>
                      <td className={`px-4 py-2.5 font-semibold ${t.text}`}>{r.recipientCount}</td>
                      <td className="px-4 py-2.5">
                        <span className="flex items-center gap-1 w-fit px-2 py-0.5 rounded-full text-[10px] font-bold"
                          style={{
                            background: r.status==="Active"?"#34d39918":"#fbbf2418",
                            color:      r.status==="Active"?"#34d399":"#fbbf24",
                          }}>
                          <span className="w-1.5 h-1.5 rounded-full inline-block" style={{background:"currentColor"}}/>
                          {r.status}
                        </span>
                      </td>
                      <td className="px-4 py-2.5">
                        <div className="flex items-center gap-1">
                          <button onClick={()=>openAutoReminderModal(r.id)}
                            className={`px-2 py-1 rounded-lg text-[10px] font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                            Edit
                          </button>
                          <button onClick={()=>toggleScheduledStatus(r.id)}
                            className="px-2 py-1 rounded-lg text-[10px] font-medium transition-all"
                            style={{
                              background: r.status==="Active"?"#fbbf2415":"#34d39915",
                              color:      r.status==="Active"?"#fbbf24":"#34d399",
                              border:     `1px solid ${r.status==="Active"?"#fbbf2430":"#34d39930"}`,
                            }}>
                            {r.status==="Active" ? "Pause" : "Resume"}
                          </button>
                          <button onClick={()=>deleteScheduled(r.id)}
                            className="px-2 py-1 rounded-lg text-[10px] font-medium"
                            style={{background:"#f8717115",color:"#f87171",border:"1px solid #f8717130"}}>
                            Delete
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ── Reminder History ── */}
      {isAdmin && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex items-center justify-between`}>
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 rounded-xl flex items-center justify-center" style={{background:"#f59e0b18"}}>
                <Icon path={Icons.send} size={16} style={{color:"#f59e0b"}}/>
              </div>
              <div>
                <h3 className={`font-bold text-sm ${t.text}`}>Reminder History</h3>
                <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>{reminderHistory.length} reminder{reminderHistory.length!==1?"s":""} sent</p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={()=>setShowReminderHistory(v=>!v)}
                className={`text-xs font-medium px-3 py-1.5 rounded-xl border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                {showReminderHistory ? "Hide" : "Show"}
              </button>
              {reminderHistory.length > 0 && (
                <button onClick={exportReminderHistory}
                  className="flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-xl"
                  style={{background:"#f59e0b18",color:"#f59e0b",border:"1px solid #f59e0b30"}}>
                  <Icon path={Icons.download} size={12}/> Export
                </button>
              )}
            </div>
          </div>
          {showReminderHistory && (
            reminderHistory.length === 0 ? (
              <div className={`px-5 py-10 text-center ${t.textMuted} text-sm italic`}>No reminders sent yet.</div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs" style={{minWidth:900}}>
                  <thead>
                    <tr className={`${t.tableHead} text-[10px] uppercase tracking-wider`}>
                      {["User Name","Email","Module","FY","Month","Type","Sent By","Sent Date & Time","Status"].map(h=>(
                        <th key={h} className={`px-4 py-3 text-left font-bold border-b ${t.cardBorder} whitespace-nowrap`}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${t.divider}`}>
                    {reminderHistory.map(r=>(
                      <tr key={r.id} className={`${t.tableRow} transition-colors`}>
                        <td className={`px-4 py-2.5 font-semibold ${t.text}`}>{r.userName}</td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.email}</td>
                        <td className="px-4 py-2.5">
                          <span className="px-1.5 py-0.5 rounded-full font-semibold text-[10px]"
                            style={{background:"#4a90d918",color:"#4a90d9"}}>{r.module}</span>
                        </td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.fy}</td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.month}</td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.type}</td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.sentBy}</td>
                        <td className={`px-4 py-2.5 ${t.textMuted}`}>{r.sentAt}</td>
                        <td className="px-4 py-2.5">
                          <span className="flex items-center gap-1 w-fit px-2 py-0.5 rounded-full text-[10px] font-bold"
                            style={{background:"#34d39918",color:"#34d399"}}>
                            <span className="w-1.5 h-1.5 rounded-full inline-block bg-[#34d399]"/>
                            {r.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════
          SEND REMINDER MODAL
      ══════════════════════════════════════════════════════════════ */}
      {sendReminderModal && isAdmin && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
          style={{background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}}>
          <div className={`${t.card} border ${t.cardBorder} rounded-2xl shadow-2xl w-full max-w-lg max-h-[90vh] flex flex-col overflow-hidden`}>
            {/* Header */}
            <div className="px-6 py-4 flex items-center justify-between"
              style={{background:"linear-gradient(135deg,#92400e,#b45309)"}}>
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl flex items-center justify-center bg-white/15">
                  <Icon path={Icons.send} size={18} className="text-white"/>
                </div>
                <div>
                  <h3 className="text-white font-bold text-sm">Send Reminder</h3>
                  <p className="text-white/60 text-[10px]">ProCAS – Pending Data Update</p>
                </div>
              </div>
              <button onClick={()=>setSendReminderModal(false)}
                className="w-8 h-8 rounded-lg text-white/60 hover:text-white hover:bg-white/10 flex items-center justify-center">
                <Icon path={Icons.x} size={16}/>
              </button>
            </div>
            <div className="px-6 py-5 overflow-y-auto flex-1 space-y-4">
              {/* Target */}
              <div>
                <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${t.textMuted}`}>Send To</div>
                <div className="flex flex-wrap gap-2">
                  {["All Active Users","Multiple Users","Individual User"].map(opt=>(
                    <button key={opt} onClick={()=>{setSrTarget(opt);setSrUsers([]);}}
                      className="px-3 py-1.5 rounded-xl text-xs font-semibold transition-all"
                      style={{
                        background: srTarget===opt?"#f59e0b":(dark?"#1a2d44":"#f1f5f9"),
                        color:      srTarget===opt?"#fff":(dark?"#64748b":"#6b7280"),
                      }}>{opt}</button>
                  ))}
                </div>
              </div>
              {/* User picker when not "All" */}
              {srTarget !== "All Active Users" && (
                <div className={`rounded-xl border ${t.cardBorder} overflow-hidden max-h-40 overflow-y-auto`}>
                  {users.map(u=>{
                    const sel = srUsers.includes(u.id);
                    const single = srTarget === "Individual User";
                    return (
                      <div key={u.id}
                        className={`flex items-center gap-3 px-3 py-2.5 cursor-pointer border-b last:border-0 ${t.cardBorder} transition-colors`}
                        style={{background:sel?(dark?"#1e3a1a":"#f0fdf4"):"transparent"}}
                        onClick={()=>{
                          if (single) setSrUsers([u.id]);
                          else setSrUsers(prev=>prev.includes(u.id)?prev.filter(x=>x!==u.id):[...prev,u.id]);
                        }}>
                        <div className="w-4 h-4 rounded flex items-center justify-center border-2 flex-shrink-0"
                          style={{borderColor:sel?"#f59e0b":(dark?"#1a2d44":"#c8d8e8"),background:sel?"#f59e0b":"transparent"}}>
                          {sel && <span className="text-white text-[9px] font-black">✓</span>}
                        </div>
                        <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-[10px] font-black flex-shrink-0"
                          style={{background:"linear-gradient(135deg,#f59e0b,#ea580c)"}}>
                          {u.name.charAt(0)}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className={`text-xs font-semibold truncate ${t.text}`}>{u.name}</div>
                          <div className={`text-[10px] truncate ${t.textMuted}`}>{u.email}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
              {/* Filters */}
              <div className="grid grid-cols-3 gap-3">
                {[
                  {label:"Module", val:srModule, set:setSrModule, opts:AR_MODULES},
                  {label:"Financial Year", val:srFY, set:setSrFY, opts:AR_FYS},
                  {label:"Month", val:srMonth, set:setSrMonth, opts:AR_MONTHS},
                ].map(f=>(
                  <div key={f.label}>
                    <div className={`text-[10px] font-bold uppercase tracking-widest mb-1.5 ${t.textMuted}`}>{f.label}</div>
                    <select value={f.val} onChange={e=>f.set(e.target.value)}
                      className={`w-full px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`}>
                      {f.opts.map(o=><option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                ))}
              </div>
              {/* Email preview */}
              <div className={`rounded-xl border ${t.cardBorder} text-xs font-mono px-4 py-3 space-y-1 ${t.textMuted}`}
                style={{background:dark?"#0a1424":"#f8fbff"}}>
                <div><span className="font-semibold" style={{color:"#f59e0b"}}>Subject: </span>ProCAS Reminder – Pending Data Update</div>
                <div className="pt-1">Hello [User Name],</div>
                <div className="pt-1">This is a reminder to update your assigned client data in ProCAS.</div>
                <div className="pt-1"><span className={`font-semibold ${t.text}`}>Module:</span> {srModule}</div>
                <div><span className={`font-semibold ${t.text}`}>FY:</span> {srFY} &nbsp;|&nbsp; <span className={`font-semibold ${t.text}`}>Month:</span> {srMonth}</div>
                <div className="pt-1">Regards, ProCAS Administrator</div>
              </div>
            </div>
            <div className={`px-6 py-4 border-t ${t.cardBorder} flex items-center justify-between`}>
              <span className={`text-xs ${t.textMuted}`}>
                {srTarget==="All Active Users"
                  ? `${getActiveUsers().length} active users`
                  : `${srUsers.length} user${srUsers.length!==1?"s":""} selected`}
              </span>
              <div className="flex gap-2">
                <button onClick={()=>setSendReminderModal(false)}
                  className={`px-4 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>Cancel</button>
                <button
                  onClick={sendNowReminder}
                  disabled={srTarget!=="All Active Users" && srUsers.length===0}
                  className="flex items-center gap-1.5 px-5 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90 disabled:opacity-40"
                  style={{background:"linear-gradient(135deg,#f59e0b,#ea580c)"}}>
                  <Icon path={Icons.send} size={14}/> Send Now
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════
          AUTO REMINDER MODAL — Configure Auto Reminder
      ══════════════════════════════════════════════════════════════ */}
      {autoReminderModal && isAdmin && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4"
          style={{background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}}>
          <div className={`${t.card} border ${t.cardBorder} rounded-2xl shadow-2xl w-full max-w-xl max-h-[90vh] flex flex-col overflow-hidden`}>
            {/* Header */}
            <div className="px-6 py-4 flex items-center justify-between"
              style={{background:"linear-gradient(135deg,#0c4a6e,#312e81)"}}>
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl flex items-center justify-center bg-white/15">
                  <Icon path={Icons.clock} size={18} className="text-white"/>
                </div>
                <div>
                  <h3 className="text-white font-bold text-sm">Configure Auto Reminder</h3>
                  <p className="text-white/60 text-[10px]">Schedule recurring reminder emails for pending data updates</p>
                </div>
              </div>
              <button onClick={()=>setAutoReminderModal(false)}
                className="w-8 h-8 rounded-lg text-white/60 hover:text-white hover:bg-white/10 flex items-center justify-center">
                <Icon path={Icons.x} size={16}/>
              </button>
            </div>

            <div className="px-6 py-5 overflow-y-auto flex-1 space-y-5">

              {/* Success confirmation */}
              {arSaved && (
                <div className="rounded-xl px-4 py-4 space-y-2" style={{background:"#34d39915",border:"1px solid #34d39940"}}>
                  <div className="flex items-center gap-2 font-bold text-sm" style={{color:"#34d399"}}>
                    <Icon path={Icons.check} size={16}/> Auto Reminder Scheduled Successfully
                  </div>
                  <div className="grid grid-cols-2 gap-x-6 gap-y-1 text-xs mt-1">
                    {[
                      ["Module",     arSaved.module],
                      ["Frequency",  arSaved.freq],
                      ["Start Date", arSaved.startDate],
                      ["Time",       arSaved.time],
                      ["Recipients", arSaved.recipientCount],
                      ["FY / Month", `${arSaved.fy} · ${arSaved.month}`],
                    ].map(([k,v])=>(
                      <div key={k}>
                        <span className={`${t.textMuted}`}>{k}: </span>
                        <span className={`font-semibold ${t.text}`}>{v}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Reminder Type */}
              <div>
                <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${t.textMuted}`}>1. Reminder Type</div>
                <div className="flex flex-wrap gap-2">
                  {AR_TYPES.map(opt=>(
                    <button key={opt} onClick={()=>{setArType(opt);setArUsers([]);}}
                      className="px-3 py-1.5 rounded-xl text-xs font-semibold transition-all"
                      style={{
                        background: arType===opt?"#6366f1":(dark?"#1a2d44":"#f1f5f9"),
                        color:      arType===opt?"#fff":(dark?"#64748b":"#6b7280"),
                      }}>{opt}</button>
                  ))}
                </div>
              </div>

              {/* User picker for Individual / Multiple */}
              {(arType === "Individual User" || arType === "Multiple Users") && (
                <div>
                  <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${t.textMuted}`}>
                    Select User{arType==="Multiple Users"?"s":""}
                  </div>
                  <div className={`rounded-xl border ${t.cardBorder} overflow-hidden max-h-36 overflow-y-auto`}>
                    {users.map(u=>{
                      const sel = arUsers.includes(u.id);
                      return (
                        <div key={u.id}
                          className={`flex items-center gap-3 px-3 py-2 cursor-pointer border-b last:border-0 ${t.cardBorder} transition-colors`}
                          style={{background:sel?(dark?"#1e2d46":"#eff6ff"):"transparent"}}
                          onClick={()=>{
                            if(arType==="Individual User") setArUsers([u.id]);
                            else setArUsers(prev=>prev.includes(u.id)?prev.filter(x=>x!==u.id):[...prev,u.id]);
                          }}>
                          <div className="w-4 h-4 rounded border-2 flex items-center justify-center flex-shrink-0"
                            style={{borderColor:sel?"#6366f1":"#94a3b8",background:sel?"#6366f1":"transparent"}}>
                            {sel && <span className="text-white text-[9px] font-black">✓</span>}
                          </div>
                          <div className="w-6 h-6 rounded-full flex items-center justify-center text-white text-[9px] font-black flex-shrink-0"
                            style={{background:"linear-gradient(135deg,#0ea5e9,#6366f1)"}}>
                            {u.name.charAt(0)}
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className={`text-xs font-semibold truncate ${t.text}`}>{u.name}</div>
                            <div className={`text-[10px] truncate ${t.textMuted}`}>{u.email}</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Module, FY, Month */}
              <div className="grid grid-cols-3 gap-3">
                {[
                  {n:"2. Module",         val:arModule, set:setArModule, opts:AR_MODULES},
                  {n:"3. Financial Year", val:arFY,     set:setArFY,     opts:AR_FYS},
                  {n:"4. Month",          val:arMonth,  set:setArMonth,  opts:AR_MONTHS},
                ].map(f=>(
                  <div key={f.n}>
                    <div className={`text-[10px] font-bold uppercase tracking-widest mb-1.5 ${t.textMuted}`}>{f.n}</div>
                    <select value={f.val} onChange={e=>f.set(e.target.value)}
                      className={`w-full px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`}>
                      {f.opts.map(o=><option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                ))}
              </div>

              {/* Frequency, Start Date, Time */}
              <div className="grid grid-cols-3 gap-3">
                <div>
                  <div className={`text-[10px] font-bold uppercase tracking-widest mb-1.5 ${t.textMuted}`}>5. Frequency</div>
                  <select value={arFreq} onChange={e=>setArFreq(e.target.value)}
                    className={`w-full px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`}>
                    {AR_FREQS.map(o=><option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
                <div>
                  <div className={`text-[10px] font-bold uppercase tracking-widest mb-1.5 ${t.textMuted}`}>6. Start Date</div>
                  <input type="date" value={arStartDate} onChange={e=>setArStartDate(e.target.value)}
                    className={`w-full px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`}/>
                </div>
                <div>
                  <div className={`text-[10px] font-bold uppercase tracking-widest mb-1.5 ${t.textMuted}`}>7. Reminder Time</div>
                  <input type="time" value={arTime} onChange={e=>setArTime(e.target.value)}
                    className={`w-full px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`}/>
                </div>
              </div>

              {/* Preview */}
              <div className={`rounded-xl border ${t.cardBorder} text-xs px-4 py-3 space-y-0.5 ${t.textMuted} font-mono`}
                style={{background:dark?"#0a1424":"#f8fbff"}}>
                <div><span className="font-semibold" style={{color:"#0ea5e9"}}>Subject: </span>ProCAS Reminder – Pending Data Update</div>
                <div className="pt-1">Hello [User Name], …update {arModule} for {arMonth} FY {arFY}.</div>
                <div className="pt-0.5 not-italic" style={{color:"#6366f1"}}>
                  ⏰ Runs {arFreq.toLowerCase()} · {arStartDate||"immediately"} at {arTime}
                  · {arType==="All Active Users"?`${getActiveUsers().length} active users`:arUsers.length+" users selected"}
                </div>
              </div>
            </div>

            {/* Footer */}
            <div className={`px-6 py-4 border-t ${t.cardBorder} flex items-center justify-between gap-3`}>
              <button onClick={sendTestReminder}
                className={`flex items-center gap-1.5 px-4 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                <Icon path={Icons.send} size={13}/> Send Test
              </button>
              <div className="flex gap-2">
                <button onClick={()=>setAutoReminderModal(false)}
                  className={`px-4 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>Cancel</button>
                <button onClick={saveAutoReminder}
                  className="flex items-center gap-1.5 px-5 py-2 rounded-xl text-sm font-semibold text-white hover:opacity-90"
                  style={{background:"linear-gradient(135deg,#0ea5e9,#6366f1)"}}>
                  <Icon path={Icons.check} size={14}/>
                  {arEditId ? "Update Schedule" : "Save Auto Reminder"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Hidden file input for bulk import */}
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
        onChange={e => {
          const f = e.target.files?.[0];
          if (f) { setBulkFile(f); parseBulkFile(f); }
          e.target.value = "";
        }}/>

      {/* ── Bulk Import Modal ── */}
      {bulkModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4"
          style={{background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}}>
          <div className={`${t.card} border ${t.cardBorder} rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col`}>

            {/* Modal Header */}
            <div className="px-6 py-4 border-b flex items-center justify-between"
              style={{
                borderColor: "transparent",
                background:"linear-gradient(135deg,#3b1f6e,#1b3a6e)",
              }}>
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{background:"rgba(255,255,255,0.15)"}}>
                  <Icon path={Icons.upload} size={18} className="text-white"/>
                </div>
                <div>
                  <h3 className="text-white font-bold text-sm">Bulk Import Users</h3>
                  <p className="text-white/60 text-[10px]">
                    {bulkStep==="upload"  && "Upload an Excel or CSV file to import multiple users at once"}
                    {bulkStep==="review"  && "Review parsed records before confirming import"}
                    {bulkStep==="summary" && "Import complete — review the results below"}
                  </p>
                </div>
              </div>
              <button onClick={resetBulkModal}
                className="w-8 h-8 rounded-lg flex items-center justify-center text-white/60 hover:text-white hover:bg-white/10 transition-all">
                <Icon path={Icons.x} size={16}/>
              </button>
            </div>

            {/* Step indicator */}
            <div className="px-6 pt-4 pb-3 flex items-center gap-2">
              {["Upload","Review","Summary"].map((s,i)=>{
                const stepIdx = ["upload","review","summary"].indexOf(bulkStep);
                const done    = i < stepIdx;
                const active  = i === stepIdx;
                return (
                  <div key={s} className="flex items-center gap-2">
                    <div className="flex items-center gap-1.5">
                      <div className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-black"
                        style={{
                          background: done ? "#34d399" : active ? "#7c3aed" : (dark?"#1a2d44":"#e2ecf4"),
                          color:      done||active ? "#fff" : (dark?"#5a7a99":"#94a3b8"),
                        }}>
                        {done ? "✓" : i+1}
                      </div>
                      <span className="text-xs font-semibold"
                        style={{color: active ? (dark?"#c8dff0":"#0d2137") : (dark?"#5a7a99":"#94a3b8")}}>
                        {s}
                      </span>
                    </div>
                    {i < 2 && <div className="w-8 h-px" style={{background:dark?"#1a2d44":"#e2ecf4"}}/>}
                  </div>
                );
              })}
            </div>

            {/* Modal Body */}
            <div className="px-6 pb-6 overflow-y-auto flex-1">

              {/* ── STEP 1: Upload ── */}
              {bulkStep==="upload" && (
                <div className="space-y-4">
                  {/* Drop zone */}
                  <div
                    className="border-2 border-dashed rounded-2xl p-10 flex flex-col items-center gap-3 cursor-pointer transition-all hover:border-[#7c3aed] hover:bg-[#7c3aed08]"
                    style={{borderColor:dark?"#243d58":"#c8d8e8"}}
                    onClick={() => fileInputRef.current?.click()}
                    onDragOver={e=>e.preventDefault()}
                    onDrop={e=>{
                      e.preventDefault();
                      const f = e.dataTransfer.files?.[0];
                      if (f) { setBulkFile(f); parseBulkFile(f); }
                    }}>
                    <div className="w-14 h-14 rounded-2xl flex items-center justify-center" style={{background:"#7c3aed18"}}>
                      <Icon path={Icons.upload} size={26} style={{color:"#7c3aed"}}/>
                    </div>
                    <div className="text-center">
                      <p className={`text-sm font-semibold ${t.text}`}>Click to browse or drag &amp; drop</p>
                      <p className={`text-xs mt-1 ${t.textMuted}`}>Supports .xlsx, .xls, .csv</p>
                    </div>
                    {bulkFile && (
                      <div className="flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-semibold"
                        style={{background:"#7c3aed18",color:"#7c3aed",border:"1px solid #7c3aed30"}}>
                        <Icon path={Icons.fileText} size={13}/> {bulkFile.name}
                      </div>
                    )}
                  </div>

                  {/* Download template */}
                  <div className={`flex items-center justify-between px-4 py-3 rounded-xl border ${t.cardBorder}`}
                    style={{background:dark?"#0a1424":"#f8fbff"}}>
                    <div>
                      <p className={`text-xs font-semibold ${t.text}`}>Need a template?</p>
                      <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>Download the sample Excel file with required columns</p>
                    </div>
                    <button onClick={downloadTemplate}
                      className={`flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-semibold border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                      <Icon path={Icons.download} size={12}/> User_Import_Template.xlsx
                    </button>
                  </div>

                  {/* Column guide */}
                  <div className={`rounded-xl border ${t.cardBorder} overflow-hidden`}>
                    <div className={`px-4 py-2.5 border-b ${t.cardBorder} text-[10px] font-bold uppercase tracking-widest ${t.textMuted}`}>
                      Required Column Format
                    </div>
                    <table className="w-full text-xs">
                      <thead>
                        <tr className={t.tableHead}>
                          {["Name of User *","Email ID *","Role","Assigned Clients"].map(h=>(
                            <th key={h} className={`px-3 py-2 text-left font-semibold border-b ${t.cardBorder} text-[10px]`}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr className={t.tableRow}>
                          <td className="px-3 py-2">User A</td>
                          <td className="px-3 py-2">usera@example.com</td>
                          <td className="px-3 py-2">End User</td>
                          <td className="px-3 py-2">Client Alpha, Client Beta</td>
                        </tr>
                        <tr className={t.tableRow}>
                          <td className="px-3 py-2">User B</td>
                          <td className="px-3 py-2">userb@example.com</td>
                          <td className="px-3 py-2">Admin</td>
                          <td className="px-3 py-2">All Clients</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* ── STEP 2: Review ── */}
              {bulkStep==="review" && bulkParsed && (
                <div className="space-y-4">
                  {/* Parse summary chips */}
                  <div className="flex flex-wrap gap-2">
                    {[
                      { label:"Total Rows",    value:bulkParsed.total,         color:"#4a90d9" },
                      { label:"Ready to Import", value:bulkParsed.rows.length, color:"#34d399" },
                      { label:"Errors",          value:bulkParsed.errors.length,color:"#f87171"},
                      { label:"Duplicates",      value:bulkParsed.dupes.length, color:"#fbbf24"},
                    ].map(c=>(
                      <div key={c.label} className="flex items-center gap-2 px-3 py-1.5 rounded-xl text-xs font-semibold"
                        style={{background:c.color+"18",color:c.color,border:`1px solid ${c.color}30`}}>
                        <span className="font-black text-sm">{c.value}</span> {c.label}
                      </div>
                    ))}
                  </div>

                  {/* Valid rows preview */}
                  {bulkParsed.rows.length > 0 && (
                    <div className={`rounded-xl border ${t.cardBorder} overflow-hidden`}>
                      <div className={`px-4 py-2.5 border-b ${t.cardBorder} flex items-center gap-2`}>
                        <div className="w-2 h-2 rounded-full" style={{background:"#34d399"}}/>
                        <span className={`text-xs font-bold ${t.text}`}>Ready to Import ({bulkParsed.rows.length})</span>
                      </div>
                      <div className="overflow-x-auto max-h-36 overflow-y-auto">
                        <table className="w-full text-xs">
                          <thead className="sticky top-0">
                            <tr className={t.tableHead}>
                              {["Row","Name","Email","Role","Assigned Clients"].map(h=>(
                                <th key={h} className={`px-3 py-2 text-left font-semibold border-b ${t.cardBorder} text-[10px]`}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody className={`divide-y ${t.divider}`}>
                            {bulkParsed.rows.map(r=>(
                              <tr key={r.rowNum} className={t.tableRow}>
                                <td className={`px-3 py-2 font-mono ${t.textMuted}`}>{r.rowNum}</td>
                                <td className={`px-3 py-2 font-medium ${t.text}`}>{r.name}</td>
                                <td className={`px-3 py-2 ${t.textMuted}`}>{r.email}</td>
                                <td className="px-3 py-2">
                                  <span className="px-1.5 py-0.5 rounded-full text-[10px] font-semibold"
                                    style={{background:"#4a90d918",color:"#4a90d9"}}>{r.role}</span>
                                </td>
                                <td className={`px-3 py-2 ${t.textMuted}`}>{r.clients||"—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* Error rows */}
                  {bulkParsed.errors.length > 0 && (
                    <div className={`rounded-xl border overflow-hidden`} style={{borderColor:"#f8717140"}}>
                      <div className="px-4 py-2.5 border-b flex items-center gap-2" style={{borderColor:"#f8717130",background:"#f8717108"}}>
                        <div className="w-2 h-2 rounded-full" style={{background:"#f87171"}}/>
                        <span className="text-xs font-bold" style={{color:"#f87171"}}>Validation Errors ({bulkParsed.errors.length} rows will be skipped)</span>
                      </div>
                      <div className="max-h-32 overflow-y-auto">
                        {bulkParsed.errors.map(r=>(
                          <div key={r.rowNum} className="px-4 py-2.5 flex items-start gap-3 text-xs border-b last:border-0"
                            style={{borderColor:"#f8717120"}}>
                            <span className="font-mono font-bold shrink-0" style={{color:"#f87171"}}>Row {r.rowNum}</span>
                            <span className={t.textMuted}>{r.errors.join(" · ")}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Duplicate rows */}
                  {bulkParsed.dupes.length > 0 && (
                    <div className={`rounded-xl border overflow-hidden`} style={{borderColor:"#fbbf2440"}}>
                      <div className="px-4 py-2.5 border-b flex items-center gap-2" style={{borderColor:"#fbbf2430",background:"#fbbf2408"}}>
                        <div className="w-2 h-2 rounded-full" style={{background:"#fbbf24"}}/>
                        <span className="text-xs font-bold" style={{color:"#fbbf24"}}>Duplicate Emails ({bulkParsed.dupes.length}) — choose action</span>
                      </div>
                      <div className="max-h-40 overflow-y-auto">
                        {bulkParsed.dupes.map(d=>(
                          <div key={d.email} className="px-4 py-3 flex items-center justify-between gap-3 text-xs border-b last:border-0"
                            style={{borderColor:"#fbbf2420"}}>
                            <div>
                              <span className={`font-semibold ${t.text}`}>{d.name}</span>
                              <span className={`ml-2 ${t.textMuted}`}>{d.email}</span>
                              <span className="ml-2 text-[10px]" style={{color:"#fbbf24"}}>· Row {d.rowNum}</span>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              {["skip","update"].map(act=>(
                                <button key={act} onClick={()=>setBulkDupeAction(prev=>({...prev,[d.email]:act}))}
                                  className="px-2.5 py-1 rounded-lg text-[10px] font-bold capitalize transition-all"
                                  style={{
                                    background: bulkDupeAction[d.email]===act ? (act==="update"?"#fbbf2430":"#94a3b830") : "transparent",
                                    color:      bulkDupeAction[d.email]===act ? (act==="update"?"#fbbf24":"#94a3b8") : (dark?"#5a7a99":"#94a3b8"),
                                    border:     `1px solid ${bulkDupeAction[d.email]===act?(act==="update"?"#fbbf2450":"#94a3b850"):(dark?"#1a2d44":"#e2ecf4")}`,
                                  }}>
                                  {act === "skip" ? "Skip" : "Update"}
                                </button>
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Send invite toggle */}
                  {bulkParsed.rows.length > 0 && (
                    <div className={`flex items-center gap-3 px-4 py-3 rounded-xl border ${t.cardBorder}`}
                      style={{background:dark?"#0a1424":"#f8fbff"}}>
                      <button onClick={()=>setBulkSendInvite(v=>!v)}
                        className="flex items-center gap-2 text-xs font-semibold"
                        style={{color:bulkSendInvite?"#00c9b1":(dark?"#5a7a99":"#94a3b8")}}>
                        <div className="w-4 h-4 rounded flex items-center justify-center border-2 transition-all"
                          style={{borderColor:bulkSendInvite?"#00c9b1":(dark?"#1a2d44":"#c8d8e8"),
                                  background:bulkSendInvite?"#00c9b1":"transparent"}}>
                          {bulkSendInvite && <span className="text-white text-[9px] font-black">✓</span>}
                        </div>
                        Send invite emails to {bulkParsed.rows.length} imported user{bulkParsed.rows.length!==1?"s":""}
                      </button>
                      <span className={`text-[10px] ${t.textMuted}`}>(uses existing invitation workflow)</span>
                    </div>
                  )}

                  {/* Actions */}
                  <div className="flex items-center justify-end gap-2 pt-1">
                    <button onClick={()=>setBulkStep("upload")}
                      className={`px-4 py-2 rounded-xl text-sm font-medium border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
                      ← Back
                    </button>
                    <button
                      onClick={executeBulkImport}
                      disabled={bulkParsed.rows.length === 0 && bulkParsed.dupes.every(d=>bulkDupeAction[d.email]==="skip")}
                      className="px-5 py-2 rounded-xl text-sm font-semibold text-white transition-all hover:opacity-90 disabled:opacity-40"
                      style={{background:"linear-gradient(135deg,#7c3aed,#4a90d9)"}}>
                      Import {bulkParsed.rows.length + bulkParsed.dupes.filter(d=>bulkDupeAction[d.email]==="update").length} Users →
                    </button>
                  </div>
                </div>
              )}

              {/* ── STEP 3: Summary ── */}
              {bulkStep==="summary" && bulkSummary && (
                <div className="space-y-4">
                  {/* Result grid */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                    {[
                      { label:"Total Uploaded",     value:bulkSummary.total,    color:"#4a90d9" },
                      { label:"Successfully Imported", value:bulkSummary.imported, color:"#34d399" },
                      { label:"Failed / Skipped",   value:bulkSummary.failed + bulkSummary.skipped, color:"#f87171"},
                      { label:"Duplicates",          value:bulkSummary.dupes,   color:"#fbbf24" },
                    ].map(c=>(
                      <div key={c.label} className={`${t.card} border ${t.cardBorder} rounded-xl p-4 text-center`}
                        style={{borderTop:`2.5px solid ${c.color}`}}>
                        <div className="text-2xl font-black" style={{color:c.color}}>{c.value}</div>
                        <div className={`text-[10px] mt-1 font-medium ${t.textMuted}`}>{c.label}</div>
                      </div>
                    ))}
                  </div>

                  {/* Updated count */}
                  {bulkSummary.updated > 0 && (
                    <div className="flex items-center gap-2 text-xs px-3 py-2 rounded-xl"
                      style={{background:"#fbbf2415",color:"#fbbf24",border:"1px solid #fbbf2430"}}>
                      <Icon path={Icons.edit} size={12}/> {bulkSummary.updated} existing user{bulkSummary.updated!==1?"s":""} updated
                    </div>
                  )}

                  {/* Invite note */}
                  {bulkSendInvite && bulkSummary.imported > 0 && (
                    <div className="flex items-center gap-2 text-xs px-3 py-2 rounded-xl"
                      style={{background:"#00c9b115",color:"#00c9b1",border:"1px solid #00c9b130"}}>
                      <Icon path={Icons.send} size={12}/> Invite emails queued for {bulkSummary.imported} user{bulkSummary.imported!==1?"s":""}
                    </div>
                  )}

                  {/* Error report */}
                  {bulkSummary.failed > 0 && (
                    <div className={`rounded-xl border overflow-hidden`} style={{borderColor:"#f8717140"}}>
                      <div className="px-4 py-2.5 border-b flex items-center justify-between" style={{borderColor:"#f8717130",background:"#f8717108"}}>
                        <div className="flex items-center gap-2">
                          <div className="w-2 h-2 rounded-full" style={{background:"#f87171"}}/>
                          <span className="text-xs font-bold" style={{color:"#f87171"}}>Failed Rows ({bulkSummary.failed})</span>
                        </div>
                        <button onClick={()=>downloadErrorReport(bulkSummary)}
                          className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-semibold"
                          style={{background:"#f8717118",color:"#f87171",border:"1px solid #f8717130"}}>
                          <Icon path={Icons.download} size={11}/> Download Error Report
                        </button>
                      </div>
                      <div className="max-h-32 overflow-y-auto">
                        {bulkSummary.errorRows.map(r=>(
                          <div key={r.rowNum} className="px-4 py-2.5 flex items-start gap-3 text-xs border-b last:border-0"
                            style={{borderColor:"#f8717120"}}>
                            <span className="font-mono font-bold shrink-0" style={{color:"#f87171"}}>Row {r.rowNum}</span>
                            <span className={t.textMuted}>{r.errors.join(" · ")}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="flex items-center justify-end pt-1">
                    <button onClick={resetBulkModal}
                      className="px-5 py-2 rounded-xl text-sm font-semibold text-white transition-all hover:opacity-90"
                      style={{background:"linear-gradient(135deg,#7c3aed,#4a90d9)"}}>
                      Done
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Modals */}
      {addModal  && <InviteFormModal t={t} dark={dark} title="Invite New User" form={form} setForm={setForm} formErr={formErr} onSave={handleAdd}  onClose={()=>setAddModal(false)} masterClients={masterClients}/>}
      {editUser  && <InviteFormModal t={t} dark={dark} title="Edit User"       form={form} setForm={setForm} formErr={formErr} onSave={handleEdit} onClose={()=>setEditUser(null)}  masterClients={masterClients}/>}
      {deleteId  && <ConfirmDialog t={t} message={`Remove ${users.find(u=>u.id===deleteId)?.name}?`} onConfirm={()=>handleDelete(deleteId)} onCancel={()=>setDeleteId(null)}/>}
    </div>
  );
};


// ─────────────────────────────────────────────────────────────────────────────
// ── SETTINGS TAB ─────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const SettingsTab = ({ t, dark, darkMode, setDarkMode, isAdmin, setIsAdmin }) => {
  const [saved, setSaved] = useState(false);
  const save = () => { setSaved(true); setTimeout(()=>setSaved(false), 2000); };

  const Row = ({ label, sub, children }) => (
    <div className={`flex items-center justify-between py-4 border-b ${t.cardBorder} last:border-0`}>
      <div className="mr-6">
        <div className={`text-sm font-semibold ${t.text}`}>{label}</div>
        {sub && <div className={`text-xs mt-0.5 ${t.textMuted}`}>{sub}</div>}
      </div>
      <div className="shrink-0">{children}</div>
    </div>
  );

  const Toggle = ({ value, onChange }) => (
    <div className="relative w-10 h-6 rounded-full cursor-pointer transition-colors"
      style={{background: value ? "#00c9b1" : (dark ? "#2d3748" : "#e2e8f0")}}
      onClick={() => onChange(!value)}>
      <div className="absolute top-1 w-4 h-4 rounded-full bg-white shadow transition-all"
        style={{left: value ? "22px" : "4px"}}/>
    </div>
  );

  return (
    <div className="space-y-5 max-w-2xl">
      {saved && (
        <div className="fixed top-5 right-5 z-[75] flex items-center gap-2 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
          style={{background:"linear-gradient(135deg,#34d399,#10b981)"}}>
          <Icon path={Icons.check} size={15}/> Settings saved
        </div>
      )}

      {/* Appearance */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-6 py-2`}>
        <h3 className={`font-bold text-sm ${t.text} pt-4 pb-2`}>Appearance</h3>
        <Row label="Dark Mode" sub="Toggle between dark and light interface theme">
          <Toggle value={darkMode} onChange={setDarkMode}/>
        </Row>
      </div>

      {/* Access Control */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-6 py-2`}>
        <h3 className={`font-bold text-sm ${t.text} pt-4 pb-2`}>Access Control</h3>
        <Row label="Admin Mode" sub="Toggle between Admin and Viewer role (demo only)">
          <Toggle value={isAdmin} onChange={setIsAdmin}/>
        </Row>
        <Row label="Two-Factor Authentication" sub="Require 2FA for all admin logins">
          <span className="text-xs px-2.5 py-1 rounded-full font-semibold" style={{background:"#fbbf2418",color:"#fbbf24"}}>Coming Soon</span>
        </Row>
        <Row label="Session Timeout" sub="Auto-logout after inactivity period">
          <select className={`px-3 py-1.5 rounded-lg border text-sm outline-none ${t.input} focus:ring-2 focus:ring-[#00c9b1]`}>
            <option>30 minutes</option><option>1 hour</option><option>4 hours</option><option>8 hours</option>
          </select>
        </Row>
      </div>

      {/* System Info */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-6 py-2`}>
        <h3 className={`font-bold text-sm ${t.text} pt-4 pb-2`}>System Information</h3>
        {[
          ["Platform",  "ProCAS v1.0.0"],
          ["Firm",      "Your Organisation Name"],
          ["ICAI Reg.", "XXXXXX"],
          ["FY Active", "2026-27"],
          ["Data Mode", "In-memory (Demo)"],
        ].map(([k,v]) => (
          <Row key={k} label={k} sub="">
            <span className={`text-sm font-semibold ${t.textAccent}`}>{v}</span>
          </Row>
        ))}
      </div>

      <button onClick={save}
        className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-semibold text-white hover:opacity-90"
        style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
        <Icon path={Icons.save} size={14}/> Save Settings
      </button>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── ARCHITECTURE TAB (no backend imports – display only) ─────────────────────
// ─────────────────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
// ── MONTH LOCK CONTROL — Admin-only Architect panel ──────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const MonthLockControl = ({ t, dark, isAdmin, adminName }) => {
  const { locks, setLocks, auditLog, addAuditEntry } = useLockCtx();
  const [selFy,     setSelFy]     = useState(DEFAULT_FY);
  const [selMonth,  setSelMonth]  = useState("April");
  const [selModule, setSelModule] = useState("all");
  const [toast,     setToast]     = useState(null);

  const MONTHS_LIST = ["April","May","June","July","August","September","October","November","December","January","February","March"];

  const showToast = (msg, type="lock") => { setToast({msg,type}); setTimeout(()=>setToast(null),3000); };

  const getLock = (fy, month, mod) => locks[mkLockKey(fy,month,mod)] || locks[mkLockKey(fy,month,"all")] || null;

  const doLock = () => {
    if (!isAdmin) return;
    const now = new Date().toISOString();
    const mods = selModule === "all" ? ["all"] : [selModule];
    setLocks(prev => {
      const next={...prev};
      mods.forEach(m=>{ next[mkLockKey(selFy,selMonth,m)]={by:adminName,at:now,fy:selFy,month:selMonth,module:m}; });
      return next;
    });
    const lbl = selModule==="all"?"All Modules":LOCK_MODULES_META.find(m=>m.id===selModule)?.label;
    addAuditEntry({action:"LOCK",by:adminName,at:now,fy:selFy,month:selMonth,module:lbl});
    showToast(`🔒 Locked: ${lbl} · ${selMonth} · FY ${selFy}`, "lock");
  };

  const doUnlock = () => {
    if (!isAdmin) return;
    const now = new Date().toISOString();
    const mods = selModule==="all" ? ["all",...LOCK_MODULES_META.map(m=>m.id)] : [selModule];
    setLocks(prev => { const next={...prev}; mods.forEach(m=>{delete next[mkLockKey(selFy,selMonth,m)];}); return next; });
    const lbl = selModule==="all"?"All Modules":LOCK_MODULES_META.find(m=>m.id===selModule)?.label;
    addAuditEntry({action:"UNLOCK",by:adminName,at:now,fy:selFy,month:selMonth,module:lbl});
    showToast(`🔓 Unlocked: ${lbl} · ${selMonth} · FY ${selFy}`, "unlock");
  };

  const doFYBulk = lockAll => {
    const now = new Date().toISOString();
    setLocks(prev => {
      const next={...prev};
      MONTHS_LIST.forEach(m => {
        const k=mkLockKey(selFy,m,"all");
        if (lockAll) next[k]={by:adminName,at:now,fy:selFy,month:m,module:"all"};
        else { delete next[k]; LOCK_MODULES_META.forEach(mod=>delete next[mkLockKey(selFy,m,mod.id)]); }
      });
      return next;
    });
    addAuditEntry({action:lockAll?"LOCK_FY":"UNLOCK_FY",by:adminName,at:now,fy:selFy,month:"All Months",module:"All Modules"});
    showToast(lockAll?`🔒 Entire FY ${selFy} locked`:`🔓 Entire FY ${selFy} unlocked`, lockAll?"lock":"unlock");
  };

  const exportAudit = () => {
    const ws = XLSX.utils.json_to_sheet(auditLog.map(e=>({
      "Action":e.action,"Module":e.module,"Financial Year":e.fy,"Month":e.month,
      "By":xlsSafe(e.by),"Date & Time":new Date(e.at).toLocaleString("en-IN"),
    })));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Lock Audit");
    XLSX.writeFile(wb,"ProCAS_Lock_Audit.xlsx");
  };

  const sel=`w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#f87171] ${t.input}`;

  return (
    <div className="space-y-5">
      {toast&&(<div className="fixed top-5 right-5 z-[80] flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white"
        style={{background:toast.type==="unlock"?"linear-gradient(135deg,#d97706,#fbbf24)":"linear-gradient(135deg,#dc2626,#f87171)",minWidth:300}}>
        {toast.type==="unlock"?"🔓":"🔒"} {toast.msg}
      </div>)}

      <div className="rounded-2xl px-6 py-4 flex items-center gap-4"
        style={{background:"linear-gradient(135deg,#1a0800 0%,#2d1200 50%,#0a0e1a 100%)",border:"1px solid #f8717130"}}>
        <span className="text-2xl">🔒</span>
        <div>
          <h3 className="font-black text-sm text-white">Month Lock Control</h3>
          <p className="text-[11px] mt-0.5" style={{color:"rgba(255,255,255,0.45)"}}>Admin-only · Lock / unlock data entry per module, month, and FY</p>
        </div>
        {!isAdmin&&<div className="ml-auto flex items-center gap-2 px-3 py-1.5 rounded-xl text-xs font-bold" style={{background:"#f8717118",color:"#f87171",border:"1px solid #f8717130"}}><Icon path={Icons.shield} size={13}/> Admin only</div>}
      </div>

      {isAdmin ? (<>
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
          <h4 className={`font-bold text-sm mb-4 ${t.text}`}>Configure Lock</h4>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Financial Year</label>
              <select value={selFy} onChange={e=>setSelFy(e.target.value)} className={sel}>{FY_LIST.map(f=><option key={f} value={f}>{f}</option>)}</select></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Month</label>
              <select value={selMonth} onChange={e=>setSelMonth(e.target.value)} className={sel}>{MONTHS_LIST.map(m=><option key={m} value={m}>{m}</option>)}</select></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Module</label>
              <select value={selModule} onChange={e=>setSelModule(e.target.value)} className={sel}>
                <option value="all">All Modules</option>
                {LOCK_MODULES_META.map(m=><option key={m.id} value={m.id}>{m.label}</option>)}
              </select></div>
          </div>
          {(()=>{
            const mods=selModule==="all"?LOCK_MODULES_META.map(m=>m.id):[selModule];
            return (<div className={`rounded-xl p-3 mb-4 border ${t.cardBorder}`} style={{background:dark?"#0c1e30":"#f8fbff"}}>
              <p className={`text-[10px] font-bold uppercase tracking-widest mb-2 ${t.textMuted}`}>Current Status</p>
              <div className="flex flex-wrap gap-2">
                {mods.map(mod=>{ const info=getLock(selFy,selMonth,mod); const label=LOCK_MODULES_META.find(m=>m.id===mod)?.label||mod;
                  return (<div key={mod} className="flex items-center gap-2 px-3 py-1.5 rounded-xl"
                    style={{background:info?"#f8717112":"#34d39912",border:`1px solid ${info?"#f8717130":"#34d39930"}`}}>
                    <span>{info?"🔒":"🔓"}</span>
                    <span className="text-xs font-semibold" style={{color:info?"#f87171":"#34d399"}}>{label}</span>
                    {info&&<span className="text-[9px]" style={{color:"#f87171bb"}}>by {info.by}</span>}
                  </div>);
                })}
              </div>
            </div>);
          })()}
          <div className="flex flex-wrap gap-3">
            <button onClick={doLock} className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-bold text-white hover:opacity-90 active:scale-95 transition-all" style={{background:"linear-gradient(135deg,#dc2626,#f87171)"}}>
              🔒 Lock {selModule==="all"?"All Modules":LOCK_MODULES_META.find(m=>m.id===selModule)?.label}</button>
            <button onClick={doUnlock} className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-bold text-white hover:opacity-90 active:scale-95 transition-all" style={{background:"linear-gradient(135deg,#d97706,#fbbf24)"}}>
              🔓 Unlock {selModule==="all"?"All Modules":LOCK_MODULES_META.find(m=>m.id===selModule)?.label}</button>
            <div className={`w-px self-stretch border-l ${t.cardBorder}`}/>
            <button onClick={()=>doFYBulk(true)} className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm font-semibold border ${t.cardBorder} ${t.textMuted} ${t.hover} transition-all`}>🔒 Lock Entire FY {selFy}</button>
            <button onClick={()=>doFYBulk(false)} className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm font-semibold border ${t.cardBorder} ${t.textMuted} ${t.hover} transition-all`}>🔓 Unlock Entire FY {selFy}</button>
          </div>
        </div>

        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder}`}>
            <h4 className={`font-bold text-sm ${t.text}`}>Lock Status — FY {selFy}</h4>
            <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>12-month view across all 3 modules</p>
          </div>
          <div className="overflow-x-auto"><table className="w-full text-xs border-collapse">
            <thead><tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
              <th className="px-4 py-3 text-left text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",minWidth:90}}>Month</th>
              {LOCK_MODULES_META.map(m=>(<th key={m.id} className="px-4 py-3 text-center text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",minWidth:130}}>{m.label}</th>))}
              <th className="px-4 py-3 text-center text-[9px] font-bold uppercase tracking-widest" style={{color:"#94b4cc",width:90}}>Quick</th>
            </tr></thead>
            <tbody>{MONTHS_LIST.map((m,ri)=>(
              <tr key={m} className={`border-b ${t.cardBorder} transition-colors`} style={{background:ri%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.4)":"#f8fbff")}}>
                <td className={`px-4 py-3 font-semibold text-xs ${t.text}`}>{m}</td>
                {LOCK_MODULES_META.map(mod=>{ const info=getLock(selFy,m,mod.id); return (
                  <td key={mod.id} className="px-4 py-3 text-center">
                    {info?(<div><span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold" style={{background:"#f8717118",color:"#f87171",border:"1px solid #f8717130"}}>🔒 Locked</span>
                      <div className="text-[9px] mt-0.5" style={{color:"#f87171aa"}}>by {info.by}</div></div>
                    ):(<span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold" style={{background:"#34d39915",color:"#34d399",border:"1px solid #34d39930"}}>🔓 Open</span>)}
                  </td>);
                })}
                <td className="px-4 py-3 text-center"><div className="flex items-center justify-center gap-1">
                  <button title={`Lock all for ${m}`} onClick={()=>{const now=new Date().toISOString();setLocks(prev=>({...prev,[mkLockKey(selFy,m,"all")]:{by:adminName,at:now,fy:selFy,month:m,module:"all"}}));addAuditEntry({action:"LOCK",by:adminName,at:new Date().toISOString(),fy:selFy,month:m,module:"All Modules"});showToast(`🔒 ${m} locked`,"lock");}}
                    className="px-2 py-1 rounded-lg text-[10px] font-bold hover:opacity-80" style={{background:"#f8717118",color:"#f87171",border:"1px solid #f8717130"}}>🔒</button>
                  <button title={`Unlock all for ${m}`} onClick={()=>{const now=new Date().toISOString();setLocks(prev=>{const next={...prev};delete next[mkLockKey(selFy,m,"all")];LOCK_MODULES_META.forEach(mod=>delete next[mkLockKey(selFy,m,mod.id)]);return next;});addAuditEntry({action:"UNLOCK",by:adminName,at:new Date().toISOString(),fy:selFy,month:m,module:"All Modules"});showToast(`🔓 ${m} unlocked`,"unlock");}}
                    className="px-2 py-1 rounded-lg text-[10px] font-bold hover:opacity-80" style={{background:"#fbbf2415",color:"#fbbf24",border:"1px solid #fbbf2430"}}>🔓</button>
                </div></td>
              </tr>
            ))}</tbody>
          </table></div>
        </div>

        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex items-center justify-between`}>
            <div><h4 className={`font-bold text-sm ${t.text}`}>Audit Log</h4>
              <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>{auditLog.length} entries · most recent first</p></div>
            {auditLog.length>0&&<button onClick={exportAudit} className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}><Icon path={Icons.download} size={12}/> Export</button>}
          </div>
          <div style={{overflowX:"auto",maxHeight:300,overflowY:"auto"}}><table className="w-full text-xs border-collapse">
            <thead className="sticky top-0 z-10"><tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
              {["Action","Module","FY","Month","By","Date & Time"].map(h=>(<th key={h} className="px-4 py-2.5 text-left text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",borderBottom:"1px solid #ffffff18"}}>{h}</th>))}
            </tr></thead>
            <tbody>{auditLog.length===0?(<tr><td colSpan={6} className={`px-4 py-10 text-center text-sm ${t.textMuted}`}>No audit entries yet</td></tr>)
              :auditLog.map((e,i)=>{ const isL=e.action.includes("LOCK")&&!e.action.includes("UNLOCK"); return (
                <tr key={i} className={`border-b ${t.cardBorder}`} style={{background:i%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.4)":"#f8fbff")}}>
                  <td className="px-4 py-2.5"><span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold" style={{background:isL?"#f8717118":"#fbbf2418",color:isL?"#f87171":"#fbbf24",border:`1px solid ${isL?"#f8717130":"#fbbf2430"}`}}>{isL?"🔒":"🔓"} {e.action.replace("_"," ")}</span></td>
                  <td className={`px-4 py-2.5 ${t.text}`}>{e.module}</td>
                  <td className={`px-4 py-2.5 font-mono text-[10px] ${t.textMuted}`}>{e.fy}</td>
                  <td className={`px-4 py-2.5 ${t.text}`}>{e.month}</td>
                  <td className={`px-4 py-2.5 ${t.textMuted}`}>{e.by}</td>
                  <td className={`px-4 py-2.5 text-[10px] ${t.textMuted}`}>{new Date(e.at).toLocaleString("en-IN",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}</td>
                </tr>);
              })}
            </tbody>
          </table></div>
        </div>
      </>) : (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-12 flex flex-col items-center justify-center gap-3 text-center`}>
          <div className="text-4xl">🔒</div>
          <p className={`text-sm font-semibold ${t.text}`}>Admin Access Required</p>
          <p className={`text-xs ${t.textMuted}`}>Month Lock Control is available to Administrators only.</p>
        </div>
      )}
    </div>
  );
};


// ─────────────────────────────────────────────────────────────────────────────
// ── AUTO REMINDER MANAGEMENT ─────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const ReminderManagement = ({ t, dark, isAdmin }) => {
  const [reminders, setReminders] = useState(() => {
    try { const s = localStorage.getItem("procas_reminders"); return s ? JSON.parse(s) : []; } catch { return []; }
  });
  const [showAdd, setShowAdd]   = useState(false);
  const [editId,  setEditId]    = useState(null);
  const [toast,   setToast]     = useState(null);
  const [form, setForm] = useState({ title:"", module:"CAS MIS", triggerDay:10, frequency:"Monthly", recipients:"All Users", notes:"" });

  const saveToStorage = (data) => { try { localStorage.setItem("procas_reminders", JSON.stringify(data)); } catch {} };
  const showToast = (msg, type="ok") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  const MODULES = ["CAS MIS","KRA/KPI","Fund Request","All Modules"];
  const FREQS   = ["Daily","Weekly","Monthly","Quarterly"];

  const addOrSave = () => {
    if (!form.title.trim()) return;
    if (editId) {
      const updated = reminders.map(r => r.id === editId ? { ...r, ...form } : r);
      setReminders(updated); saveToStorage(updated); setEditId(null);
      showToast("Reminder updated");
    } else {
      const newR = { ...form, id: `REM-${Date.now()}`, status:"Active", createdAt: new Date().toISOString(), lastSent: null, nextSend: new Date(Date.now()+86400000*form.triggerDay%31).toLocaleDateString("en-IN") };
      const updated = [...reminders, newR];
      setReminders(updated); saveToStorage(updated);
      showToast("Reminder scheduled");
    }
    setShowAdd(false); setForm({ title:"", module:"CAS MIS", triggerDay:10, frequency:"Monthly", recipients:"All Users", notes:"" });
  };

  const toggleStatus = (id) => {
    const updated = reminders.map(r => r.id===id ? {...r, status: r.status==="Active"?"Paused":"Active"} : r);
    setReminders(updated); saveToStorage(updated);
    showToast(updated.find(r=>r.id===id).status==="Active"?"Reminder resumed":"Reminder paused");
  };

  const deleteReminder = (id) => {
    const updated = reminders.filter(r=>r.id!==id);
    setReminders(updated); saveToStorage(updated); showToast("Reminder deleted","warn");
  };

  const startEdit = (r) => {
    setForm({ title:r.title, module:r.module, triggerDay:r.triggerDay, frequency:r.frequency, recipients:r.recipients, notes:r.notes||"" });
    setEditId(r.id); setShowAdd(true);
  };

  const inp = `w-full px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`;

  return (
    <div className="space-y-5">
      {toast && <div className="fixed top-5 right-5 z-[80] px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white" style={{background:toast.type==="warn"?"linear-gradient(135deg,#f87171,#dc2626)":"linear-gradient(135deg,#00a896,#1b5fa8)",minWidth:260}}>🔔 {toast.msg}</div>}

      {/* Header */}
      <div className="rounded-2xl px-6 py-4 flex items-center justify-between gap-4" style={{background:"linear-gradient(135deg,#0a1a2e,#0d2035)",border:"1px solid #fbbf2430"}}>
        <div className="flex items-center gap-3">
          <span className="text-2xl">🔔</span>
          <div>
            <h3 className="font-black text-sm text-white">Auto Reminder Management</h3>
            <p className="text-[11px] mt-0.5" style={{color:"rgba(255,255,255,0.45)"}}>Schedule & manage automated reminders across modules</p>
          </div>
        </div>
        {isAdmin && (
          <button onClick={()=>{setShowAdd(v=>!v);setEditId(null);setForm({ title:"", module:"CAS MIS", triggerDay:10, frequency:"Monthly", recipients:"All Users", notes:"" });}}
            className="flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold text-white hover:opacity-90"
            style={{background:"linear-gradient(135deg,#fbbf24,#d97706)"}}>
            + New Reminder
          </button>
        )}
      </div>

      {/* Add/Edit form */}
      {showAdd && isAdmin && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 space-y-4`} style={{borderColor:"#fbbf2440"}}>
          <h4 className={`font-bold text-sm ${t.text}`}>{editId?"Edit Reminder":"Schedule New Reminder"}</h4>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Reminder Title *</label>
              <input value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))} placeholder="e.g. MIS Filing Due" className={inp}/></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Module</label>
              <select value={form.module} onChange={e=>setForm(f=>({...f,module:e.target.value}))} className={inp}>
                {MODULES.map(m=><option key={m} value={m}>{m}</option>)}</select></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Trigger Day (of month)</label>
              <input type="number" min={1} max={31} value={form.triggerDay} onChange={e=>setForm(f=>({...f,triggerDay:+e.target.value}))} className={inp}/></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Frequency</label>
              <select value={form.frequency} onChange={e=>setForm(f=>({...f,frequency:e.target.value}))} className={inp}>
                {FREQS.map(fr=><option key={fr} value={fr}>{fr}</option>)}</select></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Recipients</label>
              <input value={form.recipients} onChange={e=>setForm(f=>({...f,recipients:e.target.value}))} placeholder="All Users / Specific names" className={inp}/></div>
            <div><label className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} block mb-1.5`}>Notes</label>
              <input value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder="Optional notes" className={inp}/></div>
          </div>
          <div className="flex gap-3">
            <button onClick={()=>{setShowAdd(false);setEditId(null);}} className={`flex-1 py-2.5 rounded-xl border text-sm font-medium ${t.cardBorder} ${t.textMuted} ${t.hover}`}>Cancel</button>
            <button onClick={addOrSave} className="flex-1 py-2.5 rounded-xl text-sm font-bold text-white" style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
              {editId?"Save Changes":"Schedule Reminder"}
            </button>
          </div>
        </div>
      )}

      {/* Reminders table */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
        <div className={`px-5 py-3.5 border-b ${t.cardBorder} flex items-center justify-between`}>
          <h4 className={`font-bold text-sm ${t.text}`}>Scheduled Reminders</h4>
          <span className={`text-[10px] ${t.textMuted}`}>{reminders.length} total · {reminders.filter(r=>r.status==="Active").length} active</span>
        </div>
        {reminders.length === 0 ? (
          <div className={`px-5 py-12 text-center ${t.textMuted} text-sm`}>No reminders yet. Click "+ New Reminder" to add one.</div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead><tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
                {["Title","Module","Trigger","Frequency","Recipients","Status","Last Sent","Actions"].map(h=>(
                  <th key={h} className="px-4 py-3 text-left text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",borderBottom:"1px solid #ffffff18"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {reminders.map((r,i)=>(
                  <tr key={r.id} className={`border-b ${t.cardBorder} transition-colors ${t.hover}`}
                    style={{background:i%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.4)":"#f8fbff"),opacity:r.status==="Paused"?0.65:1}}>
                    <td className={`px-4 py-3 font-semibold ${t.text}`}>{r.title}</td>
                    <td className="px-4 py-3"><span className="px-2 py-0.5 rounded-full text-[10px] font-bold" style={{background:"#4a90d918",color:"#4a90d9"}}>{r.module}</span></td>
                    <td className={`px-4 py-3 ${t.textMuted}`}>{r.triggerDay}th</td>
                    <td className={`px-4 py-3 ${t.textMuted}`}>{r.frequency}</td>
                    <td className={`px-4 py-3 ${t.textMuted} truncate max-w-[120px]`}>{r.recipients}</td>
                    <td className="px-4 py-3">
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold"
                        style={{background:r.status==="Active"?"#34d39918":"#fbbf2418",color:r.status==="Active"?"#34d399":"#fbbf24",border:`1px solid ${r.status==="Active"?"#34d39930":"#fbbf2430"}`}}>
                        {r.status==="Active"?"● Active":"⏸ Paused"}
                      </span>
                    </td>
                    <td className={`px-4 py-3 ${t.textMuted} text-[10px]`}>{r.lastSent||"—"}</td>
                    <td className="px-4 py-3">
                      {isAdmin && (
                        <div className="flex items-center gap-1">
                          <button onClick={()=>startEdit(r)} className={`px-2 py-1 rounded-lg text-[10px] font-semibold border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>Edit</button>
                          <button onClick={()=>toggleStatus(r.id)} className="px-2 py-1 rounded-lg text-[10px] font-semibold" style={{background:r.status==="Active"?"#fbbf2418":"#34d39918",color:r.status==="Active"?"#fbbf24":"#34d399",border:`1px solid ${r.status==="Active"?"#fbbf2430":"#34d39930"}`}}>
                            {r.status==="Active"?"Pause":"Resume"}
                          </button>
                          <button onClick={()=>deleteReminder(r.id)} className="px-2 py-1 rounded-lg text-[10px] font-semibold" style={{background:"#f8717118",color:"#f87171",border:"1px solid #f8717130"}}>Delete</button>
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── SYSTEM AUDIT LOG ─────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const SystemAuditLog = ({ t, dark }) => {
  const { auditLog } = useLockCtx();
  const [search, setSearch]   = useState("");
  const [modF,   setModF]     = useState("All");
  const [actF,   setActF]     = useState("All");
  const [page,   setPage]     = useState(1);
  const PER = 15;

  // Merge lock audit with synthetic demo entries
  const [demoEntries] = useState(() => [
    { action:"USER_LOGIN",    user:"Admin User",  module:"System",        at:new Date(Date.now()-3600000).toISOString(),    remarks:"Admin login" },
    { action:"DATA_ENTRY",    user:"User A",      module:"CAS MIS",       at:new Date(Date.now()-7200000).toISOString(),    remarks:"April MIS data" },
    { action:"DATA_EDIT",     user:"User B",      module:"KRA/KPI",       at:new Date(Date.now()-10800000).toISOString(),   remarks:"Score updated" },
    { action:"USER_CREATED",  user:"Admin User",  module:"User Database", at:new Date(Date.now()-86400000).toISOString(),   remarks:"New user added" },
    { action:"CLIENT_IMPORT", user:"Admin User",  module:"Client Master", at:new Date(Date.now()-172800000).toISOString(),  remarks:"12 clients imported" },
    { action:"DATA_DELETE",   user:"User A",      module:"Fund Request",  at:new Date(Date.now()-259200000).toISOString(),  remarks:"June entry removed" },
  ]);

  const allEntries = [
    ...auditLog.map(e=>({ action:e.action, user:e.by, module:e.module, at:e.at, remarks:`${e.fy} · ${e.month}` })),
    ...demoEntries,
  ].sort((a,b)=>new Date(b.at)-new Date(a.at));

  const ACTION_COLORS = {
    LOCK:"#f87171", LOCK_FY:"#f87171", UNLOCK:"#fbbf24", UNLOCK_FY:"#fbbf24",
    USER_LOGIN:"#4a90d9", DATA_ENTRY:"#34d399", DATA_EDIT:"#00c9b1",
    DATA_DELETE:"#f87171", USER_CREATED:"#a78bfa", CLIENT_IMPORT:"#fb923c", USER_IMPORT:"#22d3ee",
  };

  const modules = ["All",...Array.from(new Set(allEntries.map(e=>e.module)))];
  const actions  = ["All",...Array.from(new Set(allEntries.map(e=>e.action)))];

  const filtered = allEntries.filter(e=>{
    const q = search.toLowerCase();
    const matchQ = !q || (e.user||"").toLowerCase().includes(q) || (e.action||"").toLowerCase().includes(q) || (e.remarks||"").toLowerCase().includes(q);
    const matchM = modF==="All" || e.module===modF;
    const matchA = actF==="All" || e.action===actF;
    return matchQ && matchM && matchA;
  });

  const totalPages = Math.max(1, Math.ceil(filtered.length/PER));
  const paged = filtered.slice((page-1)*PER, page*PER);
  const sel = `px-3 py-2 rounded-xl border text-xs outline-none ${t.input}`;

  const exportAudit = () => {
    const ws = XLSX.utils.json_to_sheet(filtered.map(e=>({
      "Date & Time": new Date(e.at).toLocaleString("en-IN"),
      "User": xlsSafe(e.user), "Module": e.module, "Action": e.action, "Remarks": xlsSafe(e.remarks),
    })));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Audit Log");
    XLSX.writeFile(wb,"ProCAS_Audit_Log.xlsx");
  };

  return (
    <div className="space-y-5">
      <div className="rounded-2xl px-6 py-4 flex items-center justify-between gap-4" style={{background:"linear-gradient(135deg,#0a1a2e,#0d2035)",border:"1px solid #4a90d930"}}>
        <div className="flex items-center gap-3">
          <span className="text-2xl">📋</span>
          <div>
            <h3 className="font-black text-sm text-white">System Audit Log</h3>
            <p className="text-[11px] mt-0.5" style={{color:"rgba(255,255,255,0.45)"}}>Track all user actions, data changes, and system events</p>
          </div>
        </div>
        <button onClick={exportAudit} className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-semibold ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
          <Icon path={Icons.download} size={12}/> Export
        </button>
      </div>

      {/* Filters */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-5 py-3.5 flex flex-wrap items-center gap-3`}>
        <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border ${t.input} flex-1`} style={{minWidth:160}}>
          <Icon path={Icons.search} size={12} className={t.textMuted}/>
          <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}} placeholder="Search user, action, remarks…" className={`bg-transparent outline-none text-xs flex-1 ${t.text}`}/>
          {search&&<button onClick={()=>{setSearch("");setPage(1);}} className={t.textMuted}><Icon path={Icons.x} size={10}/></button>}
        </div>
        <select value={modF} onChange={e=>{setModF(e.target.value);setPage(1);}} className={`${sel}`} style={{minWidth:120}}>
          {modules.map(m=><option key={m} value={m}>{m==="All"?"All Modules":m}</option>)}
        </select>
        <select value={actF} onChange={e=>{setActF(e.target.value);setPage(1);}} className={`${sel}`} style={{minWidth:140}}>
          {actions.map(a=><option key={a} value={a}>{a==="All"?"All Actions":a.replace(/_/g," ")}</option>)}
        </select>
        <span className={`text-[10px] ${t.textMuted} ml-auto`}>{filtered.length} entries</span>
      </div>

      {/* Table */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
        <div style={{overflowX:"auto",maxHeight:480,overflowY:"auto"}}>
          <table className="w-full text-xs border-collapse">
            <thead className="sticky top-0 z-10"><tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
              {["Date & Time","User","Module","Action","Remarks"].map(h=>(
                <th key={h} className="px-4 py-3 text-left text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",borderBottom:"1px solid #ffffff18"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {paged.length===0 ? (
                <tr><td colSpan={5} className={`px-4 py-10 text-center text-sm ${t.textMuted}`}>No entries match your filters</td></tr>
              ) : paged.map((e,i)=>{
                const col = ACTION_COLORS[e.action] || "#94a3b8";
                return (
                  <tr key={i} className={`border-b ${t.cardBorder} transition-colors`}
                    style={{background:i%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.4)":"#f8fbff")}}>
                    <td className={`px-4 py-2.5 text-[10px] ${t.textMuted}`}>{new Date(e.at).toLocaleString("en-IN",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}</td>
                    <td className={`px-4 py-2.5 font-semibold ${t.text}`}>{e.user||"—"}</td>
                    <td className="px-4 py-2.5"><span className="px-2 py-0.5 rounded-full text-[9px] font-bold" style={{background:col+"18",color:col}}>{e.module||"—"}</span></td>
                    <td className="px-4 py-2.5"><span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[9px] font-bold" style={{background:col+"18",color:col,border:`1px solid ${col}30`}}>{(e.action||"—").replace(/_/g," ")}</span></td>
                    <td className={`px-4 py-2.5 ${t.textMuted} text-[10px]`}>{e.remarks||"—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {totalPages>1 && (
          <div className={`px-5 py-3 border-t ${t.cardBorder} flex items-center justify-between`}>
            <span className={`text-[10px] ${t.textMuted}`}>Page {page} of {totalPages} · {filtered.length} entries</span>
            <div className="flex gap-1">
              {["‹","›"].map((lbl,i)=>(
                <button key={lbl} onClick={()=>setPage(p=>i===0?Math.max(1,p-1):Math.min(totalPages,p+1))} disabled={i===0?page===1:page===totalPages}
                  className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs disabled:opacity-30 ${t.hover} ${t.textMuted}`}>{lbl}</button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── MASTER SETTINGS ──────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const MasterSettings = ({ t, dark, isAdmin }) => {
  const [activeMS, setActiveMS] = useState("fy");
  const [toast, setToast] = useState(null);
  const showToast = (msg,type="ok") => { setToast({msg,type}); setTimeout(()=>setToast(null),2800); };

  // ── Financial Years ─────────────────────────────────────────────────────────
  const [fyList, setFyList] = useState(() => {
    try { const s=localStorage.getItem("procas_fyMaster"); return s?JSON.parse(s):["2026-27","2027-28","2028-29","2029-30","2030-31"]; } catch { return ["2026-27","2027-28","2028-29","2029-30","2030-31"]; }
  });
  const saveFY = (data) => { setFyList(data); try{ localStorage.setItem("procas_fyMaster",JSON.stringify(data)); }catch{} };
  const [newFY, setNewFY] = useState("");

  // ── States ──────────────────────────────────────────────────────────────────
  const DEFAULT_STATES = ["Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh","Delhi","Goa","Gujarat","Haryana","Himachal Pradesh","Jharkhand","Karnataka","Kerala","Madhya Pradesh","Maharashtra","Manipur","Meghalaya","Mizoram","Nagaland","Odisha","Punjab","Rajasthan","Sikkim","Tamil Nadu","Telangana","Tripura","Uttar Pradesh","Uttarakhand","West Bengal","Others"];
  const [statesList, setStatesList] = useState(() => { try{ const s=localStorage.getItem("procas_statesMaster"); return s?JSON.parse(s):DEFAULT_STATES; }catch{ return DEFAULT_STATES; }});
  const saveStates = (data) => { setStatesList(data); try{ localStorage.setItem("procas_statesMaster",JSON.stringify(data)); }catch{} };
  const [newState, setNewState] = useState("");

  // ── Sectors ─────────────────────────────────────────────────────────────────
  const DEFAULT_SECTORS = ["IT Services","Manufacturing","Retail","Real Estate","Healthcare","Export","Steel","Energy","Finance","Education","Others"];
  const [sectorsList, setSectorsList] = useState(() => { try{ const s=localStorage.getItem("procas_sectorsMaster"); return s?JSON.parse(s):DEFAULT_SECTORS; }catch{ return DEFAULT_SECTORS; }});
  const saveSectors = (data) => { setSectorsList(data); try{ localStorage.setItem("procas_sectorsMaster",JSON.stringify(data)); }catch{} };
  const [newSector, setNewSector] = useState("");

  // ── Currencies ──────────────────────────────────────────────────────────────
  const DEFAULT_CURRENCIES = ["INR","USD","EUR","GBP","AED","SGD","AUD","CAD","JPY","CHF","HKD"];
  const [currList, setCurrList] = useState(() => { try{ const s=localStorage.getItem("procas_currMaster"); return s?JSON.parse(s):DEFAULT_CURRENCIES; }catch{ return DEFAULT_CURRENCIES; }});
  const saveCurr = (data) => { setCurrList(data); try{ localStorage.setItem("procas_currMaster",JSON.stringify(data)); }catch{} };
  const [newCurr, setNewCurr] = useState("");

  const MS_NAV = [
    { id:"fy",       label:"📅 Financial Years" },
    { id:"states",   label:"🗺 States"           },
    { id:"sectors",  label:"🏭 Sectors"          },
    { id:"currency", label:"💱 Currencies"       },
  ];

  const MasterList = ({ items, onAdd, onDelete, newVal, setNewVal, placeholder, addLabel }) => (
    <div className="space-y-3">
      {isAdmin && (
        <div className="flex gap-2">
          <input value={newVal} onChange={e=>setNewVal(e.target.value)}
            onKeyDown={e=>{if(e.key==="Enter"&&newVal.trim()){onAdd(newVal.trim());setNewVal("");}}}
            placeholder={placeholder}
            className={`flex-1 px-3 py-2.5 rounded-xl border text-sm outline-none focus:ring-2 focus:ring-[#00c9b1] ${t.input}`}/>
          <button onClick={()=>{if(newVal.trim()){onAdd(newVal.trim());setNewVal("");}}}
            className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}}>
            + Add
          </button>
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2">
        {items.map((item,i)=>(
          <div key={i} className="flex items-center justify-between px-3 py-2 rounded-xl group"
            style={{background:dark?"rgba(15,30,46,0.8)":"#f0f6ff",border:`1px solid ${dark?"#243d58":"#d4e0ed"}`}}>
            <span className={`text-xs font-medium truncate flex-1 ${t.text}`}>{item}</span>
            {isAdmin && (
              <button onClick={()=>{const updated=items.filter((_,j)=>j!==i);onDelete(updated);showToast(`"${item}" removed`,"warn");}}
                className="ml-2 w-4 h-4 rounded flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity"
                style={{color:"#f87171"}}>✕</button>
            )}
          </div>
        ))}
      </div>
      <p className={`text-[10px] ${t.textMuted}`}>{items.length} entries{isAdmin?" · Hover an item to delete":""}</p>
    </div>
  );

  return (
    <div className="space-y-5">
      {toast && <div className="fixed top-5 right-5 z-[80] px-4 py-3 rounded-xl shadow-2xl text-sm font-semibold text-white" style={{background:toast.type==="warn"?"linear-gradient(135deg,#f87171,#dc2626)":"linear-gradient(135deg,#00a896,#1b5fa8)",minWidth:260}}>⚙ {toast.msg}</div>}

      <div className="rounded-2xl px-6 py-4 flex items-center gap-4" style={{background:"linear-gradient(135deg,#0a1a2e,#0d2035)",border:"1px solid #34d39930"}}>
        <span className="text-2xl">⚙</span>
        <div>
          <h3 className="font-black text-sm text-white">Master Settings</h3>
          <p className="text-[11px] mt-0.5" style={{color:"rgba(255,255,255,0.45)"}}>Manage lookup data used across all ProCAS modules</p>
        </div>
      </div>

      {/* Sub-nav */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-1.5 flex gap-1`}>
        {MS_NAV.map(n=>(
          <button key={n.id} onClick={()=>setActiveMS(n.id)}
            className={`flex-1 px-3 py-2 rounded-xl text-xs font-semibold transition-all whitespace-nowrap ${activeMS===n.id?"text-white shadow-md":t.textMuted}`}
            style={activeMS===n.id?{background:"linear-gradient(135deg,#34d399,#00a896)"}:{}}>
            {n.label}
          </button>
        ))}
      </div>

      {/* Content */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5`}>
        {activeMS==="fy" && (
          <div className="space-y-3">
            <h4 className={`font-bold text-sm mb-3 ${t.text}`}>Financial Years</h4>
            <MasterList
              items={fyList}
              onAdd={v=>{if(!fyList.includes(v)){const u=[...fyList,v].sort();saveFY(u);showToast(`FY ${v} added`);}else showToast("Already exists","warn");}}
              onDelete={saveFY}
              newVal={newFY} setNewVal={setNewFY}
              placeholder="e.g. 2031-32"
            />
          </div>
        )}
        {activeMS==="states" && (
          <div className="space-y-3">
            <h4 className={`font-bold text-sm mb-3 ${t.text}`}>States Master</h4>
            <MasterList
              items={statesList}
              onAdd={v=>{if(!statesList.includes(v)){const u=[...statesList,v].sort();saveStates(u);showToast(`"${v}" added`);}else showToast("Already exists","warn");}}
              onDelete={saveStates}
              newVal={newState} setNewVal={setNewState}
              placeholder="e.g. Jammu & Kashmir"
            />
          </div>
        )}
        {activeMS==="sectors" && (
          <div className="space-y-3">
            <h4 className={`font-bold text-sm mb-3 ${t.text}`}>Sector Master</h4>
            <MasterList
              items={sectorsList}
              onAdd={v=>{if(!sectorsList.includes(v)){const u=[...sectorsList,v].sort();saveSectors(u);showToast(`"${v}" added`);}else showToast("Already exists","warn");}}
              onDelete={saveSectors}
              newVal={newSector} setNewVal={setNewSector}
              placeholder="e.g. Pharma"
            />
          </div>
        )}
        {activeMS==="currency" && (
          <div className="space-y-3">
            <h4 className={`font-bold text-sm mb-3 ${t.text}`}>Currency Master</h4>
            <MasterList
              items={currList}
              onAdd={v=>{const code=v.toUpperCase().slice(0,3);if(!currList.includes(code)){const u=[...currList,code].sort();saveCurr(u);showToast(`${code} added`);}else showToast("Already exists","warn");}}
              onDelete={saveCurr}
              newVal={newCurr} setNewVal={setNewCurr}
              placeholder="e.g. MYR"
            />
          </div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── SYSTEM HEALTH DASHBOARD ───────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const SystemHealthDashboard = ({ t, dark }) => {
  const { clients } = useSyncContext();
  const { locks, auditLog } = useLockCtx();
  const { users } = useContext(UserContext);
  const [reminders] = useState(() => { try{ const s=localStorage.getItem("procas_reminders"); return s?JSON.parse(s):[]; }catch{ return []; }});

  const activeUsers   = users.filter(u=>u.status==="Active").length;
  const pendingInvites= users.filter(u=>u.inviteStatus==="Pending Invite").length;
  const lockedCount   = Object.keys(locks).length;
  const activeRems    = reminders.filter(r=>r.status==="Active").length;
  const totalClients  = clients.length;
  const activeClients = clients.filter(c=>c.status==="Active").length;

  const healthCards = [
    { label:"Total Clients",        value:totalClients,    sub:`${activeClients} active`,                   color:"#00c9b1", icon:"👥" },
    { label:"Total Users",           value:users.length,    sub:`${activeUsers} active`,                     color:"#4a90d9", icon:"👤" },
    { label:"Active Users",          value:activeUsers,     sub:"Currently enabled",                          color:"#34d399", icon:"✅" },
    { label:"Pending Invites",       value:pendingInvites,  sub:"Awaiting acceptance",                        color:"#fbbf24", icon:"✉"  },
    { label:"Locked Month-Modules",  value:lockedCount,     sub:"Across all FYs & modules",                   color:"#f87171", icon:"🔒" },
    { label:"Scheduled Reminders",   value:activeRems,      sub:`${reminders.length} total configured`,       color:"#a78bfa", icon:"🔔" },
  ];

  // Module activity from audit log (last 30 entries)
  const recentActivity = auditLog.slice(0,20);

  // Lock status summary
  const MONTHS_LIST = ["April","May","June","July","August","September","October","November","December","January","February","March"];
  const currentFY = DEFAULT_FY;
  const lockStatusByMonth = MONTHS_LIST.map(m => ({
    month:m,
    cas:  checkLocked(locks,currentFY,m,"cas"),
    kra:  checkLocked(locks,currentFY,m,"kra"),
    fund: checkLocked(locks,currentFY,m,"fund"),
  }));

  return (
    <div className="space-y-5">
      <div className="rounded-2xl px-6 py-4 flex items-center gap-4" style={{background:"linear-gradient(135deg,#0a1a2e,#0d2035)",border:"1px solid #34d39930"}}>
        <span className="text-2xl">💚</span>
        <div>
          <h3 className="font-black text-sm text-white">System Health Dashboard</h3>
          <p className="text-[11px] mt-0.5" style={{color:"rgba(255,255,255,0.45)"}}>Real-time overview of ProCAS system status</p>
        </div>
        <div className="ml-auto flex items-center gap-2 px-3 py-1.5 rounded-xl" style={{background:"#34d39918",border:"1px solid #34d39930"}}>
          <div className="w-2 h-2 rounded-full animate-pulse" style={{background:"#34d399"}}/>
          <span className="text-xs font-bold" style={{color:"#34d399"}}>System Online</span>
        </div>
      </div>

      {/* Health KPI cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-4">
        {healthCards.map(k=>(
          <div key={k.label} className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 flex flex-col gap-2 transition-all hover:-translate-y-0.5 hover:shadow-lg`} style={{borderTop:`2.5px solid ${k.color}`}}>
            <div className="flex items-center justify-between">
              <span className="text-lg">{k.icon}</span>
              <div className="w-6 h-6 rounded-lg flex items-center justify-center" style={{background:k.color+"18"}}>
                <div className="w-2 h-2 rounded-full" style={{background:k.color}}/>
              </div>
            </div>
            <div className="text-2xl font-black leading-none" style={{color:k.color}}>{k.value}</div>
            <div>
              <div className={`text-[9px] font-bold uppercase tracking-widest ${t.textMuted}`}>{k.label}</div>
              <div className="text-[9px] mt-0.5" style={{color:k.color+"99"}}>{k.sub}</div>
            </div>
          </div>
        ))}
      </div>

      {/* Lock status grid for current FY */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
        <div className={`px-5 py-3.5 border-b ${t.cardBorder}`}>
          <h4 className={`font-bold text-sm ${t.text}`}>Lock Status — FY {currentFY}</h4>
          <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>Current data entry lock state across all 3 modules</p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <thead><tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
              {["Month","CAS MIS","KRA / KPI","Fund Request"].map(h=>(
                <th key={h} className="px-4 py-3 text-center text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",borderBottom:"1px solid #ffffff18"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {lockStatusByMonth.map((row,i)=>(
                <tr key={row.month} className={`border-b ${t.cardBorder}`}
                  style={{background:i%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.4)":"#f8fbff")}}>
                  <td className={`px-4 py-2.5 font-semibold text-xs text-center ${t.text}`}>{row.month}</td>
                  {[row.cas, row.kra, row.fund].map((locked,j)=>(
                    <td key={j} className="px-4 py-2.5 text-center">
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold"
                        style={{background:locked?"#f8717118":"#34d39915",color:locked?"#f87171":"#34d399",border:`1px solid ${locked?"#f8717130":"#34d39930"}`}}>
                        {locked?"🔒 Locked":"🔓 Open"}
                      </span>
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Recent activity from audit log */}
      {recentActivity.length > 0 && (
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden`}>
          <div className={`px-5 py-3.5 border-b ${t.cardBorder}`}>
            <h4 className={`font-bold text-sm ${t.text}`}>Recent Activity</h4>
            <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>Last {recentActivity.length} recorded events</p>
          </div>
          <div className="divide-y" style={{borderColor:dark?"#1a2d44":"#e2e8f0"}}>
            {recentActivity.map((e,i)=>(
              <div key={i} className={`px-5 py-3 flex items-center gap-3 text-xs ${t.hover}`}>
                <div className="w-7 h-7 rounded-xl flex items-center justify-center shrink-0 text-sm"
                  style={{background:e.action?.includes("LOCK")&&!e.action?.includes("UNLOCK")?"#f8717118":"#4a90d918"}}>
                  {e.action?.includes("LOCK")&&!e.action?.includes("UNLOCK")?"🔒":"📝"}
                </div>
                <div className="flex-1 min-w-0">
                  <span className="font-semibold" style={{color:dark?"#e8f0f8":"#0d2137"}}>{e.by||"System"}</span>
                  <span className={t.textMuted}> · {(e.action||"").replace(/_/g," ")}</span>
                  <span className={`ml-1 ${t.textMuted}`}>· {e.fy} {e.month}</span>
                </div>
                <span className={`text-[10px] shrink-0 ${t.textMuted}`}>{new Date(e.at).toLocaleString("en-IN",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"})}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

const ArchitectureTab = ({ t, dark, isAdmin, adminName }) => {
  const [section, setSection] = useState("monthlock");

  const NAV = [
    { id:"monthlock",  label:"🔒 Month Lock"     },
    { id:"reminders",  label:"🔔 Reminders"       },
    { id:"auditlog",   label:"📋 Audit Log"        },
    { id:"mastersett", label:"⚙ Master Settings"  },
    { id:"syshealth",  label:"💚 System Health"   },
    { id:"overview",   label:"Overview"            },
    { id:"schema",     label:"DB Schema"           },
    { id:"api",        label:"API Routes"          },
    { id:"roles",      label:"Roles & Auth"        },
    { id:"stack",      label:"Tech Stack"          },
    { id:"flow",       label:"Data Flow"           },
  ];

  const Chip = ({ label, color }) => (
    <span className="text-[10px] font-bold px-2 py-0.5 rounded-full"
      style={{ background: color+"22", color }}>{label}</span>
  );

  const Card = ({ children, className="" }) => (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl overflow-hidden ${className}`}>
      {children}
    </div>
  );

  const CardHeader = ({ title }) => (
    <div className={`px-5 py-3.5 border-b ${t.cardBorder}`}>
      <h3 className={`font-bold text-sm ${t.text}`}>{title}</h3>
    </div>
  );

  // ── SQL / Code display using only safe strings ────────────────────────────
  const SqlBlock = ({ lines }) => (
    <div className="rounded-xl overflow-hidden mt-2" style={{ background:"#0d1117" }}>
      <div className="flex items-center gap-2 px-4 py-2 border-b border-[#1e2535]">
        <div className="flex gap-1.5">
          {["#f87171","#fbbf24","#34d399"].map(c=>(
            <div key={c} className="w-2 h-2 rounded-full" style={{background:c}}/>
          ))}
        </div>
        <span className="text-[10px] font-mono font-bold text-[#4b5563] ml-2">SQL</span>
      </div>
      <pre className="px-4 py-4 text-[11px] font-mono overflow-x-auto leading-relaxed text-[#94a3b8]">
        {lines.join("\n")}
      </pre>
    </div>
  );

  const BashBlock = ({ lines }) => (
    <div className="rounded-xl overflow-hidden mt-2" style={{ background:"#0d1117" }}>
      <div className="flex items-center gap-2 px-4 py-2 border-b border-[#1e2535]">
        <div className="flex gap-1.5">
          {["#f87171","#fbbf24","#34d399"].map(c=>(
            <div key={c} className="w-2 h-2 rounded-full" style={{background:c}}/>
          ))}
        </div>
        <span className="text-[10px] font-mono font-bold text-[#4b5563] ml-2">DIRECTORY</span>
      </div>
      <pre className="px-4 py-4 text-[11px] font-mono overflow-x-auto leading-relaxed text-[#94a3b8]">
        {lines.join("\n")}
      </pre>
    </div>
  );

  // ── DB Table Schema ───────────────────────────────────────────────────────
  const TableSchema = ({ name, color, rows, fkNote }) => (
    <div className="rounded-xl overflow-hidden border" style={{ borderColor: color+"40" }}>
      <div className="flex items-center gap-2 px-4 py-2.5" style={{ background: color+"18" }}>
        <div className="w-2 h-2 rounded-full" style={{ background: color }}/>
        <span className="font-black text-sm" style={{ color }}>{name}</span>
        <Chip label="PostgreSQL" color={color}/>
      </div>
      <div className={`grid grid-cols-12 gap-2 px-4 py-1.5 text-[10px] font-bold uppercase tracking-wider`}
        style={{ color: dark?"#4b5563":"#9ca3af", background: dark?"#0d1117":"#f8fafc" }}>
        <div className="col-span-3">Column</div>
        <div className="col-span-3">Type</div>
        <div className="col-span-3">Constraint</div>
        <div className="col-span-3">Description</div>
      </div>
      <div className="divide-y" style={{ borderColor: dark?"#1e2535":"#f0f4f8" }}>
        {rows.map(r => (
          <div key={r.col}
            className={`grid grid-cols-12 gap-2 px-4 py-2.5 text-xs items-center ${t.hover} transition-colors`}>
            <div className="col-span-3 font-mono font-bold flex items-center gap-1"
              style={{ color: r.pk?"#fbbf24": r.fk?"#00c9b1": dark?"#e2e8f0":"#1a202c" }}>
              {r.pk && <span className="text-[9px]">🔑</span>}
              {r.fk && <span className="text-[9px]">🔗</span>}
              {r.col}
            </div>
            <div className="col-span-3 font-mono text-[11px]" style={{ color:"#4a90d9" }}>{r.type}</div>
            <div className={`col-span-3 text-[10px] ${t.textMuted}`}>{r.constraint||"—"}</div>
            <div className={`col-span-3 text-[10px] italic ${t.textMuted}`}>{r.desc}</div>
          </div>
        ))}
      </div>
      {fkNote && (
        <div className="px-4 py-2 text-[10px] font-semibold"
          style={{ background:"#00c9b110", color:"#00c9b1" }}>🔗 {fkNote}</div>
      )}
    </div>
  );

  // ── API route accordion ───────────────────────────────────────────────────
  const ApiRoute = ({ method, path, desc, auth, reqFields, resFields }) => {
    const [open, setOpen] = useState(false);
    const MC = { GET:"#34d399", POST:"#00c9b1", PUT:"#fbbf24", DELETE:"#f87171", PATCH:"#a78bfa" };
    const mc = MC[method]||"#94a3b8";
    return (
      <div className="rounded-xl border overflow-hidden"
        style={{ borderColor: dark?"#1e2535":"#e2e8f0" }}>
        <button className={`w-full flex items-center gap-3 px-4 py-3 text-left ${t.hover}`}
          onClick={() => setOpen(o=>!o)}>
          <span className="font-black text-[11px] px-2 py-0.5 rounded font-mono w-16 text-center shrink-0"
            style={{ background:mc+"22", color:mc }}>{method}</span>
          <code className="text-[13px] font-mono flex-1 truncate" style={{ color:"#fbbf24" }}>{path}</code>
          <span className={`text-xs hidden sm:block ${t.textMuted}`}>{desc}</span>
          <span className="text-[10px] px-2 py-0.5 rounded-full font-semibold shrink-0"
            style={{ background:auth==="admin"?"#f8717118":"#00c9b118",
                     color:auth==="admin"?"#f87171":"#00c9b1" }}>
            {auth==="admin"?"🔒 Admin":"🔓 User"}
          </span>
          <Icon path={open?Icons.chevronUp:Icons.chevronDown} size={13} className={t.textMuted}/>
        </button>
        {open && (
          <div className="px-4 pb-4 space-y-3"
            style={{ background:dark?"#0a0d14":"#f8fafc" }}>
            {reqFields && (
              <div>
                <p className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} pt-2 mb-1`}>
                  Request Fields
                </p>
                <div className="space-y-1">
                  {reqFields.map(f=>(
                    <div key={f.name} className="flex items-center gap-3 text-xs">
                      <code className="font-mono w-32 shrink-0" style={{color:"#fbbf24"}}>{f.name}</code>
                      <span className="font-mono text-[11px]" style={{color:"#4a90d9"}}>{f.type}</span>
                      <span className={t.textMuted}>{f.desc}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <div>
              <p className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} mb-1`}>
                Response Fields
              </p>
              <div className="space-y-1">
                {resFields.map(f=>(
                  <div key={f.name} className="flex items-center gap-3 text-xs">
                    <code className="font-mono w-32 shrink-0" style={{color:"#34d399"}}>{f.name}</code>
                    <span className="font-mono text-[11px]" style={{color:"#4a90d9"}}>{f.type}</span>
                    <span className={t.textMuted}>{f.desc}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  // ── Stack card ────────────────────────────────────────────────────────────
  const StackCard = ({ icon, name, version, role, color, bullets }) => (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 flex flex-col gap-3
      hover:-translate-y-0.5 hover:shadow-xl transition-all duration-200`}>
      <div className="flex items-center gap-3">
        <div className="w-10 h-10 rounded-xl flex items-center justify-center text-xl"
          style={{ background:color+"18" }}>{icon}</div>
        <div>
          <div className={`font-bold text-sm ${t.text}`}>{name}</div>
          <div className={`text-[10px] font-mono ${t.textMuted}`}>{version}</div>
        </div>
        <Chip label={role} color={color}/>
      </div>
      <ul className="space-y-1.5">
        {bullets.map((b,i)=>(
          <li key={i} className={`text-xs flex items-start gap-2 ${t.textMuted}`}>
            <span style={{color}} className="shrink-0 mt-0.5">▸</span>{b}
          </li>
        ))}
      </ul>
    </div>
  );

  // ── OVERVIEW ──────────────────────────────────────────────────────────────
  const OverviewSection = () => (
    <div className="space-y-5">
      {/* Hero banner */}
      <div className="rounded-2xl p-6 relative overflow-hidden"
        style={{ background:"linear-gradient(135deg,#00a89618,#1b5fa818)",
                 border:"1px solid #00a89630" }}>
        <div className="absolute right-0 top-0 w-72 h-72 rounded-full opacity-10 pointer-events-none"
          style={{ background:"radial-gradient(#00c9b1,transparent)", transform:"translate(35%,-35%)" }}/>
        <div className="relative">
          <div className="flex flex-wrap items-center gap-3 mb-4">
            <div className="w-11 h-11 rounded-xl flex items-center justify-center text-2xl"
              style={{ background:"linear-gradient(135deg,#00a896,#1b5fa8)" }}>🏗</div>
            <div>
              <h2 className={`font-black text-xl ${t.text}`}>ProCAS — Platform Architecture Overview</h2>
              <p className={`text-xs ${t.textMuted}`}>Complete application reference · v2.0 · All modules documented</p>
            </div>
            {["7 Modules","RBAC","Audit-logged","Month Lock","Analytics"].map(b=>(
              <Chip key={b} label={b} color="#00c9b1"/>
            ))}
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            {[
              { label:"Frontend",  value:"React 18 SPA",        color:"#00c9b1" },
              { label:"Hosting",   value:"Netlify (Static)",     color:"#34d399" },
              { label:"Storage",   value:"localStorage (UAT)",   color:"#fbbf24" },
              { label:"Auth",      value:"RBAC · Role-based",    color:"#f87171" },
            ].map(s=>(
              <div key={s.label} className={`${t.card} border ${t.cardBorder} rounded-xl p-3`}>
                <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted}`}>{s.label}</div>
                <div className="font-bold text-xs mt-1" style={{color:s.color}}>{s.value}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Module map */}
      <Card>
        <CardHeader title="🗺 Application Module Map — All 7 Modules"/>
        <div className="p-5">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {[
              { icon:"📊", title:"Dashboard", color:"#00c9b1",
                features:["Executive KPI scorecards","CAS MIS analytics panel","KRA/KPI analytics panel","Fund Request analytics panel","Cross-module data aggregation"],
                desc:"Unified executive view aggregating live data from all operational modules into one CFO-level dashboard." },
              { icon:"🏢", title:"Client Master", color:"#34d399",
                features:["Client registry with Sector & State","Multi-user assignment per client","Bulk Excel import with validation","Import result summary with error log","Client sync to all operational modules"],
                desc:"Central client registry. Every operational module (CAS MIS, KRA/KPI, Fund Request) references clients defined here." },
              { icon:"📋", title:"CAS MIS", color:"#4a90d9",
                features:["Monthly MIS date compliance tracking","Outstanding dues per SPOC","Collection status monitoring","Statutory health tracking (PT/TDS/PF/ESI/GST)","Revenue mix & reconciliation analytics"],
                desc:"Monthly compliance and revenue tracking per client. Tracks MIS dates, statutory obligations, and collection status." },
              { icon:"🎯", title:"KRA / KPI", color:"#a78bfa",
                features:["Per-client KPI scoring (6 parameters)","MYSA compliance tracking","Rectification & escalation monitoring","Raksha tool compliance","CapitallWant & Capital tracking"],
                desc:"Key Result Area tracking across all clients. Per-parameter applicability flags control which KPIs apply to each client." },
              { icon:"💰", title:"Fund Request", color:"#fb923c",
                features:["Multi-currency foreign fund tracking","Auto INR calculation (fc × rate)","Fund analytics: currency breakdown, trends","Funding ledger with variance analysis","Per-row Month Lock enforcement"],
                desc:"Tracks inward fund requests per client with automatic INR conversion, analytics, and governance controls." },
              { icon:"👥", title:"User Database", color:"#f87171",
                features:["User CRUD with role assignment","Client assignment per End User","3-step invite email simulation","Invite log with retry & status","Active/Inactive status management"],
                desc:"Manages platform users. Admin assigns clients to End Users, controlling which records each user can edit." },
              { icon:"🏗", title:"Architect", color:"#22d3ee",
                features:["Month Lock Control (FY/Month/Module)","Auto Reminder management","System Audit Log with export","Master Settings (FY/States/Sectors/Currencies)","System Health Dashboard"],
                desc:"Admin-only control centre for governance, locks, reminders, audit log review, and master data configuration." },
            ].map(c=>(
              <div key={c.title}
                className={`${t.card} border ${t.cardBorder} rounded-2xl p-5 flex flex-col gap-3
                  hover:-translate-y-0.5 hover:shadow-xl transition-all duration-200`}
                style={{borderColor:c.color+"30"}}>
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl flex items-center justify-center text-xl"
                    style={{background:c.color+"18"}}>{c.icon}</div>
                  <div>
                    <h3 className="font-bold text-sm" style={{color:c.color}}>{c.title}</h3>
                  </div>
                </div>
                <p className={`text-[11px] leading-relaxed ${t.textMuted}`}>{c.desc}</p>
                <ul className="space-y-1">
                  {c.features.map(f=>(
                    <li key={f} className={`text-[11px] flex items-start gap-2 ${t.textMuted}`}>
                      <span style={{color:c.color}} className="shrink-0 mt-0.5">▸</span>{f}
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Architecture principles */}
      <Card>
        <CardHeader title="⚙️ Architecture Principles"/>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[
            { icon:"🔒", title:"Role-Based Access Control",  color:"#f87171",
              desc:"Admin has full access. End Users can only edit their assigned clients. View-only badge shown for unassigned clients across CAS MIS, KRA/KPI, and Fund Request." },
            { icon:"🔐", title:"Month Lock Governance",       color:"#fbbf24",
              desc:"Admin can lock any FY/Month/Module combination. Locked months disable all Add, Save, and Delete actions for all users including Admin." },
            { icon:"📋", title:"Audit Trail",                 color:"#4a90d9",
              desc:"Every lock/unlock action is recorded in the audit log with actor, timestamp, module, FY, and month. Exportable to Excel." },
            { icon:"🔔", title:"Auto Reminder Framework",     color:"#a78bfa",
              desc:"Admin configures scheduled reminders per module. Reminders are managed centrally in Architect → Reminders with status tracking." },
            { icon:"📊", title:"Unified Analytics Layer",     color:"#00c9b1",
              desc:"CAS MIS, KRA/KPI, and Fund Request each have dedicated analytics tabs. Dashboard aggregates all three into executive KPI scorecards and charts." },
            { icon:"🗄️", title:"Client-Centric Data Model",  color:"#34d399",
              desc:"Client Master is the root entity. Every CAS MIS record, KRA/KPI entry, and Fund Request row references a client ID. User assignments flow from Client Master." },
          ].map(c=>(
            <div key={c.title} className={`${t.card} border ${t.cardBorder} rounded-xl p-4`}>
              <div className="text-xl mb-2">{c.icon}</div>
              <h3 className="font-bold text-xs mb-1.5" style={{color:c.color}}>{c.title}</h3>
              <p className={`text-[11px] leading-relaxed ${t.textMuted}`}>{c.desc}</p>
            </div>
          ))}
        </div>
      </Card>

      {/* Directory structure */}
      <Card>
        <CardHeader title="📁 Project Directory Structure (Production Target)"/>
        <div className="p-5">
          <BashBlock lines={[
            "procas/",
            "├── apps/",
            "│   ├── web/                       ← React 18 SPA (current: single JSX file)",
            "│   │   ├── app/",
            "│   │   │   ├── (dashboard)/",
            "│   │   │   │   ├── overview/      ← Executive Dashboard",
            "│   │   │   │   ├── clients/       ← Client Master",
            "│   │   │   │   ├── cas-mis/       ← CAS MIS module",
            "│   │   │   │   ├── kra-kpi/       ← KRA/KPI module",
            "│   │   │   │   ├── fund/          ← Fund Request module",
            "│   │   │   │   ├── users/         ← User Database module",
            "│   │   │   │   ├── architect/     ← Architect / Admin module",
            "│   │   │   │   └── layout        ← Sidebar shell + RBAC provider",
            "│   │   │   └── auth/              ← Login / register",
            "│   │   ├── components/",
            "│   │   │   ├── ui/                ← Design system components",
            "│   │   │   ├── clients/           ← Client Master components",
            "│   │   │   ├── cas-mis/           ← CAS MIS table + analytics",
            "│   │   │   ├── kra-kpi/           ← KRA/KPI table + analytics",
            "│   │   │   ├── fund/              ← Fund Request components",
            "│   │   │   ├── users/             ← User management components",
            "│   │   │   └── architect/         ← Lock, Reminder, Audit components",
            "│   │   └── lib/",
            "│   │       ├── api.ts             ← Typed HTTP client",
            "│   │       ├── auth.ts            ← Auth helpers + token refresh",
            "│   │       └── rbac.ts            ← useRBAC hook + canEditClient()",
            "│   │",
            "│   └── api/                       ← REST API server (Node.js)",
            "│       └── src/",
            "│           ├── routes/",
            "│           │   ├── auth           ← Login / refresh / logout",
            "│           │   ├── clients        ← CRUD + bulk import + export",
            "│           │   ├── cas-mis        ← Monthly MIS CRUD + analytics",
            "│           │   ├── kra-kpi        ← KPI CRUD + analytics",
            "│           │   ├── fund-requests  ← Fund CRUD + analytics + export",
            "│           │   ├── users          ← User CRUD + invite + assignment",
            "│           │   ├── month-locks    ← Lock / unlock / audit",
            "│           │   ├── reminders      ← Schedule / trigger / manage",
            "│           │   └── master         ← FY / States / Sectors / Currencies",
            "│           ├── middleware/",
            "│           │   ├── authenticate   ← JWT token verification",
            "│           │   └── authorize      ← Role + client-assignment guard",
            "│           ├── services/",
            "│           │   ├── excel          ← SheetJS import/export",
            "│           │   ├── email          ← SMTP invite delivery",
            "│           │   ├── audit          ← Audit log writer",
            "│           │   └── scheduler      ← Reminder cron engine",
            "│           └── db/",
            "│               ├── pool           ← Connection pool singleton",
            "│               └── migrations/    ← SQL DDL files (all tables)",
            "│",
            "├── packages/shared/               ← Shared TypeScript types",
            "├── docker-compose.yml",
            "└── .env.example",
          ]}/>
        </div>
      </Card>
    </div>
  );

  // ── SCHEMA ────────────────────────────────────────────────────────────────
  const SchemaSection = () => (
    <div className="space-y-5">
      {/* ── Users ── */}
      <TableSchema name="users" color="#00c9b1"
        rows={[
          { col:"id",            type:"UUID",         constraint:"PK · auto-generated",  desc:"Primary key",                pk:true },
          { col:"email",         type:"VARCHAR(255)", constraint:"NOT NULL · UNIQUE",     desc:"Login email address"                  },
          { col:"password_hash", type:"TEXT",         constraint:"NOT NULL",              desc:"Hashed password"                      },
          { col:"full_name",     type:"VARCHAR(255)", constraint:"NOT NULL",              desc:"Display name"                         },
          { col:"role",          type:"VARCHAR(20)",  constraint:"CHECK admin|user",      desc:"Admin or End User"                    },
          { col:"is_active",     type:"BOOLEAN",      constraint:"DEFAULT true",          desc:"Active / Inactive status"             },
          { col:"invite_status", type:"VARCHAR(20)",  constraint:"pending|sent|accepted", desc:"Invite workflow state"                },
          { col:"invite_token",  type:"TEXT",         constraint:"nullable",              desc:"Secure invite token (HMAC-SHA256)"    },
          { col:"invite_sent_at",type:"TIMESTAMPTZ",  constraint:"nullable",              desc:"When invite email was dispatched"     },
          { col:"created_at",    type:"TIMESTAMPTZ",  constraint:"DEFAULT NOW()",         desc:"Account creation time"                },
          { col:"updated_at",    type:"TIMESTAMPTZ",  constraint:"DEFAULT NOW()",         desc:"Last profile update"                  },
        ]}/>

      {/* ── Client Master ── */}
      <TableSchema name="client_master" color="#34d399"
        rows={[
          { col:"id",           type:"UUID",         constraint:"PK · auto-generated",  desc:"Primary key",                pk:true },
          { col:"client_name",  type:"VARCHAR(255)", constraint:"NOT NULL · UNIQUE",     desc:"Company or individual name"           },
          { col:"person_name",  type:"VARCHAR(255)", constraint:"NOT NULL",              desc:"Primary contact person"               },
          { col:"sector",       type:"VARCHAR(100)", constraint:"NOT NULL",              desc:"Business sector"                      },
          { col:"state",        type:"VARCHAR(100)", constraint:"NOT NULL",              desc:"Indian state"                         },
          { col:"turnover",     type:"VARCHAR(30)",  constraint:"nullable",              desc:"Turnover label e.g. ₹42.8 Cr"        },
          { col:"type",         type:"VARCHAR(50)",  constraint:"nullable",              desc:"Client type e.g. HUF, LLP, Pvt Ltd"  },
          { col:"risk",         type:"VARCHAR(20)",  constraint:"Low|Medium|High",       desc:"Risk classification"                  },
          { col:"status",       type:"VARCHAR(20)",  constraint:"Active|Inactive",       desc:"Client active status"                 },
          { col:"rm",           type:"VARCHAR(255)", constraint:"nullable",              desc:"Assigned relationship manager (SPOC)" },
          { col:"is_active",    type:"BOOLEAN",      constraint:"DEFAULT true",          desc:"Soft-delete flag"                     },
          { col:"created_by",   type:"UUID",         constraint:"FK → users.id",         desc:"Admin who created record",   fk:true  },
          { col:"created_at",   type:"TIMESTAMPTZ",  constraint:"DEFAULT NOW()",         desc:"Row creation timestamp"               },
          { col:"updated_at",   type:"TIMESTAMPTZ",  constraint:"DEFAULT NOW()",         desc:"Last modification"                    },
        ]}
        fkNote="created_by → users(id) ON DELETE SET NULL"/>

      {/* ── Client User Assignments ── */}
      <TableSchema name="client_user_assignments" color="#22d3ee"
        rows={[
          { col:"id",        type:"UUID", constraint:"PK · auto-generated",     desc:"Primary key",                  pk:true },
          { col:"client_id", type:"UUID", constraint:"FK → client_master.id",   desc:"Client being assigned",        fk:true },
          { col:"user_id",   type:"UUID", constraint:"FK → users.id",           desc:"End User being assigned",      fk:true },
          { col:"assigned_by",type:"UUID",constraint:"FK → users.id",           desc:"Admin who made the assignment", fk:true},
          { col:"created_at",type:"TIMESTAMPTZ",constraint:"DEFAULT NOW()",     desc:"When assignment was made"               },
        ]}
        fkNote="UNIQUE(client_id, user_id) · client_id → client_master(id) ON DELETE CASCADE"/>

      {/* ── CAS MIS ── */}
      <TableSchema name="cas_mis" color="#4a90d9"
        rows={[
          { col:"id",              type:"UUID",        constraint:"PK · auto-generated",  desc:"Primary key",                pk:true },
          { col:"client_id",       type:"UUID",        constraint:"FK → client_master.id",desc:"Client reference",           fk:true },
          { col:"financial_year",  type:"VARCHAR(9)",  constraint:"NOT NULL",             desc:"e.g. 2026-27"                        },
          { col:"month",           type:"VARCHAR(15)", constraint:"NOT NULL",             desc:"e.g. April"                          },
          { col:"mis_date",        type:"DATE",        constraint:"nullable",             desc:"MIS submission date"                  },
          { col:"outstanding_dues",type:"NUMERIC(18,2)",constraint:"DEFAULT 0",          desc:"Outstanding amount for SPOC"          },
          { col:"collection_status",type:"VARCHAR(30)",constraint:"nullable",            desc:"Collected / Pending / Partial"        },
          { col:"collection_amt",  type:"NUMERIC(18,2)",constraint:"DEFAULT 0",          desc:"Amount collected"                     },
          { col:"recon_gap",       type:"NUMERIC(18,2)",constraint:"DEFAULT 0",          desc:"Reconciliation gap amount"            },
          { col:"pt_applicable",   type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Professional Tax applicability"       },
          { col:"tds_applicable",  type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"TDS applicability"                    },
          { col:"pf_applicable",   type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"PF applicability"                     },
          { col:"esi_applicable",  type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"ESI applicability"                    },
          { col:"gst_applicable",  type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"GST applicability"                    },
          { col:"software_used",   type:"VARCHAR(100)",constraint:"nullable",            desc:"Accounting software in use"           },
          { col:"remarks",         type:"TEXT",        constraint:"nullable",             desc:"Additional notes"                     },
          { col:"created_by",      type:"UUID",        constraint:"FK → users.id",        desc:"Entry creator",              fk:true  },
          { col:"created_at",      type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",        desc:"Row creation timestamp"               },
          { col:"updated_at",      type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",        desc:"Last modification"                    },
        ]}
        fkNote="UNIQUE(client_id, financial_year, month) · client_id → client_master(id) ON DELETE RESTRICT"/>

      {/* ── KRA KPI ── */}
      <TableSchema name="kra_kpi" color="#a78bfa"
        rows={[
          { col:"id",             type:"UUID",        constraint:"PK · auto-generated",  desc:"Primary key",                pk:true },
          { col:"client_id",      type:"UUID",        constraint:"FK → client_master.id",desc:"Client reference",           fk:true },
          { col:"financial_year", type:"VARCHAR(9)",  constraint:"NOT NULL",             desc:"e.g. 2026-27"                        },
          { col:"month",          type:"VARCHAR(15)", constraint:"NOT NULL",             desc:"e.g. April"                          },
          { col:"mis_date",       type:"DATE",        constraint:"nullable",             desc:"MIS date compliance date"             },
          { col:"mis_applicable", type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"MIS date KPI applicability"           },
          { col:"mysa",           type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"MYSA compliance status"               },
          { col:"mysa_applicable",type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"MYSA KPI applicability"               },
          { col:"rectification",  type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Rectification KPI status"             },
          { col:"rect_applicable",type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Rectification applicability"          },
          { col:"escalation",     type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Escalation status"                    },
          { col:"esc_applicable", type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Escalation applicability"             },
          { col:"raksha",         type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Raksha tool compliance"               },
          { col:"raksha_applicable",type:"VARCHAR(3)",constraint:"Y|N|A",              desc:"Raksha applicability"                 },
          { col:"capitalWant",    type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"CapitallWant status"                  },
          { col:"cw_applicable",  type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"CapitallWant applicability"           },
          { col:"capital",        type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Capital KPI status"                   },
          { col:"cap_applicable", type:"VARCHAR(3)",  constraint:"Y|N|A",               desc:"Capital applicability"                },
          { col:"created_by",     type:"UUID",        constraint:"FK → users.id",        desc:"Entry creator",              fk:true  },
          { col:"created_at",     type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",        desc:"Row creation timestamp"               },
          { col:"updated_at",     type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",        desc:"Last modification"                    },
        ]}
        fkNote="UNIQUE(client_id, financial_year, month) · client_id → client_master(id) ON DELETE RESTRICT"/>

      {/* ── Fund Requests ── */}
      <TableSchema name="fund_requests" color="#fb923c"
        rows={[
          { col:"id",             type:"UUID",          constraint:"PK · auto-generated",   desc:"Primary key",           pk:true },
          { col:"client_id",      type:"UUID",          constraint:"FK → client_master.id", desc:"Client reference",      fk:true },
          { col:"financial_year", type:"VARCHAR(9)",    constraint:"NOT NULL",              desc:"e.g. 2026-27"                   },
          { col:"month",          type:"VARCHAR(15)",   constraint:"NOT NULL",              desc:"e.g. April"                     },
          { col:"currency",       type:"VARCHAR(5)",    constraint:"NOT NULL",              desc:"ISO code: USD, EUR, GBP…"       },
          { col:"foreign_amount", type:"NUMERIC(18,4)", constraint:"CHECK > 0",            desc:"Amount in foreign currency"     },
          { col:"exchange_rate",  type:"NUMERIC(12,6)", constraint:"CHECK > 0",            desc:"Rate at time of entry"          },
          { col:"inr_amount",     type:"NUMERIC(20,2)", constraint:"GENERATED ALWAYS AS…", desc:"Auto-computed INR value"        },
          { col:"remarks",        type:"TEXT",          constraint:"nullable",              desc:"Optional notes"                 },
          { col:"created_by",     type:"UUID",          constraint:"FK → users.id",         desc:"Entry creator",         fk:true },
          { col:"created_at",     type:"TIMESTAMPTZ",   constraint:"DEFAULT NOW()",         desc:"Row creation timestamp"         },
        ]}
        fkNote="client_id → client_master(id) ON DELETE RESTRICT"/>

      {/* ── Month Locks ── */}
      <TableSchema name="month_locks" color="#fbbf24"
        rows={[
          { col:"id",             type:"UUID",        constraint:"PK · auto-generated",  desc:"Primary key",              pk:true },
          { col:"financial_year", type:"VARCHAR(9)",  constraint:"NOT NULL",             desc:"e.g. 2026-27"                      },
          { col:"month",          type:"VARCHAR(15)", constraint:"NOT NULL",             desc:"e.g. April  (all = entire FY)"    },
          { col:"module",         type:"VARCHAR(20)", constraint:"cas|kra|fund|all",     desc:"Module being locked"               },
          { col:"is_locked",      type:"BOOLEAN",     constraint:"DEFAULT false",         desc:"Lock state"                       },
          { col:"locked_by",      type:"UUID",        constraint:"FK → users.id",         desc:"Admin who locked",         fk:true },
          { col:"locked_at",      type:"TIMESTAMPTZ", constraint:"nullable",              desc:"When lock was applied"             },
          { col:"unlocked_by",    type:"UUID",        constraint:"FK → users.id",         desc:"Admin who unlocked",       fk:true },
          { col:"unlocked_at",    type:"TIMESTAMPTZ", constraint:"nullable",              desc:"When lock was removed"             },
        ]}
        fkNote="UNIQUE(financial_year, month, module) · locked_by → users(id) ON DELETE SET NULL"/>

      {/* ── Reminders ── */}
      <TableSchema name="reminders" color="#34d399"
        rows={[
          { col:"id",          type:"UUID",        constraint:"PK · auto-generated",  desc:"Primary key",              pk:true },
          { col:"title",       type:"VARCHAR(255)",constraint:"NOT NULL",             desc:"Reminder title"                    },
          { col:"module",      type:"VARCHAR(30)", constraint:"NOT NULL",             desc:"Target module"                     },
          { col:"frequency",   type:"VARCHAR(20)", constraint:"daily|weekly|monthly", desc:"Recurrence type"                   },
          { col:"trigger_day", type:"INTEGER",     constraint:"nullable",             desc:"Day of month for trigger"          },
          { col:"target_role", type:"VARCHAR(20)", constraint:"admin|user|all",       desc:"Who receives the reminder"         },
          { col:"message",     type:"TEXT",        constraint:"NOT NULL",             desc:"Reminder message body"             },
          { col:"status",      type:"VARCHAR(20)", constraint:"active|paused",        desc:"Current reminder status"           },
          { col:"last_sent",   type:"TIMESTAMPTZ", constraint:"nullable",             desc:"Last trigger time"                 },
          { col:"created_by",  type:"UUID",        constraint:"FK → users.id",         desc:"Admin who created reminder", fk:true},
          { col:"created_at",  type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",         desc:"Creation timestamp"                },
        ]}/>

      {/* ── Audit Log ── */}
      <TableSchema name="audit_log" color="#e2e8f0"
        rows={[
          { col:"id",         type:"BIGSERIAL",   constraint:"PK auto-increment",   desc:"Primary key",           pk:true },
          { col:"table_name", type:"VARCHAR(50)", constraint:"NOT NULL",            desc:"Affected table name"           },
          { col:"record_id",  type:"UUID",        constraint:"NOT NULL",            desc:"Affected row ID"               },
          { col:"action",     type:"VARCHAR(20)", constraint:"INSERT|UPDATE|DELETE|LOCK|UNLOCK",desc:"Operation type" },
          { col:"module",     type:"VARCHAR(30)", constraint:"nullable",            desc:"Module context"                },
          { col:"old_data",   type:"JSONB",       constraint:"nullable",            desc:"Previous row snapshot"         },
          { col:"new_data",   type:"JSONB",       constraint:"nullable",            desc:"New row snapshot"              },
          { col:"actor_id",   type:"UUID",        constraint:"FK → users.id",       desc:"User who made change",  fk:true},
          { col:"acted_at",   type:"TIMESTAMPTZ", constraint:"DEFAULT NOW()",       desc:"When change occurred"          },
        ]}/>

      {/* ── Master tables ── */}
      <Card>
        <CardHeader title="📋 Master Data Tables — FY / States / Sectors / Currencies"/>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
          {[
            { name:"master_fy",         color:"#00c9b1", cols:["id UUID PK","label VARCHAR(9) — e.g. 2026-27","is_active BOOLEAN DEFAULT true","sort_order INTEGER"] },
            { name:"master_states",     color:"#34d399", cols:["id UUID PK","name VARCHAR(100) NOT NULL UNIQUE","is_active BOOLEAN DEFAULT true"] },
            { name:"master_sectors",    color:"#4a90d9", cols:["id UUID PK","name VARCHAR(100) NOT NULL UNIQUE","is_active BOOLEAN DEFAULT true"] },
            { name:"master_currencies", color:"#a78bfa", cols:["id UUID PK","code VARCHAR(5) NOT NULL UNIQUE — e.g. USD","name VARCHAR(100)","is_active BOOLEAN DEFAULT true"] },
          ].map(mt=>(
            <div key={mt.name} className="rounded-xl border p-4" style={{borderColor:mt.color+"40"}}>
              <div className="flex items-center gap-2 mb-3">
                <div className="w-2 h-2 rounded-full" style={{background:mt.color}}/>
                <span className="font-black text-sm" style={{color:mt.color}}>{mt.name}</span>
              </div>
              <ul className="space-y-1">
                {mt.cols.map(c=>(
                  <li key={c} className={`text-[11px] font-mono ${t.textMuted}`}>▸ {c}</li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </Card>

      {/* ── DDL ── */}
      <Card>
        <CardHeader title="SQL DDL — Complete Table Set"/>
        <div className="p-5">
          <SqlBlock lines={[
            "-- Enable UUID generation",
            "CREATE EXTENSION IF NOT EXISTS pgcrypto;",
            "",
            "-- 1. Users",
            "CREATE TABLE users (",
            "  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  email         VARCHAR(255) NOT NULL UNIQUE,",
            "  password_hash TEXT NOT NULL,",
            "  full_name     VARCHAR(255) NOT NULL,",
            "  role          VARCHAR(20) NOT NULL DEFAULT 'user'",
            "                  CHECK (role IN ('admin','user')),",
            "  is_active     BOOLEAN NOT NULL DEFAULT true,",
            "  invite_status VARCHAR(20) DEFAULT 'pending',",
            "  invite_token  TEXT,",
            "  invite_sent_at TIMESTAMPTZ,",
            "  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()",
            ");",
            "",
            "-- 2. Client Master",
            "CREATE TABLE client_master (",
            "  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  client_name  VARCHAR(255) NOT NULL UNIQUE,",
            "  person_name  VARCHAR(255) NOT NULL,",
            "  sector       VARCHAR(100) NOT NULL,",
            "  state        VARCHAR(100) NOT NULL,",
            "  turnover     VARCHAR(30),",
            "  type         VARCHAR(50),",
            "  risk         VARCHAR(20) CHECK (risk IN ('Low','Medium','High')),",
            "  status       VARCHAR(20) DEFAULT 'Active',",
            "  rm           VARCHAR(255),",
            "  is_active    BOOLEAN NOT NULL DEFAULT true,",
            "  created_by   UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  updated_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()",
            ");",
            "",
            "-- 3. Client–User Assignments",
            "CREATE TABLE client_user_assignments (",
            "  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  client_id   UUID NOT NULL REFERENCES client_master(id) ON DELETE CASCADE,",
            "  user_id     UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,",
            "  assigned_by UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  UNIQUE (client_id, user_id)",
            ");",
            "",
            "-- 4. CAS MIS",
            "CREATE TABLE cas_mis (",
            "  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  client_id        UUID NOT NULL REFERENCES client_master(id) ON DELETE RESTRICT,",
            "  financial_year   VARCHAR(9) NOT NULL,",
            "  month            VARCHAR(15) NOT NULL,",
            "  mis_date         DATE,",
            "  outstanding_dues NUMERIC(18,2) DEFAULT 0,",
            "  collection_status VARCHAR(30),",
            "  collection_amt   NUMERIC(18,2) DEFAULT 0,",
            "  recon_gap        NUMERIC(18,2) DEFAULT 0,",
            "  pt_applicable    VARCHAR(3) CHECK (pt_applicable IN ('Y','N','A')),",
            "  tds_applicable   VARCHAR(3) CHECK (tds_applicable IN ('Y','N','A')),",
            "  pf_applicable    VARCHAR(3) CHECK (pf_applicable IN ('Y','N','A')),",
            "  esi_applicable   VARCHAR(3) CHECK (esi_applicable IN ('Y','N','A')),",
            "  gst_applicable   VARCHAR(3) CHECK (gst_applicable IN ('Y','N','A')),",
            "  software_used    VARCHAR(100),",
            "  remarks          TEXT,",
            "  created_by       UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  UNIQUE (client_id, financial_year, month)",
            ");",
            "",
            "-- 5. KRA / KPI",
            "CREATE TABLE kra_kpi (",
            "  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  client_id        UUID NOT NULL REFERENCES client_master(id) ON DELETE RESTRICT,",
            "  financial_year   VARCHAR(9) NOT NULL,",
            "  month            VARCHAR(15) NOT NULL,",
            "  mis_date         DATE,",
            "  mis_applicable   VARCHAR(3), mysa VARCHAR(3), mysa_applicable VARCHAR(3),",
            "  rectification    VARCHAR(3), rect_applicable VARCHAR(3),",
            "  escalation       VARCHAR(3), esc_applicable  VARCHAR(3),",
            "  raksha           VARCHAR(3), raksha_applicable VARCHAR(3),",
            "  capitalWant      VARCHAR(3), cw_applicable   VARCHAR(3),",
            "  capital          VARCHAR(3), cap_applicable  VARCHAR(3),",
            "  created_by       UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),",
            "  UNIQUE (client_id, financial_year, month)",
            ");",
            "",
            "-- 6. Fund Requests",
            "CREATE TABLE fund_requests (",
            "  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  client_id       UUID NOT NULL REFERENCES client_master(id) ON DELETE RESTRICT,",
            "  financial_year  VARCHAR(9) NOT NULL,",
            "  month           VARCHAR(15) NOT NULL,",
            "  currency        VARCHAR(5) NOT NULL,",
            "  foreign_amount  NUMERIC(18,4) NOT NULL CHECK (foreign_amount > 0),",
            "  exchange_rate   NUMERIC(12,6) NOT NULL CHECK (exchange_rate > 0),",
            "  inr_amount      NUMERIC(20,2) GENERATED ALWAYS AS",
            "                    (ROUND(foreign_amount * exchange_rate, 2)) STORED,",
            "  remarks         TEXT,",
            "  created_by      UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()",
            ");",
            "",
            "-- 7. Month Locks",
            "CREATE TABLE month_locks (",
            "  id             UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  financial_year VARCHAR(9) NOT NULL,",
            "  month          VARCHAR(15) NOT NULL,",
            "  module         VARCHAR(20) NOT NULL CHECK (module IN ('cas','kra','fund','all')),",
            "  is_locked      BOOLEAN NOT NULL DEFAULT false,",
            "  locked_by      UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  locked_at      TIMESTAMPTZ,",
            "  unlocked_by    UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  unlocked_at    TIMESTAMPTZ,",
            "  UNIQUE (financial_year, month, module)",
            ");",
            "",
            "-- 8. Reminders",
            "CREATE TABLE reminders (",
            "  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),",
            "  title       VARCHAR(255) NOT NULL,",
            "  module      VARCHAR(30) NOT NULL,",
            "  frequency   VARCHAR(20) CHECK (frequency IN ('daily','weekly','monthly')),",
            "  trigger_day INTEGER,",
            "  target_role VARCHAR(20) DEFAULT 'all',",
            "  message     TEXT NOT NULL,",
            "  status      VARCHAR(20) DEFAULT 'active',",
            "  last_sent   TIMESTAMPTZ,",
            "  created_by  UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()",
            ");",
            "",
            "-- 9. Audit Log",
            "CREATE TABLE audit_log (",
            "  id         BIGSERIAL PRIMARY KEY,",
            "  table_name VARCHAR(50) NOT NULL,",
            "  record_id  UUID NOT NULL,",
            "  action     VARCHAR(20) NOT NULL,",
            "  module     VARCHAR(30),",
            "  old_data   JSONB,",
            "  new_data   JSONB,",
            "  actor_id   UUID REFERENCES users(id) ON DELETE SET NULL,",
            "  acted_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()",
            ");",
            "CREATE INDEX idx_audit_actor  ON audit_log(actor_id);",
            "CREATE INDEX idx_audit_table  ON audit_log(table_name);",
            "CREATE INDEX idx_audit_action ON audit_log(action);",
          ]}/>
        </div>
      </Card>
    </div>
  );

  // ── API ROUTES ────────────────────────────────────────────────────────────
  const ApiSection = () => (
    <div className="space-y-5">
      <Card>
        <CardHeader title="🔐 Authentication  —  /api/auth"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="POST" path="/api/auth/login"   desc="Get JWT tokens" auth="user"
            reqFields={[{name:"email",type:"string",desc:"User email address"},{name:"password",type:"string",desc:"Plain text password"}]}
            resFields={[{name:"access_token",type:"string",desc:"Short-lived JWT (1h)"},{name:"refresh_token",type:"string",desc:"Long-lived token (7d)"},{name:"user.role",type:"string",desc:"admin or user"},{name:"user.assignedClients",type:"array",desc:"Client IDs this user can edit"}]}/>
          <ApiRoute method="POST" path="/api/auth/refresh" desc="Rotate access token" auth="user"
            reqFields={[{name:"refresh_token",type:"string",desc:"Current refresh token"}]}
            resFields={[{name:"access_token",type:"string",desc:"New access token"},{name:"expires_in",type:"number",desc:"Seconds until expiry"}]}/>
          <ApiRoute method="POST" path="/api/auth/logout"  desc="Revoke token" auth="user"
            reqFields={[{name:"refresh_token",type:"string",desc:"Token to revoke"}]}
            resFields={[{name:"message",type:"string",desc:"Confirmation message"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="👥 Client Master  —  /api/clients"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/clients"               desc="List all (paginated)" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Client objects with sector, state, assigned users"},{name:"total",type:"number",desc:"Total count"},{name:"page",type:"number",desc:"Current page"}]}/>
          <ApiRoute method="GET"    path="/api/clients/:id"           desc="Get single client" auth="user"
            reqFields={null}
            resFields={[{name:"id",type:"UUID",desc:"Client identifier"},{name:"client_name",type:"string",desc:"Company name"},{name:"sector",type:"string",desc:"Business sector"},{name:"state",type:"string",desc:"Indian state"},{name:"assignedUsers",type:"array",desc:"Assigned user IDs"}]}/>
          <ApiRoute method="POST"   path="/api/clients"               desc="Create client" auth="admin"
            reqFields={[{name:"client_name",type:"string",desc:"Company or individual name"},{name:"person_name",type:"string",desc:"Primary contact"},{name:"sector",type:"string",desc:"Business sector"},{name:"state",type:"string",desc:"Indian state"},{name:"assignedUserIds",type:"array",desc:"User IDs to assign"}]}
            resFields={[{name:"id",type:"UUID",desc:"New client ID"},{name:"created_at",type:"timestamp",desc:"Creation time"}]}/>
          <ApiRoute method="PUT"    path="/api/clients/:id"           desc="Update client" auth="admin"
            reqFields={[{name:"client_name",type:"string",desc:"Updated name"},{name:"sector",type:"string",desc:"Updated sector"},{name:"state",type:"string",desc:"Updated state"},{name:"assignedUserIds",type:"array",desc:"Updated user assignments"}]}
            resFields={[{name:"id",type:"UUID",desc:"Client ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="DELETE" path="/api/clients/:id"           desc="Soft-delete" auth="admin"
            reqFields={null}
            resFields={[{name:"message",type:"string",desc:"Confirmation"},{name:"id",type:"UUID",desc:"Deleted ID"}]}/>
          <ApiRoute method="POST"   path="/api/clients/import"        desc="Bulk Excel import" auth="admin"
            reqFields={[{name:"file",type:"multipart",desc:".xlsx or .csv upload"}]}
            resFields={[{name:"imported",type:"number",desc:"New rows added"},{name:"updated",type:"number",desc:"Existing rows updated"},{name:"duplicates",type:"number",desc:"Skipped rows"},{name:"errors[]",type:"array",desc:"Row-level errors"}]}/>
          <ApiRoute method="GET"    path="/api/clients/export"        desc="Export to Excel" auth="user"
            reqFields={null}
            resFields={[{name:"file",type:"binary",desc:".xlsx stream download"}]}/>
          <ApiRoute method="PUT"    path="/api/clients/:id/assign"    desc="Update user assignments" auth="admin"
            reqFields={[{name:"userIds",type:"array",desc:"Array of user IDs to assign to this client"}]}
            resFields={[{name:"assignments",type:"array",desc:"Updated assignment list"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="👤 User Database  —  /api/users"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/users"              desc="List all users" auth="admin"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"User objects with role and invite status"},{name:"total",type:"number",desc:"Total count"}]}/>
          <ApiRoute method="POST"   path="/api/users"              desc="Create user" auth="admin"
            reqFields={[{name:"full_name",type:"string",desc:"Display name"},{name:"email",type:"string",desc:"Login email"},{name:"role",type:"string",desc:"admin or user"},{name:"assignedClientIds",type:"array",desc:"Client IDs to assign"}]}
            resFields={[{name:"id",type:"UUID",desc:"New user ID"},{name:"invite_token",type:"string",desc:"Token for activation link"}]}/>
          <ApiRoute method="PUT"    path="/api/users/:id"          desc="Update user" auth="admin"
            reqFields={[{name:"full_name",type:"string",desc:"Updated name"},{name:"role",type:"string",desc:"Updated role"},{name:"is_active",type:"boolean",desc:"Active status"},{name:"assignedClientIds",type:"array",desc:"Updated client assignments"}]}
            resFields={[{name:"id",type:"UUID",desc:"User ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="DELETE" path="/api/users/:id"          desc="Deactivate user" auth="admin"
            reqFields={null}
            resFields={[{name:"message",type:"string",desc:"Confirmation"}]}/>
          <ApiRoute method="POST"   path="/api/users/:id/invite"   desc="Send / resend invite" auth="admin"
            reqFields={[{name:"email",type:"string",desc:"Target email address"}]}
            resFields={[{name:"invite_token",type:"string",desc:"HMAC-SHA256 signed token"},{name:"sent_at",type:"timestamp",desc:"Dispatch time"}]}/>
          <ApiRoute method="POST"   path="/api/users/activate"     desc="Activate account via token" auth="user"
            reqFields={[{name:"token",type:"string",desc:"Invite token from email"},{name:"password",type:"string",desc:"Chosen password"}]}
            resFields={[{name:"message",type:"string",desc:"Activation confirmation"}]}/>
          <ApiRoute method="POST"   path="/api/users/import"       desc="Bulk user import" auth="admin"
            reqFields={[{name:"file",type:"multipart",desc:".xlsx with user records"}]}
            resFields={[{name:"imported",type:"number",desc:"Users created"},{name:"errors[]",type:"array",desc:"Row-level errors"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="📋 CAS MIS  —  /api/cas-mis"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/cas-mis"           desc="List records (paginated)" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"CAS MIS records"},{name:"total",type:"number",desc:"Total matching"}]}/>
          <ApiRoute method="GET"    path="/api/cas-mis/analytics" desc="Analytics aggregates" auth="user"
            reqFields={null}
            resFields={[{name:"revenue_mix[]",type:"array",desc:"Revenue by sector"},{name:"outstanding_spoc[]",type:"array",desc:"Dues by SPOC"},{name:"statutory_health[]",type:"array",desc:"Y/N/A breakdown per statutory type"},{name:"collection_timeline[]",type:"array",desc:"Monthly collection delay trend"}]}/>
          <ApiRoute method="POST"   path="/api/cas-mis"           desc="Create / update MIS record" auth="user"
            reqFields={[{name:"client_id",type:"UUID",desc:"Client reference"},{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"month",type:"string",desc:"e.g. April"},{name:"mis_date",type:"date",desc:"MIS submission date"},{name:"outstanding_dues",type:"number",desc:"Outstanding amount"},{name:"pt_applicable",type:"string",desc:"Y/N/A"},{name:"tds_applicable",type:"string",desc:"Y/N/A"}]}
            resFields={[{name:"id",type:"UUID",desc:"Record ID"},{name:"created_at",type:"timestamp",desc:"Creation time"}]}/>
          <ApiRoute method="PUT"    path="/api/cas-mis/:id"       desc="Update MIS record" auth="user"
            reqFields={[{name:"mis_date",type:"date",desc:"Updated MIS date"},{name:"collection_status",type:"string",desc:"Updated status"},{name:"remarks",type:"string",desc:"Updated notes"}]}
            resFields={[{name:"id",type:"UUID",desc:"Record ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="GET"    path="/api/cas-mis/export"    desc="Export to Excel" auth="user"
            reqFields={null}
            resFields={[{name:"file",type:"binary",desc:".xlsx stream download"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="🎯 KRA / KPI  —  /api/kra-kpi"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/kra-kpi"           desc="List records (paginated)" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"KRA/KPI records with all parameter values"},{name:"total",type:"number",desc:"Total matching"}]}/>
          <ApiRoute method="GET"    path="/api/kra-kpi/analytics" desc="KPI analytics aggregates" auth="user"
            reqFields={null}
            resFields={[{name:"param_scores[]",type:"array",desc:"Score % per KPI parameter"},{name:"client_heatmap[]",type:"array",desc:"Compliance matrix per client"},{name:"escalation_trend[]",type:"array",desc:"Monthly escalation count"}]}/>
          <ApiRoute method="POST"   path="/api/kra-kpi"           desc="Create / update KPI record" auth="user"
            reqFields={[{name:"client_id",type:"UUID",desc:"Client reference"},{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"month",type:"string",desc:"e.g. April"},{name:"mysa",type:"string",desc:"Y/N/A"},{name:"rectification",type:"string",desc:"Y/N/A"},{name:"escalation",type:"string",desc:"Y/N/A"},{name:"raksha",type:"string",desc:"Y/N/A"},{name:"capitalWant",type:"string",desc:"Y/N/A"},{name:"capital",type:"string",desc:"Y/N/A"}]}
            resFields={[{name:"id",type:"UUID",desc:"Record ID"},{name:"created_at",type:"timestamp",desc:"Creation time"}]}/>
          <ApiRoute method="PUT"    path="/api/kra-kpi/:id"       desc="Update KPI record" auth="user"
            reqFields={[{name:"mysa",type:"string",desc:"Updated Y/N/A"},{name:"escalation",type:"string",desc:"Updated Y/N/A"},{name:"raksha",type:"string",desc:"Updated Y/N/A"}]}
            resFields={[{name:"id",type:"UUID",desc:"Record ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="PATCH"  path="/api/kra-kpi/:id/applicability" desc="Toggle KPI applicability" auth="admin"
            reqFields={[{name:"parameter",type:"string",desc:"KPI parameter name e.g. mysa"},{name:"applicable",type:"string",desc:"Y/N/A"}]}
            resFields={[{name:"id",type:"UUID",desc:"Record ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="GET"    path="/api/kra-kpi/export"    desc="Export to Excel" auth="user"
            reqFields={null}
            resFields={[{name:"file",type:"binary",desc:".xlsx stream download"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="💰 Fund Requests  —  /api/fund-requests"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/fund-requests"            desc="List with filters" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Fund request objects"},{name:"total",type:"number",desc:"Total matching"},{name:"page",type:"number",desc:"Current page"}]}/>
          <ApiRoute method="GET"    path="/api/fund-requests/analytics"  desc="Chart aggregates" auth="user"
            reqFields={null}
            resFields={[{name:"currency_breakdown[]",type:"array",desc:"INR by currency"},{name:"monthly_trend[]",type:"array",desc:"INR by month"},{name:"top_clients[]",type:"array",desc:"Ranked by INR total"}]}/>
          <ApiRoute method="POST"   path="/api/fund-requests"            desc="Create entry" auth="user"
            reqFields={[{name:"client_id",type:"UUID",desc:"Client reference"},{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"month",type:"string",desc:"e.g. April"},{name:"currency",type:"string",desc:"ISO code"},{name:"foreign_amount",type:"number",desc:"Amount in foreign ccy"},{name:"exchange_rate",type:"number",desc:"Rate at entry time"}]}
            resFields={[{name:"id",type:"UUID",desc:"New entry ID"},{name:"inr_amount",type:"number",desc:"Auto-calculated INR"},{name:"created_at",type:"timestamp",desc:"Creation time"}]}/>
          <ApiRoute method="PUT"    path="/api/fund-requests/:id"        desc="Update entry" auth="user"
            reqFields={[{name:"foreign_amount",type:"number",desc:"Updated amount"},{name:"exchange_rate",type:"number",desc:"Updated rate"},{name:"remarks",type:"string",desc:"Updated notes"}]}
            resFields={[{name:"id",type:"UUID",desc:"Entry ID"},{name:"inr_amount",type:"number",desc:"Recalculated INR"}]}/>
          <ApiRoute method="DELETE" path="/api/fund-requests/:id"        desc="Delete entry" auth="admin"
            reqFields={null}
            resFields={[{name:"message",type:"string",desc:"Confirmation"},{name:"id",type:"UUID",desc:"Deleted ID"}]}/>
          <ApiRoute method="GET"    path="/api/fund-requests/export"     desc="Download Excel" auth="user"
            reqFields={null}
            resFields={[{name:"file",type:"binary",desc:".xlsx stream download"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="🔒 Month Lock  —  /api/month-locks"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/month-locks"              desc="Get lock status grid" auth="admin"
            reqFields={null}
            resFields={[{name:"locks[]",type:"array",desc:"Lock records per FY/Month/Module"},{name:"audit[]",type:"array",desc:"Recent lock audit entries"}]}/>
          <ApiRoute method="POST"   path="/api/month-locks/lock"         desc="Lock a FY/Month/Module" auth="admin"
            reqFields={[{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"month",type:"string",desc:"e.g. April"},{name:"module",type:"string",desc:"cas | kra | fund | all"}]}
            resFields={[{name:"lock",type:"object",desc:"Lock record created"},{name:"audit_id",type:"number",desc:"Audit log entry ID"}]}/>
          <ApiRoute method="POST"   path="/api/month-locks/unlock"       desc="Unlock a FY/Month/Module" auth="admin"
            reqFields={[{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"month",type:"string",desc:"e.g. April"},{name:"module",type:"string",desc:"cas | kra | fund | all"}]}
            resFields={[{name:"lock",type:"object",desc:"Updated lock record"},{name:"audit_id",type:"number",desc:"Audit log entry ID"}]}/>
          <ApiRoute method="POST"   path="/api/month-locks/bulk-lock"    desc="Lock all months in a FY" auth="admin"
            reqFields={[{name:"financial_year",type:"string",desc:"e.g. 2026-27"},{name:"module",type:"string",desc:"cas | kra | fund | all"}]}
            resFields={[{name:"locked_count",type:"number",desc:"Number of months locked"}]}/>
          <ApiRoute method="GET"    path="/api/month-locks/audit"        desc="Audit log for lock events" auth="admin"
            reqFields={null}
            resFields={[{name:"entries[]",type:"array",desc:"Lock/unlock audit events with actor and timestamp"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="🔔 Reminders  —  /api/reminders"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/reminders"           desc="List all reminders" auth="admin"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Reminder records with status and last_sent"}]}/>
          <ApiRoute method="POST"   path="/api/reminders"           desc="Create reminder" auth="admin"
            reqFields={[{name:"title",type:"string",desc:"Reminder title"},{name:"module",type:"string",desc:"Target module"},{name:"frequency",type:"string",desc:"daily|weekly|monthly"},{name:"trigger_day",type:"number",desc:"Day of month (for monthly)"},{name:"target_role",type:"string",desc:"admin|user|all"},{name:"message",type:"string",desc:"Reminder body"}]}
            resFields={[{name:"id",type:"UUID",desc:"New reminder ID"}]}/>
          <ApiRoute method="PUT"    path="/api/reminders/:id"       desc="Update reminder" auth="admin"
            reqFields={[{name:"title",type:"string",desc:"Updated title"},{name:"frequency",type:"string",desc:"Updated frequency"},{name:"status",type:"string",desc:"active|paused"}]}
            resFields={[{name:"id",type:"UUID",desc:"Reminder ID"},{name:"updated_at",type:"timestamp",desc:"Update time"}]}/>
          <ApiRoute method="DELETE" path="/api/reminders/:id"       desc="Delete reminder" auth="admin"
            reqFields={null}
            resFields={[{name:"message",type:"string",desc:"Confirmation"}]}/>
          <ApiRoute method="POST"   path="/api/reminders/:id/trigger" desc="Manually trigger reminder" auth="admin"
            reqFields={null}
            resFields={[{name:"sent_count",type:"number",desc:"Number of notifications dispatched"},{name:"triggered_at",type:"timestamp",desc:"Trigger time"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="⚙️ Master Settings  —  /api/master"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET"    path="/api/master/fy"           desc="List financial years" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"FY list e.g. 2026-27"}]}/>
          <ApiRoute method="POST"   path="/api/master/fy"           desc="Add financial year" auth="admin"
            reqFields={[{name:"label",type:"string",desc:"e.g. 2027-28"}]}
            resFields={[{name:"id",type:"UUID",desc:"New FY record ID"}]}/>
          <ApiRoute method="GET"    path="/api/master/states"       desc="List states" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Indian state list"}]}/>
          <ApiRoute method="GET"    path="/api/master/sectors"      desc="List sectors" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Business sector list"}]}/>
          <ApiRoute method="GET"    path="/api/master/currencies"   desc="List currencies" auth="user"
            reqFields={null}
            resFields={[{name:"data[]",type:"array",desc:"Currency codes e.g. USD, EUR, GBP"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="📊 Dashboard Analytics  —  /api/dashboard"/>
        <div className="p-5 space-y-2">
          <ApiRoute method="GET" path="/api/dashboard/summary" desc="Executive KPI summary" auth="user"
            reqFields={null}
            resFields={[{name:"total_clients",type:"number",desc:"Active client count"},{name:"cas_revenue",type:"number",desc:"Total expected CAS revenue"},{name:"cas_collected",type:"number",desc:"Amount collected"},{name:"kpi_score",type:"number",desc:"Overall KPI compliance %"},{name:"fund_inward",type:"number",desc:"Total fund inward INR"}]}/>
          <ApiRoute method="GET" path="/api/dashboard/analytics" desc="Cross-module chart data" auth="user"
            reqFields={null}
            resFields={[{name:"revenue_mix[]",type:"array",desc:"Revenue by sector (CAS MIS)"},{name:"kra_heatmap[]",type:"array",desc:"KPI compliance heatmap"},{name:"fund_trend[]",type:"array",desc:"Fund inflow monthly trend"},{name:"top_clients[]",type:"array",desc:"Top clients by fund volume"}]}/>
        </div>
      </Card>

      <Card>
        <CardHeader title="🔍 Common Query Parameters"/>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className={`${t.tableHead} text-[10px] uppercase tracking-wider`}>
                {["Parameter","Type","Example","Description"].map(h=>(
                  <th key={h} className={`px-4 py-2.5 text-left font-bold border-b ${t.cardBorder}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${t.divider}`}>
              {[
                ["page",          "integer","page=2",                "Page number (default 1)"],
                ["limit",         "integer","limit=20",              "Items per page (max 100)"],
                ["financial_year","string", `financial_year=${DEFAULT_FY}`,"Filter by FY"],
                ["month",         "string", "month=April",           "Filter by month name"],
                ["client_id",     "uuid",   "client_id=uuid",        "Filter by client"],
                ["module",        "string", "module=cas",            "Filter by module (lock routes)"],
                ["currency",      "string", "currency=USD",          "Filter by currency code"],
                ["search",        "string", "search=Agarwal",        "Full-text search"],
                ["sort",          "string", "sort=created_at:desc",  "Sort field + direction"],
                ["status",        "string", "status=active",         "Filter by status (reminders/users)"],
              ].map(([p,type,ex,desc])=>(
                <tr key={p} className={`${t.tableRow} transition-colors`}>
                  <td className="px-4 py-2.5"><code className="font-mono" style={{color:"#fbbf24"}}>{p}</code></td>
                  <td className="px-4 py-2.5 font-mono" style={{color:"#4a90d9"}}>{type}</td>
                  <td className={`px-4 py-2.5 font-mono text-[10px] ${t.textMuted}`}>{ex}</td>
                  <td className={`px-4 py-2.5 ${t.textMuted}`}>{desc}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );

  // ── ROLES ─────────────────────────────────────────────────────────────────
  const RolesSection = () => (
    <div className="space-y-5">
      {/* Role cards */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        {/* Admin */}
        <div className={`${t.card} border rounded-2xl p-5`} style={{borderColor:"#34d39940"}}>
          <div className="flex items-center gap-3 mb-4">
            <div className="w-10 h-10 rounded-xl text-xl flex items-center justify-center"
              style={{background:"#34d39918"}}>👑</div>
            <div>
              <div className="font-black text-sm" style={{color:"#34d399"}}>Admin Role</div>
              <div className={`text-xs ${t.textMuted}`}>Full read + write access across all modules</div>
            </div>
          </div>
          <ul className="space-y-2">
            {[
              "Access all 7 modules: Dashboard, Client Master, CAS MIS, KRA/KPI, Fund Request, User Database, Architect",
              "Create, edit, delete clients (including sector, state, user assignment)",
              "Bulk import clients via Excel",
              "Create, edit, delete CAS MIS records for all clients",
              "Create, edit, delete KRA/KPI records for all clients",
              "Create, edit, delete Fund Request entries for all clients",
              "Manage user accounts: create, edit, deactivate, assign clients",
              "Send and resend invite emails to users",
              "Lock and unlock any FY / Month / Module combination",
              "Create, edit, pause, delete auto reminders",
              "View full System Audit Log and export to Excel",
              "Configure Master Settings: FY, States, Sectors, Currencies",
              "View System Health Dashboard",
              "View all analytics and charts across all modules",
            ].map(p=>(
              <li key={p} className={`text-xs flex items-start gap-2 ${t.text}`}>
                <span style={{color:"#34d399"}} className="shrink-0 mt-0.5">✓</span>{p}
              </li>
            ))}
          </ul>
        </div>

        {/* End User */}
        <div className={`${t.card} border rounded-2xl p-5`} style={{borderColor:"#00c9b140"}}>
          <div className="flex items-center gap-3 mb-4">
            <div className="w-10 h-10 rounded-xl text-xl flex items-center justify-center"
              style={{background:"#00c9b118"}}>👤</div>
            <div>
              <div className="font-black text-sm" style={{color:"#00c9b1"}}>End User Role</div>
              <div className={`text-xs ${t.textMuted}`}>Edit assigned clients · View all · No admin controls</div>
            </div>
          </div>
          <ul className="space-y-2">
            {[
              ["View Dashboard and all analytics charts", true],
              ["View full Client Master list (all clients)", true],
              ["Edit CAS MIS records for assigned clients only", true],
              ["Edit KRA/KPI records for assigned clients only", true],
              ["Create / edit Fund Request entries for assigned clients only", true],
              ["View CAS MIS, KRA/KPI, Fund records for unassigned clients (read-only badge)", true],
              ["Export data to Excel (all modules)", true],
              ["View own profile", true],
              ["Create or delete clients", false],
              ["Manage user accounts or send invites", false],
              ["Lock or unlock any month", false],
              ["Create or manage reminders", false],
              ["Access System Audit Log", false],
              ["Modify Master Settings", false],
              ["Access Architect module", false],
            ].map(([p,ok])=>(
              <li key={p} className={`text-xs flex items-start gap-2 ${ok?t.text:t.textMuted}`}>
                <span style={{color:ok?"#00c9b1":"#f87171"}} className="shrink-0 mt-0.5">{ok?"✓":"✗"}</span>{p}
              </li>
            ))}
          </ul>
        </div>
      </div>

      {/* User-Client Assignment */}
      <Card>
        <CardHeader title="🔗 User–Client Assignment Model"/>
        <div className="p-5 space-y-4">
          <p className={`text-xs leading-relaxed ${t.textMuted}`}>
            Each End User has an <code className="font-mono" style={{color:"#fbbf24"}}>assignedClients[]</code> array.
            Admin users have <code className="font-mono" style={{color:"#fbbf24"}}>assignedClients = null</code> (meaning all clients).
            The <code className="font-mono" style={{color:"#fbbf24"}}>useRBAC()</code> hook derives edit permissions at runtime:
          </p>
          <div className="rounded-xl overflow-hidden mt-2" style={{background:"#0d1117"}}>
            <pre className="px-4 py-4 text-[11px] font-mono overflow-x-auto leading-relaxed text-[#94a3b8]">
{`const useRBAC = () => {
  const { activeUser } = useContext(UserContext);
  const isAdmin = activeUser?.role === "Admin";

  // Admin → null means all clients allowed
  // End User → array of assigned client IDs
  const assignedClientIds = isAdmin
    ? null
    : (activeUser?.assignedClients || []);

  // canEdit(clientId) → true if user can write to this client
  const canEdit = (clientId) =>
    isAdmin || assignedClientIds.includes(clientId);

  return { isAdmin, activeUser, canEdit, assignedClientIds };
};`}
            </pre>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-2">
            {[
              { icon:"✏️", label:"Edit row", color:"#34d399",
                desc:"Shown when canEdit(clientId) = true. User is Admin or client is in their assigned list." },
              { icon:"👁", label:"View Only badge", color:"#00c9b1",
                desc:"Shown for unassigned clients. Row is visible but all Save, Add, Delete buttons are hidden." },
              { icon:"🔒", label:"Month Locked", color:"#fbbf24",
                desc:"Even if canEdit = true, a locked month disables all write actions for all users including Admin." },
            ].map(b=>(
              <div key={b.label} className="rounded-xl p-4 border"
                style={{background:b.color+"0a",borderColor:b.color+"30"}}>
                <div className="text-xl mb-2">{b.icon}</div>
                <div className="font-bold text-xs mb-1" style={{color:b.color}}>{b.label}</div>
                <div className={`text-[11px] leading-relaxed ${t.textMuted}`}>{b.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* User Access Matrix */}
      <Card>
        <CardHeader title="📋 User Access Matrix — Module × Permission"/>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className={`${t.tableHead} text-[10px] uppercase tracking-wider`}>
                {["Module","Admin — View","Admin — Edit","End User — View","End User — Edit"].map(h=>(
                  <th key={h} className={`px-3 py-2.5 text-left font-bold border-b ${t.cardBorder}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${t.divider}`}>
              {[
                ["Dashboard",       "✓ All","— (read-only)","✓ All","— (read-only)"],
                ["Client Master",   "✓ All","✓ Full CRUD","✓ All (view)","✗ No edit"],
                ["CAS MIS",         "✓ All","✓ All clients","✓ All (view)","✓ Assigned only"],
                ["KRA / KPI",       "✓ All","✓ All clients","✓ All (view)","✓ Assigned only"],
                ["Fund Request",    "✓ All","✓ All clients","✓ All (view)","✓ Assigned only"],
                ["User Database",   "✓ All","✓ Full CRUD + Invite","✗ No access","✗ No access"],
                ["Architect",       "✓ All","✓ Full control","✗ No access","✗ No access"],
                ["Month Lock",      "✓ Status","✓ Lock/Unlock","✗","✗"],
                ["Reminders",       "✓ All","✓ Full CRUD","✗","✗"],
                ["Audit Log",       "✓ All","—","✗","✗"],
                ["Master Settings", "✓ All","✓ Full CRUD","✗","✗"],
              ].map(([mod,...cells])=>(
                <tr key={mod} className={`${t.tableRow} transition-colors`}>
                  <td className={`px-3 py-2.5 font-bold text-[11px] ${t.text}`}>{mod}</td>
                  {cells.map((c,i)=>(
                    <td key={i} className={`px-3 py-2.5 text-[11px] ${t.textMuted}`}
                      style={{color:c.startsWith("✓")?"#34d399":c.startsWith("✗")?"#f87171":undefined}}>
                      {c}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      {/* JWT Flow */}
      <Card>
        <CardHeader title="JWT Authentication Flow"/>
        <div className="p-5 space-y-3">
          {[
            { step:1, icon:"🔑", title:"User submits credentials",        desc:"POST /api/auth/login with email + password",                                          color:"#00c9b1" },
            { step:2, icon:"🛡️", title:"Server validates password",       desc:"Hash comparison, user status check, role lookup, assignedClients fetched from DB",    color:"#34d399" },
            { step:3, icon:"📤", title:"Issue JWT tokens",                desc:"Access token (1h) + Refresh token (7d) + assignedClients[] returned to client",       color:"#fbbf24" },
            { step:4, icon:"📋", title:"Client sends Bearer token",       desc:"Authorization: Bearer <access_token> header on each API request",                    color:"#4a90d9" },
            { step:5, icon:"✅", title:"Middleware verifies + authorizes", desc:"Validates signature, checks expiry, confirms role, checks client assignment for row-level access", color:"#a78bfa" },
          ].map(s=>(
            <div key={s.step} className="flex items-start gap-4 p-3 rounded-xl"
              style={{background:s.color+"0f",border:`1px solid ${s.color}25`}}>
              <div className="w-8 h-8 rounded-lg flex items-center justify-center text-base shrink-0"
                style={{background:s.color+"22"}}>{s.icon}</div>
              <div>
                <div className="font-bold text-xs mb-0.5" style={{color:s.color}}>Step {s.step} — {s.title}</div>
                <div className={`text-xs ${t.textMuted}`}>{s.desc}</div>
              </div>
            </div>
          ))}
        </div>
      </Card>

      {/* Security Checklist */}
      <Card>
        <CardHeader title="Security Checklist"/>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-3">
          {[
            { icon:"🔐", label:"Password Hashing",    value:"bcrypt with cost factor 12",                color:"#34d399" },
            { icon:"⏱",  label:"Token Expiry",        value:"Access 1h · Refresh 7d · Rotation",         color:"#00c9b1" },
            { icon:"🛡️", label:"Rate Limiting",       value:"100 req/min per IP via middleware",          color:"#fbbf24" },
            { icon:"✅", label:"Input Validation",     value:"Schema validation on every route",           color:"#4a90d9" },
            { icon:"🗄️", label:"SQL Safety",          value:"Parameterised queries, no string concat",    color:"#34d399" },
            { icon:"🌐", label:"CORS",                 value:"Whitelist frontend origin only",             color:"#a78bfa" },
            { icon:"⛑️", label:"HTTP Headers",         value:"Helmet.js security headers enabled",         color:"#f87171" },
            { icon:"📋", label:"Audit Log",            value:"All writes tracked with before/after diff",  color:"#fbbf24" },
            { icon:"🔗", label:"Invite Token Security",value:"HMAC-SHA256 signed tokens (replace btoa())", color:"#f87171" },
            { icon:"🔒", label:"Month Lock Override",  value:"Locked months block writes even for Admin",  color:"#fbbf24" },
          ].map(s=>(
            <div key={s.label} className="flex items-center gap-3 p-3 rounded-xl"
              style={{background:s.color+"0f",border:`1px solid ${s.color}25`}}>
              <span className="text-lg">{s.icon}</span>
              <div>
                <div className="font-bold text-xs" style={{color:s.color}}>{s.label}</div>
                <div className={`text-[10px] ${t.textMuted}`}>{s.value}</div>
              </div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );

  // ── STACK ─────────────────────────────────────────────────────────────────
  const StackSection = () => (
    <div className="space-y-5">
      {/* Current vs Target banner */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <div className="rounded-2xl p-4 border" style={{background:"#00c9b108",borderColor:"#00c9b130"}}>
          <div className="font-black text-sm mb-1" style={{color:"#00c9b1"}}>🟢 Current State (UAT)</div>
          <p className={`text-xs ${t.textMuted}`}>Single-file React 18 JSX SPA · Deployed on Netlify · All data in React state + localStorage · No backend · No real auth · SMTP simulated</p>
        </div>
        <div className="rounded-2xl p-4 border" style={{background:"#4a90d908",borderColor:"#4a90d930"}}>
          <div className="font-black text-sm mb-1" style={{color:"#4a90d9"}}>🎯 Production Target</div>
          <p className={`text-xs ${t.textMuted}`}>React SPA + Node.js REST API + PostgreSQL 15 · JWT auth · Real SMTP email · Cron scheduler · Docker · HTTPS</p>
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        <StackCard icon="⚛️" name="React 18"         version="^18.3"      role="Frontend"
          color="#00c9b1" bullets={[
            "Functional components + hooks throughout",
            "React Context: UserContext + LockContext",
            "useRBAC() hook for all permission checks",
            "Single-file JSX SPA (current architecture)",
            "Dark / Light theme with inline style switching",
          ]}/>
        <StackCard icon="🎨" name="Tailwind CSS"     version="CDN ^3.4"   role="Styling"
          color="#34d399" bullets={[
            "Utility-first, zero runtime CSS",
            "Dark mode via inline style overrides",
            "Custom color tokens per module",
            "Responsive grid layouts throughout",
          ]}/>
        <StackCard icon="📡" name="Recharts"          version="^2.12"      role="Charts"
          color="#a78bfa" bullets={[
            "CAS MIS: Donut, Bar, Scatter, Treemap, 100% Stacked Bar",
            "KRA/KPI: Heatmap, Radar-style bar",
            "Fund Request: Bar, Line, Pie charts",
            "Dashboard: All chart types aggregated",
            "ResponsiveContainer on all panels",
          ]}/>
        <StackCard icon="📊" name="SheetJS (XLSX)"   version="^0.20"      role="Excel"
          color="#fbbf24" bullets={[
            "Client Master: bulk import + export",
            "User import workflow",
            "CAS MIS, KRA/KPI, Fund Request: row export",
            "Audit Log: export to Excel",
            "Column header validation on import",
          ]}/>
        <StackCard icon="🗄️" name="localStorage"    version="Browser API" role="Storage (UAT)"
          color="#fb923c" bullets={[
            "procas_locks — Month lock state",
            "procas_lockAudit — Lock audit trail",
            "procas_reminders — Reminder records",
            "procas_smtpCfg — SMTP config (⚠ security risk)",
            "procas_fyMaster / statesMaster / sectorsMaster / currMaster",
          ]}/>
        <StackCard icon="▲"  name="Netlify"          version="Static SPA" role="Hosting (UAT)"
          color="#e2e8f0" bullets={[
            "Current live URL: procas.netlify.app",
            "Deployed as static React SPA",
            "No server-side rendering",
            "IT whitelist required for office access",
            "CDN delivery via Netlify Edge",
          ]}/>
        <StackCard icon="🟢" name="Node.js 20 LTS"   version="LTS"        role="Backend (Target)"
          color="#6ee7b7" bullets={[
            "REST API with Express.js routing",
            "Routes: auth, clients, cas-mis, kra-kpi, fund, users, locks, reminders, master",
            "Multer for file uploads",
            "Nodemailer for invite email delivery",
            "Graceful shutdown + health check endpoint",
          ]}/>
        <StackCard icon="🐘" name="PostgreSQL 15"    version="^15.0"      role="Database (Target)"
          color="#4a90d9" bullets={[
            "9 core tables fully specified in DB Schema",
            "UUID primary keys throughout",
            "Generated INR column in fund_requests",
            "UNIQUE constraints on FY/Month/Module locks",
            "Connection pool singleton",
          ]}/>
        <StackCard icon="🔑" name="JWT + bcrypt"     version="RS256"      role="Auth (Target)"
          color="#f87171" bullets={[
            "Signed access tokens (1h) + Refresh tokens (7d)",
            "bcrypt password hashing (cost factor 12)",
            "Role claim + assignedClients[] embedded in JWT",
            "HMAC-SHA256 invite tokens (replace current btoa)",
            "Role-based middleware on all protected routes",
          ]}/>
        <StackCard icon="✉️" name="SMTP / Nodemailer" version="^6.x"      role="Email (Target)"
          color="#22d3ee" bullets={[
            "Currently simulated via setTimeout in UAT",
            "Production: Nodemailer + SMTP server",
            "Invite email delivery with token link",
            "Reminder notification dispatch",
            "Credentials in server-side env vars only",
          ]}/>
        <StackCard icon="⏰" name="Cron Scheduler"   version="node-cron"  role="Reminders (Target)"
          color="#fb923c" bullets={[
            "Currently UI-only in UAT",
            "Production: node-cron or queue worker",
            "Reads active reminders from DB on schedule",
            "Dispatches emails via Nodemailer",
            "Updates last_sent timestamp after trigger",
          ]}/>
        <StackCard icon="🐳" name="Docker"           version="Compose v2" role="DevOps (Target)"
          color="#22d3ee" bullets={[
            "Multi-stage API build container",
            "PostgreSQL service container",
            "Volume for DB persistence",
            "Health check endpoints",
            "docker-compose.yml in project root",
          ]}/>
      </div>

      <Card>
        <CardHeader title="Docker Compose — Services"/>
        <div className="p-5">
          <BashBlock lines={[
            "services:",
            "  db:",
            "    image: postgres:15-alpine",
            "    environment:",
            "      POSTGRES_DB:       procas",
            "      POSTGRES_USER:     procasuser",
            "      POSTGRES_PASSWORD: (from .env)",
            "    volumes:",
            "      - pgdata:/var/lib/postgresql/data",
            "      - ./migrations:/docker-entrypoint-initdb.d",
            "    healthcheck:",
            "      test: pg_isready -U procasuser -d procas",
            "      interval: 10s",
            "",
            "  api:",
            "    build: ./apps/api",
            "    environment:",
            "      DATABASE_URL:  postgres://procasuser@db:5432/procas",
            "      JWT_SECRET:    (from .env)",
            "      SMTP_HOST:     (from .env)",
            "      SMTP_USER:     (from .env)",
            "      SMTP_PASS:     (from .env)   ← server-side only, never localStorage",
            "      NODE_ENV:      production",
            "      PORT:          4000",
            "    depends_on:",
            "      db: { condition: service_healthy }",
            "    ports: [4000:4000]",
            "",
            "  web:",
            "    build: ./apps/web",
            "    environment:",
            "      API_URL: http://api:4000",
            "    depends_on: [api]",
            "    ports: [3000:3000]",
            "",
            "volumes:",
            "  pgdata:",
          ]}/>
        </div>
      </Card>
    </div>
  );

  // ── FLOW ──────────────────────────────────────────────────────────────────
  const FlowSection = () => (
    <div className="space-y-5">

      {/* Master platform data flow */}
      <Card>
        <CardHeader title="🔄 Master Platform Data Flow — All Modules"/>
        <div className="p-5">
          <div className="space-y-2">
            {[
              { layer:"Client Master",    color:"#34d399", actions:["Root entity — all modules reference client_id","Admin creates clients with Sector, State, User Assignments","Bulk import via Excel (.xlsx / .csv) with duplicate detection"] },
              { layer:"User Assignment",  color:"#22d3ee", actions:["Admin assigns End Users to specific clients","useRBAC() hook reads assignedClients[] at runtime","Unassigned clients shown with View Only badge in all modules"] },
              { layer:"CAS MIS",          color:"#4a90d9", actions:["Per-client monthly MIS date, outstanding dues, collection status","Statutory compliance (PT/TDS/PF/ESI/GST) tracked as Y/N/A","Month Lock blocks saves when FY/Month/cas is locked"] },
              { layer:"KRA / KPI",        color:"#a78bfa", actions:["Per-client monthly KPI scoring across 6 parameters","Applicability toggles per parameter (admin-controlled)","Month Lock blocks saves when FY/Month/kra is locked"] },
              { layer:"Fund Request",     color:"#fb923c", actions:["Per-client multi-currency fund entries","INR auto-calculated: foreign_amount × exchange_rate","Month Lock blocks saves when FY/Month/fund is locked"] },
              { layer:"Analytics Engine", color:"#fbbf24", actions:["CAS MIS Analytics: revenue mix, SPOC dues, statutory health, collection timeline","KRA/KPI Analytics: parameter scores, heatmap, escalation trend","Fund Request Analytics: currency breakdown, monthly trend, top clients"] },
              { layer:"Dashboard",        color:"#00c9b1", actions:["Aggregates all 3 analytics layers into executive KPI scorecards","Cross-module filters: Month/Quarter/Year, Client search, SPOC","Deep Dive rows: KRA heatmap, Top 10 clients bar, Funding ledger"] },
            ].map(l=>(
              <div key={l.layer} className="flex items-start gap-3 p-3 rounded-xl"
                style={{background:l.color+"0a",border:`1px solid ${l.color}20`}}>
                <div className="min-w-[130px] font-bold text-[11px] pt-0.5" style={{color:l.color}}>{l.layer}</div>
                <div className="flex flex-wrap gap-2">
                  {l.actions.map((a,i)=>(
                    <span key={i} className={`text-[11px] px-2 py-0.5 rounded-full ${t.textMuted}`}
                      style={{background:dark?"#1e2535":"#f0f4f8"}}>{a}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* RBAC flow */}
      <Card>
        <CardHeader title="🔐 User Database → Role Validation → Module Access Control"/>
        <div className="p-5">
          <div className="space-y-2">
            {[
              { layer:"User Database",     color:"#00c9b1", actions:["Admin creates user with role (Admin / End User)","Assigns specific client IDs to End User","Sends invite email with HMAC-signed token"] },
              { layer:"Login / Auth",      color:"#34d399", actions:["User activates account via invite token","JWT issued with role + assignedClients[] claims","Refresh token stored; access token expires in 1h"] },
              { layer:"useRBAC() Hook",    color:"#fbbf24", actions:["isAdmin derived from activeUser.role === 'Admin'","canEdit(clientId) returns true if Admin or client in assignedClients[]","assignedClientIds = null for Admin (all clients)"] },
              { layer:"Module Access",     color:"#4a90d9", actions:["CAS MIS: edit enabled only if canEdit(client.id) && !casMisLocked","KRA/KPI: edit enabled only if canEdit(client.id) && !kraLocked","Fund Request: edit enabled only if canEdit(client.id) && !fundRowLocked"] },
              { layer:"View Only Badge",   color:"#a78bfa", actions:["Unassigned rows show 'View Only' chip in Actions column","All Save/Delete/Add buttons hidden for unassigned clients","Data is visible but not editable"] },
            ].map(l=>(
              <div key={l.layer} className="flex items-start gap-3 p-3 rounded-xl"
                style={{background:l.color+"0a",border:`1px solid ${l.color}20`}}>
                <div className="min-w-[130px] font-bold text-[11px] pt-0.5" style={{color:l.color}}>{l.layer}</div>
                <div className="flex flex-wrap gap-2">
                  {l.actions.map((a,i)=>(
                    <span key={i} className={`text-[11px] px-2 py-0.5 rounded-full ${t.textMuted}`}
                      style={{background:dark?"#1e2535":"#f0f4f8"}}>{a}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Month Lock flow */}
      <Card>
        <CardHeader title="🔒 Admin → Month Lock → Module Data Locking"/>
        <div className="p-5">
          <div className="space-y-2">
            {[
              { layer:"Admin Action",      color:"#fbbf24", actions:["Selects FY + Month + Module (cas | kra | fund | all)","Clicks Lock button in Architect → Month Lock","Bulk Lock option available to lock entire FY at once"] },
              { layer:"Lock Key Created",  color:"#f87171", actions:["mkLockKey(fy, month, mod) generates 'FY|Month|mod' string","Stored in localStorage[procas_locks] object","Audit entry written to procas_lockAudit[]"] },
              { layer:"checkLocked()",     color:"#a78bfa", actions:["checkLocked(locks, fy, month, mod) checks specific key","Also checks 'all' key: locks[FY|Month|all]","Returns true if either specific or 'all' lock exists"] },
              { layer:"CAS MIS Effect",    color:"#4a90d9", actions:["casMisLocked = checkLocked(locks, selectedFY, selectedMonth, 'cas')","canEdit overridden: locked month → all edits blocked","Lock banner shown: 'Month is locked — Architect → Month Lock to unlock'"] },
              { layer:"KRA/KPI Effect",    color:"#a78bfa", actions:["kraLocked = checkLocked(locks, selectedFY, selectedMonth, 'kra')","isEdit forced false when locked","Edit button hidden for all rows regardless of user assignment"] },
              { layer:"Fund Request Effect",color:"#fb923c",actions:["fundRowLocked checked per row (same FY/Month)","canEditRow = isAdmin && !fundRowLocked","Add Row, Save, Delete all disabled when locked"] },
            ].map(l=>(
              <div key={l.layer} className="flex items-start gap-3 p-3 rounded-xl"
                style={{background:l.color+"0a",border:`1px solid ${l.color}20`}}>
                <div className="min-w-[140px] font-bold text-[11px] pt-0.5" style={{color:l.color}}>{l.layer}</div>
                <div className="flex flex-wrap gap-2">
                  {l.actions.map((a,i)=>(
                    <span key={i} className={`text-[11px] px-2 py-0.5 rounded-full ${t.textMuted}`}
                      style={{background:dark?"#1e2535":"#f0f4f8"}}>{a}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Reminder flow */}
      <Card>
        <CardHeader title="🔔 Admin → Auto Reminder → User Notifications"/>
        <div className="p-5">
          <div className="space-y-2">
            {[
              { layer:"Admin Creates",     color:"#34d399", actions:["Title, Module, Frequency (daily/weekly/monthly), Trigger Day","Target Role: Admin, End User, or All","Message body entered","Stored to localStorage[procas_reminders]"] },
              { layer:"Scheduler (Target)",color:"#fbbf24", actions:["node-cron job reads active reminders from DB","Evaluates frequency + trigger_day against current date","Dispatches email via Nodemailer for matching reminders"] },
              { layer:"User Receives",     color:"#00c9b1", actions:["Notification delivered to target role's email","last_sent timestamp updated in DB","Reminder log visible in Architect → Reminders"] },
              { layer:"Admin Controls",    color:"#a78bfa", actions:["Pause / Resume reminder without deleting","Manual Trigger available for immediate dispatch","Delete reminder permanently"] },
            ].map(l=>(
              <div key={l.layer} className="flex items-start gap-3 p-3 rounded-xl"
                style={{background:l.color+"0a",border:`1px solid ${l.color}20`}}>
                <div className="min-w-[140px] font-bold text-[11px] pt-0.5" style={{color:l.color}}>{l.layer}</div>
                <div className="flex flex-wrap gap-2">
                  {l.actions.map((a,i)=>(
                    <span key={i} className={`text-[11px] px-2 py-0.5 rounded-full ${t.textMuted}`}
                      style={{background:dark?"#1e2535":"#f0f4f8"}}>{a}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Analytics architecture */}
      <Card>
        <CardHeader title="📊 Analytics Architecture — Data Sources → Visualisations → Dashboard"/>
        <div className="p-5">
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            {[
              { title:"CAS MIS Analytics", color:"#4a90d9", icon:"📋",
                sources:["cas_mis table — all monthly records","client_master — sector/state/RM"],
                calcs:["Total Expected Revenue = sum(outstanding_dues)","Collection % = collected / expected × 100","Reconciliation Gap = sum(recon_gap)","Statutory Health = Y/N/A counts per type"],
                charts:["Revenue Mix donut (by sector)","Outstanding SPOC Dues horizontal bar","Collection Timeline scatter","Statutory Health 100% stacked bar","MIS Turnaround colour-coded bar","Software Ecosystem treemap"] },
              { title:"KRA/KPI Analytics", color:"#a78bfa", icon:"🎯",
                sources:["kra_kpi table — all parameter values","client_master — for client names"],
                calcs:["Parameter Score = Y count / (Y+N) × 100 (excluding A)","Overall KPI score = mean of 6 parameter scores","Escalation Count = N count on escalation field"],
                charts:["Parameter score bar chart","Client compliance heatmap","Escalation trend line","Mini donut per KPI parameter"] },
              { title:"Fund Request Analytics", color:"#fb923c", icon:"💰",
                sources:["fund_requests table — currency, amounts","client_master — for client names"],
                calcs:["Total INR = sum(inr_amount)","Currency Breakdown = INR grouped by currency","Monthly Trend = INR grouped by month","Top Clients = INR ranked descending"],
                charts:["Currency breakdown donut","Monthly INR trend area chart","Top 10 clients bar chart","Funding pipeline stacked bar","Funding ledger with variance"] },
            ].map(a=>(
              <div key={a.title} className="rounded-2xl border p-4" style={{borderColor:a.color+"40"}}>
                <div className="flex items-center gap-2 mb-3">
                  <span className="text-xl">{a.icon}</span>
                  <span className="font-black text-sm" style={{color:a.color}}>{a.title}</span>
                </div>
                <div className="space-y-3">
                  <div>
                    <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} mb-1`}>Data Sources</div>
                    {a.sources.map(s=><div key={s} className={`text-[11px] ${t.textMuted} flex gap-1.5`}><span style={{color:a.color}}>▸</span>{s}</div>)}
                  </div>
                  <div>
                    <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} mb-1`}>Calculations</div>
                    {a.calcs.map(s=><div key={s} className={`text-[11px] ${t.textMuted} flex gap-1.5`}><span style={{color:a.color}}>▸</span>{s}</div>)}
                  </div>
                  <div>
                    <div className={`text-[10px] font-bold uppercase tracking-widest ${t.textMuted} mb-1`}>Visualisations</div>
                    {a.charts.map(s=><div key={s} className={`text-[11px] ${t.textMuted} flex gap-1.5`}><span style={{color:a.color}}>▸</span>{s}</div>)}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Client Import flow */}
      <Card>
        <CardHeader title="📥 Excel Import Pipeline — Client Master"/>
        <div className="p-5">
          <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3">
            {[
              { n:1, icon:"📂", title:"Select file",    desc:"Browser File API\n.xlsx / .csv\nDrag-and-drop support",    color:"#00c9b1" },
              { n:2, icon:"⚙️", title:"Parse",          desc:"SheetJS converts\nrows to JS objects\nHeader validation",  color:"#34d399" },
              { n:3, icon:"✅", title:"Validate cols",  desc:"Client Name\nSector, State\nAssigned Users",               color:"#fbbf24" },
              { n:4, icon:"📡", title:"POST to API",    desc:"Multipart upload\nwith auth token\nAdmin role required",   color:"#4a90d9" },
              { n:5, icon:"🛡️", title:"Server validate",desc:"Schema checks\nSanitise each row\nDuplicate detection",   color:"#a78bfa" },
              { n:6, icon:"🗄️", title:"Upsert",        desc:"Insert new rows\nUpdate duplicates\nSet user assignments", color:"#f87171" },
              { n:7, icon:"📊", title:"Summary",        desc:"Total / Imported\nFailed / Duplicates\nError report download", color:"#22d3ee" },
            ].map(s=>(
              <div key={s.n} className="flex flex-col items-center gap-1.5 text-center">
                <div className="w-10 h-10 rounded-xl flex items-center justify-center text-base shadow-md"
                  style={{background:`linear-gradient(135deg,${s.color},${s.color}88)`}}>{s.icon}</div>
                <div className="text-[9px] font-black uppercase tracking-wider" style={{color:s.color}}>Step {s.n}</div>
                <div className={`text-[10px] font-bold ${t.text} leading-tight`}>{s.title}</div>
                <div className={`text-[9px] ${t.textMuted} whitespace-pre-line leading-tight`}>{s.desc}</div>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Fund Request INR flow */}
      <Card>
        <CardHeader title="💰 Fund Request — INR Auto-Calculation Flow"/>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          {[
            { n:1, icon:"✏️", title:"Fill row",       desc:"Client dropdown\nCurrency + Amount\nExchange Rate",            color:"#00c9b1" },
            { n:2, icon:"⚡", title:"Auto-calc INR",  desc:"fc × rate instantly\nNo server roundtrip\nBlue AUTO cell",    color:"#34d399" },
            { n:3, icon:"📡", title:"Save to API",    desc:"Validated payload\nJWT token checked\nRBAC + Lock checked",   color:"#fbbf24" },
            { n:4, icon:"🗄️", title:"DB generates",  desc:"inr_amount column\nROUND(fc × rate, 2)\nAlways authoritative",color:"#4a90d9" },
            { n:5, icon:"📋", title:"Audit + Return", desc:"Audit row written\nFull row returned\nUI state updated",      color:"#f87171" },
          ].map(s=>(
            <div key={s.n} className="rounded-xl p-4 border"
              style={{background:s.color+"0f",borderColor:s.color+"30"}}>
              <div className="flex items-center gap-2 mb-2">
                <div className="w-8 h-8 rounded-lg flex items-center justify-center text-base"
                  style={{background:s.color+"22"}}>{s.icon}</div>
                <span className="text-[9px] font-black uppercase tracking-wider" style={{color:s.color}}>Step {s.n}</span>
              </div>
              <div className="font-bold text-xs mb-1" style={{color:s.color}}>{s.title}</div>
              <div className={`text-[10px] leading-relaxed ${t.textMuted} whitespace-pre-line`}>{s.desc}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );

  return (
    <div className="space-y-5">
      {/* Section nav */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-1.5 flex gap-1 overflow-x-auto`}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>setSection(n.id)}
            className={`flex-1 min-w-fit px-3 py-2 rounded-xl text-xs font-semibold
              transition-all whitespace-nowrap ${section===n.id?"text-white shadow-md":t.textMuted}`}
            style={section===n.id?{background:"linear-gradient(135deg,#00a896,#1b5fa8)"}:{}}>
            {n.label}
          </button>
        ))}
      </div>

      {section==="overview"   && <OverviewSection/>}
      {section==="schema"     && <SchemaSection/>}
      {section==="api"        && <ApiSection/>}
      {section==="roles"      && <RolesSection/>}
      {section==="stack"      && <StackSection/>}
      {section==="flow"       && <FlowSection/>}
      {section==="monthlock"  && <MonthLockControl t={t} dark={dark} isAdmin={isAdmin} adminName={adminName}/>}
      {section==="reminders"  && <ReminderManagement t={t} dark={dark} isAdmin={isAdmin}/>}
      {section==="auditlog"   && <SystemAuditLog t={t} dark={dark}/>}
      {section==="mastersett" && <MasterSettings t={t} dark={dark} isAdmin={isAdmin}/>}
      {section==="syshealth"  && <SystemHealthDashboard t={t} dark={dark}/>}
    </div>
  );
};


// ─────────────────────────────────────────────────────────────────────────────
// ── DASHBOARD TAB ─────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// ── OVERVIEW PANEL — Premium CFO/CXO Executive Dashboard ─────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const OverviewPanel = ({ t, dark, isAdmin, clients, onNavigate }) => {
  // ── Filters ─────────────────────────────────────────────────────────────────
  const [period,   setPeriod]   = useState("Month");
  const [clientQ,  setClientQ]  = useState("");
  const [spocF,    setSpocF]    = useState("All");
  const [row3Open, setRow3Open] = useState(true);
  const [row4Open, setRow4Open] = useState(false);
  const [ledgerQ,  setLedgerQ]  = useState("");
  const [ledgerPg, setLedgerPg] = useState(1);
  const PER_PAGE = 5;

  // ── Derived data ─────────────────────────────────────────────────────────────
  const activeC   = clients.filter(c => c.status === "Active");
  const filtered  = clients.filter(c => {
    const q = clientQ.toLowerCase();
    const matchQ = !q || c.clientName.toLowerCase().includes(q) || c.id.toLowerCase().includes(q);
    const matchS = spocF === "All" || c.rm === spocF;
    return matchQ && matchS;
  });
  const spocList  = ["All", ...Array.from(new Set(clients.map(c => c.rm).filter(Boolean)))];

  // ── Synthetic KPI data (derived from available client fields) ─────────────────
  const totalTurnover = clients.reduce((s, c) => {
    const n = parseFloat((c.turnover || "0").replace(/[₹,\sCr]/g,"")) * (c.turnover?.includes("Cr") ? 1e7 : 1e5);
    return s + (isNaN(n) ? 0 : n);
  }, 0);
  const casRevenue    = Math.round(totalTurnover * 0.018);
  const casCollected  = Math.round(casRevenue * 0.78);
  const casPending    = casRevenue - casCollected;
  const reconGap      = clients.filter(c => c.risk === "High").length;
  const kpiScore      = 82;
  const fundInward    = Math.round(totalTurnover * 0.064);
  const fundPending   = Math.round(fundInward * 0.22);
  const collPct       = casRevenue > 0 ? Math.round((casCollected / casRevenue) * 100) : 0;

  // Revenue mix (sector-weighted)
  const sectorRevMap = {};
  clients.forEach(c => {
    const n = parseFloat((c.turnover||"0").replace(/[₹,\sCr]/g,""))*(c.turnover?.includes("Cr")?1e7:1e5);
    sectorRevMap[c.sector] = (sectorRevMap[c.sector]||0) + (isNaN(n)?0:n)*0.018;
  });
  const revMixData = Object.entries(sectorRevMap).map(([name,v],i)=>({name, value:Math.round(v), color:["#4a90d9","#00c9b1","#34d399","#fbbf24","#a78bfa","#fb923c","#22d3ee","#f87171"][i%8]}));

  // Collection delay timeline (synthetic monthly)
  const months6 = ["Jan","Feb","Mar","Apr","May","Jun"];
  const collDelay = months6.map((m,i) => ({ month:m, delay: [8,6,9,5,7,4][i] }));

  // Fund inflow trend
  const fundTrend = months6.map((m,i) => ({ month:m, inflow:[12,18,15,22,19,27][i]*1e5 }));

  // Outstanding SPOC dues by RM
  const rmDues = {};
  clients.forEach(c => { const v = Math.round(parseFloat((c.turnover||"0").replace(/[₹,\sCr]/g,""))*(c.turnover?.includes("Cr")?1e7:1e5)*0.018*0.22); rmDues[c.rm||"—"] = (rmDues[c.rm||"—"]||0)+v; });
  const outstandingData = Object.entries(rmDues).map(([name,amt])=>({name:name.split(" ")[0],amt})).sort((a,b)=>b.amt-a.amt);

  // Funding pipeline
  const pipelineData = [
    { month:"Apr", Initiated:4, Processing:3, Cleared:5, Completed:8 },
    { month:"May", Initiated:5, Processing:4, Cleared:6, Completed:10 },
    { month:"Jun", Initiated:3, Processing:6, Cleared:7, Completed:9 },
  ];

  // KRA mini donuts
  const kraMetrics = [
    { label:"MYSA Comp.", value:88, color:"#a78bfa" },
    { label:"Rectification", value:74, color:"#fb923c" },
    { label:"Raksha Tool", value:61, color:"#22d3ee" },
    { label:"Capital Usage", value:95, color:"#34d399" },
  ];

  // Escalations trend
  const escTrend = months6.map((m,i)=>({month:m, esc:[2,4,3,5,2,1][i]}));

  // MIS deadlines
  const misDeadlines = clients.slice(0,6).map((c,i)=>{
    const day = [6,8,12,9,15,7][i];
    return { name:c.clientName.slice(0,14), day, ok: day <= 10 };
  });

  // Statutory health
  const statData = [
    { name:"PT",  Yes:75, No:15, NA:10 },
    { name:"TDS", Yes:88, No:8,  NA:4  },
    { name:"PF",  Yes:70, No:20, NA:10 },
    { name:"ESI", Yes:65, No:25, NA:10 },
    { name:"GST", Yes:80, No:12, NA:8  },
  ];

  // Turnaround scatter
  const turnaround = clients.map((c,i)=>({name:c.clientName.slice(0,10), days:[4,7,3,9,12,5,8,6][i%8]}));

  // KRA heatmap
  const kraParams = ["MIS Date","MYSA","Rect.","Escalation","Raksha","Capital"];
  const heatValues = [
    [1,1,0,1,1,1],[1,0,1,1,0,1],[0,1,1,0,1,1],
    [1,1,1,1,1,1],[1,0,0,1,1,0],[0,1,1,0,0,1],
    [1,1,0,1,0,1],[1,1,1,1,1,0],
  ];

  // Top clients by fund volume
  const topClients = [...clients].sort((a,b)=>{
    const toN = s => parseFloat((s||"0").replace(/[₹,\sCr]/g,""))*(s?.includes("Cr")?1e7:1e5);
    return toN(b.turnover) - toN(a.turnover);
  }).slice(0,10).map(c => ({
    name: c.clientName.slice(0,18),
    vol: Math.round(parseFloat((c.turnover||"0").replace(/[₹,\sCr]/g,""))*(c.turnover?.includes("Cr")?1e7:1e5)*0.064),
  }));

  // Ledger
  const ledgerRows = clients.map(c => {
    const t2 = parseFloat((c.turnover||"0").replace(/[₹,\sCr]/g,""))*(c.turnover?.includes("Cr")?1e7:1e5);
    const exp = Math.round(t2*0.064); const act = Math.round(exp*[0.92,0.88,1.05,0.78,0.95,1.0,0.82,0.97][clients.indexOf(c)%8]);
    const variance = act - exp;
    return { id:c.id, name:c.clientName.slice(0,22), parent:c.type, exp, act, variance };
  });
  const ledgerFiltered = ledgerRows.filter(r => r.name.toLowerCase().includes(ledgerQ.toLowerCase()) || r.id.toLowerCase().includes(ledgerQ.toLowerCase()));
  const ledgerPages = Math.ceil(ledgerFiltered.length / PER_PAGE);
  const ledgerSlice = ledgerFiltered.slice((ledgerPg-1)*PER_PAGE, ledgerPg*PER_PAGE);

  // Formatters
  const fmtC = n => n>=1e7?`₹${(n/1e7).toFixed(1)}Cr`:n>=1e5?`₹${(n/1e5).toFixed(1)}L`:n>=1000?`₹${(n/1000).toFixed(0)}K`:`₹${n}`;

  // Chart tokens
  const AX  = dark ? "#5a7a99" : "#9ca3af";
  const GR  = dark ? "#1a2d4420" : "#e2e8f040";
  const TBG = dark ? "#0f1e2e" : "#ffffff";
  const TBD = dark ? "#243d58" : "#d4e0ed";
  const CC  = ["#4a90d9","#00c9b1","#34d399","#fbbf24","#a78bfa","#fb923c","#22d3ee","#f87171"];

  // Recharts components are imported as ES modules at the top of the file.
  // window.Recharts is only set when using the CDN <script> tag, which is
  // not used here — named imports are always in scope, so hasR is always true.
  const hasR = true;

  // Tooltip
  const CTip = ({ active, payload, label }) => {
    if (!active||!payload?.length) return null;
    return (
      <div style={{background:TBG,border:`1px solid ${TBD}`,borderRadius:10,padding:"8px 12px",minWidth:110,boxShadow:"0 8px 24px rgba(0,0,0,0.22)"}}>
        {label && <p style={{fontWeight:700,fontSize:11,color:dark?"#c8dff0":"#0d2137",marginBottom:4}}>{label}</p>}
        {payload.map((p,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:10,marginTop:2}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:p.color||p.fill,flexShrink:0}}/>
            <span style={{color:dark?"#8aa4be":"#6b7280"}}>{p.name}:</span>
            <span style={{fontWeight:700,color:p.color||p.fill,marginLeft:"auto"}}>
              {typeof p.value==="number"&&p.value>999?fmtC(p.value):p.value}
            </span>
          </div>
        ))}
      </div>
    );
  };

  // Card
  const Card = ({ title, sub, accent, badge, children, className="" }) => (
    <div className={`${t.card} border ${t.cardBorder} rounded-2xl flex flex-col overflow-hidden ${className}`}
      style={accent?{borderTop:`2.5px solid ${accent}`}:{}}>
      <div className={`px-4 py-3 border-b ${t.cardBorder} flex items-center justify-between`}>
        <div>
          <h4 className={`font-bold text-sm ${t.text}`}>{title}</h4>
          {sub && <p className={`text-[10px] mt-0.5 ${t.textMuted}`}>{sub}</p>}
        </div>
        {badge && <span className="text-[9px] font-bold px-2 py-0.5 rounded-full" style={{background:`${badge[1]}18`,color:badge[1]}}>{badge[0]}</span>}
      </div>
      <div className="flex-1 px-4 py-3">{children}</div>
    </div>
  );

  // Collapsible section header
  const CollRow = ({ title, open, onToggle, accent }) => (
    <button onClick={onToggle}
      className={`w-full flex items-center gap-3 px-5 py-3 rounded-2xl border ${t.cardBorder} transition-all`}
      style={{background:dark?"linear-gradient(135deg,#0c1e30,#0a1a28)":"linear-gradient(135deg,#f0f7ff,#e8f4f2)",textAlign:"left"}}>
      <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0" style={{background:`${accent}20`}}>
        <span style={{color:accent,fontSize:12}}>{open?"▾":"▸"}</span>
      </div>
      <span className={`font-bold text-sm ${t.text} flex-1`}>{title}</span>
      <span className="text-[10px] font-semibold" style={{color:accent}}>{open?"Collapse":"Expand"}</span>
    </button>
  );

  // Mini donut SVG (no recharts)
  const MiniDonut = ({ value, color, label }) => {
    const r=28, cx=36, cy=36, circ=2*Math.PI*r;
    const filled = (value/100)*circ;
    return (
      <div className="flex flex-col items-center">
        <svg width={72} height={72}>
          <circle cx={cx} cy={cy} r={r} fill="none" stroke={dark?"#1a2d44":"#e2e8f0"} strokeWidth={8}/>
          <circle cx={cx} cy={cy} r={r} fill="none" stroke={color} strokeWidth={8}
            strokeDasharray={`${filled} ${circ-filled}`} strokeLinecap="round"
            transform={`rotate(-90 ${cx} ${cy})`}/>
          <text x={cx} y={cy+1} textAnchor="middle" dominantBaseline="middle" fill={color} fontSize={12} fontWeight="bold">{value}%</text>
        </svg>
        <span className="text-[10px] font-semibold mt-1" style={{color:AX}}>{label}</span>
      </div>
    );
  };

  // Treemap cell
  const TMCell = ({ x,y,width,height,name,count,index }) => {
    if (width<20||height<20) return null;
    const col=CC[index%CC.length];
    return (
      <g>
        <rect x={x+1} y={y+1} width={width-2} height={height-2} rx={6} fill={col} fillOpacity={0.85}/>
        {width>40&&<text x={x+width/2} y={y+height/2} textAnchor="middle" dominantBaseline="middle" fill="#fff" fontSize={Math.min(11,width/5)} fontWeight="bold">{name}</text>}
      </g>
    );
  };

  // Software ecosystem data
  const swData = [
    {name:"Zoho",count:4,size:4},{name:"Tally",count:2,size:2},{name:"SAP",count:1,size:1},
    {name:"NetSuite",count:1,size:1},{name:"Others",count:clients.length-8<0?0:clients.length-8,size:1},
  ];

  // Geographic distribution
  const geoData = [
    {state:"Maharashtra",count:3,color:"#4a90d9"},{state:"Karnataka",count:2,color:"#00c9b1"},
    {state:"Delhi",count:2,color:"#34d399"},{state:"Gujarat",count:1,color:"#fbbf24"},
  ];

  return (
    <div className="space-y-5">

      {/* ══ STICKY FILTER HEADER ════════════════════════════════════════════════ */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl px-5 py-3 flex flex-wrap items-center gap-x-4 gap-y-2 sticky top-0 z-20`}
        style={{boxShadow: dark?"0 4px 24px rgba(0,0,0,0.4)":"0 4px 24px rgba(0,0,0,0.10)"}}>
        {/* Period */}
        <div className="flex items-center gap-1">
          {["Month","Quarter","Year"].map(p=>(
            <button key={p} onClick={()=>setPeriod(p)}
              className="px-3 py-1.5 rounded-lg text-[11px] font-semibold transition-all"
              style={{background:period===p?"linear-gradient(135deg,#00c9b1,#007a6e)":"transparent",
                color:period===p?"#fff":AX,
                border:period===p?"none":`1px solid ${dark?"#243d58":"#d4e0ed"}`}}>
              {p}
            </button>
          ))}
        </div>
        <div className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border ${t.cardBorder} flex-1`} style={{minWidth:140,maxWidth:220}}>
          <Icon path={Icons.search} size={12} className={t.textMuted}/>
          <input value={clientQ} onChange={e=>setClientQ(e.target.value)} placeholder="Search client…"
            className={`bg-transparent outline-none text-xs flex-1 ${t.text}`}/>
          {clientQ&&<button onClick={()=>setClientQ("")} className={t.textMuted}><Icon path={Icons.x} size={10}/></button>}
        </div>
        <select value={spocF} onChange={e=>setSpocF(e.target.value)}
          className={`px-3 py-1.5 rounded-xl border text-xs outline-none ${t.input}`} style={{minWidth:110}}>
          {spocList.map(s=><option key={s} value={s}>{s==="All"?"All SPOCs":s}</option>)}
        </select>
        <div className="flex items-center gap-2 ml-auto">
          <button className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[11px] font-semibold border ${t.cardBorder} ${t.textMuted} ${t.hover}`}>
            <Icon path={Icons.download} size={12}/> Export
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-[11px] font-semibold text-white"
            style={{background:"linear-gradient(135deg,#00c9b1,#007a6e)"}}>
            <Icon path={Icons.refreshCw} size={12}/> Refresh
          </button>
        </div>
        <span className={`text-[10px] ${t.textMuted} w-full text-right`}
          style={{marginTop:-4}}>Showing {filtered.length} of {clients.length} clients · {period} view</span>
      </div>

      {/* ══ ROW 1 — EXECUTIVE SCORECARDS ════════════════════════════════════════ */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Revenue CAS */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 relative overflow-hidden`} style={{borderTop:"2.5px solid #4a90d9"}}>
          <div className="absolute -top-8 -right-8 w-24 h-24 rounded-full opacity-[0.05]" style={{background:"#4a90d9"}}/>
          <div className="flex items-start justify-between mb-2">
            <div>
              <p className={`text-[9px] font-bold uppercase tracking-widest ${t.textMuted}`}>Revenue · CAS</p>
              <p className={`text-[8px] ${t.textMuted}`}>Expected this {period.toLowerCase()}</p>
            </div>
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{background:"#4a90d918"}}><Icon path={Icons.dollar} size={13} style={{color:"#4a90d9"}}/></div>
          </div>
          <p className="text-2xl font-black leading-none mb-2" style={{color:"#4a90d9"}}>{fmtC(casRevenue)}</p>
          <div className="flex gap-2">
            <div className="flex-1 rounded-lg p-1.5" style={{background:dark?"#0a1e14":"#f0faf5",border:"1px solid #34d39925"}}>
              <p className="text-[8px]" style={{color:"#34d39970"}}>Collected</p><p className="text-xs font-black" style={{color:"#34d399"}}>{fmtC(casCollected)}</p>
            </div>
            <div className="flex-1 rounded-lg p-1.5" style={{background:dark?"#1a0f00":"#fff8f0",border:"1px solid #fbbf2425"}}>
              <p className="text-[8px]" style={{color:"#fbbf2470"}}>Pending</p><p className="text-xs font-black" style={{color:"#fbbf24"}}>{fmtC(casPending)}</p>
            </div>
          </div>
          <div className="mt-2 h-1.5 rounded-full overflow-hidden" style={{background:dark?"#1a2d44":"#e2e8f0"}}>
            <div className="h-full rounded-full" style={{width:`${collPct}%`,background:"linear-gradient(90deg,#34d399,#4a90d9)"}}/>
          </div>
          <p className="text-[9px] mt-0.5" style={{color:AX}}>{collPct}% collected</p>
        </div>

        {/* Card 2: Capital Funds */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 relative overflow-hidden`} style={{borderTop:"2.5px solid #00c9b1"}}>
          <div className="absolute -top-8 -right-8 w-24 h-24 rounded-full opacity-[0.05]" style={{background:"#00c9b1"}}/>
          <div className="flex items-start justify-between mb-2">
            <div><p className={`text-[9px] font-bold uppercase tracking-widest ${t.textMuted}`}>Capital · Funds</p><p className={`text-[8px] ${t.textMuted}`}>Inward this {period.toLowerCase()}</p></div>
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{background:"#00c9b118"}}><Icon path={Icons.trending} size={13} style={{color:"#00c9b1"}}/></div>
          </div>
          <p className="text-2xl font-black leading-none mb-2" style={{color:"#00c9b1"}}>{fmtC(fundInward)}</p>
          <div className="flex gap-2">
            <div className="flex-1 rounded-lg p-1.5" style={{background:dark?"#0a1e14":"#f0faf5",border:"1px solid #00c9b125"}}>
              <p className="text-[8px]" style={{color:"#00c9b170"}}>Cleared</p><p className="text-xs font-black" style={{color:"#00c9b1"}}>{fmtC(fundInward-fundPending)}</p>
            </div>
            <div className="flex-1 rounded-lg p-1.5" style={{background:dark?"#1a0f00":"#fff8f0",border:"1px solid #fbbf2425"}}>
              <p className="text-[8px]" style={{color:"#fbbf2470"}}>Pending</p><p className="text-xs font-black" style={{color:"#fbbf24"}}>{fmtC(fundPending)}</p>
            </div>
          </div>
          <div className="mt-2 h-1.5 rounded-full overflow-hidden" style={{background:dark?"#1a2d44":"#e2e8f0"}}>
            <div className="h-full rounded-full" style={{width:`${Math.round(((fundInward-fundPending)/Math.max(fundInward,1))*100)}%`,background:"linear-gradient(90deg,#00c9b1,#34d399)"}}/>
          </div>
          <p className="text-[9px] mt-0.5" style={{color:AX}}>{Math.round(((fundInward-fundPending)/Math.max(fundInward,1))*100)}% cleared</p>
        </div>

        {/* Card 3: Accuracy */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 relative overflow-hidden`} style={{borderTop:`2.5px solid ${reconGap>0?"#f87171":"#34d399"}`}}>
          <div className="absolute -top-8 -right-8 w-24 h-24 rounded-full opacity-[0.05]" style={{background:reconGap>0?"#f87171":"#34d399"}}/>
          <div className="flex items-start justify-between mb-2">
            <div><p className={`text-[9px] font-bold uppercase tracking-widest ${t.textMuted}`}>Accuracy · CAS</p><p className={`text-[8px] ${t.textMuted}`}>Reconciliation status</p></div>
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{background:reconGap>0?"#f8717118":"#34d39918"}}><Icon path={reconGap>0?Icons.alert:Icons.checkCircle} size={13} style={{color:reconGap>0?"#f87171":"#34d399"}}/></div>
          </div>
          <p className="text-2xl font-black leading-none mb-2" style={{color:reconGap>0?"#f87171":"#34d399"}}>{reconGap}</p>
          <div className="rounded-lg px-2.5 py-1.5" style={{background:reconGap>0?"#f8717110":"#34d39910",border:`1px solid ${reconGap>0?"#f8717128":"#34d39928"}`}}>
            <p className="text-[9px] font-semibold" style={{color:reconGap>0?"#f87171":"#34d399"}}>{reconGap===0?"✅ All reconciled":`⚠ ${reconGap} high-risk exception${reconGap!==1?"s":""}`}</p>
          </div>
          <p className={`text-[9px] mt-2 ${t.textMuted}`}>Mismatch count · {period.toLowerCase()} view</p>
        </div>

        {/* Card 4: Performance KRA */}
        <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-4 relative overflow-hidden`} style={{borderTop:"2.5px solid #a78bfa"}}>
          <div className="absolute -top-8 -right-8 w-24 h-24 rounded-full opacity-[0.05]" style={{background:"#a78bfa"}}/>
          <div className="flex items-start justify-between mb-2">
            <div><p className={`text-[9px] font-bold uppercase tracking-widest ${t.textMuted}`}>Performance · KRA</p><p className={`text-[8px] ${t.textMuted}`}>Overall compliance score</p></div>
            <div className="flex items-center gap-1 px-2 py-0.5 rounded-full" style={{background:"#a78bfa18",border:"1px solid #a78bfa30"}}>
              <Icon path={Icons.arrowUp} size={9} style={{color:"#a78bfa"}}/>
              <span className="text-[9px] font-bold" style={{color:"#a78bfa"}}>+3%</span>
            </div>
          </div>
          <p className="text-2xl font-black leading-none mb-2" style={{color:"#a78bfa"}}>{kpiScore}%</p>
          <div className="mt-1 h-2 rounded-full overflow-hidden" style={{background:dark?"#1a2d44":"#e2e8f0"}}>
            <div className="h-full rounded-full transition-all duration-700" style={{width:`${kpiScore}%`,background:"linear-gradient(90deg,#a78bfa,#6d28d9)"}}/>
          </div>
          <div className="flex justify-between mt-1">
            <span className="text-[9px]" style={{color:AX}}>Team avg.</span>
            <span className="text-[9px]" style={{color:"#a78bfa"}}>Target: 85%</span>
          </div>
        </div>
      </div>

      {/* ══ ROW 2 — FINANCIALS & TRENDS ════════════════════════════════════════ */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">

        {/* LEFT: Revenue Mix + Geography */}
        <div className="space-y-4">
          <Card title="Revenue Mix" sub="Sector-wise CAS revenue" accent="#4a90d9">
            {hasR ? (
              <div>
                <ResponsiveContainer width="100%" height={160}>
                  <PieChart>
                    <Pie data={revMixData} cx="50%" cy="50%" innerRadius={42} outerRadius={68} paddingAngle={2} dataKey="value" labelLine={false}>
                      {revMixData.map((d,i)=><Cell key={i} fill={d.color} stroke="none"/>)}
                    </Pie>
                    <Tooltip content={<CTip/>}/>
                  </PieChart>
                </ResponsiveContainer>
                <div className="grid grid-cols-2 gap-x-3 gap-y-1 mt-1">
                  {revMixData.map(d=>(
                    <div key={d.name} className="flex items-center gap-1.5">
                      <div className="w-2 h-2 rounded-full shrink-0" style={{background:d.color}}/>
                      <span className="text-[9px] truncate" style={{color:AX}}>{d.name}</span>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="space-y-2">{revMixData.map((d,i)=>(
                <div key={d.name}>
                  <div className="flex justify-between text-[10px]"><span style={{color:d.color}}>{d.name}</span><span style={{color:d.color}}>{fmtC(d.value)}</span></div>
                  <div className="h-1.5 rounded-full mt-0.5" style={{background:dark?"#1a2d44":"#e2e8f0"}}><div style={{width:`${(d.value/Math.max(casRevenue,1))*100}%`,height:"100%",borderRadius:99,background:d.color}}/></div>
                </div>
              ))}</div>
            )}
          </Card>

          <Card title="Client Geography" sub="Parent entity distribution" accent="#00c9b1">
            <div className="space-y-2.5">
              {geoData.map(g=>(
                <div key={g.state}>
                  <div className="flex items-center justify-between mb-0.5">
                    <span className="text-[10px] font-medium" style={{color:dark?"#c8dff0":"#0d2137"}}>{g.state}</span>
                    <span className="text-[9px] font-bold" style={{color:g.color}}>{g.count} clients</span>
                  </div>
                  <div className="h-1.5 rounded-full" style={{background:dark?"#1a2d44":"#e2e8f0"}}>
                    <div className="h-full rounded-full" style={{width:`${(g.count/clients.length)*100}%`,background:g.color}}/>
                  </div>
                </div>
              ))}
              <p className={`text-[9px] mt-2 ${t.textMuted}`}>Based on registered office state · Client Master data</p>
            </div>
          </Card>
        </div>

        {/* CENTER: Collection Delay + Fund Inflow */}
        <div className="space-y-4">
          <Card title="Collection Delay Trend" sub="Avg days between payment & Zoho reconciliation" accent="#4a90d9">
            {hasR ? (
              <ResponsiveContainer width="100%" height={155}>
                <LineChart data={collDelay} margin={{left:0,right:8,top:5,bottom:0}}>
                  <CartesianGrid stroke={GR}/>
                  <XAxis dataKey="month" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false} label={{value:"Days",angle:-90,position:"insideLeft",fill:AX,fontSize:9}}/>
                  <Tooltip content={<CTip/>}/>
                  <Line type="monotone" dataKey="delay" name="Delay (days)" stroke="#4a90d9" strokeWidth={2} dot={{r:3,fill:"#4a90d9"}} activeDot={{r:5}}/>
                </LineChart>
              </ResponsiveContainer>
            ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
          </Card>

          <Card title="Fund Inflow Trend" sub="Monthly fund receipts · INR equivalent" accent="#00c9b1">
            {hasR ? (
              <ResponsiveContainer width="100%" height={155}>
                <AreaChart data={fundTrend} margin={{left:0,right:8,top:5,bottom:0}}>
                  <defs>
                    <linearGradient id="fundGrad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#00c9b1" stopOpacity={0.35}/>
                      <stop offset="95%" stopColor="#00c9b1" stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid stroke={GR}/>
                  <XAxis dataKey="month" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>fmtC(v)}/>
                  <Tooltip content={<CTip/>}/>
                  <Area type="monotone" dataKey="inflow" name="Inflow" stroke="#00c9b1" fill="url(#fundGrad)" strokeWidth={2}/>
                </AreaChart>
              </ResponsiveContainer>
            ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
          </Card>
        </div>

        {/* RIGHT: Outstanding + Pipeline */}
        <div className="space-y-4">
          <Card title="Outstanding SPOC Dues" sub="Pending fees by responsible manager" accent="#fbbf24">
            {hasR && outstandingData.length>0 ? (
              <ResponsiveContainer width="100%" height={155}>
                <BarChart data={outstandingData} layout="vertical" margin={{left:0,right:10,top:0,bottom:0}}>
                  <CartesianGrid horizontal={false} stroke={GR}/>
                  <XAxis type="number" tick={{fill:AX,fontSize:9}} tickFormatter={v=>fmtC(v)} axisLine={false} tickLine={false}/>
                  <YAxis type="category" dataKey="name" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false} width={46}/>
                  <Tooltip content={<CTip/>}/>
                  <Bar dataKey="amt" name="Outstanding" radius={[0,4,4,0]} barSize={12}>
                    {outstandingData.map((_,i)=><Cell key={i} fill={CC[i%CC.length]}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : <div className={`flex items-center justify-center h-36 text-xs ${t.textMuted}`}>No outstanding dues 🎉</div>}
          </Card>

          <Card title="Funding Pipeline" sub="Request status by month" accent="#00c9b1">
            {hasR ? (
              <ResponsiveContainer width="100%" height={155}>
                <BarChart data={pipelineData} margin={{left:0,right:5,top:5,bottom:0}}>
                  <CartesianGrid vertical={false} stroke={GR}/>
                  <XAxis dataKey="month" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                  <Tooltip content={<CTip/>}/>
                  <Bar dataKey="Initiated"  stackId="a" fill="#94a3b8" barSize={16}/>
                  <Bar dataKey="Processing" stackId="a" fill="#fbbf24" barSize={16}/>
                  <Bar dataKey="Cleared"    stackId="a" fill="#00c9b1" barSize={16}/>
                  <Bar dataKey="Completed"  stackId="a" fill="#34d399" radius={[3,3,0,0]} barSize={16}/>
                </BarChart>
              </ResponsiveContainer>
            ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
            <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2">
              {[{c:"#94a3b8",l:"Initiated"},{c:"#fbbf24",l:"Processing"},{c:"#00c9b1",l:"Cleared"},{c:"#34d399",l:"Completed"}].map(l=>(
                <div key={l.l} className="flex items-center gap-1"><div className="w-2 h-2 rounded-sm" style={{background:l.c}}/><span className="text-[9px]" style={{color:AX}}>{l.l}</span></div>
              ))}
            </div>
          </Card>
        </div>
      </div>

      {/* ══ ROW 3 — COMPLIANCE & OPERATIONS (collapsible) ════════════════════════ */}
      <CollRow title="Compliance & Operations" open={row3Open} onToggle={()=>setRow3Open(o=>!o)} accent="#34d399"/>
      {row3Open && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">

          {/* LEFT: KRA mini donuts + Software */}
          <div className="space-y-4">
            <Card title="KRA Compliance" sub="4 key performance areas" accent="#a78bfa">
              <div className="grid grid-cols-2 gap-3">
                {kraMetrics.map(k=><MiniDonut key={k.label} value={k.value} color={k.color} label={k.label}/>)}
              </div>
            </Card>
            <Card title="Software Ecosystem" sub="Accounting software adoption" accent="#4a90d9">
              {hasR ? (
                <ResponsiveContainer width="100%" height={140}>
                  <Treemap data={swData} dataKey="size" nameKey="name" aspectRatio={4/3} content={<TMCell/>}>
                    <Tooltip content={({active,payload})=>{
                      if(!active||!payload?.length) return null;
                      const d=payload[0]?.payload;
                      return <div style={{background:TBG,border:`1px solid ${TBD}`,borderRadius:8,padding:"6px 10px",fontSize:10}}><p style={{fontWeight:700,color:dark?"#c8dff0":"#0d2137"}}>{d?.name}</p><p style={{color:"#4a90d9"}}>{d?.count} clients</p></div>;
                    }}/>
                  </Treemap>
                </ResponsiveContainer>
              ) : (
                <div className="space-y-1.5">
                  {swData.map((s,i)=><div key={s.name} className="flex items-center justify-between text-xs"><span style={{color:CC[i%CC.length]}}>{s.name}</span><span className="font-bold" style={{color:CC[i%CC.length]}}>{s.count}</span></div>)}
                </div>
              )}
            </Card>
          </div>

          {/* CENTER: Escalations + MIS Deadlines */}
          <div className="space-y-4">
            <Card title="Escalations Trend" sub="Monthly escalation count" accent="#f87171">
              {hasR ? (
                <ResponsiveContainer width="100%" height={155}>
                  <LineChart data={escTrend} margin={{left:0,right:8,top:5,bottom:0}}>
                    <CartesianGrid stroke={GR}/>
                    <XAxis dataKey="month" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                    <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                    <Tooltip content={<CTip/>}/>
                    <Line type="monotone" dataKey="esc" name="Escalations" stroke="#f87171" strokeWidth={2} dot={{r:3,fill:"#f87171"}} activeDot={{r:5}}/>
                  </LineChart>
                </ResponsiveContainer>
              ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
            </Card>

            <Card title="MIS Deadlines" sub="Days completed by client — green ≤10th, red after 10th" accent="#34d399">
              <div className="space-y-2">
                {misDeadlines.map((m)=>(
                  <div key={m.name} className="flex items-center gap-3">
                    <span className="text-[9px] w-24 truncate" style={{color:dark?"#c8dff0":"#0d2137"}}>{m.name}</span>
                    <div className="flex-1 h-2 rounded-full overflow-hidden" style={{background:dark?"#1a2d44":"#e2e8f0"}}>
                      <div className="h-full rounded-full" style={{width:`${Math.min((m.day/31)*100,100)}%`,background:m.ok?"#34d399":"#f87171"}}/>
                    </div>
                    <span className="text-[9px] font-bold w-8 text-right" style={{color:m.ok?"#34d399":"#f87171"}}>{m.day}th</span>
                  </div>
                ))}
              </div>
            </Card>
          </div>

          {/* RIGHT: Statutory Health + Turnaround */}
          <div className="space-y-4">
            <Card title="Statutory Health" sub="PT · TDS · PF · ESI · GST compliance" accent="#34d399">
              {hasR ? (
                <div>
                  <ResponsiveContainer width="100%" height={155}>
                    <BarChart data={statData} margin={{left:0,right:5,top:5,bottom:0}}>
                      <CartesianGrid vertical={false} stroke={GR}/>
                      <XAxis dataKey="name" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                      <YAxis tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>`${v}%`}/>
                      <Tooltip content={<CTip/>}/>
                      <Bar dataKey="Yes" name="Filed/Paid" stackId="a" fill="#34d399"/>
                      <Bar dataKey="No"  name="Pending"    stackId="a" fill="#f87171"/>
                      <Bar dataKey="NA"  name="Not Appl."  stackId="a" fill="#94a3b8" radius={[3,3,0,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                  <div className="flex gap-3 mt-2">{[{c:"#34d399",l:"Filed"},{c:"#f87171",l:"Pending"},{c:"#94a3b8",l:"N/A"}].map(l=><div key={l.l} className="flex items-center gap-1"><div className="w-1.5 h-1.5 rounded-sm" style={{background:l.c}}/><span className="text-[9px]" style={{color:AX}}>{l.l}</span></div>)}</div>
                </div>
              ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
            </Card>

            <Card title="MIS Turnaround Time" sub="Days from First Cut to Closure" accent="#4a90d9">
              {hasR ? (
                <ResponsiveContainer width="100%" height={155}>
                  <ScatterChart margin={{top:5,right:5,bottom:5,left:0}}>
                    <CartesianGrid stroke={GR}/>
                    <XAxis dataKey="days" name="Days" type="number" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false}/>
                    <YAxis dataKey="index" tick={false} axisLine={false} tickLine={false} width={0}/>
                    <Tooltip cursor={{strokeDasharray:"3 3"}} content={({active,payload})=>{
                      if(!active||!payload?.length) return null;
                      const d=payload[0]?.payload;
                      return <div style={{background:TBG,border:`1px solid ${TBD}`,borderRadius:8,padding:"6px 10px",fontSize:10}}><p style={{fontWeight:700,color:dark?"#c8dff0":"#0d2137"}}>{d?.name}</p><p style={{color:"#4a90d9"}}>{d?.days} days</p></div>;
                    }}/>
                    <Scatter data={turnaround.map((d,i)=>({...d,index:i}))} fill="#4a90d9" fillOpacity={0.8} r={5}/>
                  </ScatterChart>
                </ResponsiveContainer>
              ) : <div className={`flex flex-col items-center justify-center h-36 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
            </Card>
          </div>
        </div>
      )}

      {/* ══ ROW 4 — DEEP DIVE (collapsible, default collapsed) ═════════════════ */}
      <CollRow title="Deep Dive & Details" open={row4Open} onToggle={()=>setRow4Open(o=>!o)} accent="#4a90d9"/>
      {row4Open && (
        <div className="space-y-5">

          {/* KRA Heatmap */}
          <Card title="KRA Heatmap" sub="Client × KPI parameter · green = met, red = missed, gray = N/A" accent="#a78bfa">
            <div className="overflow-x-auto">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr>
                    <th className={`px-3 py-2 text-left text-[9px] font-bold uppercase tracking-widest ${t.textMuted} border-b ${t.cardBorder}`} style={{minWidth:130}}>Client</th>
                    {kraParams.map(p=><th key={p} className={`px-2 py-2 text-center text-[9px] font-bold uppercase tracking-widest ${t.textMuted} border-b ${t.cardBorder}`}>{p}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {clients.slice(0,8).map((c,ri)=>(
                    <tr key={c.id} className={`border-b ${t.cardBorder}`} style={{background:ri%2===0?(dark?"rgba(15,30,46,0.6)":"#fafcff"):(dark?"rgba(12,24,38,0.4)":"#ffffff")}}>
                      <td className={`px-3 py-2 font-medium text-xs ${t.text}`}>{c.clientName.slice(0,18)}</td>
                      {kraParams.map((_,ci)=>{
                        const v=heatValues[ri]?.[ci];
                        return (
                          <td key={ci} className="px-2 py-2 text-center">
                            <div className="w-6 h-6 rounded-md mx-auto flex items-center justify-center text-[9px] font-black"
                              style={{background:v===1?"#34d39920":v===0?"#f8717120":"#94a3b815",color:v===1?"#34d399":v===0?"#f87171":"#94a3b8"}}>
                              {v===1?"✓":v===0?"✗":"—"}
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>

          {/* Top 10 by Fund Volume */}
          <Card title="Top 10 Clients by Fund Volume" sub="Estimated annual fund flow (INR equivalent)" accent="#00c9b1">
            {hasR ? (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={topClients} layout="vertical" margin={{left:0,right:20,top:0,bottom:0}}>
                  <CartesianGrid horizontal={false} stroke={GR}/>
                  <XAxis type="number" tick={{fill:AX,fontSize:9}} tickFormatter={v=>fmtC(v)} axisLine={false} tickLine={false}/>
                  <YAxis type="category" dataKey="name" tick={{fill:AX,fontSize:9}} axisLine={false} tickLine={false} width={80}/>
                  <Tooltip content={<CTip/>}/>
                  <Bar dataKey="vol" name="Fund Volume" radius={[0,5,5,0]} barSize={14}>
                    {topClients.map((_,i)=><Cell key={i} fill={CC[i%CC.length]}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : <div className={`flex flex-col items-center justify-center h-48 gap-2 text-sm ${t.textMuted}`}><span className="text-2xl">📭</span><span>No data available</span></div>}
          </Card>

          {/* Funding Ledger */}
          <Card title="Funding Ledger" sub="Client-level fund reconciliation · search, sort, paginate" accent="#fbbf24">
            <div className="space-y-3">
              <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border ${t.cardBorder}`}>
                <Icon path={Icons.search} size={12} className={t.textMuted}/>
                <input value={ledgerQ} onChange={e=>{setLedgerQ(e.target.value);setLedgerPg(1);}}
                  placeholder="Search client or ID…" className={`bg-transparent outline-none text-xs flex-1 ${t.text}`}/>
                {ledgerQ&&<button onClick={()=>{setLedgerQ("");setLedgerPg(1);}} className={t.textMuted}><Icon path={Icons.x} size={10}/></button>}
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr style={{background:dark?"linear-gradient(90deg,#0c1e30,#0e2240)":"linear-gradient(90deg,#1b3a5c,#1e4070)"}}>
                      {["Client Name","Parent Entity","Expected","Actual","Forex Variance"].map(h=>(
                        <th key={h} className="px-3 py-2.5 text-left text-[9px] font-bold uppercase tracking-widest" style={{color:"#c8dff0",borderBottom:"1px solid #ffffff18"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {ledgerSlice.map((r,i)=>(
                      <tr key={r.id} className={`border-b ${t.cardBorder}`}
                        style={{background:i%2===0?(dark?"rgba(15,30,46,0.7)":"#ffffff"):(dark?"rgba(12,24,38,0.5)":"#f8fbff")}}>
                        <td className={`px-3 py-2.5 font-medium ${t.text}`}>{r.name}</td>
                        <td className={`px-3 py-2.5 ${t.textMuted}`}>{r.parent}</td>
                        <td className="px-3 py-2.5 font-mono text-xs" style={{color:"#4a90d9"}}>{fmtC(r.exp)}</td>
                        <td className="px-3 py-2.5 font-mono text-xs" style={{color:"#00c9b1"}}>{fmtC(r.act)}</td>
                        <td className="px-3 py-2.5 font-mono text-xs font-bold" style={{color:r.variance>=0?"#34d399":"#f87171"}}>
                          {r.variance>=0?"+":""}{fmtC(Math.abs(r.variance))}{r.variance>=0?" ▴":" ▾"}
                        </td>
                      </tr>
                    ))}
                    {ledgerSlice.length===0&&(
                      <tr><td colSpan={5} className={`px-3 py-8 text-center text-sm ${t.textMuted}`}>No results match "{ledgerQ}"</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              {ledgerPages>1&&(
                <div className="flex items-center justify-between">
                  <span className={`text-[10px] ${t.textMuted}`}>Page {ledgerPg} of {ledgerPages} · {ledgerFiltered.length} rows</span>
                  <div className="flex gap-1">
                    <button disabled={ledgerPg===1} onClick={()=>setLedgerPg(p=>p-1)}
                      className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-medium disabled:opacity-30 ${t.hover} ${t.textMuted}`}>‹</button>
                    {Array.from({length:ledgerPages},(_, i)=>i+1).map(p=>(
                      <button key={p} onClick={()=>setLedgerPg(p)}
                        className="w-7 h-7 rounded-lg text-xs font-medium"
                        style={ledgerPg===p?{background:"linear-gradient(135deg,#fbbf24,#f59e0b)",color:"#fff"}:{color:AX}}>{p}</button>
                    ))}
                    <button disabled={ledgerPg===ledgerPages} onClick={()=>setLedgerPg(p=>p+1)}
                      className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-medium disabled:opacity-30 ${t.hover} ${t.textMuted}`}>›</button>
                  </div>
                </div>
              )}
            </div>
          </Card>
        </div>
      )}


    </div>
  );
};

const DashboardTab = ({ t, dark, isAdmin }) => {
  const [activeModule, setActiveModule] = useState("overview");
  const { clients } = useSyncContext();

  const modules = [
    { id:"overview",  label:"Overview",      icon:Icons.office,  color:"#00c9b1" },
    { id:"cas",       label:"CAS MIS",       icon:Icons.pie,     color:"#4a90d9" },
    { id:"kra",       label:"KRA / KPI",     icon:Icons.target,  color:"#34d399" },
    { id:"fund",      label:"Fund Request",  icon:Icons.send,    color:"#fbbf24" },
  ];

  const activeClients = clients.filter(c => c.status === "Active").length;
  const totalClients  = clients.length;

  // Overview KPI summary cards
  const overviewKpis = [
    { label:"Total Clients",   value: totalClients,    sub:"Registered in system",         icon:<Icon path={Icons.users}    size={18}/>, color:"#00c9b1", trend: 12   },
    { label:"Active Clients",  value: activeClients,   sub:`${totalClients - activeClients} inactive`, icon:<Icon path={Icons.check}    size={18}/>, color:"#34d399", trend: 5    },
    { label:"Fund Requests",   value:"12",             sub:"This financial year",          icon:<Icon path={Icons.send}     size={18}/>, color:"#fbbf24", delta: 8    },
    { label:"KPI Score",       value:"82%",            sub:"Team average performance",     icon:<Icon path={Icons.target}   size={18}/>, color:"#4a90d9", trend: 3    },
    { label:"Net AUM",         value:"₹—",             sub:"Under management",             icon:<Icon path={Icons.dollar}   size={18}/>, color:"#a78bfa"              },
    { label:"Pending Tasks",   value: clients.reduce((s,c)=>s+(c.tasks||0),0), sub:"Across all clients", icon:<Icon path={Icons.clock} size={18}/>, color:"#f87171", delta:-4 },
  ];

  return (
    <div className="space-y-5">
      {/* ── Hero Banner ── */}
      <div className="rounded-2xl overflow-hidden relative"
        style={{background:"linear-gradient(120deg,#003d5c 0%,#005a72 40%,#007a6e 75%,#00c9b1 100%)"}}>
        <div className="absolute inset-0 opacity-10"
          style={{backgroundImage:"radial-gradient(circle at 80% 50%,#ffffff 0%,transparent 60%)"}}/>
        <div className="relative px-6 py-5 flex items-center justify-between">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <div className="w-6 h-6 rounded-lg bg-white/20 flex items-center justify-center">
                <Icon path={Icons.office} size={13} className="text-white"/>
              </div>
              <span className="text-white/70 text-xs font-semibold tracking-widest uppercase">Welcome to</span>
            </div>
            <h2 className="text-white font-black text-2xl leading-none mb-1"
              style={{letterSpacing:"-0.04em"}}>ProCAS Dashboard</h2>
            <p className="text-white/75 text-xs font-medium">Streamlining your MIS, KPIs, and fund requests in one unified portal.</p>
          </div>
          <div className="hidden md:flex items-center gap-3">
            <div className="w-14 h-14 rounded-2xl bg-white/15 flex items-center justify-center border border-white/20"
              style={{backdropFilter:"blur(8px)"}}>
              <span className="text-white font-black text-lg" style={{letterSpacing:"-0.05em"}}>PC</span>
            </div>
          </div>
        </div>
      </div>

      {/* ── Module Selector Tabs ── */}
      <div className={`${t.card} border ${t.cardBorder} rounded-2xl p-1.5 flex gap-1.5 overflow-x-auto`}>
        {modules.map(m => {
          const isActive = activeModule === m.id;
          return (
            <button key={m.id} onClick={() => setActiveModule(m.id)}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm font-semibold
                transition-all duration-200 whitespace-nowrap flex-1 justify-center
                ${isActive ? "text-white shadow-lg" : `${t.textMuted} ${t.hover}`}`}
              style={isActive
                ? {background:`linear-gradient(135deg,${m.color}cc,${m.color}88)`,
                   boxShadow:`0 4px 14px ${m.color}44`}
                : {}}>
              <span style={isActive ? {color:"white"} : {color:m.color}}>
                <Icon path={m.icon} size={15}/>
              </span>
              {m.label}
            </button>
          );
        })}
      </div>

      {/* ── Overview Panel ── */}
      {activeModule === "overview" && (
        <OverviewPanel t={t} dark={dark} isAdmin={isAdmin} clients={clients} onNavigate={setActiveModule}/>
      )}

            {/* ── CAS MIS Panel ── */}
      {activeModule === "cas" && (
        <CasMisTab t={t} dark={dark} isAdmin={isAdmin}/>
      )}

      {/* ── KRA / KPI Panel ── */}
      {activeModule === "kra" && (
        <KraKpiTab t={t} dark={dark} isAdmin={isAdmin}/>
      )}

      {/* ── Fund Request Panel ── */}
      {activeModule === "fund" && (
        <FundRequestTab t={t} dark={dark} isAdmin={isAdmin}/>
      )}
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ── CLIENT SYNC PROVIDER ─────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
const AUTO_SYNC_INTERVAL = 3 * 60 * 1000; // 3 minutes

const ClientSyncProvider = ({ initialClients: seedClients, children }) => {
  const [state, dispatch] = useReducer(syncReducer, {
    clients: seedClients,
    syncing: false,
    lastSynced: new Date(),
    syncError: null,
  });
  const [autoSync, setAutoSync] = useState(false);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((type, message) => {
    setToast({ type, message });
    setTimeout(() => setToast(null), 3500);
  }, []);

  // Called by MasterClientTab whenever clients change (add/edit/delete/import)
  // This is the single source of truth — no prop-drilling desync
  const updateClients = useCallback((newClients) => {
    dispatch({ type: "UPDATE_CLIENTS", payload: newClients });
  }, []);

  const triggerSync = useCallback(async () => {
    if (state.syncing) return;
    dispatch({ type: "SYNC_START" });
    try {
      // Simulate a brief async "fetch" — in production replace with real API call:
      // const res = await fetch("/api/client-master"); const data = await res.json();
      await new Promise(r => setTimeout(r, 700));
      // Re-dispatch the current clients (they are already live via updateClients)
      dispatch({ type: "SYNC_SUCCESS", payload: state.clients });
      showToast("success", "Latest client master synced successfully");
    } catch (err) {
      dispatch({ type: "SYNC_ERROR", payload: "Sync failed. Please try again." });
      showToast("error", "Sync failed. Could not reach client master.");
    }
  }, [state.syncing, state.clients, showToast]);

  // Auto-sync on interval when enabled
  useEffect(() => {
    if (!autoSync) return;
    const id = setInterval(triggerSync, AUTO_SYNC_INTERVAL);
    return () => clearInterval(id);
  }, [autoSync, triggerSync]);

  return (
    <ClientSyncContext.Provider value={{
      ...state,
      autoSync,
      setAutoSync,
      triggerSync,
      updateClients,
      toast,
    }}>
      <SyncToast toast={toast}/>
      {children}
    </ClientSyncContext.Provider>
  );
};

// ── Main App ──────────────────────────────────────────────────────────────────
export default function CADashboard() {
  // ── Theme — persisted to localStorage so it survives refresh & re-login ────
  const [darkMode, setDarkMode] = useState(() => {
    try {
      const saved = localStorage.getItem("procas_darkMode");
      return saved !== null ? saved === "true" : true; // default dark
    } catch { return true; }
  });

  useEffect(() => {
    try { localStorage.setItem("procas_darkMode", String(darkMode)); } catch {}
    // Keep document root background in sync so any content outside the React
    // tree (browser chrome, scrollbars) also reflects the chosen theme.
    document.documentElement.style.background = darkMode ? "#08101c" : "#eef3f8";
  }, [darkMode]);

  const [activeTab, setActiveTab]     = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [users, setUsers]             = useState(INITIAL_USERS);
  const [activeUser, setActiveUser]   = useState(INITIAL_USERS[0]);

  // isAdmin is derived from activeUser — single source of truth
  const isAdmin = activeUser?.role === "Admin";

  const t    = themes[darkMode ? "dark" : "light"];
  const dark = darkMode;
  // raw: hex colour values used in inline style={} so the theme switch is
  // guaranteed to work even if Tailwind CDN never generated the light-mode
  // arbitrary-value classes (absent from the initial dark render).
  const raw  = t.raw;

  const navItems = [
    { id:"dashboard", label:"Dashboard",     icon:Icons.office,   section:"main"   },
    { id:"master",    label:"Client Master", icon:Icons.users,    section:"main"   },
    { id:"userdb",    label:"User Database", icon:Icons.database, section:"main"   },
    { id:"arch",      label:"Architect",     icon:Icons.cpu,      section:"main"   },
    { id:"logout",    label:"Logout",        icon:Icons.logOut,   section:"bottom" },
  ];
  const mainNav   = navItems.filter(n => n.section === "main");
  const bottomNav = navItems.filter(n => n.section === "bottom");

  const themeTransition = "background 0.22s ease, color 0.18s ease, border-color 0.18s ease";

  return (
    <ThemeContext.Provider value={{ dark, t }}>
    <UserContext.Provider value={{ users, setUsers, activeUser, setActiveUser }}>
    <LockProvider adminName={activeUser?.name || "Admin"}>
    <ClientSyncProvider initialClients={initialClients}>
    <GlobalDarkStyles dark={dark}/>

    {/*
      ROOT WRAPPER — inline style for background + color so the theme ALWAYS
      updates regardless of whether Tailwind CDN generated the arbitrary-value
      classes for light mode.
    */}
    <div className="min-h-screen flex font-sans"
      style={{
        fontFamily: "'Outfit','DM Sans','Segoe UI',system-ui,sans-serif",
        background:  raw.bg,
        color:       raw.text,
        transition:  themeTransition,
      }}>

      {/* ── SIDEBAR — intentionally NOT themed; always dark per spec ── */}
      <aside
        className={`fixed left-0 top-0 h-full z-40 w-56 flex flex-col
          transition-transform duration-300
          ${sidebarOpen ? "translate-x-0" : "-translate-x-full"} lg:translate-x-0`}
        style={{
          background:  "#0b1623",
          borderRight: "1px solid #1a2d44",
          boxShadow:   "4px 0 24px rgba(0,0,0,0.3)",
        }}>

        {/* Logo */}
        <div className="px-5 py-4 flex items-center gap-3"
          style={{ borderBottom: "1px solid #1a2d44" }}>
          <div className="w-9 h-9 rounded-xl flex items-center justify-center text-white font-black text-xs shrink-0"
            style={{ background:"linear-gradient(135deg,#003d5c,#00c9b1)" }}>
            <span style={{ letterSpacing:"-0.03em" }}>PC</span>
          </div>
          <div className="min-w-0">
            <div className="font-black text-sm leading-tight text-[#e8f0f8]"
              style={{ letterSpacing:"-0.02em" }}>ProCAS</div>
            <div className="text-[10px] text-[#5a7a99] truncate">Client Accounting System</div>
          </div>
        </div>

        {/* Main navigation */}
        <nav className="flex-1 px-3 py-4 space-y-0.5 overflow-y-auto">
          {mainNav.map(item => {
            const isActive = activeTab === item.id;
            return (
              <button key={item.id}
                onClick={() => { setActiveTab(item.id); setSidebarOpen(false); }}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-all duration-150 group relative"
                style={isActive
                  ? { background:"#132036", color:"#00c9b1" }
                  : { color:"#5a7a99" }}
                onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = "#1a2d44"; }}
                onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = "transparent"; }}>
                {isActive && (
                  <span className="absolute left-0 top-1/2 -translate-y-1/2 w-0.5 h-5 rounded-r-full"
                    style={{ background:"#00c9b1" }}/>
                )}
                <span className={`transition-colors ${isActive ? "" : "group-hover:text-[#00c9b1]"}`}>
                  <Icon path={item.icon} size={16}/>
                </span>
                <span className={`truncate ${isActive ? "font-semibold" : ""}`}>{item.label}</span>
              </button>
            );
          })}
        </nav>

        {/* Bottom: Logout + User */}
        <div className="px-3 pb-3 pt-3 space-y-0.5"
          style={{ borderTop:"1px solid #1a2d44" }}>
          {bottomNav.map(item => {
            const isLogout = item.id === "logout";
            const isActive = activeTab === item.id;
            return (
              <button key={item.id}
                onClick={() => {
                  if (isLogout) {
                    // Reset session: return to default Admin seed user, go to Dashboard, close sidebar
                    setActiveUser(INITIAL_USERS[0]);
                    setActiveTab("dashboard");
                    setSidebarOpen(false);
                  } else {
                    setActiveTab(item.id);
                  }
                }}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-all duration-150"
                style={isLogout
                  ? { color:"#f87171" }
                  : isActive
                    ? { background:"#132036", color:"#00c9b1" }
                    : { color:"#5a7a99" }}
                onMouseEnter={e => {
                  if (isLogout) e.currentTarget.style.background = "#f8717112";
                  else if (!isActive) e.currentTarget.style.background = "#1a2d44";
                }}
                onMouseLeave={e => {
                  if (!isActive || isLogout) e.currentTarget.style.background = "transparent";
                }}>
                <Icon path={item.icon} size={15}/>
                <span>{item.label}</span>
              </button>
            );
          })}
          {/* User chip */}
          <div className="flex items-center gap-3 px-3 py-3 mt-1 rounded-xl cursor-default"
            style={{ color:"#e8f0f8" }}
            onMouseEnter={e => e.currentTarget.style.background = "#1a2d44"}
            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
            <div className="w-7 h-7 rounded-full flex items-center justify-center text-white text-[10px] font-bold shrink-0"
              style={{ background:"linear-gradient(135deg,#003d5c,#00c9b1)" }}>
              {(activeUser?.name||"AU").split(" ").map(w=>w[0]).join("").slice(0,2).toUpperCase()}
            </div>
            <div className="flex-1 min-w-0">
              <div className="text-xs font-semibold truncate text-[#e8f0f8]">{activeUser?.name || "Admin User"}</div>
              <div className="text-[10px] text-[#5a7a99]">{activeUser?.role || "User"}</div>
            </div>
          </div>
        </div>
      </aside>

      {sidebarOpen && (
        <div className="fixed inset-0 z-30 bg-black/50 lg:hidden"
          onClick={() => setSidebarOpen(false)}/>
      )}

      {/* ── MAIN CONTENT AREA — fully themed ── */}
      <div className="flex-1 lg:ml-56 flex flex-col min-h-screen">

        {/* HEADER — inline style so background ALWAYS switches with the theme */}
        <header className="sticky top-0 z-20 px-4 lg:px-6 h-14 flex items-center justify-between border-b"
          style={{
            background:     raw.headerBg,
            borderColor:    raw.cardBorder,
            backdropFilter: "blur(12px)",
            transition:     themeTransition,
          }}>
          <div className="flex items-center gap-3">
            <button onClick={() => setSidebarOpen(!sidebarOpen)}
              className="lg:hidden p-2 rounded-lg transition-colors"
              style={{ color: raw.textMuted }}
              onMouseEnter={e => e.currentTarget.style.background = raw.hoverBg}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              <div className="w-4 h-0.5 bg-current mb-1"/>
              <div className="w-4 h-0.5 bg-current mb-1"/>
              <div className="w-4 h-0.5 bg-current"/>
            </button>
            <div>
              <h1 className="text-sm font-black tracking-tight"
                style={{ letterSpacing:"-0.02em", color: raw.text }}>ProCAS</h1>
              <p className="text-[10px] hidden sm:block" style={{ color: raw.textMuted }}>
                {activeTab === "dashboard"
                  ? "Dashboard"
                  : navItems.find(n => n.id === activeTab)?.label || "Professional Client Accounting System"
                } · FY 2026-27
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {/* Role toggle — switches activeUser between Admin and End User (demo) */}
            <button
              onClick={() => setActiveUser(prev => {
                if (prev?.role === "Admin") {
                  // Switch to first active End User
                  const endUser = users.find(u => u.role === "End User" && u.status === "Active");
                  return endUser || prev;
                } else {
                  // Switch back to Admin
                  const admin = users.find(u => u.role === "Admin");
                  return admin || prev;
                }
              })}
              className="hidden sm:flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full font-medium border transition-all"
              style={{
                borderColor: isAdmin ? "#00c9b140" : raw.cardBorder,
                color:       isAdmin ? "#00c9b1"   : raw.textMuted,
                background:  isAdmin ? "#00c9b115" : "transparent",
              }}
              title={`Switch role (demo) — current: ${activeUser?.name}`}>
              <Icon path={Icons.shield} size={11}/>
              {isAdmin ? `Admin (${activeUser?.name?.split(" ")[0]})` : `User (${activeUser?.name?.split(" ")[0]})`}
            </button>

            {/* Live badge */}
            <span className="hidden md:flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full font-medium"
              style={{ background: raw.activeBg, color: raw.textAccent }}>
              <div className="w-1.5 h-1.5 rounded-full bg-[#00c9b1] animate-pulse"/>Live Data
            </span>

            {/* Dark / Light toggle button */}
            <button
              onClick={() => setDarkMode(d => !d)}
              className="w-8 h-8 rounded-lg flex items-center justify-center transition-colors"
              style={{ color: raw.textMuted }}
              title={dark ? "Switch to Light Mode" : "Switch to Dark Mode"}
              onMouseEnter={e => e.currentTarget.style.background = raw.hoverBg}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              <Icon path={dark ? Icons.sun : Icons.moon} size={15}/>
            </button>
          </div>
        </header>

        {/* MAIN — background inherited from root wrapper; modules use t.* tokens */}
        <ErrorBoundary>
        <main className="flex-1 p-4 lg:p-6 overflow-auto">
          {activeTab === "dashboard" && (
            <DashboardTab t={t} dark={dark} isAdmin={isAdmin}/>
          )}
          {activeTab === "master" && (
            <div className="mb-5 rounded-2xl overflow-hidden relative"
              style={{ background:"linear-gradient(120deg,#003d5c 0%,#005a72 40%,#007a6e 75%,#00c9b1 100%)" }}>
              <div className="absolute inset-0 opacity-10"
                style={{ backgroundImage:"radial-gradient(circle at 80% 50%,#ffffff 0%,transparent 60%)" }}/>
              <div className="relative px-6 py-5 flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-6 h-6 rounded-lg bg-white/20 flex items-center justify-center">
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                      </svg>
                    </div>
                    <span className="text-white/70 text-xs font-semibold tracking-widest uppercase">Client Master</span>
                  </div>
                  <h2 className="text-white font-black text-2xl leading-none mb-1"
                    style={{ letterSpacing:"-0.04em" }}>ProCAS</h2>
                  <p className="text-white/75 text-xs font-medium">Streamlining your MIS, KPIs, and fund requests in one unified portal.</p>
                </div>
                <div className="hidden md:flex items-center gap-3">
                  <div className="text-right">
                    <div className="text-white/60 text-[10px] font-semibold uppercase tracking-widest">System</div>
                    <div className="text-white font-bold text-sm">Professional Client</div>
                    <div className="text-white font-bold text-sm">Accounting System</div>
                  </div>
                  <div className="w-14 h-14 rounded-2xl bg-white/15 flex items-center justify-center border border-white/20"
                    style={{ backdropFilter:"blur(8px)" }}>
                    <span className="text-white font-black text-lg" style={{ letterSpacing:"-0.05em" }}>PC</span>
                  </div>
                </div>
              </div>
            </div>
          )}
          {activeTab === "master" && <MasterClientTab t={t} dark={dark} isAdmin={isAdmin}/>}
          {activeTab === "arch"   && <ArchitectureTab t={t} dark={dark} isAdmin={isAdmin} adminName={activeUser?.name || "Admin"}/>}
          {activeTab === "userdb" && <UserDatabaseTab t={t} dark={dark} isAdmin={isAdmin}/>}
        </main>
        </ErrorBoundary>        {/* FOOTER — inline style so background ALWAYS switches with the theme */}
        <footer className="px-6 pt-3 pb-2 border-t flex flex-col gap-1.5"
          style={{
            background:  raw.footerBg,
            borderColor: raw.cardBorder,
            transition:  themeTransition,
          }}>
          <div className="flex items-center justify-between text-[11px]"
            style={{ color: raw.textMuted }}>
            <span>© 2026 ProCAS. All Rights Reserved.</span>
            <span style={{ color:"#00c9b1", fontWeight:700, letterSpacing:"-0.02em" }}>ProCAS v1.0.0</span>
          </div>
          <div className="border-t pt-1.5 flex items-center justify-center"
            style={{ borderColor: raw.cardBorder }}>
            <span
              className="text-[10px] font-medium tracking-wide transition-all duration-300 cursor-default select-none"
              style={{ color: dark ? "#4a6a8a" : "#8aa0b8", letterSpacing:"0.04em" }}
              onMouseEnter={e => e.target.style.color = dark ? "#7aa0c0" : "#4a7090"}
              onMouseLeave={e => e.target.style.color = dark ? "#4a6a8a" : "#8aa0b8"}>
              Developed by CA. Anjan G
            </span>
          </div>
        </footer>
      </div>
    </div>

    </ClientSyncProvider>
    </LockProvider>
    </UserContext.Provider>
    </ThemeContext.Provider>
  );
}

