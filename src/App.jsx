import { useState, useMemo, useEffect, useRef } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, HeadingLevel, PageBreak,
} from 'docx';
import './App.css';

// crypto.randomUUID is only available in secure contexts (HTTPS / localhost),
// so it fails over LAN IP on phone. Fallback to a v4-shaped random id.
const uuid = () => {
  try {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
      return crypto.randomUUID();
    }
  } catch {}
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
};

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const SUBJECT_PALETTE = [
  '#334155', '#1e40af', '#0f766e', '#166534', '#9a3412', '#991b1b',
  '#6b21a8', '#831843', '#3730a3', '#115e59', '#92400e', '#1f2937',
];

const DEFAULT_PDF_STYLE = {
  headerBg: '#2c3e50',
  cornerBg: '#1a1a1a',
  headerText: '#ffffff',
  titleColor: '#1a1a1a',
  subtitleColor: '#444444',
};

// Sample seed — used the first time the app runs (or after a reset).
const buildDefaultState = () => {
  const subjects = [
    { id: 'sub-math',  name: 'Mathematics', color: '#1e40af' },
    { id: 'sub-eng',   name: 'English',     color: '#0f766e' },
    { id: 'sub-sci',   name: 'Science',     color: '#166534' },
    { id: 'sub-his',   name: 'History',     color: '#9a3412' },
    { id: 'sub-geo',   name: 'Geography',   color: '#6b21a8' },
    { id: 'sub-art',   name: 'Art',         color: '#831843' },
  ];
  const teachers = [
    { id: 't-mitchell', name: 'Sarah Mitchell',  subjectIds: ['sub-math'] },
    { id: 't-chen',     name: 'David Chen',      subjectIds: ['sub-eng'] },
    { id: 't-rahman',   name: 'Aisha Rahman',    subjectIds: ['sub-sci'] },
    { id: 't-oconnor',  name: "James O'Connor",  subjectIds: ['sub-his'] },
    { id: 't-sharma',   name: 'Priya Sharma',    subjectIds: ['sub-geo'] },
    { id: 't-becker',   name: 'Thomas Becker',   subjectIds: ['sub-art'] },
  ];
  const classes = ['Grade 6', 'Grade 7', 'Grade 8'];
  const periods = ['08:30 — 09:15', '09:20 — 10:05', '10:10 — 10:55', '11:15 — 12:00', '12:05 — 12:50'];
  const cell = (t, s, text = '') => ({ teacherId: t, subjectId: s, text });
  const grid = [
    [cell('t-mitchell','sub-math'), cell('t-chen','sub-eng'),    cell('t-rahman','sub-sci')],
    [cell('t-chen','sub-eng'),      cell('t-rahman','sub-sci'),  cell('t-mitchell','sub-math')],
    [cell('t-oconnor','sub-his'),   cell('t-sharma','sub-geo'),  cell('t-becker','sub-art')],
    [cell('t-sharma','sub-geo'),    cell('t-becker','sub-art'),  cell('t-oconnor','sub-his')],
    [cell('t-becker','sub-art'),    cell('t-mitchell','sub-math'),cell('t-chen','sub-eng')],
  ];
  return {
    schoolName: 'Greenwood International School',
    subtitle: 'Weekly Class Timetable — 2026',
    classes, periods, teachers, subjects, grid,
    footerText: 'All teachers are required to report to their assigned classroom 5 minutes before the period begins. Any schedule change must be approved by the academic coordinator.',
    pdfStyle: DEFAULT_PDF_STYLE,
  };
};

export default function App() {
  // ---- Theme ----
  const [theme, setTheme] = useState(() => localStorage.getItem('tt-theme') || 'light');
  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('tt-theme', theme);
  }, [theme]);

  // ---- App state (empty until server load completes; prevents flash of seed) ----
  const [schoolName, setSchoolName] = useState('');
  const [subtitle, setSubtitle]     = useState('');
  const [classes, setClasses]       = useState([]);
  const [periods, setPeriods]       = useState([]);
  const [teachers, setTeachers]     = useState([]);
  const [subjects, setSubjects]     = useState([]);
  const [grid, setGrid]             = useState([]);
  const [footerText, setFooterText] = useState('');
  const [pdfStyle, setPdfStyle]     = useState(DEFAULT_PDF_STYLE);

  const applyState = (s) => {
    setSchoolName(s.schoolName ?? '');
    setSubtitle(s.subtitle ?? '');
    setClasses(s.classes ?? []);
    setPeriods(s.periods ?? []);
    setTeachers(s.teachers ?? []);
    setSubjects(s.subjects ?? []);
    setGrid(s.grid ?? []);
    setFooterText(s.footerText ?? '');
    setPdfStyle(s.pdfStyle ?? DEFAULT_PDF_STYLE);
  };

  const [teacherInput, setTeacherInput] = useState('');
  const [subjectInput, setSubjectInput] = useState('');
  const [selectedTeacher, setSelectedTeacher] = useState(null);
  const [sidebarTab, setSidebarTab] = useState('teachers');
  const [filter, setFilter] = useState('');
  const [dragged, setDragged] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [exportMenu, setExportMenu] = useState(false);
  const importInputRef = useRef(null);

  // close export menu on outside click
  useEffect(() => {
    if (!exportMenu) return;
    const close = (e) => {
      if (!e.target.closest('.export-menu-wrap')) setExportMenu(false);
    };
    window.addEventListener('mousedown', close);
    return () => window.removeEventListener('mousedown', close);
  }, [exportMenu]);

  // ---- Persistence (localStorage) ----
  // Two slots:
  //   tt-state    — working/autosaved state (every edit)
  //   tt-snapshot — explicit "Save" snapshot (user-confirmed default)
  //   tt-savedAt  — timestamp of last explicit save
  const STORAGE_STATE = 'tt-state';
  const STORAGE_SNAPSHOT = 'tt-snapshot';
  const STORAGE_SAVED_AT = 'tt-savedAt';

  const [loaded, setLoaded] = useState(false);
  const [savedAt, setSavedAt] = useState(null);
  const [dirty, setDirty] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle'); // idle|saving|saved|error
  const lastSavedJson = useRef('');

  // Initial load — synchronous from localStorage.
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_STATE) || localStorage.getItem(STORAGE_SNAPSHOT);
      const stored = raw ? JSON.parse(raw) : null;
      applyState(stored || buildDefaultState());
      const ts = localStorage.getItem(STORAGE_SAVED_AT);
      setSavedAt(ts || null);
      const snap = localStorage.getItem(STORAGE_SNAPSHOT);
      if (snap) lastSavedJson.current = snap;
    } catch (e) {
      console.warn('Could not load from localStorage; using seed.', e);
      applyState(buildDefaultState());
    } finally {
      setLoaded(true);
    }
  }, []);

  const stateBundle = useMemo(() => ({
    schoolName, subtitle, classes, periods, teachers, subjects, grid, footerText, pdfStyle,
  }), [schoolName, subtitle, classes, periods, teachers, subjects, grid, footerText, pdfStyle]);

  // Autosave working state to localStorage on every change (debounced).
  const autosaveTimer = useRef(null);
  useEffect(() => {
    if (!loaded) return;
    const json = JSON.stringify(stateBundle);
    setDirty(json !== lastSavedJson.current);
    clearTimeout(autosaveTimer.current);
    autosaveTimer.current = setTimeout(() => {
      try { localStorage.setItem(STORAGE_STATE, json); } catch {}
    }, 250);
    return () => clearTimeout(autosaveTimer.current);
  }, [stateBundle, loaded]);

  const [confirmSave, setConfirmSave] = useState(false);

  const requestSave = () => {
    // First save: no confirmation needed. Subsequent saves: confirm overwrite.
    if (savedAt) setConfirmSave(true);
    else doSave();
  };

  const doSave = async () => {
    setConfirmSave(false);
    setSaveStatus('saving');
    try {
      const json = JSON.stringify(stateBundle);
      const ts = new Date().toISOString();
      localStorage.setItem(STORAGE_SNAPSHOT, json);
      localStorage.setItem(STORAGE_STATE, json);
      localStorage.setItem(STORAGE_SAVED_AT, ts);
      lastSavedJson.current = json;
      setSavedAt(ts);
      setDirty(false);
      setSaveStatus('saved');
      setTimeout(() => setSaveStatus('idle'), 1500);
    } catch (e) {
      console.error(e);
      setSaveStatus('error');
      setTimeout(() => setSaveStatus('idle'), 2000);
    }
  };

  // legacy alias for any existing call sites
  const saveSnapshot = requestSave;

  // Ctrl/Cmd+S
  useEffect(() => {
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        saveSnapshot();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  // ---- CRUD ----
  const addTeacher = () => {
    const name = teacherInput.trim();
    if (!name) return;
    const newT = { id: uuid(), name, subjectIds: [] };
    setTeachers((t) => [...t, newT]);
    setTeacherInput('');
    setSelectedTeacher(newT.id);
  };
  const removeTeacher = (id) => {
    setTeachers((t) => t.filter((x) => x.id !== id));
    setGrid((g) => g.map((row) => row.map((c) => (c.teacherId === id ? { ...c, teacherId: null } : c))));
    if (selectedTeacher === id) setSelectedTeacher(null);
  };
  const renameTeacher = (id, name) =>
    setTeachers((ts) => ts.map((t) => (t.id === id ? { ...t, name } : t)));
  const toggleTeacherSubject = (teacherId, subjectId) =>
    setTeachers((ts) => ts.map((t) =>
      t.id !== teacherId ? t : {
        ...t,
        subjectIds: t.subjectIds.includes(subjectId)
          ? t.subjectIds.filter((s) => s !== subjectId)
          : [...t.subjectIds, subjectId],
      }
    ));

  const addSubject = () => {
    const name = subjectInput.trim();
    if (!name) return;
    const color = SUBJECT_PALETTE[subjects.length % SUBJECT_PALETTE.length];
    setSubjects((s) => [...s, { id: uuid(), name, color }]);
    setSubjectInput('');
  };
  const removeSubject = (id) => {
    setSubjects((s) => s.filter((x) => x.id !== id));
    setTeachers((ts) => ts.map((t) => ({ ...t, subjectIds: t.subjectIds.filter((x) => x !== id) })));
    setGrid((g) => g.map((row) => row.map((c) => (c.subjectId === id ? { ...c, subjectId: null } : c))));
  };
  const renameSubject = (id, name) =>
    setSubjects((ss) => ss.map((s) => (s.id === id ? { ...s, name } : s)));

  const resizeGrid = (np, nc) => {
    setGrid((g) => {
      const out = [];
      for (let p = 0; p < np; p++) {
        const row = [];
        for (let c = 0; c < nc; c++) row.push(g[p]?.[c] ?? { teacherId: null, subjectId: null, text: '' });
        out.push(row);
      }
      return out;
    });
  };
  const addClass = () => {
    const next = [...classes, `Class ${String.fromCharCode(65 + classes.length)}`];
    setClasses(next); resizeGrid(periods.length, next.length);
  };
  const removeClass = (idx) => {
    if (classes.length <= 1) return;
    setClasses(classes.filter((_, i) => i !== idx));
    setGrid((g) => g.map((row) => row.filter((_, i) => i !== idx)));
  };
  const addPeriod = () => {
    const next = [...periods, `Period ${periods.length + 1}`];
    setPeriods(next); resizeGrid(next.length, classes.length);
  };
  const removePeriod = (idx) => {
    if (periods.length <= 1) return;
    setPeriods(periods.filter((_, i) => i !== idx));
    setGrid((g) => g.filter((_, i) => i !== idx));
  };
  const updateClassName = (idx, val) => setClasses((c) => c.map((x, i) => (i === idx ? val : x)));
  const updatePeriodName = (idx, val) => setPeriods((p) => p.map((x, i) => (i === idx ? val : x)));

  const updateCell = (p, c, patch) =>
    setGrid((g) => g.map((row, ri) => row.map((cell, ci) => (ri === p && ci === c ? { ...cell, ...patch } : cell))));

  // ---- Conflicts ----
  const teacherConflicts = useMemo(() => {
    const set = new Set();
    grid.forEach((row, p) => {
      const seen = {};
      row.forEach((cell, c) => {
        if (cell.teacherId) (seen[cell.teacherId] = seen[cell.teacherId] || []).push(c);
      });
      Object.values(seen).forEach((cols) => {
        if (cols.length > 1) cols.forEach((c) => set.add(`${p}:${c}`));
      });
    });
    return set;
  }, [grid]);
  const subjectConflicts = useMemo(() => {
    const set = new Set();
    if (!classes.length) return set;
    for (let c = 0; c < classes.length; c++) {
      const seen = {};
      grid.forEach((row, p) => {
        const cell = row[c];
        if (cell?.subjectId) (seen[cell.subjectId] = seen[cell.subjectId] || []).push(p);
      });
      Object.values(seen).forEach((rows) => {
        if (rows.length > 1) rows.forEach((p) => set.add(`${p}:${c}`));
      });
    }
    return set;
  }, [grid, classes]);

  // ---- Availability ----
  const availability = useMemo(() => periods.map((per, pi) => {
    const used = new Set();
    (grid[pi] || []).forEach((c) => { if (c.teacherId) used.add(c.teacherId); });
    return { period: per, free: teachers.filter((t) => !used.has(t.id)), usedCount: used.size };
  }), [periods, teachers, grid]);

  // ---- Subject coverage ----
  const subjectCoverage = useMemo(() => {
    const map = {};
    subjects.forEach((s) => (map[s.id] = 0));
    teachers.forEach((t) => t.subjectIds.forEach((sid) => { if (map[sid] != null) map[sid]++; }));
    return map;
  }, [subjects, teachers]);

  const teacherById = (id) => teachers.find((t) => t.id === id);
  const subjectById = (id) => subjects.find((s) => s.id === id);
  const initials = (name) => name.split(/\s+/).filter(Boolean).slice(0, 2).map((p) => p[0].toUpperCase()).join('') || '?';

  const onDropCell = (p, c) => {
    if (!dragged) return;
    if (dragged.type === 'teacher') updateCell(p, c, { teacherId: dragged.id });
    else if (dragged.type === 'subject') updateCell(p, c, { subjectId: dragged.id });
    setDragged(null);
  };

  // ---- PDF ----
  const buildHeader = (subLabel) => {
    const header = document.createElement('div');
    header.style.cssText = 'text-align:center;border-bottom:3px double #1a1a1a;padding-bottom:16px;margin-bottom:24px;';
    if (schoolName.trim()) {
      const h1 = document.createElement('div');
      h1.textContent = schoolName.trim();
      h1.style.cssText = `font-size:32px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${pdfStyle.titleColor};`;
      header.appendChild(h1);
    }
    if (subtitle.trim() || subLabel) {
      const h2 = document.createElement('div');
      h2.textContent = subLabel || subtitle.trim();
      h2.style.cssText = `font-size:16px;font-style:italic;color:${pdfStyle.subtitleColor};margin-top:6px;`;
      header.appendChild(h2);
    }
    const dateLine = document.createElement('div');
    dateLine.textContent = `Issued: ${new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })}`;
    dateLine.style.cssText = 'font-size:12px;color:#666;margin-top:8px;';
    header.appendChild(dateLine);
    return header;
  };

  const buildScheduleNode = () => {
    const root = document.createElement('div');
    root.style.cssText = 'width:1200px;padding:40px;background:#fff;font-family:"Times New Roman",Georgia,serif;color:#1a1a1a;';
    root.appendChild(buildHeader());
    const table = document.createElement('table');
    table.style.cssText = 'width:100%;border-collapse:collapse;font-family:"Helvetica Neue",Arial,sans-serif;font-size:13px;';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    ['Period / Class', ...classes].forEach((label, i) => {
      const th = document.createElement('th');
      th.textContent = label;
      th.style.cssText = `border:1px solid #1a1a1a;padding:10px 8px;background:${i === 0 ? pdfStyle.cornerBg : pdfStyle.headerBg};color:${pdfStyle.headerText};font-weight:600;font-size:13px;letter-spacing:0.3px;text-align:center;`;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    periods.forEach((per, pi) => {
      const tr = document.createElement('tr');
      const rh = document.createElement('th');
      rh.textContent = per;
      rh.style.cssText = 'border:1px solid #1a1a1a;padding:10px 8px;background:#ecf0f1;font-weight:600;text-align:center;color:#1a1a1a;';
      tr.appendChild(rh);
      classes.forEach((_, ci) => {
        const cell = grid[pi]?.[ci] ?? { teacherId: null, subjectId: null, text: '' };
        const t = cell.teacherId ? teacherById(cell.teacherId) : null;
        const s = cell.subjectId ? subjectById(cell.subjectId) : null;
        const tConflict = teacherConflicts.has(`${pi}:${ci}`);
        const sConflict = subjectConflicts.has(`${pi}:${ci}`);
        const bg = tConflict ? '#fde8e8' : sConflict ? '#fef3c7' : (pi % 2 === 0 ? '#fff' : '#fafbfc');
        const td = document.createElement('td');
        td.style.cssText = `border:1px solid #1a1a1a;padding:8px;vertical-align:middle;text-align:center;height:64px;background:${bg};`;
        if (t) {
          const name = document.createElement('div');
          name.textContent = t.name;
          name.style.cssText = 'font-weight:700;color:#1a1a1a;font-size:13px;line-height:1.25;';
          td.appendChild(name);
        }
        if (s) {
          const sub = document.createElement('div');
          sub.textContent = s.name;
          sub.style.cssText = `font-size:11px;font-weight:600;color:${s.color};margin-top:${t ? '2px' : '0'};letter-spacing:0.2px;`;
          td.appendChild(sub);
        }
        if (cell.text) {
          const txt = document.createElement('div');
          txt.textContent = cell.text;
          txt.style.cssText = `font-size:10.5px;color:#555;margin-top:${t || s ? '3px' : '0'};font-style:italic;`;
          td.appendChild(txt);
        }
        if (!t && !s && !cell.text) {
          const dash = document.createElement('div');
          dash.textContent = '—';
          dash.style.color = '#bbb';
          td.appendChild(dash);
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    root.appendChild(table);

    if (footerText.trim()) {
      const f = document.createElement('div');
      f.textContent = footerText.trim();
      f.style.cssText = 'margin-top:24px;padding:14px 16px;border-left:3px solid #2c3e50;background:#f7f9fa;font-size:12px;line-height:1.6;color:#333;white-space:pre-wrap;font-family:"Helvetica Neue",Arial,sans-serif;';
      root.appendChild(f);
    }
    const sig = document.createElement('div');
    sig.style.cssText = 'margin-top:40px;display:flex;justify-content:space-between;font-size:11px;color:#666;font-family:"Helvetica Neue",Arial,sans-serif;';
    sig.innerHTML = `
      <div style="border-top:1px solid #1a1a1a;padding-top:6px;width:200px;text-align:center;">Prepared By</div>
      <div style="border-top:1px solid #1a1a1a;padding-top:6px;width:200px;text-align:center;">Principal / Head</div>`;
    root.appendChild(sig);
    return root;
  };

  const buildAvailabilityNode = () => {
    const root = document.createElement('div');
    root.style.cssText = 'width:1200px;padding:40px;background:#fff;font-family:"Times New Roman",Georgia,serif;color:#1a1a1a;';
    root.appendChild(buildHeader('Teacher Availability — by Period'));
    const table = document.createElement('table');
    table.style.cssText = 'width:100%;border-collapse:collapse;font-family:"Helvetica Neue",Arial,sans-serif;font-size:13px;';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    [
      { label: 'Period', w: '180px' },
      { label: 'Free / Total', w: '140px' },
      { label: 'Available Teachers', w: 'auto' },
    ].forEach(({ label, w }, i) => {
      const th = document.createElement('th');
      th.textContent = label;
      th.style.cssText = `border:1px solid #1a1a1a;padding:10px 12px;background:${i === 0 ? pdfStyle.cornerBg : pdfStyle.headerBg};color:${pdfStyle.headerText};font-weight:600;font-size:13px;letter-spacing:0.3px;text-align:left;width:${w};`;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    availability.forEach((row, pi) => {
      const tr = document.createElement('tr');
      const bg = pi % 2 === 0 ? '#fff' : '#fafbfc';
      const tdP = document.createElement('td');
      tdP.textContent = row.period;
      tdP.style.cssText = `border:1px solid #1a1a1a;padding:10px 12px;background:${bg};font-weight:600;color:#1a1a1a;`;
      tr.appendChild(tdP);
      const tdC = document.createElement('td');
      const fracBg = row.free.length === 0 ? '#fee2e2' : row.free.length === teachers.length ? '#ecf0f1' : '#dcfce7';
      const fracColor = row.free.length === 0 ? '#991b1b' : row.free.length === teachers.length ? '#374151' : '#166534';
      tdC.innerHTML = `<span style="display:inline-block;padding:3px 12px;border-radius:12px;background:${fracBg};color:${fracColor};font-weight:700;font-size:12px;">${row.free.length} / ${teachers.length}</span>`;
      tdC.style.cssText = `border:1px solid #1a1a1a;padding:10px 12px;background:${bg};`;
      tr.appendChild(tdC);
      const tdL = document.createElement('td');
      tdL.style.cssText = `border:1px solid #1a1a1a;padding:10px 12px;background:${bg};color:#374151;line-height:1.7;`;
      if (teachers.length === 0) tdL.innerHTML = '<span style="color:#9ca3af;font-style:italic;">No teachers registered.</span>';
      else if (row.free.length === 0) tdL.innerHTML = '<span style="color:#9ca3af;font-style:italic;">All teachers are assigned this period.</span>';
      else tdL.innerHTML = row.free.map((t) => `<span style="display:inline-block;padding:3px 10px;background:#f1f5f9;border:1px solid #d1d5db;border-radius:12px;font-size:12px;font-weight:500;color:#1f2937;margin:2px 4px 2px 0;">${escapeHtml(t.name)}</span>`).join('');
      tr.appendChild(tdL);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    root.appendChild(table);
    const note = document.createElement('div');
    note.textContent = 'A teacher is considered "available" in a period when they are not assigned to any class in that period.';
    note.style.cssText = 'margin-top:18px;font-size:11px;color:#6b7280;font-style:italic;font-family:"Helvetica Neue",Arial,sans-serif;text-align:center;';
    root.appendChild(note);
    return root;
  };

  const renderNodeToCanvas = async (node) => {
    node.style.position = 'fixed'; node.style.left = '-10000px'; node.style.top = '0';
    document.body.appendChild(node);
    try { return await html2canvas(node, { scale: 2, backgroundColor: '#ffffff', useCORS: true }); }
    finally { document.body.removeChild(node); }
  };
  const addCanvasAsPage = (pdf, canvas, isFirst) => {
    if (!isFirst) pdf.addPage('a4', 'landscape');
    const pw = pdf.internal.pageSize.getWidth(), ph = pdf.internal.pageSize.getHeight(), margin = 28;
    const ratio = Math.min((pw - margin * 2) / canvas.width, (ph - margin * 2) / canvas.height);
    const w = canvas.width * ratio, h = canvas.height * ratio;
    pdf.addImage(canvas.toDataURL('image/png'), 'PNG', (pw - w) / 2, (ph - h) / 2, w, h);
  };
  const exportPDF = async () => {
    setExporting(true);
    try {
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      addCanvasAsPage(pdf, await renderNodeToCanvas(buildScheduleNode()), true);
      addCanvasAsPage(pdf, await renderNodeToCanvas(buildAvailabilityNode()), false);
      pdf.save(`${(schoolName.trim() || 'timetable').replace(/\s+/g, '_').toLowerCase()}.pdf`);
    } finally { setExporting(false); }
  };

  const baseFilename = () => (schoolName.trim() || 'timetable').replace(/\s+/g, '_').toLowerCase();

  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportJSON = () => {
    const payload = { ...stateBundle, exportedAt: new Date().toISOString(), version: 1 };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    downloadBlob(blob, `${baseFilename()}.json`);
  };

  const importJSON = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!Array.isArray(parsed.classes) || !Array.isArray(parsed.periods)) {
          throw new Error('Invalid file: missing classes or periods.');
        }
        applyState(parsed);
        alert('Timetable imported successfully. Click Save to persist it as the new default.');
      } catch (e) {
        alert(`Could not import file: ${e.message}`);
      }
    };
    reader.readAsText(file);
  };

  const exportDOCX = async () => {
    setExporting(true);
    try {
      const cellShading = (fill) => ({ fill, type: 'clear', color: 'auto' });
      const allBorders = (() => {
        const b = { style: BorderStyle.SINGLE, size: 6, color: '1a1a1a' };
        return { top: b, bottom: b, left: b, right: b };
      })();

      const headerCell = (text, isCorner = false) => new TableCell({
        shading: cellShading(isCorner ? pdfStyle.cornerBg.replace('#','') : pdfStyle.headerBg.replace('#','')),
        borders: allBorders,
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: true, color: pdfStyle.headerText.replace('#',''), size: 22 })],
        })],
      });

      const periodCell = (text) => new TableCell({
        shading: cellShading('ecf0f1'),
        borders: allBorders,
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: true, size: 22, color: '1a1a1a' })],
        })],
      });

      const dataCell = (cell, pi) => {
        const t = cell.teacherId ? teacherById(cell.teacherId) : null;
        const s = cell.subjectId ? subjectById(cell.subjectId) : null;
        const tConflict = teacherConflicts.has(`${pi}:${cell._ci}`);
        const sConflict = subjectConflicts.has(`${pi}:${cell._ci}`);
        const fill = tConflict ? 'fde8e8' : sConflict ? 'fef3c7' : (pi % 2 === 0 ? 'ffffff' : 'fafbfc');
        const paragraphs = [];
        if (t) paragraphs.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: t.name, bold: true, size: 22, color: '1a1a1a' })],
        }));
        if (s) paragraphs.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: s.name, bold: true, size: 18, color: s.color.replace('#','') })],
        }));
        if (cell.text) paragraphs.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: cell.text, italics: true, size: 18, color: '555555' })],
        }));
        if (!t && !s && !cell.text) paragraphs.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '—', color: 'bbbbbb' })],
        }));
        return new TableCell({
          shading: cellShading(fill),
          borders: allBorders,
          margins: { top: 140, bottom: 140, left: 120, right: 120 },
          children: paragraphs,
        });
      };

      const scheduleHeaderRow = new TableRow({
        children: [
          headerCell('Period / Class', true),
          ...classes.map((c) => headerCell(c, false)),
        ],
      });
      const scheduleBodyRows = periods.map((p, pi) => new TableRow({
        children: [
          periodCell(p),
          ...classes.map((_, ci) => dataCell({ ...(grid[pi]?.[ci] ?? { teacherId: null, subjectId: null, text: '' }), _ci: ci }, pi)),
        ],
      }));
      const scheduleTable = new Table({
        rows: [scheduleHeaderRow, ...scheduleBodyRows],
        width: { size: 100, type: WidthType.PERCENTAGE },
      });

      // Availability table
      const avHeader = new TableRow({
        children: [
          headerCell('Period', true),
          headerCell('Free / Total'),
          headerCell('Available Teachers'),
        ],
      });
      const avRows = availability.map((row, pi) => {
        const fill = pi % 2 === 0 ? 'ffffff' : 'fafbfc';
        const fracBg = row.free.length === 0 ? 'fee2e2' : row.free.length === teachers.length ? 'ecf0f1' : 'dcfce7';
        const fracColor = row.free.length === 0 ? '991b1b' : row.free.length === teachers.length ? '374151' : '166534';
        return new TableRow({
          children: [
            new TableCell({
              shading: cellShading(fill), borders: allBorders,
              margins: { top: 120, bottom: 120, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: row.period, bold: true, size: 22 })] })],
            }),
            new TableCell({
              shading: cellShading(fracBg), borders: allBorders,
              margins: { top: 120, bottom: 120, left: 140, right: 140 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: `${row.free.length} / ${teachers.length}`, bold: true, size: 22, color: fracColor })],
              })],
            }),
            new TableCell({
              shading: cellShading(fill), borders: allBorders,
              margins: { top: 120, bottom: 120, left: 140, right: 140 },
              children: [new Paragraph({
                children: row.free.length === 0
                  ? [new TextRun({ text: 'All teachers are assigned this period.', italics: true, color: '9ca3af', size: 20 })]
                  : row.free.map((t, i) => new TextRun({ text: (i ? '   •   ' : '') + t.name, size: 20, color: '1f2937' })),
              })],
            }),
          ],
        });
      });
      const availabilityTable = new Table({
        rows: [avHeader, ...avRows],
        width: { size: 100, type: WidthType.PERCENTAGE },
      });

      const titleP = () => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({
          text: (schoolName.trim() || 'Timetable').toUpperCase(),
          bold: true, size: 44, color: pdfStyle.titleColor.replace('#',''),
        })],
      });
      const subtitleP = () => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({
          text: subtitle.trim() || 'Weekly Class Timetable',
          italics: true, size: 26, color: pdfStyle.subtitleColor.replace('#',''),
        })],
      });
      const dateP = () => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [new TextRun({
          text: `Issued: ${new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })}`,
          size: 18, color: '666666',
        })],
      });

      const footerP = footerText.trim()
        ? new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: footerText.trim(), size: 20, color: '333333' })] })
        : null;

      const availabilityTitle = new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 200 },
        children: [
          new TextRun({ break: 1 }),
          new TextRun({ text: 'TEACHER AVAILABILITY — BY PERIOD', bold: true, size: 28, color: pdfStyle.titleColor.replace('#','') }),
        ],
      });

      const doc = new Document({
        sections: [{
          properties: { page: { size: { orientation: 'landscape' }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
          children: [
            titleP(), subtitleP(), dateP(),
            scheduleTable,
            ...(footerP ? [footerP] : []),
            new Paragraph({ children: [new PageBreak()] }),
            titleP(), subtitleP(), availabilityTitle,
            availabilityTable,
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      downloadBlob(blob, `${baseFilename()}.docx`);
    } finally {
      setExporting(false);
    }
  };

  // ---- Filtered lists for sidebar ----
  const filteredTeachers = useMemo(
    () => teachers.filter((t) => t.name.toLowerCase().includes(filter.toLowerCase())),
    [teachers, filter]
  );
  const filteredSubjects = useMemo(
    () => subjects.filter((s) => s.name.toLowerCase().includes(filter.toLowerCase())),
    [subjects, filter]
  );
  const selT = teachers.find((t) => t.id === selectedTeacher);

  const formatSavedAt = (iso) => {
    if (!iso) return 'Never saved';
    const d = new Date(iso);
    return `Saved ${d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
  };

  // ---- Render ----
  if (!loaded) {
    return (
      <div className="app">
        <div className="boot-screen">
          <div className="boot-spinner" />
          <div className="boot-text">Loading timetable…</div>
        </div>
      </div>
    );
  }
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <div className="logo-mark">TT</div>
          <div>
            <h1>Timetable Studio</h1>
            <div className="tag">Professional class scheduling</div>
          </div>
        </div>
        <div className="topbar-actions">
          <div className={`save-status ${dirty ? 'dirty' : ''}`}>
            {saveStatus === 'saving' ? 'Saving…' :
             saveStatus === 'saved'  ? '✓ Saved' :
             saveStatus === 'error'  ? '✗ Save failed' :
             dirty ? '• Unsaved changes' : formatSavedAt(savedAt)}
          </div>
          <button className="ghost" onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')} title="Toggle theme">
            {theme === 'light' ? '🌙 Dark' : '☀ Light'}
          </button>
          <button className="primary outline" onClick={requestSave} title="Save snapshot (Ctrl+S)" disabled={saveStatus === 'saving'}>
            Save
          </button>
          <button className="ghost" onClick={() => importInputRef.current?.click()} title="Import JSON">
            Import
          </button>
          <input
            ref={importInputRef}
            type="file"
            accept="application/json,.json"
            style={{ display: 'none' }}
            onChange={(e) => { importJSON(e.target.files?.[0]); e.target.value = ''; }}
          />
          <div className="export-menu-wrap">
            <button className="primary" onClick={() => setExportMenu((v) => !v)} disabled={exporting}>
              {exporting ? 'Exporting…' : 'Export ▾'}
            </button>
            {exportMenu && (
              <div className="export-menu">
                <button onClick={() => { setExportMenu(false); exportPDF(); }}>
                  <span className="ext-tag pdf">PDF</span> Schedule + Availability
                </button>
                <button onClick={() => { setExportMenu(false); exportDOCX(); }}>
                  <span className="ext-tag docx">DOCX</span> Editable Word document
                </button>
                <button onClick={() => { setExportMenu(false); exportJSON(); }}>
                  <span className="ext-tag json">JSON</span> Full state (re-importable)
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      <div className="layout">
        <main className="main">
          <div className="card identity">
            <label>
              <span>School / Institution</span>
              <input value={schoolName} onChange={(e) => setSchoolName(e.target.value)} />
            </label>
            <label>
              <span>Subtitle</span>
              <input value={subtitle} onChange={(e) => setSubtitle(e.target.value)} />
            </label>
          </div>

          <section className="controls">
            <button onClick={addClass}>+ Class column</button>
            <button onClick={addPeriod}>+ Period row</button>
            {teacherConflicts.size > 0 && (
              <span className="warn-inline">⚠ {teacherConflicts.size} teacher conflict{teacherConflicts.size > 1 ? 's' : ''}</span>
            )}
            {subjectConflicts.size > 0 && (
              <span className="warn-inline subject-warn">⚠ {subjectConflicts.size} subject conflict{subjectConflicts.size > 1 ? 's' : ''}</span>
            )}
          </section>

          <div className="schedule-card">
            <div className="schedule-header">
              <div className="title">{schoolName.trim() || 'Your School Name'}</div>
              <div className="sub">{subtitle.trim() || 'Weekly Class Timetable'}</div>
            </div>
            <div className="table-wrap">
              <table className="schedule">
                <thead>
                  <tr>
                    <th className="corner">
                      <div className="corner-inner">
                        <span className="corner-class">CLASS</span>
                        <svg className="corner-line" viewBox="0 0 100 100" preserveAspectRatio="none"><line x1="0" y1="0" x2="100" y2="100" stroke="currentColor" strokeWidth="1" /></svg>
                        <span className="corner-period">PERIOD</span>
                      </div>
                    </th>
                    {classes.map((cls, ci) => (
                      <th key={ci}>
                        <div className="th-row">
                          <input value={cls} onChange={(e) => updateClassName(ci, e.target.value)} />
                          <button className="x" onClick={() => removeClass(ci)}>×</button>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {periods.map((per, pi) => (
                    <tr key={pi}>
                      <th className="row-head">
                        <div className="th-row">
                          <input value={per} onChange={(e) => updatePeriodName(pi, e.target.value)} />
                          <button className="x" onClick={() => removePeriod(pi)}>×</button>
                        </div>
                      </th>
                      {classes.map((_, ci) => {
                        const cell = grid[pi]?.[ci] ?? { teacherId: null, subjectId: null, text: '' };
                        const tConflict = teacherConflicts.has(`${pi}:${ci}`);
                        const sConflict = subjectConflicts.has(`${pi}:${ci}`);
                        const t = cell.teacherId ? teacherById(cell.teacherId) : null;
                        const s = cell.subjectId ? subjectById(cell.subjectId) : null;
                        return (
                          <td key={ci}
                            className={`cell ${tConflict ? 'conflict' : ''} ${sConflict ? 'subject-conflict' : ''} ${dragged ? 'drop-ready' : ''}`}
                            onDragOver={(e) => e.preventDefault()}
                            onDrop={() => onDropCell(pi, ci)}>
                            {s && (
                              <div className="chip subject" style={{ background: s.color }}>
                                <span>{s.name}</span>
                                <button className="x" onClick={() => updateCell(pi, ci, { subjectId: null })}>×</button>
                              </div>
                            )}
                            {t && (
                              <div className="chip teacher">
                                <span>{t.name}</span>
                                <button className="x" onClick={() => updateCell(pi, ci, { teacherId: null })}>×</button>
                              </div>
                            )}
                            {!t && cell.teacherId && (
                              <div className="chip missing">
                                <span>Removed</span>
                                <button className="x" onClick={() => updateCell(pi, ci, { teacherId: null })}>×</button>
                              </div>
                            )}
                            <input className="cell-text" placeholder="note…" value={cell.text}
                              onChange={(e) => updateCell(pi, ci, { text: e.target.value })} />
                            <div className="tag-stack">
                              {tConflict && <div className="conflict-tag">Teacher</div>}
                              {sConflict && <div className="conflict-tag subject">Subject</div>}
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="card availability-card">
            <div className="availability-head">
              <div>
                <h3 className="card-title">Teacher Availability</h3>
                <p className="card-sub">Teachers free in each period (not assigned to any class).</p>
              </div>
              <div className="legend">
                <span className="legend-pill avail">Available</span>
                <span className="legend-pill busy">Assigned</span>
              </div>
            </div>
            <div className="table-wrap">
              <table className="availability">
                <thead>
                  <tr><th style={{ width: 180 }}>Period</th><th style={{ width: 130 }}>Free / Total</th><th>Available teachers</th></tr>
                </thead>
                <tbody>
                  {availability.map((row, i) => (
                    <tr key={i}>
                      <td className="av-period">{row.period}</td>
                      <td className="av-count">
                        <span className={`fraction ${row.free.length === 0 ? 'none' : row.free.length === teachers.length ? 'all' : ''}`}>
                          {row.free.length} / {teachers.length}
                        </span>
                      </td>
                      <td className="av-list">
                        {teachers.length === 0 ? <span className="muted">No teachers registered.</span>
                          : row.free.length === 0 ? <span className="muted">All teachers are assigned this period.</span>
                          : row.free.map((t) => <span key={t.id} className="av-chip">{t.name}</span>)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="card">
            <label className="block">
              <span className="block-label">Footer notes (appear below schedule in PDF)</span>
              <textarea value={footerText} onChange={(e) => setFooterText(e.target.value)} />
            </label>
          </div>

          <div className="card">
            <div className="pdf-style-head">
              <div>
                <h3 className="card-title">PDF Export Style</h3>
                <p className="card-sub">Customize header colors of the exported PDF.</p>
              </div>
              <button className="link-btn" onClick={() => setPdfStyle(DEFAULT_PDF_STYLE)}>Reset to default</button>
            </div>
            <div className="pdf-style-grid">
              {[
                { key: 'headerBg', label: 'Class header background' },
                { key: 'headerText', label: 'Class header text' },
                { key: 'cornerBg', label: 'Corner cell background' },
                { key: 'titleColor', label: 'School name color' },
                { key: 'subtitleColor', label: 'Subtitle color' },
              ].map(({ key, label }) => (
                <label key={key} className="color-field">
                  <span>{label}</span>
                  <div className="color-input">
                    <input type="color" value={pdfStyle[key]} onChange={(e) => setPdfStyle((s) => ({ ...s, [key]: e.target.value }))} />
                    <input type="text" value={pdfStyle[key]} onChange={(e) => setPdfStyle((s) => ({ ...s, [key]: e.target.value }))} />
                  </div>
                </label>
              ))}
            </div>
            <div className="pdf-preview" style={{ background: pdfStyle.headerBg, color: pdfStyle.headerText }}>
              <div className="pdf-preview-corner" style={{ background: pdfStyle.cornerBg, color: pdfStyle.headerText }}>P / C</div>
              <div className="pdf-preview-cells"><span>Class A</span><span>Class B</span><span>Class C</span></div>
            </div>
            <div className="pdf-preview-titles">
              <div style={{ color: pdfStyle.titleColor, fontWeight: 700, letterSpacing: 1, textTransform: 'uppercase', fontFamily: 'Georgia, serif', fontSize: 18 }}>{schoolName.trim() || 'Your School Name'}</div>
              <div style={{ color: pdfStyle.subtitleColor, fontStyle: 'italic', fontFamily: 'Georgia, serif', fontSize: 13, marginTop: 2 }}>{subtitle.trim() || 'Weekly Class Timetable'}</div>
            </div>
          </div>
        </main>

        {/* ============== Sidebar (master/detail) ============== */}
        <aside className="sidebar">
          <div className="tabs">
            <button className={sidebarTab === 'teachers' ? 'tab active' : 'tab'} onClick={() => setSidebarTab('teachers')}>
              Teachers <span className="count">{teachers.length}</span>
            </button>
            <button className={sidebarTab === 'subjects' ? 'tab active' : 'tab'} onClick={() => setSidebarTab('subjects')}>
              Subjects <span className="count">{subjects.length}</span>
            </button>
          </div>

          <div className="filter-row">
            <input
              className="filter-input"
              placeholder={sidebarTab === 'teachers' ? 'Search teachers…' : 'Search subjects…'}
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
            />
          </div>

          <div className="add-row">
            {sidebarTab === 'teachers' ? (
              <>
                <input placeholder="Add teacher name" value={teacherInput}
                  onChange={(e) => setTeacherInput(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && addTeacher()} />
                <button className="primary-mini" onClick={addTeacher}>Add</button>
              </>
            ) : (
              <>
                <input placeholder="Add subject" value={subjectInput}
                  onChange={(e) => setSubjectInput(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && addSubject()} />
                <button className="primary-mini" onClick={addSubject}>Add</button>
              </>
            )}
          </div>

          <p className="hint">
            {sidebarTab === 'teachers'
              ? 'Drag a teacher onto a cell. Click to select and edit subjects.'
              : 'Drag a subject onto a cell.'}
          </p>

          <div className="list-scroll">
            {sidebarTab === 'teachers' ? (
              <ul className="row-list">
                {filteredTeachers.map((t) => {
                  const isSel = selectedTeacher === t.id;
                  return (
                    <li
                      key={t.id}
                      className={`row teacher-row ${isSel ? 'selected' : ''}`}
                      draggable
                      onDragStart={(e) => { setDragged({ type: 'teacher', id: t.id }); e.dataTransfer.effectAllowed = 'copy'; }}
                      onDragEnd={() => setDragged(null)}
                      onClick={() => setSelectedTeacher(isSel ? null : t.id)}
                    >
                      <span className="grip" aria-hidden>⋮⋮</span>
                      <div className="avatar">{initials(t.name)}</div>
                      <div className="row-main">
                        <div className="row-name">{t.name}</div>
                        <div className="row-meta">
                          {t.subjectIds.length === 0
                            ? <span className="muted">No subjects</span>
                            : <span className="row-meta-text">{t.subjectIds.length} subject{t.subjectIds.length>1?'s':''}</span>}
                        </div>
                      </div>
                      <button className="x" onClick={(e) => { e.stopPropagation(); removeTeacher(t.id); }}>×</button>
                    </li>
                  );
                })}
                {filteredTeachers.length === 0 && (
                  <li className="empty">{teachers.length === 0 ? 'No teachers yet — add one above.' : 'No matches.'}</li>
                )}
              </ul>
            ) : (
              <ul className="row-list">
                {filteredSubjects.map((s) => {
                  const cov = subjectCoverage[s.id] || 0;
                  return (
                    <li
                      key={s.id}
                      className="row subject-row"
                      draggable
                      onDragStart={(e) => { setDragged({ type: 'subject', id: s.id }); e.dataTransfer.effectAllowed = 'copy'; }}
                      onDragEnd={() => setDragged(null)}
                    >
                      <span className="grip" aria-hidden>⋮⋮</span>
                      <span className="swatch" style={{ background: s.color }} />
                      <div className="row-main">
                        <input
                          className="row-name-input"
                          value={s.name}
                          onChange={(e) => renameSubject(s.id, e.target.value)}
                          onClick={(e) => e.stopPropagation()}
                        />
                        <div className="row-meta">
                          <span className={`coverage ${cov === 0 ? 'cov-zero' : cov === 1 ? 'cov-one' : 'cov-many'}`}>
                            {cov} teacher{cov !== 1 ? 's' : ''}
                          </span>
                        </div>
                      </div>
                      <button className="x" onClick={() => removeSubject(s.id)}>×</button>
                    </li>
                  );
                })}
                {filteredSubjects.length === 0 && (
                  <li className="empty">{subjects.length === 0 ? 'No subjects yet — add one above.' : 'No matches.'}</li>
                )}
              </ul>
            )}
          </div>

          {/* Detail panel — only on teachers tab */}
          {sidebarTab === 'teachers' && selT && (
            <div className="detail-panel">
              <div className="detail-head">
                <div className="avatar lg">{initials(selT.name)}</div>
                <div className="detail-name">
                  <input
                    value={selT.name}
                    onChange={(e) => renameTeacher(selT.id, e.target.value)}
                  />
                  <div className="detail-sub">Assign subjects this teacher can teach</div>
                </div>
                <button className="x" onClick={() => setSelectedTeacher(null)}>×</button>
              </div>
              <div className="detail-body">
                {subjects.length === 0 ? (
                  <div className="muted small">Add subjects in the Subjects tab first.</div>
                ) : (
                  <div className="toggle-grid">
                    {subjects.map((s) => {
                      const on = selT.subjectIds.includes(s.id);
                      return (
                        <button
                          key={s.id}
                          className={`toggle ${on ? 'on' : ''}`}
                          style={on ? { background: s.color, borderColor: s.color, color: '#fff' } : { borderColor: s.color, color: s.color }}
                          onClick={() => toggleTeacherSubject(selT.id, s.id)}
                        >
                          {s.name}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>
          )}
        </aside>
      </div>

      {confirmSave && (
        <div className="modal-backdrop" onClick={() => setConfirmSave(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>Overwrite saved snapshot?</h3>
            </div>
            <div className="modal-body">
              <p>
                This will replace the snapshot saved on the server
                {savedAt ? <> (last saved <strong>{new Date(savedAt).toLocaleString()}</strong>)</> : ''}.
                The new snapshot becomes the default the app loads on next server restart.
              </p>
              <p className="modal-note">This action cannot be undone.</p>
            </div>
            <div className="modal-foot">
              <button onClick={() => setConfirmSave(false)}>Cancel</button>
              <button className="primary-mini danger" onClick={doSave}>Overwrite & Save</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
