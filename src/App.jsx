import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  Download,
  Upload,
  Bold,
  Italic,
  Plus,
  Type,
  Percent,
  IndianRupee,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Trash2,
  Sheet as SheetIcon,
  Check,
  X,
  ListChecks,
  ChevronDown,
  Grid2x2Plus,
  Snowflake,
  Filter as FilterIcon,
  FunctionSquare,
} from "lucide-react";
import { HyperFormula } from "hyperformula";

/* ===================== Config ===================== */
const STORAGE_KEY = "mini-excel-grid-v4";
const DEFAULT_COLS = 12;
const DEFAULT_ROWS = 30;

/* ===================== Helpers ===================== */
const colToName = (n) => {
  let s = "";
  n += 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
const addr = (r, c) => `${colToName(c)}${r + 1}`;
const parseA1 = (a1) => {
  const re = new RegExp("([A-Z]+)(\\d+)", "i");
  const m = a1.match(re);
  if (!m) return null;
  const col =
    m[1]
      .toUpperCase()
      .split("")
      .reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  const row = parseInt(m[2], 10) - 1;
  return { r: row, c: col };
};

const defaultCell = () => ({
  input: "",
  value: "",
  fmt: { bold: false, italic: false, align: "left", type: "text" },
  validation: null, // {type:'list', values:[...]}
});
const newGrid = (rows = DEFAULT_ROWS, cols = DEFAULT_COLS) =>
  Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => defaultCell())
  );

const serialize = (state) => JSON.stringify(state);
const deserialize = (s) => {
  try {
    return JSON.parse(s);
  } catch {
    return null;
  }
};

/* ===================== App ===================== */
export default function App() {
  // Multiple sheets
  const [sheets, setSheets] = useState(() => {
    const saved = deserialize(localStorage.getItem(STORAGE_KEY));
    if (saved?.sheets && saved?.activeIndex >= 0) return saved;
    return {
      sheets: [{ name: "Sheet1", grid: newGrid(), names: {} }],
      activeIndex: 0,
    };
  });
  const active = sheets.sheets[sheets.activeIndex];

  const [selection, setSelection] = useState({ r: 0, c: 0 });
  const [editVal, setEditVal] = useState("");
  const [isEditing, setIsEditing] = useState(false);
  const inputRef = useRef(null);

  // Named range modal
  const [showNameModal, setShowNameModal] = useState(false);
  const [nameKey, setNameKey] = useState("");
  const [nameRef, setNameRef] = useState("");

  // Validation UI
  const [showValidation, setShowValidation] = useState(false);
  const [valList, setValList] = useState("");

  // Pivot UI
  const [pivot, setPivot] = useState({
    range: "",
    rowCol: 0,
    valCol: 1,
    agg: "SUM",
  });
  const [pivotData, setPivotData] = useState(null);

  // Freeze panes toggles
  const [freezeFirstRow, setFreezeFirstRow] = useState(true);
  const [freezeFirstCol, setFreezeFirstCol] = useState(true);

  // Column filters
  const [filters, setFilters] = useState(
    Array.from({ length: DEFAULT_COLS }, () => "")
  );
  useEffect(() => {
    const cols = active.grid[0]?.length || DEFAULT_COLS;
    setFilters((prev) =>
      prev.length === cols
        ? prev
        : Array.from({ length: cols }, (_, i) => prev[i] ?? "")
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheets.activeIndex, active.grid[0]?.length]);

  // Function wizard
  const [showWizard, setShowWizard] = useState(false);
  const functionCatalog = useMemo(
    () => ({
      Basics: [
        { name: "SUM", tpl: "=SUM(A1:A10)", hint: "Adds numbers in a range" },
        {
          name: "AVERAGE",
          tpl: "=AVERAGE(A1:A10)",
          hint: "Mean of numbers",
        },
        {
          name: "COUNTIF",
          tpl: '=COUNTIF(A1:A100, "x")',
          hint: "Counts cells matching criteria",
        },
        {
          name: "IF",
          tpl: '=IF(A1>0, "Yes", "No")',
          hint: "Conditional logic",
        },
      ],
      Lookup: [
        {
          name: "VLOOKUP",
          tpl: "=VLOOKUP(A2, A1:D100, 3, FALSE)",
          hint: "Lookup by first column",
        },
        {
          name: "HLOOKUP",
          tpl: "=HLOOKUP(B1, A1:Z10, 2, FALSE)",
          hint: "Horizontal lookup",
        },
        {
          name: "INDEX+MATCH",
          tpl: '=INDEX(C1:C100, MATCH(A2, A1:A100, 0))',
          hint: "Flexible lookup",
        },
      ],
      DateTime: [
        { name: "TODAY", tpl: "=TODAY()", hint: "Current date" },
        { name: "DATE", tpl: "=DATE(2025,9,6)", hint: "Build a date" },
        {
          name: "EOMONTH",
          tpl: "=EOMONTH(TODAY(),0)",
          hint: "End of month",
        },
      ],
      Text: [
        {
          name: "TEXT",
          tpl: '=TEXT(A1, "0.00")',
          hint: "Format number as text",
        },
        { name: "CONCAT", tpl: '=CONCAT(A1, " ", B1)', hint: "Join strings" },
        { name: "LEFT", tpl: "=LEFT(A1,3)", hint: "Substring" },
        { name: "RIGHT", tpl: "=RIGHT(A1,3)", hint: "Substring" },
      ],
    }),
    []
  );

  /* -------- HyperFormula workbook -------- */
  const engine = useMemo(() => {
    const hf = HyperFormula.buildEmpty({ licenseKey: "gpl-v3" });

    sheets.sheets.forEach((s, idx) => {
      const sheetName = s.name || `Sheet${idx + 1}`;

      // Find an existing sheet by name, otherwise create one.
      let id;
      try {
        id = hf.getSheetId(sheetName);
      } catch {
        id = undefined;
      }
      if (typeof id !== "number" || id < 0) {
        id = hf.addSheet(sheetName); // returns numeric sheetId
      }

      const data = s.grid.map((row) => row.map((cell) => cell.input ?? ""));
      hf.setSheetContent(id, data);

      // Named expressions scoped to the sheet
      Object.entries(s.names || {}).forEach(([k, ref]) => {
        try {
          hf.addNamedExpression(k, ref, id);
        } catch {}
      });
    });

    return hf;
  }, [sheets]);

  /* -------- Recompute display values for active sheet -------- */
  useEffect(() => {
    const id = sheets.activeIndex;
    const g = active.grid;
    const nextGrid = g.map((row, r) =>
      row.map((cell, c) => {
        let v = engine.getCellValue({ sheet: id, col: c, row: r });
        if (v && typeof v === "object" && "type" in v) v = "#ARRAY";
        if (v && typeof v === "object" && "value" in v) v = "#ERR";
        const t = cell.fmt.type;
        const num = Number(v);
        let display = v;
        if (t === "number" && Number.isFinite(num)) display = num.toLocaleString();
        if (t === "currency" && Number.isFinite(num))
          display = new Intl.NumberFormat(undefined, {
            style: "currency",
            currency: "INR",
          }).format(num);
        if (t === "percent" && Number.isFinite(num))
          display = `${(num * 100).toFixed(2)}%`;
        return { ...cell, value: display };
      })
    );
    const newsheets = {
      ...sheets,
      sheets: sheets.sheets.map((s, i) =>
        i === sheets.activeIndex ? { ...s, grid: nextGrid } : s
      ),
    };
    setSheets(newsheets);
    localStorage.setItem(STORAGE_KEY, serialize(newsheets));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine]);

  const activeGrid = sheets.sheets[sheets.activeIndex].grid;
  const selectedCell = activeGrid[selection.r]?.[selection.c] ?? defaultCell();

  /* -------- Filtered view of rows -------- */
  const rowsToRender = useMemo(() => {
    const f = filters;
    return activeGrid.filter((row) =>
      row.every((cell, i) => {
        const q = (f[i] || "").toLowerCase();
        if (!q) return true;
        const v = String(cell.value ?? cell.input ?? "").toLowerCase();
        return v.includes(q);
      })
    );
  }, [activeGrid, filters]);

  /* -------- Keyboard navigation + edit -------- */
  const handleKeyDown = (e) => {
    if (isEditing) {
      if (e.key === "Enter") commitEdit();
      return;
    }
    let { r, c } = selection;
    if (e.key === "Enter" || e.key === "ArrowDown")
      r = Math.min(activeGrid.length - 1, r + 1);
    else if (e.key === "ArrowUp") r = Math.max(0, r - 1);
    else if (e.key === "ArrowRight") c = Math.min(activeGrid[0].length - 1, c + 1);
    else if (e.key === "ArrowLeft") c = Math.max(0, c - 1);
    else if (e.key.length === 1) {
      setIsEditing(true);
      setEditVal(e.key);
      setTimeout(() => inputRef.current?.focus(), 0);
      e.preventDefault();
      return;
    }
    setSelection({ r, c });
  };

  const commitEdit = () => {
    const { r, c } = selection;
    const val = editVal;
    const rule = activeGrid[r][c].validation;
    if (rule?.type === "list") {
      const allowed = rule.values || [];
      if (!allowed.includes(val) && !val.startsWith("=")) {
        setIsEditing(false);
        flashInvalid(r, c);
        return;
      }
    }
    mutateActiveGrid((grid) => {
      grid[r][c].input = val;
    });
    setIsEditing(false);
  };

  const flashInvalid = (r, c) => {
    const el = document.querySelector(`[data-rc="${r}-${c}"]`);
    if (!el) return;
    el.animate(
      [{ outlineColor: "#ef4444" }, { outlineColor: "#ef4444" }],
      { duration: 300, iterations: 2 }
    );
  };

  const startEdit = () => {
    setEditVal(selectedCell.input || "");
    setIsEditing(true);
    setTimeout(() => inputRef.current?.focus(), 0);
  };

  const setFmt = (patch) =>
    mutateActiveGrid((grid) => {
      grid[selection.r][selection.c].fmt = {
        ...grid[selection.r][selection.c].fmt,
        ...patch,
      };
    });

  const mutateActiveGrid = (fn) => {
    setSheets((s) => {
      const next = {
        ...s,
        sheets: s.sheets.map((sh, i) => {
          if (i !== s.activeIndex) return sh;
          const grid = sh.grid.map((row) => row.map((cell) => ({ ...cell })));
          fn(grid);
          return { ...sh, grid };
        }),
      };
      localStorage.setItem(STORAGE_KEY, serialize(next));
      return next;
    });
  };

  /* -------- Row/Col & Sheet ops -------- */
  const addRow = () =>
    mutateActiveGrid((g) =>
      g.push(Array.from({ length: g[0].length }, () => defaultCell()))
    );
  const addCol = () => mutateActiveGrid((g) => g.forEach((r) => r.push(defaultCell())));
  const delRow = () => mutateActiveGrid((g) => { if (g.length > 1) g.pop(); });
  const delCol = () => mutateActiveGrid((g) => { if (g[0].length > 1) g.forEach((r) => r.pop()); });
  const newSheet = () =>
    setSheets((s) => ({
      ...s,
      sheets: [
        ...s.sheets,
        { name: `Sheet${s.sheets.length + 1}`, grid: newGrid(), names: {} },
      ],
      activeIndex: s.sheets.length,
    }));
  const renameSheet = (idx, name) =>
    setSheets((s) => ({
      ...s,
      sheets: s.sheets.map((sh, i) => (i === idx ? { ...sh, name } : sh)),
    }));
  const deleteSheet = (idx) =>
    setSheets((s) => {
      if (s.sheets.length <= 1) return s;
      const arr = s.sheets.slice();
      arr.splice(idx, 1);
      const activeIndex = Math.max(0, Math.min(s.activeIndex, arr.length - 1));
      const next = { sheets: arr, activeIndex };
      localStorage.setItem(STORAGE_KEY, serialize(next));
      return next;
    });

  /* -------- CSV Export / Import -------- */
  const exportCSV = () => {
    const rows = activeGrid.map((row) =>
      row.map((c) => String(c.input ?? "").replaceAll('"', '""'))
    );
    const csv = rows.map((r) => r.map((x) => `"${x}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${active.name || "sheet"}.csv`;
    a.click();
  };

  const importCSV = async (file) => {
    if (!file) return;
    const text = await file.text();

    // Normalize CRLF/CR, then split
    const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");

    // CSV parser with quotes
    const rows = lines.map((line) => {
      const out = [];
      let cur = "";
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQ && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQ = !inQ;
          }
        } else if (ch === "," && !inQ) {
          out.push(cur);
          cur = "";
        } else {
          cur += ch;
        }
      }
      out.push(cur);
      return out;
    });

    mutateActiveGrid((g) => {
      const rowsN = Math.max(rows.length, g.length);
      const colsN = Math.max(Math.max(...rows.map((r) => r.length)), g[0].length);

      while (g.length < rowsN)
        g.push(Array.from({ length: g[0].length }, () => defaultCell()));

      if (g[0].length < colsN)
        g.forEach((r) => {
          while (r.length < colsN) r.push(defaultCell());
        });

      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[r].length; c++) {
          g[r][c].input = rows[r][c];
        }
      }
    });
  };

  /* -------- Named Ranges -------- */
  const openNameModalFromSelection = () => {
    setNameKey("");
    setNameRef(addr(selection.r, selection.c));
    setShowNameModal(true);
  };
  const saveNamedRange = () => {
    if (!nameKey || !nameRef) {
      setShowNameModal(false);
      return;
    }
    setSheets((s) => {
      const next = {
        ...s,
        sheets: s.sheets.map((sh, i) =>
          i === s.activeIndex
            ? { ...sh, names: { ...(sh.names || {}), [nameKey]: nameRef } }
            : sh
        ),
      };
      localStorage.setItem(STORAGE_KEY, serialize(next));
      return next;
    });
    setShowNameModal(false);
  };

  /* -------- Validation (list) -------- */
  const applyValidationList = () => {
    const vals = valList.split(",").map((x) => x.trim()).filter(Boolean);
    mutateActiveGrid((g) => {
      g[selection.r][selection.c].validation = { type: "list", values: vals };
    });
    setShowValidation(false);
    setValList("");
  };

  /* -------- Pivot (simple) -------- */
  const buildPivot = () => {
    try {
      const id = sheets.activeIndex;
      const reRange = new RegExp("([A-Z]+\\d+):([A-Z]+\\d+)", "i");
      const m = pivot.range.match(reRange);
      if (!m) {
        setPivotData({ error: "Range must be like A1:C100" });
        return;
      }
      const s = parseA1(m[1]);
      const e = parseA1(m[2]);
      const r0 = Math.min(s.r, e.r),
        r1 = Math.max(s.r, e.r);
      const c0 = Math.min(s.c, e.c),
        c1 = Math.max(s.c, e.c);
      const rows = [];
      for (let r = r0; r <= r1; r++) {
        const row = [];
        for (let c = c0; c <= c1; c++) {
          let v = engine.getCellValue({ sheet: id, col: c, row: r });
          if (v && typeof v === "object") v = "";
          row.push(v);
        }
        rows.push(row);
      }
      const idxRow = pivot.rowCol;
      const idxVal = pivot.valCol;
      const map = new Map();
      for (const row of rows) {
        const k = String(row[idxRow]);
        const val = Number(row[idxVal]);
        const obj = map.get(k) || { count: 0, sum: 0 };
        obj.count += 1;
        if (Number.isFinite(val)) obj.sum += val;
        map.set(k, obj);
      }
      const out = Array.from(map.entries()).map(([k, v]) => ({
        key: k,
        COUNT: v.count,
        SUM: v.sum,
      }));
      setPivotData({ rows: out });
    } catch (e) {
      setPivotData({ error: "Pivot error" });
    }
  };

  /* -------- Wizard insert -------- */
  const insertTemplate = (tpl) => {
    setIsEditing(true);
    setEditVal(tpl);
    setTimeout(() => inputRef.current?.focus(), 0);
  };

  /* ===================== Render ===================== */
  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4">
      <div className="max-w-[1380px] mx-auto">
        {/* Sheet Tabs */}
        <div className="flex items-center gap-2 mb-3 overflow-x-auto">
          {sheets.sheets.map((sh, i) => (
            <div
              key={i}
              className={`flex items-center gap-2 px-3 py-1.5 rounded-2xl ${
                i === sheets.activeIndex
                  ? "bg-sky-700"
                  : "bg-slate-800 hover:bg-slate-700"
              } shadow whitespace-nowrap`}
            >
              <button
                onClick={() => setSheets((s) => ({ ...s, activeIndex: i }))}
                className="flex items-center gap-2"
              >
                <SheetIcon className="w-4 h-4" />
                <input
                  className="bg-transparent outline-none w-28"
                  value={sh.name}
                  onChange={(e) => renameSheet(i, e.target.value)}
                />
              </button>
              <button
                onClick={() => deleteSheet(i)}
                className="opacity-70 hover:opacity-100"
              >
                ✕
              </button>
            </div>
          ))}
          <button
            onClick={newSheet}
            className="px-3 py-1.5 rounded-2xl bg-slate-800 hover:bg-slate-700 shadow flex items-center gap-2"
          >
            <Plus className="w-4 h-4" />
            New Sheet
          </button>
        </div>

        {/* Toolbar */}
        <header className="flex flex-wrap items-center justify-between mb-3 gap-2">
          <h1 className="text-xl font-semibold tracking-tight">
            Mini Excel — Sheets, Names, Validation, Filters, Freeze, Pivot, Wizard
          </h1>
          <div className="flex flex-wrap gap-2 items-center">
            <button
              onClick={() => setFmt({ bold: !selectedCell.fmt.bold })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.bold ? "ring-2 ring-slate-400" : ""
              }`}
              title="Bold"
            >
              <Bold className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ italic: !selectedCell.fmt.italic })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.italic ? "ring-2 ring-slate-400" : ""
              }`}
              title="Italic"
            >
              <Italic className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ align: "left" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.align === "left" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Align left"
            >
              <AlignLeft className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ align: "center" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.align === "center" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Align center"
            >
              <AlignCenter className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ align: "right" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.align === "right" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Align right"
            >
              <AlignRight className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ type: "text" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.type === "text" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Text"
            >
              <Type className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ type: "number" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.type === "number" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Number"
            >
              123
            </button>
            <button
              onClick={() => setFmt({ type: "currency" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.type === "currency" ? "ring-2 ring-slate-400" : ""
              }`}
              title="INR"
            >
              <IndianRupee className="w-4 h-4" />
            </button>
            <button
              onClick={() => setFmt({ type: "percent" })}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 ${
                selectedCell.fmt.type === "percent" ? "ring-2 ring-slate-400" : ""
              }`}
              title="Percent"
            >
              <Percent className="w-4 h-4" />
            </button>

            <div className="w-px bg-slate-700 mx-1" />
            <button
              onClick={addRow}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Plus className="w-4 h-4" />
              Row
            </button>
            <button
              onClick={addCol}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Plus className="w-4 h-4" />
              Col
            </button>
            <button
              onClick={delRow}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Trash2 className="w-4 h-4" />
              Row
            </button>
            <button
              onClick={delCol}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Trash2 className="w-4 h-4" />
              Col
            </button>

            <div className="w-px bg-slate-700 mx-1" />
            <button
              onClick={exportCSV}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Download className="w-4 h-4" />
              CSV
            </button>
            <label className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2 cursor-pointer">
              <Upload className="w-4 h-4" />
              CSV
              <input
                type="file"
                accept=".csv"
                className="hidden"
                onChange={(e) => importCSV(e.target.files?.[0])}
              />
            </label>

            <div className="w-px bg-slate-700 mx-1" />
            <button
              onClick={openNameModalFromSelection}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <ListChecks className="w-4 h-4" />
              Named Range
            </button>
            <button
              onClick={() => setShowValidation(true)}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
            >
              <Check className="w-4 h-4" />
              Validation
            </button>

            <div className="w-px bg-slate-700 mx-1" />
            <button
              onClick={() => setFreezeFirstRow((v) => !v)}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2 ${
                freezeFirstRow ? "ring-2 ring-slate-400" : ""
              }`}
              title="Freeze first row"
            >
              <Snowflake className="w-4 h-4" />
              Freeze Row
            </button>
            <button
              onClick={() => setFreezeFirstCol((v) => !v)}
              className={`px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2 ${
                freezeFirstCol ? "ring-2 ring-slate-400" : ""
              }`}
              title="Freeze first column"
            >
              <Snowflake className="w-4 h-4" />
              Freeze Col
            </button>

            <div className="w-px bg-slate-700 mx-1" />
            <details className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700">
              <summary className="cursor-pointer flex items-center gap-2">
                <Grid2x2Plus className="w-4 h-4" /> Pivot{" "}
                <ChevronDown className="w-4 h-4" />
              </summary>
              <div className="pt-2 space-y-2 text-sm">
                <div className="flex gap-2 items-center">
                  <span className="w-24 text-slate-400">Range</span>
                  <input
                    value={pivot.range}
                    onChange={(e) => setPivot({ ...pivot, range: e.target.value })}
                    placeholder="A1:C50"
                    className="px-2 py-1 rounded bg-slate-900 w-40"
                  />
                </div>
                <div className="flex gap-2 items-center">
                  <span className="w-24 text-slate-400">Row field (index)</span>
                  <input
                    type="number"
                    value={pivot.rowCol}
                    onChange={(e) =>
                      setPivot({ ...pivot, rowCol: parseInt(e.target.value) })
                    }
                    className="px-2 py-1 rounded bg-slate-900 w-24"
                  />
                </div>
                <div className="flex gap-2 items-center">
                  <span className="w-24 text-slate-400">Value field (index)</span>
                  <input
                    type="number"
                    value={pivot.valCol}
                    onChange={(e) =>
                      setPivot({ ...pivot, valCol: parseInt(e.target.value) })
                    }
                    className="px-2 py-1 rounded bg-slate-900 w-24"
                  />
                </div>
                <div className="flex gap-2 items-center">
                  <span className="w-24 text-slate-400">Agg</span>
                  <select
                    value={pivot.agg}
                    onChange={(e) => setPivot({ ...pivot, agg: e.target.value })}
                    className="px-2 py-1 rounded bg-slate-900 w-28"
                  >
                    <option>SUM</option>
                    <option>COUNT</option>
                  </select>
                </div>
                <button
                  onClick={buildPivot}
                  className="px-3 py-1.5 rounded bg-sky-700 hover:bg-sky-600"
                >
                  Build Pivot
                </button>
              </div>
            </details>

            <div className="w-px bg-slate-700 mx-1" />
            <button
              onClick={() => setShowWizard(true)}
              className="px-3 py-2 rounded-2xl bg-slate-800 hover:bg-slate-700 flex items-center gap-2"
              title="Function Wizard"
            >
              <FunctionSquare className="w-4 h-4" />
              Wizard
            </button>
          </div>
        </header>

        {/* Formula bar */}
        <div className="flex items-center gap-2 mb-2">
          <div className="w-24 text-sm text-slate-400">Cell</div>
          <div className="w-28 px-3 py-2 rounded-xl bg-slate-800">
            {addr(selection.r, selection.c)}
          </div>
          <div className="flex-1">
            <input
              ref={inputRef}
              value={isEditing ? editVal : selectedCell.input}
              onChange={(e) => setEditVal(e.target.value)}
              onFocus={() => setIsEditing(true)}
              onBlur={commitEdit}
              className="w-full px-3 py-2 rounded-xl bg-slate-800 outline-none"
              placeholder="Type value or Excel formula, e.g. =SUM(A1:B2)"
            />
          </div>
        </div>

        {/* Grid */}
        <div
          className="overflow-auto rounded-2xl bg-slate-900 ring-1 ring-slate-700 shadow-lg"
          tabIndex={0}
          onKeyDown={handleKeyDown}
          onDoubleClick={startEdit}
        >
          <table className="table-fixed select-none w-full border-collapse">
            <thead>
              <tr>
                <th
                  className={`sticky ${freezeFirstRow ? "top-0" : ""} ${
                    freezeFirstCol ? "left-0" : ""
                  } z-20 bg-slate-800 w-14 h-10`}
                ></th>
                {activeGrid[0].map((_, c) => (
                  <th
                    key={c}
                    className={`h-10 font-semibold text-slate-300 border-b border-slate-700 bg-slate-800 ${
                      freezeFirstRow ? "sticky top-0 z-10" : ""
                    }`}
                  >
                    {colToName(c)}
                  </th>
                ))}
              </tr>
              {/* Filter row */}
              <tr>
                <th
                  className={`${
                    freezeFirstCol ? "sticky left-0 z-10 bg-slate-800" : ""
                  }`}
                ></th>
                {activeGrid[0].map((_, c) => (
                  <th key={`f-${c}`} className="bg-slate-900 p-1 border-b border-slate-800">
                    <div className="flex items-center gap-1 bg-slate-800 rounded-xl px-2 py-1">
                      <FilterIcon className="w-3 h-3 opacity-70" />
                      <input
                        value={filters[c] || ""}
                        onChange={(e) =>
                          setFilters((arr) => {
                            const n = [...arr];
                            n[c] = e.target.value;
                            return n;
                          })
                        }
                        placeholder="filter"
                        className="w-full bg-transparent outline-none text-xs"
                      />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rowsToRender.map((row, r) => (
                <tr key={r}>
                  <th
                    className={`text-slate-300 border-r border-slate-700 font-medium w-14 ${
                      freezeFirstCol ? "sticky left-0 z-10 bg-slate-800" : ""
                    }`}
                  >
                    {r + 1}
                  </th>
                  {row.map((cell, c) => {
                    const isSel = selection.r === r && selection.c === c;
                    const invalid =
                      cell.validation?.type === "list" &&
                      !cell.input.startsWith("=") &&
                      (cell.validation.values?.length || 0) > 0 &&
                      !cell.validation.values.includes(cell.input);
                    return (
                      <td
                        key={`${r}-${c}`}
                        data-rc={`${r}-${c}`}
                        className={`h-10 px-3 whitespace-nowrap border border-slate-800/60 ${
                          isSel ? "outline outline-2 outline-sky-400" : ""
                        } ${invalid ? "ring-2 ring-red-500" : ""}`}
                        onClick={() => setSelection({ r, c })}
                        style={{
                          fontWeight: cell.fmt.bold ? 700 : 400,
                          fontStyle: cell.fmt.italic ? "italic" : "normal",
                          textAlign: cell.fmt.align,
                        }}
                        title={addr(r, c)}
                      >
                        {String(cell.value ?? "")}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Pivot output */}
        {pivotData && (
          <div className="mt-4 p-3 rounded-xl bg-slate-900 ring-1 ring-slate-700">
            <h2 className="font-semibold mb-2">Pivot Result</h2>
            {pivotData.error ? (
              <div className="text-red-400 text-sm">{pivotData.error}</div>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr>
                    <th className="text-left p-1">Key</th>
                    <th className="text-right p-1">COUNT</th>
                    <th className="text-right p-1">SUM</th>
                  </tr>
                </thead>
                <tbody>
                  {pivotData.rows.map((r, i) => (
                    <tr key={i} className="border-t border-slate-800">
                      <td className="p-1">{String(r.key)}</td>
                      <td className="p-1 text-right">{r.COUNT}</td>
                      <td className="p-1 text-right">
                        {Number(r.SUM).toLocaleString()}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Footnotes */}
        <div className="mt-4 text-sm text-slate-400 space-y-1">
          <p>
            ✅ Excel-compatible formulas powered by{" "}
            <code className="bg-slate-800 px-1 rounded">HyperFormula</code>.
            Hundreds of functions (SUMIFS, XLOOKUP*, INDEX/MATCH, IF,
            TEXT/DATE/TIME, stats, financial, etc.). *XLOOKUP availability
            depends on HF version.
          </p>
          <p>
            💡 Features: <strong>Freeze</strong> row/col, <strong>Filters</strong> per
            column, <strong>Function Wizard</strong>, <strong>Named Ranges</strong>,{" "}
            <strong>Validation</strong>, and a simple <strong>Pivot</strong>.
          </p>
        </div>
      </div>

      {/* Named Range Modal */}
      {showNameModal && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center">
          <div className="bg-slate-900 ring-1 ring-slate-700 rounded-2xl p-4 w-[420px]">
            <h3 className="font-semibold mb-2">Create Named Range</h3>
            <div className="space-y-2 text-sm">
              <div className="flex items-center gap-2">
                <span className="w-24 text-slate-400">Name</span>
                <input
                  value={nameKey}
                  onChange={(e) => setNameKey(e.target.value)}
                  placeholder="Sales"
                  className="flex-1 px-2 py-1 rounded bg-slate-800"
                />
              </div>
              <div className="flex items-center gap-2">
                <span className="w-24 text-slate-400">Refers to</span>
                <input
                  value={nameRef}
                  onChange={(e) => setNameRef(e.target.value)}
                  placeholder="A1:B10 or =A1:A10"
                  className="flex-1 px-2 py-1 rounded bg-slate-800"
                />
              </div>
              <div className="flex justify-end gap-2 mt-2">
                <button
                  onClick={() => setShowNameModal(false)}
                  className="px-3 py-1.5 rounded bg-slate-800"
                >
                  Cancel
                </button>
                <button
                  onClick={saveNamedRange}
                  className="px-3 py-1.5 rounded bg-sky-700"
                >
                  Save
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Validation Modal */}
      {showValidation && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center">
          <div className="bg-slate-900 ring-1 ring-slate-700 rounded-2xl p-4 w-[420px]">
            <h3 className="font-semibold mb-2">Data Validation (List)</h3>
            <div className="space-y-2 text-sm">
              <p className="text-slate-400">
                Enter comma-separated allowed values for this cell. Non-matching inputs
                (except formulas) will be rejected.
              </p>
              <input
                value={valList}
                onChange={(e) => setValList(e.target.value)}
                placeholder="High, Medium, Low"
                className="w-full px-2 py-1 rounded bg-slate-800"
              />
              <div className="flex justify-end gap-2 mt-2">
                <button
                  onClick={() => setShowValidation(false)}
                  className="px-3 py-1.5 rounded bg-slate-800"
                >
                  <X className="w-4 h-4 inline" /> Cancel
                </button>
                <button
                  onClick={applyValidationList}
                  className="px-3 py-1.5 rounded bg-sky-700"
                >
                  <Check className="w-4 h-4 inline" /> Apply
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Function Wizard */}
      {showWizard && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-end">
          <aside className="h-full w-[380px] bg-slate-950 ring-1 ring-slate-700 p-4 overflow-y-auto">
            <div className="flex items-center justify-between mb-3">
              <h3 className="font-semibold flex items-center gap-2">
                <FunctionSquare className="w-4 h-4" /> Function Wizard
              </h3>
              <button
                onClick={() => setShowWizard(false)}
                className="px-2 py-1 rounded bg-slate-800"
              >
                Close
              </button>
            </div>
            {Object.entries(functionCatalog).map(([group, funcs]) => (
              <div key={group} className="mb-4">
                <div className="text-slate-300 font-medium mb-2">{group}</div>
                <div className="space-y-2">
                  {funcs.map((f) => (
                    <div
                      key={f.name}
                      className="p-2 rounded-xl bg-slate-900 ring-1 ring-slate-800"
                    >
                      <div className="flex items-center justify-between">
                        <div className="font-semibold">{f.name}</div>
                        <button
                          onClick={() => insertTemplate(f.tpl)}
                          className="px-2 py-1 rounded bg-sky-700 hover:bg-sky-600 text-xs"
                        >
                          Insert
                        </button>
                      </div>
                      <div className="text-slate-400 text-xs mt-1">{f.hint}</div>
                      <code className="text-xs bg-slate-800 px-1 py-0.5 rounded mt-1 inline-block">
                        {f.tpl}
                      </code>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </aside>
        </div>
      )}
    </div>
  );
}
