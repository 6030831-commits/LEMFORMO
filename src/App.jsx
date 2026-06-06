import React, { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Профессии → операции ──────────────────────────────────────
// Зарплаты подобраны так, чтобы ставки совпадали с CLAUDE.md:
// wage/22/8/60×1.307 → A=12.07, B=21.04, C=10.52 ₽/мин
const DEFAULT_PROFESSIONS = [
  { id: "laser_op",  name: "Лазерщик",     wage: 97500,  workDays: 22 }, // A 12.07
  { id: "welder",    name: "Сварщик",      wage: 97500,  workDays: 22 }, // A 12.07
  { id: "painter",   name: "Маляр",        wage: 170000, workDays: 22 }, // B 21.04
  { id: "grinder",   name: "Шлифовщик",    wage: 85000,  workDays: 22 }, // C 10.52
  { id: "assembler", name: "Сборщик",      wage: 85000,  workDays: 22 }, // C 10.52
  { id: "general",   name: "Разнорабочий", wage: 75000,  workDays: 22 },
  { id: "mechanic",  name: "Слесарь",      wage: 85000,  workDays: 22 }, // C 10.52
  { id: "cnc_op",    name: "Оп. ЧПУ",      wage: 100000, workDays: 22 },
];

// rateGroup: A=лазер/сварка, B=покраска/тонировка, C=остальное
const WORKS = [
  { id: "laser",   name: "Лазер (резка)",        profId: "laser_op",  rateGroup: "A" },
  { id: "bandsaw", name: "Резка ленточной пилой", profId: "mechanic",  rateGroup: "C" },
  { id: "bend",    name: "Гибка",                profId: "mechanic",  rateGroup: "C" },
  { id: "weld",    name: "Сварка",               profId: "welder",    rateGroup: "A" },
  { id: "grind",   name: "Шлифовка / зачистка",  profId: "grinder",   rateGroup: "C" },
  { id: "tone",    name: "Тонировка",             profId: "painter",   rateGroup: "B" },
  { id: "paint",   name: "Покраска",             profId: "painter",   rateGroup: "B" },
  { id: "assy",    name: "Сборка",               profId: "assembler", rateGroup: "C" },
];

const CONSUMABLES = [
  { id: "wire",        name: "Сварочная проволока", unit: "кг"     },
  { id: "gas",         name: "Газ (сварка)",         unit: "баллон" },
  { id: "abrasive",    name: "Абразив (круг)",       unit: "шт"     },
  { id: "scotchbrite", name: "Скотчбрайт",           unit: "шт"     },
  { id: "paint_c",     name: "Краска",               unit: "кг"     },
  { id: "electricity", name: "Электроэнергия",       unit: "кВт"    },
];

const SALE_COEF = 2.0;

const NON_METAL_KW = ["фанер","мдф","лдсп","пластик","поликарбонат","пвх","дерев","древесин","хдф","оргстекл"];
function isNonMetal(part) {
  const hit = s => s && NON_METAL_KW.some(kw => s.toLowerCase().includes(kw));
  return hit(part.материал_тип) || hit(part.материал_сорт) || hit(part.наименование);
}

const SYSTEM_PROMPT = `Ты — анализатор технических чертежей ЛЕМФОРМО (металлическая мебель).
По чертежу извлеки ВСЕ детали и материалы (включая крепёж).

ПРАВИЛО ПРО ЛИСТОВЫЕ ДЕТАЛИ:
длина_мм и ширина_мм — ставить ТОЛЬКО если размеры явно указаны в спецификации или на размерных линиях чертежа.
НЕ вычислять, НЕ предполагать, НЕ выводить косвенно. Если явного указания нет → длина_мм: null, ширина_мм: null.

ТИПЫ ТРУБ И ПРОФИЛЕЙ:
- "труба круглая" — круглое сечение (⌀)
- "профиль квадратный" — квадратное замкнутое сечение
- "профиль прямоугольный" — прямоугольное замкнутое сечение
Всегда указывай точный тип. поле толщина_мм — стенка трубы/профиля.

НОРМАТИВЫ ТРУДОЗАТРАТ — используй для оценки время_мин (черновик, _draft:true):
- laser:   0.5 мин × кол-во листовых деталей (съём + сортировка)
- bandsaw: 0.8 мин × кол-во резов + 0.5 мин установка партии
- bend:    0.7 мин × кол-во гибов + 0.5 мин наладка профиля
- weld:    1.2 мин × кол-во коротких швов (прихватка входит)
- grind:   0.7 мин × кол-во швов
- tone:    только нержавейка/декор-покрытие; если сталь — 0; если неизвестно — null
- paint:   8 мин × 1 изделие (обезжир + грунт + краска; сушка не в ФОТ)
- assy:    0.4 мин × кол-во крепежей + 3 мин за фанеру/вкладыши (если есть)
Подсчитай фичи из чертежа (детали, резы, гибы, швы, крепежи), умножь на норматив.

ОСТАЛЬНЫЕ ПРАВИЛА:
- НЕ выдумывай данные без основания. Что не читается и не вычислимо — null + в требует_уточнения.
- Лист Ст3: полный лист 1500×3000 мм. Нержавейка (AISI 304/430): 1250×2500 мм. КИМ=0 для листа.
- Труба/профиль: КИМ=1.3. Масса = плотность × объём × КИМ.
- Плотности г/см³: Ст3 7.85, AISI 304 7.9, AISI 430 7.7.
- Крепёж/метизы: материал_тип="крепёж", масса_заготовки_кг=null, примечание — тип и кол-во.
- Каждая деталь/материал — отдельный объект в массиве "части".

Верни СТРОГО JSON без markdown, без пояснений:
{
  "наименование_изделия": "string",
  "обозначение": "string",
  "количество_шт": number,
  "части": [
    {
      "наименование": "string",
      "материал_тип": "лист|труба круглая|профиль квадратный|профиль прямоугольный|крепёж|прочее",
      "материал_сорт": "Ст3|AISI 304|AISI 430|string",
      "толщина_мм": number|null,
      "длина_мм": number|null,
      "ширина_мм": number|null,
      "количество_шт": number|null,
      "масса_заготовки_кг": number|null,
      "ким": number,
      "метраж_м": number|null,
      "примечание": "string|null"
    }
  ],
  "операции": [{"работа":"laser|bandsaw|bend|weld|grind|tone|paint|assy","время_мин":number|null}],
  "требует_уточнения": ["string"]
}`;

// ── Claude API ────────────────────────────────────────────────
async function callClaude(messages, system) {
  const apiKey = import.meta.env.VITE_ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("API ключ не задан. Создайте .env.local с VITE_ANTHROPIC_API_KEY=sk-ant-…");
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 16000, system, messages }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `HTTP ${res.status}`);
  }
  const data = await res.json();
  if (data.stop_reason === "max_tokens") {
    throw new Error("Ответ модели обрезан (превышен лимит токенов). Попробуйте упростить чертёж или разбить на части.");
  }
  return data.content.filter(b => b.type === "text").map(b => b.text).join("\n");
}

function extractBalancedJSON(text) {
  const start = text.indexOf("{");
  if (start === -1) return null;
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "{") depth++;
    else if (text[i] === "}") { depth--; if (depth === 0) return text.slice(start, i + 1); }
  }
  return null;
}

function parseJSON(text) {
  const clean = text.replace(/```json\s*/gi, "").replace(/```/g, "").trim();
  try { return JSON.parse(clean); } catch {}
  const extracted = extractBalancedJSON(clean);
  if (extracted) return JSON.parse(extracted);
  throw new Error("JSON не найден в ответе. Начало: " + clean.slice(0, 120));
}

// ── Авто-масса листа из размеров ─────────────────────────────
const DENSITIES = { "ст3": 7.85, "aisi 304": 7.90, "aisi 430": 7.70 };
function autoMassSheet(part) {
  if ((part.материал_тип ?? "").toLowerCase() !== "лист") return null;
  const { длина_мм: L, ширина_мм: W, толщина_мм: T, материал_сорт: сорт } = part;
  if (!L || !W || !T) return null;
  const key = Object.keys(DENSITIES).find(k => (сорт ?? "").toLowerCase().includes(k));
  if (!key) return null;
  return +(L * W * T / 1e6 * DENSITIES[key]).toFixed(3);
}

// ── Расчёт стоимости детали (за 1 шт) ────────────────────────
// Лист: (1/N) × масса_полного_листа × цена; фолбэк — масса_детали × цена
function calcPartCost(part, pricesPerKg) {
  if (isNonMetal(part)) return null;
  const тип = (part.материал_тип ?? "").toLowerCase();
  if (тип === "крепёж" || тип === "прочее") return null;
  const sort = (part.материал_сорт ?? "").trim();
  const priceKey = Object.keys(pricesPerKg).find(k => k.toLowerCase() === sort.toLowerCase());
  const price = priceKey ? pricesPerKg[priceKey] : null;
  if (!price) return null;

  if (тип === "лист") {
    const N = calcSheetShare(part);
    const T = part.толщина_мм;
    if (N != null && N > 0 && T) {
      const isInox = sort.toUpperCase().includes("AISI");
      const sL = isInox ? 1250 : 1500;
      const sW = isInox ? 2500 : 3000;
      const densKey = Object.keys(DENSITIES).find(k => sort.toLowerCase().includes(k));
      if (!densKey) return null;
      const sheetMass = sL * sW * T / 1e6 * DENSITIES[densKey];
      return +((1 / N) * sheetMass * price).toFixed(2);
    }
    const масса = part.масса_заготовки_кг ?? autoMassSheet(part);
    if (масса == null) return null;
    return +(масса * price).toFixed(2);
  }

  const масса = part.масса_заготовки_кг;
  if (масса == null) return null;
  if (тип.includes("труба") || тип.includes("профиль")) {
    return +(масса * (part.ким ?? 1.3) * price).toFixed(2);
  }
  return +(масса * price).toFixed(2);
}

// ── Доля листа ────────────────────────────────────────────────
function calcSheetShare(part) {
  if ((part.материал_тип ?? "").toLowerCase() !== "лист") return null;
  const dL = part.длина_мм, dW = part.ширина_мм;
  if (!dL || !dW) return null;
  const isInox = (part.материал_сорт ?? "").toUpperCase().includes("AISI");
  const sL = isInox ? 1250 : 1500;
  const sW = isInox ? 2500 : 3000;
  const n1 = Math.floor(sL / dL) * Math.floor(sW / dW);
  const n2 = Math.floor(sL / dW) * Math.floor(sW / dL);
  const N = Math.max(n1, n2);
  return N <= 0 ? -1 : N;
}

const SECTION_NAMES = ["Лист", "Труба круглая", "Профиль квадратный", "Профиль прямоугольный", "Крепёж / прочее"];
const MATERIAL_TYPES  = ["лист", "труба круглая", "профиль квадратный", "профиль прямоугольный", "крепёж", "прочее"];
const MATERIAL_GRADES = ["Ст3", "AISI 304", "AISI 430"];
function partSection(part) {
  const т = (part.материал_тип ?? "").toLowerCase();
  if (т === "лист") return 0;
  if (т.includes("круглая") || (т.includes("труба") && !т.includes("квадрат") && !т.includes("прямоуголь"))) return 1;
  if (т.includes("квадрат")) return 2;
  if (т.includes("прямоуголь") || т.includes("профиль")) return 3;
  return 4;
}

// ── Стили ─────────────────────────────────────────────────────
const S = {
  inp:  { background: "#0c0c0c", border: "1px solid #2a2a2a", color: "#e0e0e0", padding: "2px 5px", borderRadius: 3, fontSize: 11, fontFamily: "inherit" },
  td:   { padding: "3px 5px", borderBottom: "1px solid #1a1a1a", fontSize: 11 },
  th:   { padding: "4px 5px", fontSize: 10, color: "#555", fontWeight: 400, textAlign: "left", background: "#161616" },
  sec:  { fontSize: 10, color: "#ff6b35", textTransform: "uppercase", letterSpacing: 1 },
  card: { background: "#1a1a1a", borderRadius: 6 },
};

// ── Компонент ─────────────────────────────────────────────────
export default function App() {
  const [stage, setStage]       = useState("upload");
  const [fileName, setFileName] = useState("");
  const [busy, setBusy]         = useState(false);
  const [err, setErr]           = useState("");
  const fileB64  = useRef(null);
  const fileMime = useRef(null);

  const drawingUrlRef = useRef(null);
  const bodyGridRef   = useRef(null);
  const splitDragging = useRef(false);
  const [colSplit, setColSplit]         = useState(58);
  const [drawingUrl, setDrawingUrl]     = useState(null);
  const [drawingMime, setDrawingMime]   = useState(null);
  const [drawingModal, setDrawingModal] = useState(false);

  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") setDrawingModal(false); };
    const onMove = (e) => {
      if (!splitDragging.current || !bodyGridRef.current) return;
      const rect = bodyGridRef.current.getBoundingClientRect();
      const ratio = Math.max(25, Math.min(75, (e.clientX - rect.left) / rect.width * 100));
      setColSplit(+ratio.toFixed(1));
    };
    const onUp = () => {
      if (splitDragging.current) {
        splitDragging.current = false;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
      }
    };
    window.addEventListener("keydown", onKey);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, []);

  const [assembly, setAssembly] = useState({ наименование_изделия: "", обозначение: "", количество_шт: 1 });
  const [parts, setParts]       = useState([]);
  const [ops, setOps]           = useState(WORKS.map(w => ({ работа: w.id, время_мин: null, profId: w.profId })));
  const [clarifications, setClarifications] = useState([]);
  const [partCostOverrides, setPartCostOverrides] = useState({});

  const [professions, setProfessions]       = useState(DEFAULT_PROFESSIONS);
  const [metalTonnePrices, setMetalTonnePrices] = useState({ "Ст3": 60000, "AISI 304": 267000, "AISI 430": 166000 });
  const [consumablesData, setConsumablesData]   = useState(
    Object.fromEntries(CONSUMABLES.map(c => [c.id, { qty: "", unitCost: "" }]))
  );

  const [chat, setChat]         = useState([]);
  const [chatInput, setChatInput] = useState("");

  // ── Derived ──────────────────────────────────────────────────
  const pricesPerKg = Object.fromEntries(
    Object.entries(metalTonnePrices).map(([k, v]) => [k, +(v / 1000).toFixed(4)])
  );

  const profNM = (profId) => {
    const p = professions.find(p => p.id === profId);
    return p ? +(p.wage / p.workDays / 8 / 60 * 1.307).toFixed(2) : 0;
  };

  const getPartCost = (part, idx) => {
    if (partCostOverrides[idx] != null) return +partCostOverrides[idx];
    return calcPartCost(part, pricesPerKg);
  };

  const totalMatCost = +parts.reduce((acc, p, i) => {
    const isKr = (p.материал_тип ?? "").toLowerCase() === "крепёж";
    const c = getPartCost(p, i);
    if (c == null) return acc;
    return acc + (isKr ? c : c * (p.количество_шт ?? 1));
  }, 0).toFixed(2);

  let fot = 0;
  ops.forEach(op => {
    if (op.время_мин != null) {
      const profId = op.profId ?? WORKS.find(w => w.id === op.работа)?.profId;
      fot += op.время_мин * profNM(profId);
    }
  });
  fot = +fot.toFixed(2);

  let consTotal = 0;
  CONSUMABLES.forEach(c => {
    const d = consumablesData[c.id];
    consTotal += (parseFloat(d.qty) || 0) * (parseFloat(d.unitCost) || 0);
  });
  consTotal = +consTotal.toFixed(2);

  const cost = totalMatCost > 0 ? +(totalMatCost + fot + consTotal).toFixed(2) : null;
  const sale = cost != null ? +(cost * SALE_COEF).toFixed(2) : null;
  const today = new Date().toLocaleDateString("ru-RU");

  // ── Handlers ─────────────────────────────────────────────────
  function loadResult(r) {
    setAssembly({ наименование_изделия: r.наименование_изделия || "", обозначение: r.обозначение || "", количество_шт: r.количество_шт || 1 });
    const processedParts = (r.части || [])
      .filter(p => !isNonMetal(p))
      .map(p => {
        const т = (p.материал_тип ?? "").toLowerCase();
        let q = { ...p };
        if ((т.includes("труба") || т.includes("профиль")) && (!q.ким || q.ким === 0)) q = { ...q, ким: 1.3 };
        if (т === "лист" && q.масса_заготовки_кг == null) { const m = autoMassSheet(q); if (m != null) q = { ...q, масса_заготовки_кг: m }; }
        return q;
      });
    setParts(processedParts);
    setOps(WORKS.map(w => {
      const found = (r.операции || []).find(o => o.работа === w.id);
      return { работа: w.id, время_мин: found?.время_мин ?? null, profId: w.profId };
    }));
    setClarifications(r.требует_уточнения || []);
    setPartCostOverrides({});
  }

  const onFile = useCallback((e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setFileName(f.name); setErr("");
    if (drawingUrlRef.current) URL.revokeObjectURL(drawingUrlRef.current);
    const url = URL.createObjectURL(f);
    drawingUrlRef.current = url;
    setDrawingUrl(url);
    setDrawingMime(f.type);
    const reader = new FileReader();
    reader.onload = () => { fileB64.current = reader.result.split(",")[1]; fileMime.current = f.type; };
    reader.readAsDataURL(f);
  }, []);

  async function analyze() {
    if (!fileB64.current) { setErr("Сначала загрузите чертёж"); return; }
    setStage("analyzing"); setBusy(true); setErr("");
    try {
      const isPdf = fileMime.current === "application/pdf";
      const content = [
        isPdf
          ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: fileB64.current } }
          : { type: "image",    source: { type: "base64", media_type: fileMime.current,  data: fileB64.current } },
        { type: "text", text: "Извлеки все детали и материалы. Для листовых деталей укажи длина_мм и ширина_мм только если они явно видны в спецификации или на размерных линиях чертежа — иначе null." },
      ];
      const out = await callClaude([{ role: "user", content }], SYSTEM_PROMPT);
      loadResult(parseJSON(out));
      setStage("review");
    } catch (e) {
      setErr("Ошибка анализа: " + e.message); setStage("upload");
    } finally { setBusy(false); }
  }

  function editPart(idx, field, value) {
    setParts(prev => prev.map((p, i) => i === idx ? { ...p, [field]: value } : p));
  }

  function editPartType(idx, newType) {
    setParts(prev => prev.map((p, i) => {
      if (i !== idx) return p;
      const upd = { ...p, материал_тип: newType || null };
      const т = (newType ?? "").toLowerCase();
      if (т.includes("труба") || т.includes("профиль")) { if (!upd.ким || upd.ким === 0) upd.ким = 1.3; }
      else if (т === "лист") { upd.ким = 0; }
      return upd;
    }));
  }

  function addPart() {
    setParts(prev => [...prev, { наименование: "", материал_тип: "лист", материал_сорт: "Ст3", толщина_мм: null, длина_мм: null, ширина_мм: null, количество_шт: 1, масса_заготовки_кг: null, ким: 0, метраж_м: null, примечание: null }]);
  }

  function removePart(idx) {
    setParts(prev => prev.filter((_, i) => i !== idx));
    setPartCostOverrides(prev => {
      const next = {};
      Object.entries(prev).forEach(([k, v]) => { if (+k !== idx) next[+k > idx ? +k - 1 : +k] = v; });
      return next;
    });
  }

  function editOp(id, value) {
    setOps(prev => prev.map(o => o.работа === id ? { ...o, время_мин: value } : o));
  }

  async function sendChat() {
    if (!chatInput.trim()) return;
    const userMsg = chatInput.trim();
    setChat(c => [...c, { role: "user", text: userMsg }]);
    setChatInput(""); setBusy(true);
    try {
      const fullJson = { ...assembly, части: parts, операции: ops, требует_уточнения: clarifications };
      const prompt = `Текущая спецификация (JSON):\n${JSON.stringify(fullJson, null, 2)}\n\nПравка: "${userMsg}"\n\nВерни ОБНОВЛЁННЫЙ полный JSON в том же формате, строго JSON без markdown.`;
      const out = await callClaude([{ role: "user", content: prompt }], SYSTEM_PROMPT);
      loadResult(parseJSON(out));
      setChat(c => [...c, { role: "assistant", text: "Спецификация обновлена." }]);
    } catch (e) {
      setChat(c => [...c, { role: "assistant", text: "Ошибка: " + e.message }]);
    } finally { setBusy(false); }
  }

  function resetAll() {
    setStage("upload"); setFileName(""); setErr("");
    setAssembly({ наименование_изделия: "", обозначение: "", количество_шт: 1 });
    setParts([]); setOps(WORKS.map(w => ({ работа: w.id, время_мин: null })));
    setClarifications([]); setPartCostOverrides({}); setChat([]);
    if (drawingUrlRef.current) { URL.revokeObjectURL(drawingUrlRef.current); drawingUrlRef.current = null; }
    setDrawingUrl(null); setDrawingMime(null);
    fileB64.current = null; fileMime.current = null;
  }

  const onSplitterDown = (e) => {
    splitDragging.current = true;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    e.preventDefault();
  };

  // ── Excel export ─────────────────────────────────────────────
  function exportXLSX() {
    const wb = XLSX.utils.book_new();
    const ws = {};
    let R = 0;
    const s = (c, r, v) => { ws[XLSX.utils.encode_cell({ c, r })] = { t: "s", v: String(v ?? "") }; };
    const n = (c, r, v) => { if (v != null && v !== "") ws[XLSX.utils.encode_cell({ c, r })] = { t: "n", v: +v }; };
    const f = (c, r, fm) => { ws[XLSX.utils.encode_cell({ c, r })] = { t: "n", f: fm }; };

    s(0, R, `ЛЕМФОРМО · ${assembly.наименование_изделия || ""} · ${assembly.обозначение || ""} · ${today}`);
    R += 2;
    s(0, R, "Цена Ст3 ₽/кг");      n(1, R, metalTonnePrices["Ст3"]      / 1000);
    s(2, R, "Цена AISI 304 ₽/кг"); n(3, R, metalTonnePrices["AISI 304"] / 1000);
    s(4, R, "Цена AISI 430 ₽/кг"); n(5, R, metalTonnePrices["AISI 430"] / 1000);
    const PR = R + 1;
    const pF = r1 => `IF(C${r1}="Ст3",$B$${PR},IF(C${r1}="AISI 304",$D$${PR},IF(C${r1}="AISI 430",$F$${PR},0)))`;
    R += 2;
    ["Наименование","Тип","Марка","δ мм","Длина мм","Ширина мм","шт","Масса кг","Метраж м","КИМ","Доля листа","Стоимость ₽"]
      .forEach((h, c) => s(c, R, h));
    R++;

    const costCells = [];
    [0, 1, 2, 3, 4].forEach(secId => {
      const items = parts.filter(p => partSection(p) === secId);
      if (!items.length) return;
      s(0, R, `— ${SECTION_NAMES[secId]} —`); R++;
      items.forEach(part => {
        const r1 = R + 1;
        const т = (part.материал_тип ?? "").toLowerCase();
        const isSheet = т === "лист";
        const isTube  = т.includes("труба") || т.includes("профиль");
        s(0, R, part.наименование ?? ""); s(1, R, part.материал_тип ?? ""); s(2, R, part.материал_сорт ?? "");
        if (part.толщина_мм    != null) n(3, R, part.толщина_мм);
        if (part.длина_мм      != null) n(4, R, part.длина_мм);
        if (part.ширина_мм     != null) n(5, R, part.ширина_мм);
        if (part.количество_шт != null) n(6, R, part.количество_шт);
        if (isSheet) {
          const den = `IF(C${r1}="Ст3",7.85,IF(C${r1}="AISI 304",7.9,7.7))`;
          f(7, R, `IF(AND(E${r1}<>"",F${r1}<>"",D${r1}<>""),E${r1}*F${r1}*D${r1}/1000000*${den},"")`);
          const sL = `IF(C${r1}="Ст3",1500,1250)`, sW = `IF(C${r1}="Ст3",3000,2500)`;
          f(10, R, `IF(AND(E${r1}<>"",F${r1}<>""),1/MAX(FLOOR(${sL}/E${r1},1)*FLOOR(${sW}/F${r1},1),FLOOR(${sL}/F${r1},1)*FLOOR(${sW}/E${r1},1)),"")`);
          // Стоимость: (1/N) × масса_полного_листа × цена × кол-во; фолбэк: масса_детали × цена × кол-во
          const sheetMassF = `${sL}*${sW}*D${r1}/1000000*${den}`;
          f(11, R, `IF(K${r1}<>"",K${r1}*${sheetMassF}*${pF(r1)}*G${r1},IF(H${r1}<>"",H${r1}*${pF(r1)}*G${r1},""))`);
        } else {
          if (part.масса_заготовки_кг != null) n(7, R, part.масса_заготовки_кг);
          if (part.метраж_м           != null) n(8, R, part.метраж_м);
          n(9, R, part.ким ?? 1.3);
          f(11, R, `IF(H${r1}<>"",H${r1}*IF(J${r1}<>"",J${r1},1)*${pF(r1)}*G${r1},"")`);
        }
        costCells.push(`L${r1}`); R++;
      });
    });

    R++;
    const sumF  = costCells.length ? `SUM(${costCells.join(",")})` : "0";
    const mR = R + 1;
    s(10, R, "Итого металл");        f(11, R, sumF);                                   R++;
    s(10, R, "ФОТ");                 n(11, R, fot);                                    R++;
    s(10, R, "Расходники");          n(11, R, consTotal);                              R++;
    s(10, R, "Себестоимость");       f(11, R, `L${mR}+L${mR+1}+L${mR+2}`);            R++;
    s(10, R, "Цена продажи (×2)");  f(11, R, `L${mR+3}*2`);                           R++;

    ws["!ref"]  = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 11, r: R } });
    ws["!cols"] = [25,20,10,6,8,8,5,10,8,6,12,14].map(wch => ({ wch }));
    XLSX.utils.book_append_sheet(wb, ws, "Расчёт");
    XLSX.writeFile(wb, `${(assembly.обозначение || assembly.наименование_изделия || "расчёт").replace(/\s+/g,"_")}_себестоимость.xlsx`);
  }

  // ── Print ─────────────────────────────────────────────────────
  function printTable() {
    const px = Object.fromEntries(Object.entries(metalTonnePrices).map(([k,v]) => [k, v/1000]));
    const sectionRows = [0, 1, 2, 3, 4].map(secId => {
      const items = parts.filter(p => partSection(p) === secId);
      if (!items.length) return "";
      return `<tr><td colspan="5" style="background:#f0f0f0;font-weight:700;padding:3px 6px;font-size:10px">${SECTION_NAMES[secId]}</td></tr>` +
        items.map(part => {
          const c = calcPartCost(part, px);
          const qty = part.количество_шт ?? 1;
          const total = c != null ? c * qty : null;
          return `<tr>
            <td>${part.наименование ?? ""}</td>
            <td>${part.материал_тип ?? ""} · ${part.материал_сорт ?? ""}${part.толщина_мм ? " δ"+part.толщина_мм : ""}</td>
            <td>${qty}</td>
            <td>${part.масса_заготовки_кг ?? (autoMassSheet(part) != null ? "≈"+autoMassSheet(part) : "—")}</td>
            <td style="text-align:right">${total != null ? total.toLocaleString("ru-RU")+" ₽" : "—"}</td>
          </tr>`;
        }).join("");
    }).join("");
    const fotRows = WORKS.map(w => {
      const op     = ops.find(o => o.работа === w.id);
      const t      = op?.время_мин;
      const profId = op?.profId ?? w.profId;
      if (!t) return "";
      const nm  = profNM(profId);
      const sum = +(t * nm).toFixed(2);
      const pName = professions.find(p => p.id === profId)?.name ?? profId;
      return `<tr><td>${w.name}</td><td>${t} мин · ${pName} · ${nm} ₽/мин</td><td style="text-align:right">${sum.toLocaleString("ru-RU")} ₽</td></tr>`;
    }).join("");
    const win = window.open("", "_blank", "width=900,height=700");
    win.document.write(`<!DOCTYPE html><html lang="ru"><head><title>ЛЕМФОРМО · ${assembly.наименование_изделия || ""}</title>
      <style>body{font-family:monospace;font-size:11px;margin:20px}table{border-collapse:collapse;width:100%;margin-bottom:14px}th,td{border:1px solid #ccc;padding:3px 6px}th{background:#f5f5f5}h2{font-size:13px;margin:0 0 3px}p{font-size:10px;color:#666;margin:0 0 10px}.sum td{font-weight:700}</style>
      </head><body>
      <h2>ЛЕМФОРМО · ${assembly.наименование_изделия || "—"} · ${assembly.обозначение || ""}</h2>
      <p>${today} · ООО ЛЕМ-СТИЛЬ</p>
      <table><thead><tr><th>Наименование</th><th>Тип · Марка · δ</th><th>шт</th><th>Масса кг</th><th>Стоимость</th></tr></thead>
      <tbody>${sectionRows}<tr class="sum"><td colspan="4">Итого металл</td><td style="text-align:right">${totalMatCost.toLocaleString("ru-RU")} ₽</td></tr></tbody></table>
      ${fotRows ? `<table><thead><tr><th>Операция</th><th>Время</th><th>Стоимость</th></tr></thead><tbody>${fotRows}<tr class="sum"><td colspan="2">Итого ФОТ</td><td style="text-align:right">${fot.toLocaleString("ru-RU")} ₽</td></tr></tbody></table>` : ""}
      <table class="sum" style="width:280px"><tbody>
        <tr><td>Металл</td><td style="text-align:right">${totalMatCost.toLocaleString("ru-RU")} ₽</td></tr>
        <tr><td>ФОТ</td><td style="text-align:right">${fot.toLocaleString("ru-RU")} ₽</td></tr>
        <tr><td>Расходники</td><td style="text-align:right">${consTotal.toLocaleString("ru-RU")} ₽</td></tr>
        <tr><td>Себестоимость</td><td style="text-align:right">${(cost ?? 0).toLocaleString("ru-RU")} ₽</td></tr>
        <tr style="font-size:14px"><td>Цена продажи</td><td style="text-align:right">${(sale ?? 0).toLocaleString("ru-RU")} ₽</td></tr>
      </tbody></table>
      <script>window.onload=()=>window.print();</script></body></html>`);
    win.document.close();
  }

  // ── Render ───────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'IBM Plex Mono', ui-monospace, monospace", background: "#111", color: "#e0e0e0", height: "100vh", overflow: "hidden", padding: "8px 14px", display: "flex", flexDirection: "column" }}>
      <div style={{ width: "100%", maxWidth: 1800, margin: "0 auto", flex: 1, minHeight: 0, display: "flex", flexDirection: "column" }}>

        {/* Header */}
        <header style={{ borderBottom: "2px solid #ff6b35", paddingBottom: 5, marginBottom: 7, flexShrink: 0, display: "flex", alignItems: "baseline", gap: 20 }}>
          <h1 style={{ margin: 0, fontSize: 17, letterSpacing: 1 }}>ЛЕМФОРМО · Расчёт себестоимости</h1>
          <div style={{ fontSize: 10, color: "#444" }}>чертёж → спецификация → себестоимость → КП</div>
        </header>

        {/* Steps */}
        <div style={{ display: "flex", gap: 6, marginBottom: 7, fontSize: 10, flexShrink: 0 }}>
          {[["upload","Чертёж"],["review","Правки"],["final","Таблица"]].map(([s, label], i) => (
            <div key={s} style={{ padding: "2px 10px", borderRadius: 3,
              background: (stage === s || (stage === "analyzing" && s === "upload")) ? "#ff6b35" : "#1e1e1e",
              color:      (stage === s || (stage === "analyzing" && s === "upload")) ? "#111" : "#555" }}>
              {i + 1}. {label}
            </div>
          ))}
        </div>

        {err && <div style={{ background: "#3a1a1a", color: "#ff8a8a", padding: "5px 8px", borderRadius: 4, marginBottom: 6, fontSize: 11, flexShrink: 0 }}>{err}</div>}

        {/* ══ STAGE 1: Upload ══ */}
        {(stage === "upload" || stage === "analyzing") && (
          <div style={{ border: "1px dashed #333", borderRadius: 6, padding: 40, textAlign: "center" }}>
            <input type="file" accept="image/*,application/pdf" onChange={onFile} style={{ marginBottom: 16, color: "#aaa" }} />
            {fileName && <div style={{ fontSize: 12, color: "#555", marginBottom: 16 }}>{fileName}</div>}
            <button onClick={analyze} disabled={busy || !fileName}
              style={{ background: "#ff6b35", color: "#111", border: "none", padding: "10px 28px", borderRadius: 4, cursor: busy ? "wait" : "pointer", fontWeight: 700, fontSize: 14 }}>
              {busy ? "Анализирую чертёж…" : "Анализировать"}
            </button>
          </div>
        )}

        {/* ══ STAGE 2: Review ══ */}
        {stage === "review" && (
          <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", gap: 6 }}>

            {/* TOP BAR */}
            <div style={{ display: "flex", gap: 7, flexShrink: 0, alignItems: "stretch" }}>

              {/* Metal prices */}
              <div style={{ ...S.card, padding: "6px 10px", flexShrink: 0 }}>
                <div style={{ ...S.sec, marginBottom: 4 }}>Металл ₽/т</div>
                {Object.entries(metalTonnePrices).map(([grade, val]) => (
                  <div key={grade} style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 2 }}>
                    <span style={{ fontSize: 10, color: "#666", width: 54 }}>{grade}</span>
                    <input type="number" value={val} min="0"
                      onChange={e => setMetalTonnePrices(p => ({ ...p, [grade]: +e.target.value || 0 }))}
                      style={{ ...S.inp, width: 74 }} />
                    <span style={{ fontSize: 9, color: "#444", width: 44 }}>{(val/1000).toFixed(0)} ₽/кг</span>
                  </div>
                ))}
              </div>

              {/* Assembly */}
              <div style={{ ...S.card, padding: "6px 10px", flexShrink: 0 }}>
                <div style={{ ...S.sec, marginBottom: 4 }}>Изделие</div>
                <input type="text" value={assembly.наименование_изделия}
                  onChange={e => setAssembly(a => ({ ...a, наименование_изделия: e.target.value }))}
                  style={{ ...S.inp, width: 180, display: "block", marginBottom: 3 }} placeholder="наименование" />
                <div style={{ display: "flex", gap: 4 }}>
                  <input type="text" value={assembly.обозначение}
                    onChange={e => setAssembly(a => ({ ...a, обозначение: e.target.value }))}
                    style={{ ...S.inp, width: 108 }} placeholder="обозначение" />
                  <input type="number" value={assembly.количество_шт} min="1"
                    onChange={e => setAssembly(a => ({ ...a, количество_шт: +e.target.value || 1 }))}
                    style={{ ...S.inp, width: 40 }} />
                  <span style={{ fontSize: 9, color: "#444", alignSelf: "center" }}>шт</span>
                </div>
              </div>

              {/* Professions */}
              <div style={{ ...S.card, padding: "6px 10px", flex: 1, overflow: "hidden" }}>
                <div style={{ ...S.sec, marginBottom: 4 }}>Профессии · ₽/мин = ЗП / дни / 8 / 60 × 1.307</div>
                <div style={{ display: "flex", gap: 5, overflowX: "auto", paddingBottom: 2 }}>
                  {professions.map((prof, pi) => (
                    <div key={prof.id} style={{ background: "#141414", borderRadius: 3, padding: "4px 7px", flexShrink: 0 }}>
                      <input type="text" value={prof.name}
                        onChange={e => setProfessions(prev => prev.map((p, i) => i === pi ? { ...p, name: e.target.value } : p))}
                        style={{ ...S.inp, width: 88, fontWeight: 600, display: "block", marginBottom: 2 }} />
                      <div style={{ display: "flex", gap: 2, alignItems: "center" }}>
                        <input type="number" value={prof.wage}
                          onChange={e => setProfessions(prev => prev.map((p, i) => i === pi ? { ...p, wage: +e.target.value || 0 } : p))}
                          style={{ ...S.inp, width: 62, fontSize: 10 }} />
                        <span style={{ fontSize: 9, color: "#444" }}>/</span>
                        <input type="number" value={prof.workDays} min="1" placeholder="22"
                          onChange={e => setProfessions(prev => prev.map((p, i) => i === pi ? { ...p, workDays: +e.target.value || 1 } : p))}
                          style={{ ...S.inp, width: 40, fontSize: 10 }} />
                        <span style={{ fontSize: 11, fontWeight: 700, color: "#ff6b35", marginLeft: 3 }}>{profNM(prof.id)}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* BODY */}
            <div ref={bodyGridRef} style={{ flex: 1, minHeight: 0, display: "flex", gap: 0 }}>

              {/* LEFT: thumbnail + materials + consumables */}
              <div style={{ width: `${colSplit}%`, display: "flex", flexDirection: "column", gap: 5, minHeight: 0, paddingRight: 3 }}>

                {/* Thumbnail row */}
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                  {drawingUrl && (
                    <div onClick={() => setDrawingModal(true)}
                      style={{ cursor: "zoom-in", width: 88, height: 60, flexShrink: 0, borderRadius: 4, overflow: "hidden", border: "1px solid #2a2a2a", background: "#0c0c0c" }}>
                      {drawingMime?.startsWith("image/")
                        ? <img src={drawingUrl} style={{ width: "100%", height: "100%", objectFit: "contain" }} alt="" />
                        : <iframe src={drawingUrl + "#zoom=page-width&toolbar=0"} style={{ width: "100%", height: "100%", border: "none", pointerEvents: "none" }} title="" />
                      }
                    </div>
                  )}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 10, color: "#666", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{fileName}</div>
                    {clarifications.length > 0 && (
                      <div style={{ fontSize: 9, color: "#d4a017", marginTop: 2, lineHeight: 1.4 }}>
                        {clarifications.map((c, i) => <div key={i}>⚠ {c}</div>)}
                      </div>
                    )}
                  </div>
                  {drawingUrl && (
                    <button onClick={() => setDrawingModal(true)}
                      style={{ background: "#1e1e1e", color: "#ff6b35", border: "1px solid #333", padding: "2px 8px", borderRadius: 3, cursor: "pointer", fontSize: 10, flexShrink: 0 }}>
                      ↗
                    </button>
                  )}
                </div>

                {/* Materials */}
                <div style={{ ...S.card, flex: 1, minHeight: 0, display: "flex", flexDirection: "column" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 10px", flexShrink: 0 }}>
                    <div style={{ ...S.sec }}>Материалы · 1 деталь = 1 строка</div>
                    <button onClick={addPart} style={{ background: "#222", color: "#ff6b35", border: "1px solid #ff6b35", padding: "1px 8px", borderRadius: 3, cursor: "pointer", fontSize: 10 }}>+ строка</button>
                  </div>
                  <div style={{ flex: 1, minHeight: 0, overflowY: "auto", overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 820 }}>
                      <thead>
                        <tr>
                          {["Наименование","Тип","Марка","δмм","Дл.мм","Шир.мм","шт","Масса кг","Метр.м","КИМ","Доля/длина","Стоимость ₽",""].map(h => (
                            <th key={h} style={S.th}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {parts.length === 0 && (
                          <tr><td colSpan={13} style={{ ...S.td, color: "#444", textAlign: "center", padding: 12 }}>Нет деталей — загрузите чертёж</td></tr>
                        )}
                        {[0, 1, 2, 3, 4].flatMap(secId => {
                          const items = parts.map((p, i) => [p, i]).filter(([p]) => partSection(p) === secId);
                          const hdr = (
                            <tr key={`s${secId}`}>
                              <td colSpan={13} style={{ background: "#191919", padding: "3px 8px", fontSize: 9, color: "#ff6b35", textTransform: "uppercase", letterSpacing: 1 }}>
                                {SECTION_NAMES[secId]}
                              </td>
                            </tr>
                          );
                          const emptyRow = items.length === 0 ? (
                            <tr key={`se${secId}`}><td colSpan={13} style={{ ...S.td, color: "#333", fontStyle: "italic", padding: "3px 8px" }}>— нет деталей —</td></tr>
                          ) : null;
                          return [hdr, ...(emptyRow ? [emptyRow] : []), ...items.map(([part, idx]) => {
                            const тип   = (part.материал_тип ?? "").toLowerCase();
                            const c     = getPartCost(part, idx);
                            const isKr  = тип === "крепёж";
                            const isNM  = isNonMetal(part);
                            const autoM = autoMassSheet(part);
                            const warn  = { background: "#2a1a00" };
                            return (
                              <tr key={`p${idx}`} style={{ background: idx % 2 ? "#141414" : "transparent", opacity: isNM ? 0.3 : 1 }}>
                                <td style={S.td}><input type="text" value={part.наименование ?? ""} placeholder="—" onChange={e => editPart(idx, "наименование", e.target.value || null)} style={{ ...S.inp, width: 108 }} /></td>
                                <td style={S.td}>
                                  <select value={part.материал_тип ?? ""} onChange={e => editPartType(idx, e.target.value || null)} style={{ ...S.inp, width: 96 }}>
                                    <option value="">—</option>
                                    {MATERIAL_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                                  </select>
                                </td>
                                <td style={S.td}>
                                  <select value={part.материал_сорт ?? ""} onChange={e => editPart(idx, "материал_сорт", e.target.value || null)} style={{ ...S.inp, width: 72 }}>
                                    <option value="">—</option>
                                    {MATERIAL_GRADES.map(g => <option key={g} value={g}>{g}</option>)}
                                  </select>
                                </td>
                                <td style={S.td}><input type="number" value={part.толщина_мм ?? ""} placeholder="—" onChange={e => editPart(idx, "толщина_мм", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 40, ...(part.толщина_мм == null && !isKr ? warn : {}) }} /></td>
                                <td style={S.td}><input type="number" value={part.длина_мм ?? ""} placeholder="—" onChange={e => editPart(idx, "длина_мм", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 50 }} /></td>
                                <td style={S.td}><input type="number" value={part.ширина_мм ?? ""} placeholder="—" onChange={e => editPart(idx, "ширина_мм", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 50 }} /></td>
                                <td style={S.td}><input type="number" value={part.количество_шт ?? ""} placeholder="—" onChange={e => editPart(idx, "количество_шт", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 34 }} /></td>
                                <td style={S.td}><input type="number" value={part.масса_заготовки_кг ?? ""} placeholder={autoM != null ? `≈${autoM}` : "—"} onChange={e => editPart(idx, "масса_заготовки_кг", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 54, ...(part.масса_заготовки_кг == null && !isKr && autoM == null ? warn : {}) }} /></td>
                                <td style={S.td}><input type="number" value={part.метраж_м ?? ""} placeholder="—" onChange={e => editPart(idx, "метраж_м", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 46 }} /></td>
                                <td style={S.td}><input type="number" value={part.ким ?? ""} placeholder="—" onChange={e => editPart(idx, "ким", e.target.value === "" ? null : +e.target.value)} style={{ ...S.inp, width: 34 }} /></td>
                                <td style={S.td}>
                                  {тип === "лист" ? (() => {
                                    const N = calcSheetShare(part);
                                    if (N == null) return <span style={{ color: "#444" }}>—</span>;
                                    if (N === -1) return <span style={{ color: "#d4a017", fontSize: 10 }}>!&gt;л</span>;
                                    return <span style={{ color: "#ff6b35", fontWeight: 700 }}>{(1/N).toFixed(2)}</span>;
                                  })() : (тип.includes("труба") || тип.includes("профиль")) ? (
                                    part.метраж_м != null ? <span>{part.метраж_м} м</span> : <span style={{ color: "#444" }}>—</span>
                                  ) : <span style={{ color: "#444" }}>—</span>}
                                </td>
                                <td style={S.td}>
                                  {isKr ? (
                                    <div style={{ display: "flex", gap: 2, alignItems: "center" }}>
                                      <input type="number" value={partCostOverrides[idx] ?? ""} placeholder="ввод"
                                        onChange={e => setPartCostOverrides(prev => { const n = { ...prev }; e.target.value === "" ? delete n[idx] : n[idx] = +e.target.value; return n; })}
                                        style={{ ...S.inp, width: 56 }} />
                                      <span style={{ fontSize: 9, color: "#555" }}>₽</span>
                                    </div>
                                  ) : (
                                    <span style={{ color: c == null ? "#442200" : "#e0e0e0" }}>
                                      {c == null ? "нет" : (c * (part.количество_шт ?? 1)).toLocaleString("ru-RU") + " ₽"}
                                    </span>
                                  )}
                                </td>
                                <td style={S.td}>
                                  <button onClick={() => removePart(idx)} style={{ background: "none", border: "none", color: "#554", cursor: "pointer", fontSize: 13, lineHeight: 1 }}>×</button>
                                </td>
                              </tr>
                            );
                          })];
                        })}
                        <tr style={{ background: "#1a2a1a" }}>
                          <td colSpan={11} style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>Итого металл</td>
                          <td style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>{totalMatCost.toLocaleString("ru-RU")} ₽</td>
                          <td style={S.td} />
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>{/* end LEFT column */}

              {/* Splitter */}
              <div onMouseDown={onSplitterDown}
                style={{ width: 6, cursor: "col-resize", flexShrink: 0, background: "#1a1a1a",
                  borderLeft: "1px solid #222", borderRight: "1px solid #222" }}
                onMouseEnter={e => e.currentTarget.style.background = "#ff6b35"}
                onMouseLeave={e => e.currentTarget.style.background = "#1a1a1a"}
              />

              {/* RIGHT: stacked blocks */}
              <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 6, overflowY: "auto", paddingLeft: 3 }}>

                {/* Summary — at top */}
                <div style={{ ...S.card, flexShrink: 0, padding: "8px 10px" }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
                    <div style={{ flexShrink: 0, textAlign: "center", minWidth: 128 }}>
                      <div style={{ fontSize: 9, color: "#444" }}>{today}</div>
                      <div style={{ fontSize: 9, color: "#555", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 128 }}>{assembly.наименование_изделия || "—"}</div>
                      <div style={{ fontSize: 9, color: "#444", marginTop: 4 }}>себестоимость</div>
                      <div style={{ fontSize: 13, fontWeight: 700, color: cost == null ? "#554" : "#ddd" }}>
                        {cost == null ? "—" : cost.toLocaleString("ru-RU") + " ₽"}
                      </div>
                      <div style={{ fontSize: 9, color: "#444" }}>цена × 2.0</div>
                      <div style={{ fontSize: 20, fontWeight: 700, color: sale == null ? "#554" : "#ff6b35" }}>
                        {sale == null ? "—" : sale.toLocaleString("ru-RU") + " ₽"}
                      </div>
                      <button onClick={() => setStage("final")}
                        style={{ marginTop: 6, width: "100%", background: "#ff6b35", color: "#111", border: "none", padding: "5px 0", borderRadius: 4, cursor: "pointer", fontWeight: 700, fontSize: 11 }}>
                        Готовая таблица →
                      </button>
                    </div>
                    <div style={{ flex: 1, fontSize: 10, color: "#555", lineHeight: 1.9 }}>
                      <div>Металл: {totalMatCost.toLocaleString("ru-RU")} ₽</div>
                      <div>ФОТ: {fot.toLocaleString("ru-RU")} ₽</div>
                      <div>Расходники: {consTotal.toLocaleString("ru-RU")} ₽</div>
                      <div style={{ display: "flex", gap: 4, marginTop: 6 }}>
                        <input value={chatInput} onChange={e => setChatInput(e.target.value)}
                          onKeyDown={e => e.key === "Enter" && !busy && sendChat()}
                          placeholder="правка через чат…" style={{ ...S.inp, flex: 1 }} />
                        <button onClick={sendChat} disabled={busy}
                          style={{ background: "#222", color: "#fff", border: "none", padding: "0 8px", borderRadius: 3, cursor: "pointer", fontSize: 11 }}>
                          {busy ? "…" : "↵"}
                        </button>
                      </div>
                      {chat.slice(-2).map((m, i) => (
                        <div key={i} style={{ fontSize: 9, color: m.role === "user" ? "#666" : "#8ad", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {m.role === "user" ? "▸" : "◂"} {m.text}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Operations */}
                <div style={{ ...S.card, flexShrink: 0, resize: "vertical", overflow: "auto", minHeight: 80, height: 210, display: "flex", flexDirection: "column" }}>
                  <div style={{ ...S.sec, padding: "4px 10px 2px", flexShrink: 0 }}>Операции ФОТ</div>
                  <div style={{ flex: 1, minHeight: 0, overflowY: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead>
                        <tr>
                          <th style={S.th}>Операция</th>
                          <th style={{ ...S.th, width: 20 }}>Гр.</th>
                          <th style={{ ...S.th, width: 64 }}>Профессия</th>
                          <th style={{ ...S.th, width: 44 }}>₽/мин</th>
                          <th style={{ ...S.th, width: 70 }}>Мин</th>
                          <th style={{ ...S.th, width: 74 }}>₽</th>
                        </tr>
                      </thead>
                      <tbody>
                        {WORKS.map(w => {
                          const op     = ops.find(o => o.работа === w.id);
                          const t      = op?.время_мин ?? null;
                          const profId = op?.profId ?? w.profId;
                          const nm     = profNM(profId);
                          const sum    = t != null ? +(t * nm).toFixed(2) : null;
                          const grpColor = w.rateGroup === "A" ? "#f87" : w.rateGroup === "B" ? "#fd7" : "#8cf";
                          return (
                            <tr key={w.id}>
                              <td style={S.td}>{w.name}</td>
                              <td style={{ ...S.td, fontWeight: 700, color: grpColor, fontSize: 10, textAlign: "center" }}>{w.rateGroup}</td>
                              <td style={{ ...S.td, padding: "2px 3px" }}>
                                <select value={profId}
                                  onChange={e => setOps(prev => prev.map(o => o.работа === w.id ? { ...o, profId: e.target.value } : o))}
                                  style={{ ...S.inp, width: 86, fontSize: 10 }}>
                                  {professions.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                                </select>
                              </td>
                              <td style={{ ...S.td, color: "#ff6b35" }}>{nm}</td>
                              <td style={S.td}>
                                <input type="number" min="0" value={t ?? ""} placeholder="—"
                                  onChange={e => editOp(w.id, e.target.value === "" ? null : +e.target.value)}
                                  style={{ ...S.inp, width: 58 }} />
                              </td>
                              <td style={{ ...S.td, color: sum == null ? "#333" : "#e0e0e0" }}>
                                {sum == null ? "—" : sum.toLocaleString("ru-RU")}
                              </td>
                            </tr>
                          );
                        })}
                        <tr style={{ background: "#1a2a1a" }}>
                          <td colSpan={5} style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>Итого ФОТ</td>
                          <td style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>{fot.toLocaleString("ru-RU")} ₽</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Consumables */}
                <div style={{ ...S.card, flexShrink: 0, resize: "vertical", overflow: "auto", minHeight: 60 }}>
                  <div style={{ ...S.sec, padding: "4px 10px 2px" }}>Расходники</div>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={S.th}>Наименование</th>
                        <th style={{ ...S.th, width: 86 }}>Кол-во</th>
                        <th style={{ ...S.th, width: 68 }}>₽/ед.</th>
                        <th style={{ ...S.th, width: 74 }}>₽</th>
                      </tr>
                    </thead>
                    <tbody>
                      {CONSUMABLES.map(c => {
                        const d   = consumablesData[c.id];
                        const qty = parseFloat(d.qty) || 0;
                        const uc  = parseFloat(d.unitCost) || 0;
                        const sum = qty && uc ? +(qty * uc).toFixed(2) : null;
                        return (
                          <tr key={c.id}>
                            <td style={S.td}>{c.name}</td>
                            <td style={S.td}>
                              <div style={{ display: "flex", alignItems: "center", gap: 2 }}>
                                <input type="number" min="0" value={d.qty} placeholder="—"
                                  onChange={e => setConsumablesData(prev => ({ ...prev, [c.id]: { ...prev[c.id], qty: e.target.value } }))}
                                  style={{ ...S.inp, width: 46 }} />
                                <span style={{ fontSize: 9, color: "#555" }}>{c.unit}</span>
                              </div>
                            </td>
                            <td style={S.td}>
                              <input type="number" min="0" value={d.unitCost} placeholder="—"
                                onChange={e => setConsumablesData(prev => ({ ...prev, [c.id]: { ...prev[c.id], unitCost: e.target.value } }))}
                                style={{ ...S.inp, width: 62 }} />
                            </td>
                            <td style={{ ...S.td, color: sum == null ? "#333" : "#e0e0e0" }}>
                              {sum == null ? "—" : sum.toLocaleString("ru-RU") + " ₽"}
                            </td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: "#1a2a1a" }}>
                        <td colSpan={3} style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>Итого расходники</td>
                        <td style={{ ...S.td, fontWeight: 700, color: "#8f8" }}>{consTotal.toLocaleString("ru-RU")} ₽</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

              </div>
            </div>
          </div>
        )}

        {/* ══ STAGE 3: Final table ══ */}
        {stage === "final" && (
          <div style={{ flex: 1, minHeight: 0, overflowY: "auto" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 200px 1fr", gap: 14 }}>

              <div style={S.card}>
                <div style={{ ...S.sec, padding: "6px 12px 3px" }}>Материалы</div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <tbody>
                    {parts.filter(p => !isNonMetal(p)).map((part, idx) => {
                      const c = getPartCost(part, idx);
                      const тип = (part.материал_тип ?? "").toLowerCase();
                      const isSheet = тип === "лист";
                      const isTube  = тип === "труба" || тип === "профиль";
                      const N = isSheet ? calcSheetShare(part) : null;
                      return (
                        <tr key={idx}>
                          <td style={{ ...S.td, fontSize: 10 }}>
                            {part.наименование}
                            {part.примечание && <span style={{ color: "#555" }}> · {part.примечание}</span>}
                            <div style={{ color: "#555", fontSize: 9 }}>
                              {part.материал_тип} · {part.материал_сорт}
                              {part.толщина_мм ? ` · δ${part.толщина_мм}` : ""}
                              {isSheet && part.длина_мм && part.ширина_мм ? ` · ${part.длина_мм}×${part.ширина_мм} мм` : ""}
                              {isTube  && part.метраж_м != null ? ` · ${part.метраж_м} м` : ""}
                              {part.количество_шт ? ` · ${part.количество_шт} шт` : ""}
                              {isSheet && N != null && N > 0 && <span style={{ color: "#ff6b35" }}> · {(1/N).toFixed(2)} листа</span>}
                              {isSheet && N === -1  && <span style={{ color: "#d4a017" }}> · !&gt;листа</span>}
                            </div>
                          </td>
                          <td style={{ ...S.td, width: 90, fontSize: 10, textAlign: "right" }}>
                            {c == null ? <span style={{ background: "#fff3b0", color: "#111", padding: "1px 3px" }}>?</span> : c.toLocaleString("ru-RU") + " ₽"}
                          </td>
                        </tr>
                      );
                    })}
                    <tr><td style={{ ...S.td, fontWeight: 700 }}>Итого металл</td><td style={{ ...S.td, fontWeight: 700, textAlign: "right" }}>{totalMatCost.toLocaleString("ru-RU")} ₽</td></tr>
                  </tbody>
                </table>
                <div style={{ ...S.sec, padding: "6px 12px 3px" }}>Расходники</div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <tbody>
                    {CONSUMABLES.map(c => {
                      const d   = consumablesData[c.id];
                      const sum = (parseFloat(d.qty)||0) && (parseFloat(d.unitCost)||0) ? +((parseFloat(d.qty)||0)*(parseFloat(d.unitCost)||0)).toFixed(2) : null;
                      return (
                        <tr key={c.id}>
                          <td style={{ ...S.td, fontSize: 10 }}>{c.name}</td>
                          <td style={{ ...S.td, fontSize: 10, color: "#555" }}>{d.qty || "—"} {c.unit}</td>
                          <td style={{ ...S.td, fontSize: 10, textAlign: "right" }}>{sum == null ? "—" : sum.toLocaleString("ru-RU") + " ₽"}</td>
                        </tr>
                      );
                    })}
                    <tr><td colSpan={2} style={{ ...S.td, fontWeight: 700 }}>Итого расходники</td><td style={{ ...S.td, fontWeight: 700, textAlign: "right" }}>{consTotal.toLocaleString("ru-RU")} ₽</td></tr>
                  </tbody>
                </table>
              </div>

              <div style={{ ...S.card, padding: 12, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: "#444" }}>{today}</div>
                <div style={{ fontSize: 11, color: "#666", marginTop: 4 }}>{assembly.наименование_изделия || "—"}</div>
                <div style={{ fontSize: 10, color: "#444" }}>{assembly.обозначение || ""}</div>
                <div style={{ fontSize: 28, fontWeight: 700, color: sale == null ? "#554" : "#fff", margin: "14px 0 3px" }}>
                  {sale == null ? "?" : sale.toLocaleString("ru-RU") + " ₽"}
                </div>
                <div style={{ fontSize: 10, color: "#444" }}>с/с {cost == null ? "—" : cost.toLocaleString("ru-RU")} ₽ × {SALE_COEF}</div>
                <div style={{ marginTop: 10, fontSize: 9, color: "#444", textAlign: "left", lineHeight: 2 }}>
                  <div>Металл: {totalMatCost.toLocaleString("ru-RU")} ₽</div>
                  <div>ФОТ: {fot.toLocaleString("ru-RU")} ₽</div>
                  <div>Расходники: {consTotal.toLocaleString("ru-RU")} ₽</div>
                  <div style={{ borderTop: "1px solid #222", paddingTop: 4, marginTop: 4, color: "#333" }}>накладные 20% — v2</div>
                </div>
              </div>

              <div style={S.card}>
                <div style={{ ...S.sec, padding: "6px 12px 3px" }}>Работы (ФОТ)</div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <tbody>
                    {WORKS.map(w => {
                      const op     = ops.find(o => o.работа === w.id);
                      const t      = op?.время_мин ?? null;
                      const profId = op?.profId ?? w.profId;
                      const sum    = t != null ? +(t * profNM(profId)).toFixed(2) : null;
                      const pName  = professions.find(p => p.id === profId)?.name ?? profId;
                      return (
                        <tr key={w.id}>
                          <td style={{ ...S.td, fontSize: 10 }}>{w.name}</td>
                          <td style={{ ...S.td, fontSize: 10, color: "#555", width: 40 }}>{t ?? "—"} мин</td>
                          <td style={{ ...S.td, fontSize: 9, color: "#555", width: 70 }}>{pName} · {profNM(profId)} ₽/мин</td>
                          <td style={{ ...S.td, fontSize: 10, textAlign: "right", width: 80 }}>{sum == null ? "—" : sum.toLocaleString("ru-RU") + " ₽"}</td>
                        </tr>
                      );
                    })}
                    <tr><td style={{ ...S.td, fontWeight: 700 }}>Итого ФОТ</td><td style={S.td} /><td style={{ ...S.td, fontWeight: 700, textAlign: "right" }}>{fot.toLocaleString("ru-RU")} ₽</td></tr>
                  </tbody>
                </table>
              </div>
            </div>

            <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => setStage("review")} style={{ background: "#1e1e1e", color: "#ccc", border: "none", padding: "8px 16px", borderRadius: 4, cursor: "pointer" }}>
                ← Вернуть правки
              </button>
              <button onClick={exportXLSX} style={{ background: "#1a3a1a", color: "#8f8", border: "1px solid #2a4a2a", padding: "8px 16px", borderRadius: 4, cursor: "pointer", fontWeight: 700 }}>
                ↓ Сохранить Excel
              </button>
              <button onClick={printTable} style={{ background: "#1a1a3a", color: "#88f", border: "1px solid #2a2a4a", padding: "8px 16px", borderRadius: 4, cursor: "pointer", fontWeight: 700 }}>
                ⎙ Печать
              </button>
              <button onClick={resetAll} style={{ background: "#ff6b35", color: "#111", border: "none", padding: "8px 16px", borderRadius: 4, cursor: "pointer", fontWeight: 700 }}>
                Новый чертёж
              </button>
            </div>
          </div>
        )}

      </div>

      {/* ══ Drawing modal ══ */}
      {drawingModal && drawingUrl && (
        <div onClick={() => setDrawingModal(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.93)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={e => e.stopPropagation()}
            style={{ width: "94vw", height: "94vh", display: "flex", flexDirection: "column", gap: 6 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 11, color: "#666" }}>{fileName}</span>
              <button onClick={() => setDrawingModal(false)}
                style={{ background: "none", border: "1px solid #444", color: "#ff6b35", padding: "3px 12px", borderRadius: 4, cursor: "pointer", fontSize: 12 }}>
                ✕ закрыть (Esc)
              </button>
            </div>
            {drawingMime?.startsWith("image/")
              ? <img src={drawingUrl} style={{ flex: 1, width: "100%", objectFit: "contain", borderRadius: 6, minHeight: 0 }} alt="чертёж" />
              : <iframe src={drawingUrl + "#zoom=page-width&toolbar=0"} style={{ flex: 1, width: "100%", border: "none", borderRadius: 6 }} title="чертёж" />
            }
          </div>
        </div>
      )}

    </div>
  );
}
