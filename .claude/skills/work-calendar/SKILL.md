---
name: work-calendar
description: >
  Creates an interactive monthly work calendar widget for employee shift scheduling directly in chat.
  Use this skill whenever the user wants to plan employee schedules, assign shifts, build a work timetable,
  manage staff rosters, or track who works when. Triggers on: "рабочий календарь", "график смен",
  "расписание сотрудников", "планирование смен", "work schedule", "shift calendar", "staff roster",
  "employee timetable", or any request to organize/visualize when employees work.
  Always use this skill — even for simple requests like "сделай график на месяц" or "покажи расписание".
---

# Work Calendar Skill

Builds a **fully interactive monthly shift calendar** as an inline widget using `visualize:show_widget`.
No files needed — the calendar renders directly in the conversation.

---

## What this skill produces

A single-page HTML widget containing:
- **Monthly grid** — employees as rows, days as columns, today highlighted
- **Shift assignment** — click any cell to open a modal and assign: 5/2 (8h, standard week), 2/2 (12h, rotating), Отпуск (Vacation), Больничный (Sick leave), Выходной (Day off)
- **Auto stats** — live counters for each shift type across the month
- **Add employee** — inline form to add name + role
- **Month navigation** — prev/next arrows
- **CSV export** — UTF-8 download with hours totals
- **Google Sheets panel** — step-by-step guide + ready Apps Script code
- **Telegram panel** — step-by-step guide + ready Node.js bot code
- **Legend** — color-coded shift key

---

## Implementation guide

### 1. Always call `visualize:read_me` first
Load modules `["mockup", "interactive"]` before generating the widget.

### 2. Data model

```js
// State
let employees = [{ id, name, role }, ...]
let schedule = {}  // key: `${year}-${month}-${empId}-${day}` → 'A'|'B'|'V'|'S'|'O'

// Shift codes
const SHIFTS = {
  A: { label: '5/2', hours: 8,  css: 's-A', time: '8 ч / день' },  // График 5 через 2
  B: { label: '2/2', hours: 12, css: 's-B', time: '12 ч / день' }, // График 2 через 2
  V: { label: 'От',  hours: 0,  css: 's-V', time: 'Отпуск'      },
  S: { label: 'Б',   hours: 0,  css: 's-S', time: 'Больничный'  },
  O: { label: '—',   hours: 0,  css: 's-O', time: 'Выходной'    }
}
```

### 3. Color palette (CSS classes on `.shift-inner`)

| Shift | Background | Text color | Label |
|-------|-----------|------------|-------|
| `s-A` | `#dbeafe` | `#1e40af`  | 5/2 (8 ч) |
| `s-B` | `#ede9fe` | `#5b21b6`  | 2/2 (12 ч) |
| `s-V` | `#dcfce7` | `#166534`  | Отпуск |
| `s-S` | `#fee2e2` | `#991b1b`  | Больничный |
| `s-O` | `transparent` | secondary text | Выходной |

### 4. Table layout

```
<table>
  <thead>
    <tr>
      <th class="emp-col">Сотрудник</th>  <!-- sticky, min-width 130px -->
      <th class="day-col">1<br>Пн</th>    <!-- width 32px per day -->
      ...
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="emp-cell">name + role</td>
      <td class="shift-cell" onclick="openModal(empId, day)">
        <div class="shift-inner s-D">Д</div>
      </td>
      ...
    </tr>
  </tbody>
</table>
```

Key CSS rules:
- `border-collapse: collapse` on table
- Weekend columns: add class `.weekend` → `background: rgba(0,0,0,0.03)`
- Today column: `.today-col` → `outline: 1.5px solid #2563eb`
- `.shift-cell:hover { filter: brightness(0.88) }`

### 5. Modal (shift picker)

Use `position: fixed` with `inset: 0` only when rendering inside the widget's own iframe context.
Grid of 2 shift buttons (5/2 — 8ч, 2/2 — 12ч) + "Отпуск" + "Больничный" + "Выходной / снять" + "Отмена".
Store `modalCtx = { empId, day }` and call `setSchedule()` on selection.

### 6. Stats row

4 metric cards (5/2 shifts / 2/2 shifts / Vacation / Sick) in a `grid-template-columns: repeat(4, minmax(0, 1fr))` grid.
Recompute on every `render()` call.

### 7. Integration panels

When the user clicks **Google Sheets** or **Telegram** buttons, show a collapsible panel below the calendar with:
- Two tabs: "Инструкция" (numbered steps) and "Код" (code block with copy button)
- A "Закрыть" button

**Google Sheets — визуальный календарь (Apps Script):**

Apps Script получает данные из виджета через Web App и рисует полноценный цветной календарь прямо в Google Sheets: шапка с датами и днями недели, ячейки с цветом смены, итоговая колонка с часами, выходные затемнены.

```js
// ===== ВСТАВИТЬ В Apps Script =====
const SHEET_NAME = "График";

const SHIFT_COLORS = {
  "5/2": { bg: "#DBEAFE", fg: "#1E40AF", label: "8 ч/день"  },
  "2/2": { bg: "#EDE9FE", fg: "#5B21B6", label: "12 ч/день" },
  "От": { bg: "#DCFCE7", fg: "#166534", label: "Отпуск"     },
  "Б":  { bg: "#FEE2E2", fg: "#991B1B", label: "Больничный" },
  "—":  { bg: "#F8F8F8", fg: "#9CA3AF", label: "Выходной"   },
};
const WEEKEND_BG = "#F3F4F6";
const HEADER_BG  = "#1A7340";
const HOURS = { "5/2": 8, "2/2": 12, "От": 0, "Б": 0, "—": 0 };

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

// GET → возвращает текущие данные таблицы в JSON
function doGet() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST → рисует визуальный календарь
function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  // Обновление одной ячейки (real-time клик в виджете)
  if (body.action === "cell") {
    const sheet = getOrCreateSheet();
    const cell = sheet.getRange(body.row, body.col);
    const val = body.value;
    const sc = SHIFT_COLORS[val] || SHIFT_COLORS["—"];
    cell.setValue(val)
        .setBackground(sc.bg)
        .setFontColor(sc.fg)
        .setHorizontalAlignment("center")
        .setFontWeight(val === "—" ? "normal" : "bold");
    updateTotals(sheet);
    return ok();
  }

  // Полная перерисовка календаря
  renderCalendar(body.rows, body.year, body.month);
  return ok();
}

function renderCalendar(rows, year, month) {
  const sheet = getOrCreateSheet();
  sheet.clearContents();
  sheet.clearFormats();

  const days = new Date(year, month + 1, 0).getDate();
  const totalCols = days + 3; // имя + должность + дни + итого

  // ── Строка 1: заголовок месяца ──
  const MONTH_NAMES = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                       "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
  sheet.getRange(1, 1, 1, totalCols).merge()
    .setValue(`${MONTH_NAMES[month]} ${year}`)
    .setBackground(HEADER_BG).setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 32);

  // ── Строка 2: шапка дней ──
  const DOW_RU = ["Вс","Пн","Вт","Ср","Чт","Пт","Сб"];
  const hdrRow = ["Сотрудник", "Должность"];
  for (let d = 1; d <= days; d++) {
    const dw = new Date(year, month, d).getDay();
    hdrRow.push(`${d}\n${DOW_RU[dw]}`);
  }
  hdrRow.push("Часов");

  const hdrRange = sheet.getRange(2, 1, 1, totalCols);
  hdrRange.setValues([hdrRow])
    .setBackground("#2D6A4F").setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setRowHeight(2, 38);

  // Выходные — подсветить заголовок
  for (let d = 1; d <= days; d++) {
    const dw = new Date(year, month, d).getDay();
    if (dw === 0 || dw === 6) {
      sheet.getRange(2, d + 2).setBackground("#1A5C3A");
    }
  }

  // ── Строки сотрудников ──
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];  // [name, role, d1, d2, ..., hours]
    const sheetRow = r + 3;
    sheet.setRowHeight(sheetRow, 26);

    // Имя + должность
    sheet.getRange(sheetRow, 1).setValue(row[0])
      .setFontWeight("bold").setFontSize(11)
      .setVerticalAlignment("middle");
    sheet.getRange(sheetRow, 2).setValue(row[1])
      .setFontColor("#6B7280").setFontSize(10)
      .setVerticalAlignment("middle");

    // Ячейки смен
    for (let d = 0; d < days; d++) {
      const val = row[d + 2];
      const dw = new Date(year, month, d + 1).getDay();
      const isWe = dw === 0 || dw === 6;
      const sc = SHIFT_COLORS[val] || SHIFT_COLORS["—"];
      const cell = sheet.getRange(sheetRow, d + 3);
      cell.setValue(val === "—" ? "" : val)
        .setBackground(isWe && val === "—" ? WEEKEND_BG : sc.bg)
        .setFontColor(sc.fg)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setFontWeight(val === "—" ? "normal" : "bold")
        .setFontSize(10);
    }

    // Итого часов
    const totalHours = row.slice(2, 2 + days)
      .reduce((s, v) => s + (HOURS[v] || 0), 0);
    sheet.getRange(sheetRow, days + 3).setValue(totalHours)
      .setFontWeight("bold").setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#F0FDF4").setFontColor("#166534");
  }

  // ── Ширина колонок ──
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 110);
  for (let d = 1; d <= days; d++) sheet.setColumnWidth(d + 2, 30);
  sheet.setColumnWidth(days + 3, 60);

  // ── Границы ──
  const fullRange = sheet.getRange(1, 1, rows.length + 2, totalCols);
  fullRange.setBorder(true, true, true, true, true, true,
    "#D1D5DB", SpreadsheetApp.BorderStyle.SOLID);

  // Заморозить первые 2 строки и 2 колонки
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
}

function updateTotals(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) return;
  for (let r = 3; r <= lastRow; r++) {
    const vals = sheet.getRange(r, 3, 1, lastCol - 3).getValues()[0];
    const total = vals.reduce((s, v) => s + (HOURS[v] || 0), 0);
    sheet.getRange(r, lastCol).setValue(total);
  }
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 График")
    .addItem("Показать Web App URL", "showUrl")
    .addToUi();
}

function showUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert("Скопируйте этот URL и вставьте в виджет:\n\n" + url);
}
```

**Что рисует скрипт в таблице:**
- Строка 1 — название месяца (тёмно-зелёная шапка, объединённая)
- Строка 2 — числа дней + день недели двумя строками, выходные темнее
- Строки сотрудников — ячейки с цветом смены (У=синий, Д=жёлтый, Отпуск=зелёный, Больничный=красный, выходной=серый)
- Последняя колонка — итого рабочих часов за месяц
- Первые 2 строки и 2 колонки заморожены для прокрутки

**Передавать из виджета нужно:**
```js
// При pushToSheets — добавить year и month в тело запроса
body: JSON.stringify({ rows, year, month })
```

**Telegram bot code** (Node.js):
```js
const BOT_TOKEN = "ВАШ_ТОКЕН";
const CHAT_ID = "ВАШ_CHAT_ID";
async function sendSchedule(text) {
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: "POST", headers: {"Content-Type":"application/json"},
    body: JSON.stringify({ chat_id: CHAT_ID, text, parse_mode: "HTML" })
  });
}
```

Telegram setup steps:
1. Создайте бота через @BotFather → /newbot → получите токен
2. Добавьте бота в чат, получите CHAT_ID через @userinfobot
3. Разверните на Node.js или Cloudflare Workers
4. Настройте cron на 18:00 для отправки завтрашнего графика

### 8. CSV export

```js
function exportCSV() {
  const days = daysInMonth(year, month);
  let csv = 'Сотрудник,Должность';
  for (let d = 1; d <= days; d++) csv += `,${d}`;
  csv += ',Итого часов\n';
  for (const emp of employees) {
    let hours = 0, row = `"${emp.name}","${emp.role}"`;
    for (let d = 1; d <= days; d++) {
      const s = getShift(emp.id, d);
      row += `,${SHIFTS[s].label}`;
      hours += SHIFTS[s].hours;
    }
    csv += row + `,${hours}\n`;
  }
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `schedule_${year}_${month+1}.csv`;
  a.click();
}
```

---

## Default employees

Start with an empty list — no seed data:

```js
let employees = [];
let eid = 1;
```

If the user provides their own employee list, pre-populate from it.

---

## Customisation hints

| User request | How to adapt |
|---|---|
| Different shift times | Update labels and `shiftHours` values |
| More shift types | Add new codes to `SHIFTS`, add modal buttons |
| Custom employee list | Pre-populate `employees` array from user input |
| Pre-filled schedule | Seed `schedule` map before first `render()` |
| Different language | Swap Russian labels to requested locale |
| Dark/print color scheme | Adjust the `s-*` CSS classes |

---

## Loading messages (use these for `visualize:show_widget`)

```
["Расставляем смены по дням...", "Загружаем сотрудников...", "Настраиваем интеграции..."]
```
