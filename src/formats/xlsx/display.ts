/**
 * Grid math shared by the xlsx and xls engines: A1 references, Excel serial
 * dates, numFmt-aware display rendering. Pure functions - no library imports.
 */

// -- A1 reference math ---------------------------------------------------------

/** 1-based everywhere: 1 -> A, 26 -> Z, 27 -> AA. */
export function colName(col: number): string {
	let n = col;
	let out = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out || "A";
}

/** Letters to 1-based column index: A -> 1, AA -> 27. */
export function colIndexFromName(name: string): number {
	let n = 0;
	for (const ch of name.toUpperCase()) {
		const code = ch.charCodeAt(0);
		if (code < 65 || code > 90) {
			return 0;
		}
		n = n * 26 + (code - 64);
	}
	return n;
}

export function toA1(row: number, col: number): string {
	return `${colName(col)}${row}`;
}

export function parseA1(addr: string): { row: number; col: number } {
	const m = /^([A-Za-z]+)(\d+)$/.exec(addr);
	if (!m) {
		return { row: 0, col: 0 };
	}
	return { col: colIndexFromName(m[1]), row: Number.parseInt(m[2], 10) };
}

// -- Excel serial dates (1900 system, incl. the 1900 leap-year offset) ---------

export const EXCEL_EPOCH_DAYS = 25569; // 1970-01-01 as an Excel serial.

export function excelSerialToDate(serial: number): Date {
	const ms = Math.round((serial - EXCEL_EPOCH_DAYS) * 86400000);
	return new Date(ms);
}

export function dateToExcelSerial(date: Date): number {
	return date.getTime() / 86400000 + EXCEL_EPOCH_DAYS;
}

// -- numFmt display (best-effort; the file keeps the real numFmt) --------------

const BUILTIN_CODES: Record<number, string> = {
	9: "0%",
	10: "0.00%",
	14: "yyyy-mm-dd",
	15: "yyyy-mm-dd",
	16: "yyyy-mm-dd",
	17: "yyyy-mm-dd",
	22: "yyyy-mm-dd hh:mm",
	18: "hh:mm",
	19: "hh:mm:ss",
	20: "hh:mm",
	21: "hh:mm:ss",
	46: "hh:mm:ss",
	47: "hh:mm:ss",
};

const MONTHS_EN = [
	"Jan", "Feb", "Mar", "Apr", "May", "Jun",
	"Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

/** True for date/time numFmts (bracketed colors like [Red] are stripped first). */
export function isDateFmt(fmt: string): boolean {
	if (!fmt || fmt === "General" || fmt === "@") {
		return false;
	}
	const bare = fmt.replace(/\[[^\]]*\]/g, "");
	return /[ydhs]/i.test(bare) && !/^[#0?,. "%$£€¥-]+$/.test(bare);
}

/** Map builtin numFmt ids to a code we can render. */
export function resolveFmt(numFmt: string | number | undefined): string {
	if (numFmt === undefined || numFmt === null) {
		return "General";
	}
	if (typeof numFmt === "number") {
		return BUILTIN_CODES[numFmt] ?? "General";
	}
	return numFmt;
}

/** Renders a date-formatted number or Date loosely: y/m/d/h/s token soup. */
/** Renders a date-formatted number or Date loosely: y/m/d/h/s token soup. */
export function formatDateTokens(date: Date, fmt: string): string {
	const pad = (n: number, w = 2) => String(Math.abs(n)).padStart(w, "0");
	const h24 = date.getHours();
	const isPm = h24 >= 12;
	const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
	let out = "";
	let i = 0;
	while (i < fmt.length) {
		const ch = fmt[i];
		let run = 1;
		while (i + run < fmt.length && fmt[i + run] === ch) {
			run++;
		}
		if (ch === "y") {
			out += run >= 3 ? String(date.getFullYear()).padStart(4, "0") : pad(date.getFullYear() % 100);
		} else if (ch === "m" && looksLikeMinutes(fmt, i)) {
			out += run >= 2 ? pad(date.getMinutes()) : String(date.getMinutes());
		} else if (ch === "m") {
			out += run >= 2 ? pad(date.getMonth() + 1) : String(date.getMonth() + 1);
		} else if (ch === "d") {
			out += run >= 3 ? MONTHS_EN[date.getMonth()] : run === 2 ? pad(date.getDate()) : String(date.getDate());
		} else if (ch === "h") {
			const use12 = /am\/pm/i.test(fmt);
			out += use12 ? pad(h12) : run >= 2 ? pad(h24) : String(h24);
		} else if (ch === "s") {
			out += run >= 2 ? pad(date.getSeconds()) : String(date.getSeconds());
		} else if ("#?0.,%".includes(ch)) {
			out += ch;
		} else if (ch === "G" && fmt.slice(i, i + 7) === "General") {
			out += "General";
			i += 7;
			continue;
		} else if (/[A-Za-z]/.test(ch)) {
			// month/day names and other literal text are dropped
		} else {
			out += ch;
		}
		i += run;
	}
	return out.replace(/am\/pm/i, isPm ? "PM" : "AM");
}

function looksLikeMinutes(fmt: string, at: number): boolean {
	// Skip separator characters: in hh:mm the m follows h across a colon.
	let i = at - 1;
	while (i >= 0 && !/[a-z]/i.test(fmt[i])) {
		i--;
	}
	let j = at + 1;
	while (j < fmt.length && !/[a-z]/i.test(fmt[j])) {
		j++;
	}
	const prev = i >= 0 ? fmt[i] : "";
	const next = j < fmt.length ? fmt[j] : "";
	return prev === "h" || prev === "s" || next === "s";
}
/** Numeric numFmt rendering: percent, thousands separators, decimal places. */
export function formatNumberWithFmt(value: number, fmt: string): string {
	if (fmt.includes("%")) {
		const scaled = value * 100;
		return applySeparators(decimals(scaled, fmt), fmt) + "%";
	}
	return applySeparators(decimals(value, fmt), fmt);
}

function decimals(value: number, fmt: string): string {
	const dot = fmt.indexOf(".");
	const count = dot === -1 ? 0 : (fmt.slice(dot + 1).match(/0/g) ?? []).length;
	if (count === 0 && !Number.isInteger(value)) {
		return String(Number(value.toPrecision(11)));
	}
	return value.toFixed(count);
}

function applySeparators(text: string, fmt: string): string {
	if (!fmt.includes(",")) {
		return text;
	}
	const parts = text.split(".");
	const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	return parts.length > 1 ? `${intPart}.${parts[1]}` : intPart;
}

/** General-format rendering (what Excel shows with no explicit numFmt). */
export function formatGeneralNumber(value: number): string {
	if (Number.isInteger(value)) {
		return String(value);
	}
	return String(Number(value.toPrecision(11)));
}

/** Top-level cell display renderer used by both grid engines. */
export function formatGridValue(value: unknown, numFmt?: string | number): string {
	if (value === null || value === undefined) {
		return "";
	}
	if (typeof value === "boolean") {
		return value ? "TRUE" : "FALSE";
	}
	if (value instanceof Date) {
		return formatDateTokens(value, resolveFmt(numFmt));
	}
	if (typeof value === "number") {
		const fmt = resolveFmt(numFmt);
		if (isDateFmt(fmt)) {
			return formatDateTokens(excelSerialToDate(value), fmt);
		}
		if (fmt === "General" || fmt === "@") {
			return formatGeneralNumber(value);
		}
		return formatNumberWithFmt(value, fmt);
	}
	if (typeof value === "string") {
		return value;
	}
	if (typeof value === "number" || typeof value === "boolean") {
		return String(value);
	}
	return "";
}
