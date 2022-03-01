/*! xlsxspread.js (C) SheetJS LLC -- https://sheetjs.com/ */
/* eslint-env browser */
/*global XLSX */
/*exported stox, xtos */
// import * as XLSX from 'xlsx';

/**
 * Converts data from SheetJS to x-spreadsheet
 *
 * @param  {Object} wb SheetJS workbook object
 *
 * @returns {Object[]} An x-spreadsheet data
 */
export const stox = (text) => {
    if (!text) return '';
    // const wb = XLSX.read(text, {
    //     type: 'string'
    // });
    const wb = html_to_book(text);
    
    const out = [];
    
    wb.SheetNames.forEach(function (name) {
      const o = { name: name, rows: {} };
      const ws = wb.Sheets[name];
      const range = decode_range(ws['!ref']);
      // sheet_to_json will lost empty row and col at begin as default
      range.s = { r: 0, c: 0 };
      const aoa = sheet_to_json(ws, {
        raw: false,
        header: 1,
        range: range
      });
  
      aoa.forEach(function (r, i) {
        const cells = {};
        r.forEach(function (c, j) {
          cells[j] = { text: c };
  
          const cellRef = encode_cell({ r: i, c: j });
  
          if ( ws[cellRef] != null && ws[cellRef].f != null) {
            cells[j].text = "=" + ws[cellRef].f;
          }
        });
        o.rows[i] = { cells: cells };
      });
  
      o.merges = [];
      (ws["!merges"]||[]).forEach(function (merge, i) {
        //Needed to support merged cells with empty content
        if (o.rows[merge.s.r] == null) {
          o.rows[merge.s.r] = { cells: {} };
        }
        if (o.rows[merge.s.r].cells[merge.s.c] == null) {
          o.rows[merge.s.r].cells[merge.s.c] = {};
        }
  
        o.rows[merge.s.r].cells[merge.s.c].merge = [
          merge.e.r - merge.s.r,
          merge.e.c - merge.s.c
        ];
  
        o.merges[i] = encode_range(merge);
      });
  
      out.push(o);
    });
  
    return out;
}
  
/**
 * Converts data from x-spreadsheet to SheetJS
 *
 * @param  {Object[]} sdata An x-spreadsheet data object
 *
 * @returns {Object} A SheetJS workbook object
 */
export const xtos = (sdata) => {
  let html = '';
  // const out = XLSX.utils.book_new();
  sdata.forEach(function (xws) {
    const ws = {};
    const rowobj = xws.rows;
    for (let ri = 0; ri < rowobj.len; ++ri) {
      const row = rowobj[ri];
      if (!row) continue;

      let minCoord;
      let maxCoord;
      Object.keys(row.cells).forEach(function (k) {
        const idx = +k;
        if (isNaN(idx)) return;

        const lastRef = encode_cell({ r: ri, c: idx });
        if (minCoord == null) {
          minCoord = { r: ri, c: idx };
        } else {
          if (ri < minCoord.r) minCoord.r = ri;
          if (idx < minCoord.c) minCoord.c = idx;
        }
        if (maxCoord == undefined) {
          maxCoord = { r: ri, c: idx };
        } else {
          if (ri > maxCoord.r) maxCoord.r = ri;
          if (idx > maxCoord.c) maxCoord.c = idx;
        }

        let cellText = row.cells[k].text, type = "s";
        if (!cellText) {
          cellText = "";
          type = "z";
        } else if (!isNaN(parseFloat(cellText))) {
          cellText = parseFloat(cellText);
          type = "n";
        } else if (cellText.toLowerCase() === "true" || cellText.toLowerCase() === "false") {
          cellText = Boolean(cellText);
          type = "b";
        }

        ws[lastRef] = { v: cellText, t: type };
        
        if (type == "s" && cellText[0] == "=") {
          ws[lastRef].f = cellText.slice(1);
        }

        if (row.cells[k].merge != null) {
          if (ws["!merges"] == null) ws["!merges"] = [];

          ws["!merges"].push({
            s: { r: ri, c: idx },
            e: {
              r: ri + row.cells[k].merge[0],
              c: idx + row.cells[k].merge[1]
            }
          });
        }
      });

      ws["!ref"] = encode_range({
        s: { r: minCoord.r, c: minCoord.c },
        e: { r: maxCoord.r, c: maxCoord.c }
      });
    }

    html = sheet_to_html(ws, {
      header: '<head><meta charset="utf-8"/></head>'
    });
  });

  return html;
}
/* simple blank workbook object */
const book_new = () => {
	return { SheetNames: [], Sheets: {} };
};

/* add a worksheet to the end of a given workbook */
const book_append_sheet = (wb, ws, name) => {
	if(!name) for(const i = 1; i <= 0xFFFF; ++i, name = undefined) if(wb.SheetNames.indexOf(name = "Sheet" + i) == -1) break;
	if(!name || wb.SheetNames.length >= 0xFFFF) throw new Error("Too many worksheets");
	check_ws_name(name);
	if(wb.SheetNames.indexOf(name) >= 0) throw new Error("Worksheet with name |" + name + "| already exists!");

	wb.SheetNames.push(name);
	wb.Sheets[name] = ws;
};

const badchars = "][*?\/\\".split("");
const check_ws_name = (n, safe) => {
	if(n.length > 31) { if(safe) return false; throw new Error("Sheet names cannot exceed 31 chars"); }
	const _good = true;
	badchars.forEach(function(c) {
		if(n.indexOf(c) == -1) return;
		if(!safe) throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
		_good = false;
	});
	return _good;
}

const safe_split_regex = "abacaba".split(/(:?b)/i).length == 5;
const split_regex = (str, re, def) => {
	if(safe_split_regex || typeof re == "string") return str.split(re);
	const p = str.split(re), o = [p[0]];
	for(let i = 1; i < p.length; ++i) { o.push(def); o.push(p[i]); }
	return o;
}

const html_to_book = (str, opts) => {
  const mtch = str.match(/<table.*?>[\s\S]*?<\/table>/gi);
  if(!mtch || mtch.length == 0) throw new Error("Invalid HTML: could not find <table>");
  if(mtch.length == 1) return sheet_to_workbook(html_to_sheet(mtch[0], opts), opts);
  const wb = book_new();
  mtch.forEach(function(s, idx) { book_append_sheet(wb, html_to_sheet(s, opts), "Sheet" + (idx+1)); });
  return wb;
}

const sheet_to_workbook = (sheet, opts) => {
  const n = opts && opts.sheet ? opts.sheet : "Sheet1";
  const sheets = {}; sheets[n] = sheet;
  return { SheetNames: [n], Sheets: sheets };
}
let DENSE = null;
const lower_months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const fuzzydate = (s) => {
	const o = new Date(s), n = new Date(NaN);
	const y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	const lower = s.toLowerCase();
	if(lower.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
		lower = lower.replace(/[^a-z]/g,"").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/,"");
		if(lower.length > 3 && lower_months.indexOf(lower) == -1) return n;
	} else if(lower.match(/[a-z]/)) return n;
	if(y < 0 || y > 8099) return n;
	if((m > 0 || d > 1) && y != 101) return o;
	if(s.match(/[^-0-9:,\/\\]/)) return n;
	return o;
}
const html_to_sheet = (str, _opts)=> {
  const opts = _opts || {};
  if(DENSE != null && opts.dense == null) opts.dense = DENSE;
  const ws = opts.dense ? ([]) : ({});
  str = str.replace(/<!--.*?-->/g, "");
  const mtch = str.match(/<table/i);
  if(!mtch) throw new Error("Invalid HTML: could not find <table>");
  const mtch2 = str.match(/<\/table/i);
  const i = mtch.index, j = mtch2 && mtch2.index || str.length;
  const rows = split_regex(str.slice(i, j), /(:?<tr[^>]*>)/i, "<tr>");
  let R = -1, C = 0, RS = 0, CS = 0;
  const range = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
  const merges = [];
  for(let i = 0; i < rows.length; ++i) {
    const row = rows[i].trim();
    const hd = row.slice(0,3).toLowerCase();
    if(hd == "<tr") { ++R; if(opts.sheetRows && opts.sheetRows <= R) { --R; break; } C = 0; continue; }
    if(hd != "<td" && hd != "<th") continue;
    const cells = row.split(/<\/t[dh]>/i);
    
    for(let j = 0; j < cells.length; ++j) {
      const cell = cells[j].trim()
      if(!cell.match(/<t[dh]/i)) continue;
      let m = cell, cc = 0;
      m = m.replace(/<td (.*)>/, '');
      /* TODO: parse styles etc */
      // while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
      for(let midx = 0; midx < merges.length; ++midx) {
        const _merge = merges[midx];
        if(_merge.s.c == C && _merge.s.r < R && R <= _merge.e.r) { C = _merge.e.c + 1; midx = -1; }
      }
      const tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
      CS = tag.colspan ? +tag.colspan : 1;
      if((RS = +tag.rowspan)>1 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
      const _t = tag.t || tag["data-t"] || "";
      /* TODO: generate stub cells */
      if(!m.length) { C += CS; continue; }
      m = htmldecode(m);
      if(range.s.r > R) range.s.r = R; if(range.e.r < R) range.e.r = R;
      if(range.s.c > C) range.s.c = C; if(range.e.c < C) range.e.c = C;
      if(!m.length) continue;
      const o = {t:'s', v:m};
      if(opts.raw || !m.trim().length || _t == 's'){}
      else if(m === 'TRUE') o = {t:'b', v:true};
      else if(m === 'FALSE') o = {t:'b', v:false};
      else if(!isNaN(fuzzynum(m))) o = {t:'n', v:fuzzynum(m)};
      else if(!isNaN(fuzzydate(m).getDate())) {
        o = ({t:'d', v:parseDate(m)});
        if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)});
        o.z = opts.dateNF || SSF._table[14];
      }
      if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = o; }
      else ws[encode_cell({r:R, c:C})] = o;
      C += CS;
    }
  }
  ws['!ref'] = encode_range(range);
  if(merges.length) ws["!merges"] = merges;
  return ws;
}

const htmldecode = (function() {
	const entities = [
		['nbsp', ' '], ['middot', '·'],
		['quot', '"'], ['apos', "'"], ['gt',   '>'], ['lt',   '<'], ['amp',  '&']
	].map(function(x) { return [new RegExp('&' + x[0] + ';', "ig"), x[1]]; });
	return function htmldecode(str) {
		let o = str
				// Remove new lines and spaces from start of content
				.replace(/^[\t\n\r ]+/, "")
				// Remove new lines and spaces from end of content
				.replace(/[\t\n\r ]+$/,"")
				// Added line which removes any white space characters after and before html tags
				.replace(/>\s+/g,">").replace(/\s+</g,"<")
				// Replace remaining new lines and spaces with space
				.replace(/[\t\n\r ]+/g, " ")
				// Replace <br> tags with new lines
				.replace(/<\s*[bB][rR]\s*\/?>/g,"\n")
				// Strip HTML elements
				.replace(/<[^>]*>/g,"");
		for(let i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
		return o;
	};
})();

/* TODO: stress test */
const fuzzynum = (s) => {
	let v = Number(s);
	if(isFinite(v)) return v;
	if(!isNaN(v)) return NaN;
	if(!/\d/.test(s)) return v;
	let wt = 1;
	let ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
	if(!isNaN(v = Number(ss))) return v / wt;
	ss = ss.replace(/[(](.*)[)]/,function($$, $1) { wt = -wt; return $1;});
	if(!isNaN(v = Number(ss))) return v / wt;
	return v;
}
const attregexg=/([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
const parsexmltag = (tag, skip_root, skip_LC) => {
	const z = ({});
	let eq = 0, c = 0;
	for(; eq !== tag.length; ++eq) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
	if(!skip_root) z[0] = tag.slice(0, eq);
	if(eq === tag.length) return z;
	const m = tag.match(attregexg);
    let j=0, v="", i=0, q="", cc="", quot = 1;
	if(m) for(let i = 0; i != m.length; ++i) {
		cc = m[i];
		for(let c=0; c != cc.length; ++c) if(cc.charCodeAt(c) === 61) break;
		q = cc.slice(0,c).trim();
		while(cc.charCodeAt(c+1) == 32) ++c;
		quot = ((eq=cc.charCodeAt(c+1)) == 34 || eq == 39) ? 1 : 0;
		v = cc.slice(c+1+quot, cc.length-quot);
		for(let j=0;j!=q.length;++j) if(q.charCodeAt(j) === 58) break;
		if(j===q.length) {
			if(q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_")); // from ods
			z[q] = v;
			if(!skip_LC) z[q.toLowerCase()] = v;
		}
		else {
			const k = (j===5 && q.slice(0,5)==="xmlns"?"xmlns":"")+q.slice(j+1);
			if(z[k] && q.slice(j-3,j) == "ext") continue; // from ods
			z[k] = v;
			if(!skip_LC) z[k.toLowerCase()] = v;
		}
	}
	return z;
}

const parseDate = (str, fixdate) => {
	const d = new Date(str);
	if(good_pd) {
    if(fixdate > 0) d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
		else if(fixdate < 0) d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
		return d;
	}
	if(str instanceof Date) return str;
	if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
		const s = d.getFullYear();
		if(str.indexOf("" + s) > -1) return d;
		d.setFullYear(d.getFullYear() + 100); return d;
	}
	const n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
	const out = new Date(+n[0], +n[1] - 1, +n[2], (+n[3]||0), (+n[4]||0), (+n[5]||0));
	if(str.indexOf("Z") > -1) out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
	return out;
}

const datenum = (v, date1904) => {
	const epoch = v.getTime();
	if(date1904) epoch -= 1462*24*60*60*1000;
	const dnthresh = basedate.getTime() + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}

const encode_cell = (cell) => {
	let col = cell.c + 1;
	let s="";
	for(; col; col=((col-1)/26)|0) s = String.fromCharCode(((col-1)%26) + 65) + s;
	return s + (cell.r + 1);
}

const decode_cell = (cstr) => {
	let R = 0, C = 0;
	for(let i = 0; i < cstr.length; ++i) {
		const cc = cstr.charCodeAt(i);
		if(cc >= 48 && cc <= 57) R = 10 * R + (cc - 48);
		else if(cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
	}
	return { c: C - 1, r:R - 1 };
}

const encode_range = (cs,ce) => {
	if(typeof ce === 'undefined' || typeof ce === 'number') {
        return encode_range(cs.s, cs.e);
    }
    if(typeof cs !== 'string') cs = encode_cell((cs));
    if(typeof ce !== 'string') ce = encode_cell((ce));
    return cs == ce ? cs : cs + ":" + ce;
}

const decode_range = (range) => {
	const idx = range.indexOf(":");
	if(idx == -1) return { s: decode_cell(range), e: decode_cell(range) };
	return { s: decode_cell(range.slice(0, idx)), e: decode_cell(range.slice(idx + 1)) };
}

const make_html_row = (ws, r, R, o) => {
    const M = (ws['!merges'] ||[]);
    const oo = [];
    for(let C = r.s.c; C <= r.e.c; ++C) {
       
        let RS = 0, CS = 0;
        for(let j = 0; j < M.length; ++j) {
            if(M[j].s.r > R || M[j].s.c > C) continue;
            if(M[j].e.r < R || M[j].e.c < C) continue;
            if(M[j].s.r < R || M[j].s.c < C) { RS = -1; break; }
            RS = M[j].e.r - M[j].s.r + 1; CS = M[j].e.c - M[j].s.c + 1; break;
        }
        if(RS < 0) continue;
        
        const coord = encode_cell({r:R,c:C});
        const cell = o.dense ? (ws[R]||[])[C] : ws[coord];
        /* TODO: html entities */
        const w = (cell && cell.v != null) && (cell.h || escapehtml(cell.v || (format_cell(cell), cell.v) || "")) || "";
        const sp = ({});
        if(RS > 1) sp.rowspan = RS;
        if(CS > 1) sp.colspan = CS;
        if(o.editable) w = '<span contenteditable="true">' + w + '</span>';
        else if(cell) {
            sp["data-t"] = cell && cell.t || 'z';
            if(cell.v != null) sp["data-v"] = cell.v;
            if(cell.z != null) sp["data-z"] = cell.z;
            if(cell.l && (cell.l.Target || "#").charAt(0) != "#") w = '<a href="' + cell.l.Target +'">' + w + '</a>';
        }
        sp.id = (o.id || "sjs") + "-" + coord;
        
        oo.push(writextag('td', w, sp));
    }
    const preamble = "<tr>";
    return preamble + oo.join("") + "</tr>";
}
const make_html_preamble = (ws, R, o) => {
    const out = [];
    return out.join("") + '<table' + (o && o.id ? ' id="' + o.id + '"' : "") + '>';
}
const _BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
const _END = '</body></html>';
const sheet_to_html = (ws, opts/*, wb:?Workbook*/) => {
    const o = opts || {};
    const header = o.header != null ? o.header : _BEGIN;
    const footer = o.footer != null ? o.footer : _END;
    const out = [header];
    const r = decode_range(ws['!ref']);
    o.dense = Array.isArray(ws);
    out.push(make_html_preamble(ws, r, o));
    for(let R = 0; R <= r.e.r; ++R) {
        out.push(make_html_row(ws, r, R, o));
    }
    // for(let R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
    out.push("</table>" + footer);
    return out.join("");
}

const decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
const htmlcharegex = /[\u0000-\u001f]/g;
const escapehtml = (text) => {
	const s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(/\n/g, "<br/>").replace(htmlcharegex,function(s) { return "&#x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + ";"; });
}

const format_cell = (cell, v, o) => {
	if(cell == null || cell.t == null || cell.t == 'z') return "";
	if(cell.w !== undefined) return cell.w;
	if(cell.t == 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
	if(cell.t == "e") return BErr[cell.v] || cell.v;
	if(v == undefined) return safe_format_cell(cell, cell.v);
	return safe_format_cell(cell, v);
}

const safe_format_cell = (cell, v) => {
	const q = (cell.t == 'd' && v instanceof Date);
	if(cell.z != null) try { return (cell.w = SSF.format(cell.z, q ? datenum(v) : v)); } catch(e) { }
	try { return (cell.w = SSF.format((cell.XF||{}).numFmtId||(q ? 14 : 0),  q ? datenum(v) : v)); } catch(e) { return ''+v; }
}

const keys = (o) => {
	const ks = Object.keys(o), o2 = [];
	for(let i = 0; i < ks.length; ++i) if(Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
	return o2;
}

const wtregex = /(^\s|\s$|\n)/;
const wxt_helper = (h) => { return keys(h).map(function(k) { return " " + k + '="' + h[k] + '"';}).join(""); }
const writextag = (f,g,h) => { return '<' + f + ((h != null) ? wxt_helper(h) : "") + ((g != null) ? (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f : "/") + '>';}

const sheet_to_json = (sheet, opts) => {
	if(sheet == null || sheet["!ref"] == null) return [];
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	else if(o.header == null) header = 0;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols = [];
	var out = [];
	var outi = 0, counter = 0;
	var dense = Array.isArray(sheet);
	var R = r.s.r, C = 0, CC = 0;
	if(dense && !sheet[R]) sheet[R] = [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encode_col(C);
		val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = 0;
				for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
				hdr[C] = vv;
		}
	}
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
		if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
	}
	out.length = outi;
	return out;
}

const encode_row = (row) => { return "" + (row + 1); }
const encode_col = (col) => { if(col < 0) throw new Error("invalid column " + col); var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }

const make_json_row = (sheet, r, R, cols, header, hdr, dense, o) => {
	var rr = encode_row(R);
	var defval = o.defval, raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
	var isempty = true;
	var row = (header === 1) ? [] : {};
	if(header !== 1) {
		if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
		else row.__rowNum__ = R;
	}
	if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
		var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		if(val === undefined || val.t === undefined) {
			if(defval === undefined) continue;
			if(hdr[C] != null) { row[hdr[C]] = defval; }
			continue;
		}
		var v = val.v;
		switch(val.t){
			case 'z': if(v == null) break; continue;
			case 'e': v = (v == 0 ? null : void 0); break;
			case 's': case 'd': case 'b': case 'n': break;
			default: throw new Error('unrecognized type ' + val.t);
		}
		if(hdr[C] != null) {
			if(v == null) {
				if(val.t == "e" && v === null) row[hdr[C]] = null;
				else if(defval !== undefined) row[hdr[C]] = defval;
				else if(raw && v === null) row[hdr[C]] = null;
				else continue;
			} else {
				row[hdr[C]] = raw || (o.rawNumbers && val.t == "n") ? v : format_cell(val,v,o);
			}
			if(v != null) isempty = false;
		}
	}
	return { row: row, isempty: isempty };
}