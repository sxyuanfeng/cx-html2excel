/*
 * @Author: xujiang
 * @Date: 2024-02-28 15:59:48
 * @LastEditors: xujiang
 * Copyright (c) 2024 by xujiang/cicc, All Rights Reserved.
 */

const ExcelJS = require('exceljs');

var pdre1 = /^(\d+):(\d+)(:\d+)?(\.\d+)?$/; // HH:MM[:SS[.UUU]]
var pdre2 = /^(\d+)-(\d+)-(\d+)$/; // YYYY-mm-dd
var pdre3 = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)?(\.\d+)?$/; // YYYY-mm-dd(T or space)HH:MM[:SS[.UUU]], sans "Z"

var dnthresh  = Date.UTC(1899, 11, 30, 0, 0, 0); // -2209161600000
var dnthresh1 = Date.UTC(1899, 11, 31, 0, 0, 0); // -2209075200000
var dnthresh2 = Date.UTC(1904, 0, 1, 0, 0, 0); // -2209075200000

var FDRE1 = /^(0?\d|1[0-2])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))\s+([ap])m?$/;
var FDRE2 = /^([01]?\d|2[0-3])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))$/;
var FDISO = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)(\.\d+)?[Z]?$/; // YYYY-mm-dd(T or space)HH:MM:SS[.UUU][Z]

var utc_append_works = new Date("6/9/69 00:00 UTC").valueOf() == -17798400000;

var lower_months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

var table_fmt = {
	0:  'General',
	1:  '0',
	2:  '0.00',
	3:  '#,##0',
	4:  '#,##0.00',
	9:  '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'm/d/yy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yy h:mm',
	37: '#,##0 ;(#,##0)',
	38: '#,##0 ;[Red](#,##0)',
	39: '#,##0.00;(#,##0.00)',
	40: '#,##0.00;[Red](#,##0.00)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	47: 'mmss.0',
	48: '##0.0E+0',
	49: '@',
	56: '"上午/下午 "hh"時"mm"分"ss"秒 "'
};

const defaultOptions = {
	rows: true,
	font: true,
	alignment: true,
	border: true,
	numfmt: true,
	chart: true
}

export function generateExcel(tables, userOptions) {
	const ops = { ...defaultOptions, ...userOptions };
	let workbook = new ExcelJS.Workbook();
	for (let table of tables) {
		let xlsx = parse_dom_table(table.elet, ops);
		let worksheet = workbook.addWorksheet(table.name);
		worksheet.columns = xlsx["!cols"];
		if (ops.rows) {
			let rows = xlsx["!rows"];
			for (let i = 0; i < rows.length; i++) {
				let row = rows[i];
				worksheet.getRow(i + 1).height = row.hpx;
			}
		}
		let merges = xlsx["!merges"];
		for (let merge of merges) {
			worksheet.mergeCells(encode_range(merge));
		}
		let data = xlsx["!data"];
		for (let key in data) {
			let cell = worksheet.getCell(key);
			cell.value = data[key].v;
			if (ops.font) cell.font = data[key].s.font;
			if (ops.alignment) cell.alignment = data[key].s.alignment;
			if (ops.border) cell.border = data[key].s.border;
			if (ops.numfmt) cell.numFmt = data[key].s.numfmt;
			if (ops.chart && data[key].chart) {
				const imageId = workbook.addImage({
					base64: data[key].chart.base64,
					extension: 'png',
				});
				worksheet.addImage(imageId, {
					tl: {col: decode_cell(key).c, row: decode_cell(key).r},
					ext: { width: data[key].chart.width, height: data[key].chart.height }
				});
				// worksheet.addImage(imageId, data[key].chart.range);
			}
		}
	}
	return workbook.xlsx.writeBuffer().then(buffer => {
    return buffer;
  });
}

function sheet_add_dom(ws, table, _opts) {
	var rows = table.rows;
	if(!rows) {
		/* not an HTML TABLE */
		throw "Unsupported origin when " + table.tagName + " is not a TABLE";
	}

	var opts = _opts || {};
	var dense = ws["!data"] != null;
	var or_R = 0, or_C = 0;
	if(opts.origin != null) {
		if(typeof opts.origin == 'number') or_R = opts.origin;
		else {
			var _origin = typeof opts.origin == "string" ? decode_cell(opts.origin) : opts.origin;
			or_R = _origin.r; or_C = _origin.c;
		}
	}

	var sheetRows = Math.min(opts.sheetRows||10000000, rows.length);
	var range = {s:{r:0,c:0},e:{r:or_R,c:or_C}};
	if(ws["!ref"]) {
		var _range = decode_range(ws["!ref"]);
		range.s.r = Math.min(range.s.r, _range.s.r);
		range.s.c = Math.min(range.s.c, _range.s.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		range.e.c = Math.max(range.e.c, _range.e.c);
		if(or_R == -1) range.e.r = or_R = _range.e.r + 1;
	}
	var merges = [], midx = 0;
	var rowinfo = ws["!rows"] || (ws["!rows"] = []);
	var _R = 0, R = 0, _C = 0, C = 0, RS = 0, CS = 0;
	if(!ws["!cols"]) ws['!cols'] = [];
	let colgroupElem = table.getElementsByTagName('colgroup')[0];
	let colElems = colgroupElem.getElementsByTagName('col');
	for (let i = 0; i < colElems.length; i++) {
		let colElem = colElems[i];
		let get_computed_style = get_get_computed_style_function(colElem);
		let width = 0;
		if(get_computed_style) width = get_computed_style(colElem).getPropertyValue('width').match(/\d+/g)[0];
		ws['!cols'].push({key: String(i), width: (Number(width) / 5).toFixed(6)});
	}
	for(; _R < rows.length && R < sheetRows; ++_R) {
		var row = rows[_R];
		if (is_dom_element_hidden(row)) {
			if (opts.display) continue;
			rowinfo[R] = {hidden: true};
		}
		if (rowinfo[R]) {
			rowinfo[R].hpx = row.clientHeight;
		} else {
			rowinfo[R] = {hpx: row.clientHeight};
		}
		let trHeights = []
		for (let tr of rows) {
			trHeights.push({hpx: tr.clientHeight})
		}
		ws["!rows"] = trHeights
		var elts = (row.cells);
		for(_C = C = 0; _C < elts.length; ++_C) {
			var elt = elts[_C];
			if (opts.display && is_dom_element_hidden(elt)) continue;
			var v = elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmldecode(elt.innerHTML);
			var z = elt.getAttribute('data-z') || elt.getAttribute('z');
			var numfmt = elt.getAttribute('data-numfmt') || elt.getAttribute('numfmt');
			for(midx = 0; midx < merges.length; ++midx) {
				var m = merges[midx];
				if(m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) { C = m.e.c+1 - or_C; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if( ((RS = (+elt.getAttribute("rowspan") || 1)))>1 || CS>1) merges.push({s:{r:R + or_R,c:C + or_C},e:{r:R + or_R + (RS||1) - 1, c:C + or_C + (CS||1) - 1}});
			var o = {t:'s', v:v};
			var _t = elt.getAttribute("data-t") || elt.getAttribute("t") || "";
			if(v != null) {
				if(v.length == 0) o.t = _t || 'z';
				else if(opts.raw || v.trim().length == 0 || _t == "s"){}
				else if(v === 'TRUE') o = {t:'b', v:true};
				else if(v === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(v))) o = {t:'n', v:fuzzynum(v)};
				else if(!isNaN(fuzzydate(v).getDate())) {
					o = ({t:'d', v:parseDate(v)});
					if(opts.UTC) o.v = local_to_utc(o.v);
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)});
					o.z = opts.dateNF || table_fmt[14];
				}
			}
			if(o.z === undefined && z != null) o.z = z;
			/* The first link is used.  Links are assumed to be fully specified.
			 * TODO: The right way to process relative links is to make a new <a> */
			var l = "", Aelts = elt.getElementsByTagName("A");
			if(Aelts && Aelts.length) for(var Aelti = 0; Aelti < Aelts.length; ++Aelti)	if(Aelts[Aelti].hasAttribute("href")) {
				l = Aelts[Aelti].getAttribute("href"); if(l.charAt(0) != "#") break;
			}
			if(l && l.charAt(0) != "#" &&	l.slice(0, 11).toLowerCase() != 'javascript:') o.l = ({ Target: l });
			o.s = setCellStyle(elt);
			if (parseChart(elt)) {
				o.chart = parseChart(elt, C + or_C, R + or_R);
			}
			o.s.numfmt = numfmt;
			// if(dense) { if(!ws["!data"][R + or_R]) ws["!data"][R + or_R] = []; ws["!data"][R + or_R][C + or_C] = o; }
			// else ws[encode_cell({c:C + or_C, r:R + or_R})] = o;
			if (!ws["!data"]) {
				ws["!data"] = {};
			}
			ws["!data"][encode_cell({c:C + or_C, r:R + or_R})] = o;
			if(range.e.c < C + or_C) range.e.c = C + or_C;
			C += CS;
		}
		++R;
	}
	if(merges.length) ws['!merges'] = (ws["!merges"] || []).concat(merges);
	range.e.r = Math.max(range.e.r, R - 1 + or_R);
	ws['!ref'] = encode_range(range);
	if(R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length-_R+R-1 + or_R,range)); // We can count the real number of rows to parse but we don't to improve the performance
	console.log(ws)
	return ws;
}

function parseChart(dom, cs, rs) {
	let get_computed_style = get_get_computed_style_function(dom);
	let rowSpan = dom.getAttribute("rowspan") || 1;
	let colSpan = dom.getAttribute("colspan") || 1;
	let re = rs + Number(rowSpan);
	let ce = cs + Number(colSpan);
	let range = encode_range({s:{c: cs, r: rs}, e:{c: ce, r: re}});
	let imgEles = dom.getElementsByTagName('img');
	let canvasEles = dom.getElementsByTagName('canvas');
	if (imgEles.length > 0) {
		let imgEle = imgEles[0];
		let width = 0, height = 0;
		if(get_computed_style) width = get_computed_style(imgEle).getPropertyValue('width').match(/\d+/g)[0];
		if(get_computed_style) height = get_computed_style(imgEle).getPropertyValue('height').match(/\d+/g)[0];
		// 创建 Canvas 元素
		const canvas = document.createElement('canvas');
		canvas.width = imgEle.width;
		canvas.height = imgEle.height;

		// 将图像绘制到 Canvas 上
		const ctx = canvas.getContext('2d');
		ctx.drawImage(imgEle, 0, 0);

		// 获取 base64 编码的图像数据
		const base64Image = canvas.toDataURL('image/png');
		return {width: width, height: height, range: range, base64: base64Image};
	}
	if (canvasEles.length > 0) {
		let canvasEle = canvasEles[0];
		let width = 0, height = 0;
		if(get_computed_style) width = get_computed_style(canvasEle).getPropertyValue('width').match(/\d+/g)[0];
		if(get_computed_style) height = get_computed_style(canvasEle).getPropertyValue('height').match(/\d+/g)[0];
		const base64Image = canvasEle.toDataURL('image/png');
		return {width: width, height: height, range: range, base64: base64Image};
	}
	return undefined;
}

function setCellStyle(dom) {
	let get_computed_style = get_get_computed_style_function(dom);
	let o = {};
	if (get_computed_style) {
		let font = {};
		font.size = get_computed_style(dom).getPropertyValue('font-size').match(/\d+/g)[0];
		let fontName = get_computed_style(dom).getPropertyValue('font-family');
		let fontNameDict = {SimSun: "SimSun"};
		font.name = fontNameDict[fontName] ?? "SimSun";
		font.bold = get_computed_style(dom).getPropertyValue('font-weight') === '700';

		let hex = RGBToHex(get_computed_style(dom).getPropertyValue('color')).color;
		font.color = { argb: 'ff' + hex };
		o.font = font;
		let alignment = {};
		alignment.wrapText = get_computed_style(dom).getPropertyValue('white-space') !== 'nowrap';
		//todo 部分css和excel属性值未匹配
		alignment.horizontal = get_computed_style(dom).getPropertyValue('text-align');
		alignment.vertical = get_computed_style(dom).getPropertyValue('vertical-align');
		o.alignment = alignment;
		let border = {
			top: parseBorder(get_computed_style(dom).getPropertyValue('border-top')),
			bottom: parseBorder(get_computed_style(dom).getPropertyValue('border-bottom')),
			left: parseBorder(get_computed_style(dom).getPropertyValue('border-left')),
			right: parseBorder(get_computed_style(dom).getPropertyValue('border-right')),
		};
		o.border = JSON.parse(JSON.stringify(border));
	}
	return o;
}

function svg2base64(svgElement) {
	const serializer = new XMLSerializer();
	const svgXML = serializer.serializeToString(svgElement);
	return btoa(svgXML);
}

function parse_dom_table(table, _opts) {
	var opts = _opts || {};
	var ws = ({}); if(opts.dense) ws["!data"] = [];
	return sheet_add_dom(ws, table, _opts);
}

function decode_range(range) {
	var idx = range.indexOf(":");
	if(idx == -1) return { s: decode_cell(range), e: decode_cell(range) };
	return { s: decode_cell(range.slice(0, idx)), e: decode_cell(range.slice(idx + 1)) };
}

function encode_range(cs,ce) {
	if(typeof ce === 'undefined' || typeof ce === 'number') {
return encode_range(cs.s, cs.e);
	}
if(typeof cs !== 'string') cs = encode_cell((cs));
	if(typeof ce !== 'string') ce = encode_cell((ce));
return cs == ce ? cs : cs + ":" + ce;
}

function parseBorder(s) {
	let str = s.replace(new RegExp(', ', 'g'), ',')
	let arr = str.split(' ');
	let dict = {solid: 'thin'};
	let style = dict[arr[1]];
	if (!style) return undefined;
	let color = RGBToHex(arr[2]);
	return {style, color: {argb: 'ff' + color.color}};
}

function RGBToHex(rgba) {
	if (rgba.startsWith('rgba')) {
		let str = rgba.slice(5, rgba.length - 1),
			arry = str.split(','),
			opa = Number(arry[3].trim()) * 100,
			strHex = "",
			r = Number(arry[0].trim()),
			g = Number(arry[1].trim()),
			b = Number(arry[2].trim());

		strHex += ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);

		return { color: strHex, opacity: opa };
	}
	let str = rgba.slice(4, rgba.length - 1),
			arry = str.split(','),
			strHex = "",
			r = Number(arry[0].trim()),
			g = Number(arry[1].trim()),
			b = Number(arry[2].trim());

	strHex += ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);

	return { color: strHex };
}

function decode_cell(cstr) {
	var R = 0, C = 0;
	for(var i = 0; i < cstr.length; ++i) {
		var cc = cstr.charCodeAt(i);
		if(cc >= 48 && cc <= 57) R = 10 * R + (cc - 48);
		else if(cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
	}
	return { c: C - 1, r:R - 1 };
}

function encode_cell(cell) {
	var col = cell.c + 1;
	var s="";
	for(; col; col=((col-1)/26)|0) s = String.fromCharCode(((col-1)%26) + 65) + s;
	return s + (cell.r + 1);
}

function is_dom_element_hidden(element) {
	var display = '';
	var get_computed_style = get_get_computed_style_function(element);
	if(get_computed_style) display = get_computed_style(element).getPropertyValue('display');
	if(!display) display = element.style && element.style.display;
	return display === 'none';
}

function datenum(v, date1904) {
	var epoch = v.getTime();
	var res = (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	if(date1904) { res -= 1462; return res < -1402 ? res - 1 : res; }
	return res < 60 ? res - 1 : res;
}

function get_get_computed_style_function(element) {
	// The proper getComputedStyle implementation is the one defined in the element window
	if(element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
	// If it is not available, try to get one from the global namespace
	if(typeof getComputedStyle === 'function') return getComputedStyle;
	return null;
}

function local_to_utc(local) {
	return new Date(Date.UTC(local.getFullYear(), local.getMonth(), local.getDate(), local.getHours(), local.getMinutes(), local.getSeconds(), local.getMilliseconds()));
}

function parseDate(str, date1904) {
	if(str instanceof Date) return str;
	var m = str.match(pdre1);
	if(m) return new Date((date1904 ? dnthresh2 : dnthresh1) + ((parseInt(m[1], 10)*60 + parseInt(m[2], 10))*60 + (m[3] ? parseInt(m[3].slice(1), 10) : 0))*1000 + (m[4] ? parseInt((m[4]+"000").slice(1,4), 10) : 0));
	m = str.match(pdre2);
	if(m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3], 0, 0, 0, 0));
	/* TODO: 1900-02-29T00:00:00.000 should return a flag to treat as a date code (affects xlml) */
	m = str.match(pdre3);
	if(m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3], +m[4], +m[5], ((m[6] && parseInt(m[6].slice(1), 10))|| 0), ((m[7] && parseInt(m[7].slice(1), 10))||0)));
	var d = new Date(str);
	return d;
}

function fuzzynum(s) {
	var v = Number(s);
	if(!isNaN(v)) return isFinite(v) ? v : NaN;
	if(!/\d/.test(s)) return v;
	var wt = 1;
	var ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
	if(!isNaN(v = Number(ss))) return v / wt;
	ss = ss.replace(/[(](.*)[)]/,function($$, $1) { wt = -wt; return $1;});
	if(!isNaN(v = Number(ss))) return v / wt;
	return v;
}

function fuzzytime1(M)  {
	if(!M[2]) return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), 0, 0, 0));
	if(M[3]) {
			if(M[4]) return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], +M[4], parseFloat(M[3])*1000));
			else return new Date(Date.UTC(1899,11,31,(M[7] == "p" ? 12 : 0), +M[1], +M[2], parseFloat(M[3])*1000));
	}
	else if(M[5]) return new Date(Date.UTC(1899,11,31, (+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], +M[5], M[6] ? parseFloat(M[6]) * 1000 : 0));
	else return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], 0, 0));
}
function fuzzytime2(M)  {
	if(!M[2]) return new Date(Date.UTC(1899,11,31,+M[1], 0, 0, 0));
	if(M[3]) {
			if(M[4]) return new Date(Date.UTC(1899,11,31,+M[1], +M[2], +M[4], parseFloat(M[3])*1000));
			else return new Date(Date.UTC(1899,11,31,0, +M[1], +M[2], parseFloat(M[3])*1000));
	}
	else if(M[5]) return new Date(Date.UTC(1899,11,31, +M[1], +M[2], +M[5], M[6] ? parseFloat(M[6]) * 1000 : 0));
	else return new Date(Date.UTC(1899,11,31,+M[1], +M[2], 0, 0));
}

function fuzzydate(s) {
	// See issue 2863 -- this is technically not supported in Excel but is otherwise useful
	if(FDISO.test(s)) return s.indexOf("Z") == -1 ? local_to_utc(new Date(s)) : new Date(s);
	var lower = s.toLowerCase();
	var lnos = lower.replace(/\s+/g, " ").trim();
	var M = lnos.match(FDRE1);
	if(M) return fuzzytime1(M);
	M = lnos.match(FDRE2);
	if(M) return fuzzytime2(M);
	M = lnos.match(pdre3);
	if(M) return new Date(Date.UTC(+M[1], +M[2]-1, +M[3], +M[4], +M[5], ((M[6] && parseInt(M[6].slice(1), 10))|| 0), ((M[7] && parseInt(M[7].slice(1), 10))||0)));
	var o = new Date(utc_append_works && s.indexOf("UTC") == -1 ? s + " UTC": s), n = new Date(NaN);
	var y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	if(lower.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
		lower = lower.replace(/[^a-z]/g,"").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/,"");
		if(lower.length > 3 && lower_months.indexOf(lower) == -1) return n;
	} else if(lower.replace(/[ap]m?/, "").match(/[a-z]/)) return n;
	if(y < 0 || y > 8099 || s.match(/[^-0-9:,\/\\\ ]/)) return n;
	return o;
}

var htmldecode = (function() {
	var entities = [
		['nbsp', ' '], ['middot', '·'],
		['quot', '"'], ['apos', "'"], ['gt',   '>'], ['lt',   '<'], ['amp',  '&']
	].map(function(x) { return [new RegExp('&' + x[0] + ';', "ig"), x[1]]; });
	return function htmldecode(str) {
		var o = str
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
		for(var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
		return o;
	};
})();
